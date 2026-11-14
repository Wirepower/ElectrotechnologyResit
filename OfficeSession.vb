Option Strict Off
Option Explicit On

Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Threading
Imports ElectrotechnologyResit.Core

Friend Class OfficeCallResult
    Public Property Kind As OfficeWaitKind
    Public Property Value As Object
    Public Property Problem As Exception

    Public ReadOnly Property Succeeded As Boolean
        Get
            Return Kind = OfficeWaitKind.Completed AndAlso Problem Is Nothing AndAlso Value IsNot Nothing
        End Get
    End Property
End Class

''' <summary>Runs Outlook and Word COM calls on one dedicated STA thread. The UI thread only waits asynchronously.</summary>
Friend Class OfficeStaHost
    Private ReadOnly _queue As New BlockingCollection(Of Action)()
    Private _thread As Thread
    Private ReadOnly _ready As New ManualResetEventSlim(False)
    Private _abandoned As Boolean

    Public ReadOnly Property Abandoned As Boolean
        Get
            Return _abandoned
        End Get
    End Property

    Public Sub Start()
        _thread = New Thread(AddressOf Run)
        _thread.IsBackground = True
        _thread.Name = "Office STA"
        _thread.SetApartmentState(ApartmentState.STA)
        _thread.Start()
        If Not _ready.Wait(TimeSpan.FromSeconds(5)) Then
            Throw New InvalidOperationException("The Office thread did not start.")
        End If
    End Sub

    Private Sub Run()
        _ready.Set()
        For Each action In _queue.GetConsumingEnumerable()
            action()
        Next
    End Sub

    Public Function InvokeAsync(work As Func(Of Object), timeout As TimeSpan, cancel As CancellationToken) As Task(Of OfficeCallResult)
        If _abandoned Then
            Return Task.FromResult(New OfficeCallResult With {.Kind = OfficeWaitKind.TimedOut})
        End If
        Return Task.Run(Function() Invoke(work, timeout, cancel))
    End Function

    Private Function Invoke(work As Func(Of Object), timeout As TimeSpan, cancel As CancellationToken) As OfficeCallResult
        Dim done As New ManualResetEventSlim(False)
        Dim value As Object = Nothing
        Dim problem As Exception = Nothing
        Try
            _queue.Add(Sub()
                           Try
                               value = work()
                           Catch ex As Exception
                               problem = ex
                           Finally
                               done.Set()
                           End Try
                       End Sub)
        Catch ex As Exception
            Return New OfficeCallResult With {.Kind = OfficeWaitKind.Completed, .Problem = ex}
        End Try
        Dim kind = OfficeCallWait.Wait(done, timeout, cancel, TimeSpan.FromMilliseconds(50))
        If kind <> OfficeWaitKind.Completed Then
            _abandoned = True
            Return New OfficeCallResult With {.Kind = kind}
        End If
        Return New OfficeCallResult With {.Kind = OfficeWaitKind.Completed, .Value = value, .Problem = problem}
    End Function

    Public Sub StopPump()
        _queue.CompleteAdding()
    End Sub
End Class

Friend NotInheritable Class ComActivation
    Private Sub New()
    End Sub

    <DllImport("ole32.dll", CharSet:=CharSet.Unicode, PreserveSig:=True)>
    Private Shared Function CLSIDFromProgID(progId As String, ByRef clsid As Guid) As Integer
    End Function

    <DllImport("oleaut32.dll", PreserveSig:=True)>
    Private Shared Function GetActiveObject(ByRef rclsid As Guid, reserved As IntPtr, <Out> ByRef ppunk As Object) As Integer
    End Function

    Friend Shared Function TryAttach(progId As String) As Object
        Dim running = TryGetActive(progId)
        If running IsNot Nothing Then Return running
        Dim kind = Type.GetTypeFromProgID(progId, False)
        If kind Is Nothing Then Return Nothing
        Return Activator.CreateInstance(kind)
    End Function

    Private Shared Function TryGetActive(progId As String) As Object
        Try
            Dim clsid As Guid
            If CLSIDFromProgID(progId, clsid) <> 0 Then Return Nothing
            Dim active As Object = Nothing
            If GetActiveObject(clsid, IntPtr.Zero, active) <> 0 Then Return Nothing
            Return active
        Catch
            Return Nothing
        End Try
    End Function
End Class

Friend Class OfficeSession
    Implements IDisposable

    Private ReadOnly _host As OfficeStaHost
    Private _application As Object
    Private _released As Boolean

    Public ReadOnly Property Application As Object
        Get
            Return _application
        End Get
    End Property

    Private Sub New(host As OfficeStaHost, application As Object)
        _host = host
        _application = application
    End Sub

    Public Shared Async Function TryAttachAsync(progId As String, timeout As TimeSpan, cancel As CancellationToken, logStep As Action(Of String)) As Task(Of OfficeSession)
        Dim host As New OfficeStaHost()
        host.Start()
        If logStep IsNot Nothing Then logStep(progId & ": waiting to attach")
        Dim result = Await host.InvokeAsync(Function() ComActivation.TryAttach(progId), timeout, cancel).ConfigureAwait(True)
        If result.Kind = OfficeWaitKind.Cancelled Then
            If logStep IsNot Nothing Then logStep(progId & ": cancelled")
            host.StopPump()
            Return Nothing
        End If
        If result.Kind = OfficeWaitKind.TimedOut OrElse result.Problem IsNot Nothing OrElse result.Value Is Nothing Then
            If logStep IsNot Nothing Then logStep(progId & ": did not start")
            host.StopPump()
            Return Nothing
        End If
        If logStep IsNot Nothing Then logStep(progId & ": attached")
        Return New OfficeSession(host, result.Value)
    End Function

    Public Function InvokeAsync(work As Action, timeout As TimeSpan, cancel As CancellationToken) As Task(Of OfficeCallResult)
        Return _host.InvokeAsync(Function()
                                      work()
                                      Return True
                                  End Function, timeout, cancel)
    End Function

    Public Async Function ReleaseAsync() As Task
        If _released OrElse _host Is Nothing OrElse _host.Abandoned Then
            _released = True
            Return
        End If
        _released = True
        Dim app = _application
        _application = Nothing
        Await _host.InvokeAsync(Function()
                                     Try
                                         If app IsNot Nothing AndAlso Marshal.IsComObject(app) Then
                                             Marshal.FinalReleaseComObject(app)
                                         End If
                                     Catch
                                     End Try
                                     Return True
                                 End Function, TimeSpan.FromSeconds(5), CancellationToken.None).ConfigureAwait(True)
        _host.StopPump()
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        ReleaseAsync().GetAwaiter().GetResult()
    End Sub
End Class

Friend Class OfficeWaitDialog
    Inherits Form

    Private ReadOnly _cancel As New CancellationTokenSource()

    Public ReadOnly Property Token As CancellationToken
        Get
            Return _cancel.Token
        End Get
    End Property

    Public Sub New(message As String)
        Text = message
        FormBorderStyle = FormBorderStyle.FixedDialog
        ControlBox = False
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(420, 120)
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        Dim label As New Label With {
            .Text = message,
            .AutoSize = False,
            .Location = New Point(16, 16),
            .Size = New Size(388, 40)
        }
        Dim cancel As New Button With {
            .Text = "Cancel",
            .Location = New Point(160, 72),
            .Size = New Size(100, 28)
        }
        AddHandler cancel.Click, Sub()
                                      _cancel.Cancel()
                                      cancel.Enabled = False
                                      label.Text = "Cancelling..."
                                  End Sub
        Controls.Add(label)
        Controls.Add(cancel)
        CancelButton = cancel
    End Sub
End Class
