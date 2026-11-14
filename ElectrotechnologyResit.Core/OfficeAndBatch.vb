Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading

Namespace ElectrotechnologyResit.Core

    Public Enum OfficeWaitKind
        Completed = 0
        TimedOut = 1
        Cancelled = 2
    End Enum

    ''' <summary>Polls a wait handle so a caller can stop waiting without blocking a UI message pump.</summary>
    Public NotInheritable Class OfficeCallWait
        Private Sub New()
        End Sub

        Public Shared Function Wait(done As ManualResetEventSlim, timeout As TimeSpan, cancel As CancellationToken, poll As TimeSpan) As OfficeWaitKind
            If done Is Nothing Then Throw New ArgumentNullException(NameOf(done))
            If poll <= TimeSpan.Zero Then poll = TimeSpan.FromMilliseconds(50)
            If timeout < TimeSpan.Zero Then timeout = TimeSpan.Zero
            Dim started = Stopwatch.StartNew()
            While True
                If done.IsSet Then Return OfficeWaitKind.Completed
                If cancel.IsCancellationRequested Then Return OfficeWaitKind.Cancelled
                If started.Elapsed >= timeout Then Return OfficeWaitKind.TimedOut
                Dim slice = timeout - started.Elapsed
                If slice > poll Then slice = poll
                If slice < TimeSpan.Zero Then slice = TimeSpan.Zero
                done.Wait(slice)
            End While
            Return OfficeWaitKind.TimedOut
        End Function
    End Class

    ''' <summary>Step log for Outlook and Word. The text is reduced to a short label so a student id or address cannot be written.</summary>
    Public NotInheritable Class OfficeStepLog
        Private Shared ReadOnly Gate As New Object()

        Private Sub New()
        End Sub

        Public Shared Function Sanitise(stepName As String) As String
            Dim text = If(stepName, "").Trim()
            If text.Length > 80 Then text = text.Substring(0, 80)
            Dim cleaned As New StringBuilder()
            For Each ch In text
                If Char.IsLetterOrDigit(ch) OrElse ch = " "c OrElse ch = ":"c OrElse ch = "-"c OrElse ch = "/"c OrElse ch = "."c Then
                    cleaned.Append(ch)
                End If
            Next
            Dim compact = System.Text.RegularExpressions.Regex.Replace(cleaned.ToString(), "\d{6,}", "")
            Return compact.Trim()
        End Function

        Public Shared Sub Write(folder As String, stepName As String)
            Dim clean = Sanitise(stepName)
            If clean.Length = 0 OrElse String.IsNullOrWhiteSpace(folder) Then Return
            Dim line = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", Globalization.CultureInfo.InvariantCulture) & " " & clean & Environment.NewLine
            Try
                If Not Threading.Monitor.TryEnter(Gate, 200) Then Return
                Try
                    Directory.CreateDirectory(folder)
                    File.AppendAllText(Path.Combine(folder, "office-steps.log"), line)
                Finally
                    Threading.Monitor.Exit(Gate)
                End Try
            Catch
            End Try
        End Sub
    End Class

    Public NotInheritable Class OldResitPurge
        Private Sub New()
        End Sub

        Public Shared Function Cutoff(today As Date) As Date
            Return today.Date.AddDays(-14)
        End Function

        Public Shared Function Message(removed As Integer) As String
            Return "Removed " & removed.ToString() & " old resits"
        End Function
    End Class

    Public NotInheritable Class GraphBatch
        Public Const MaxRequests As Integer = 20

        Private Sub New()
        End Sub

        Public Shared Function DeleteBody(siteId As String, listId As String, itemIds As IEnumerable(Of String)) As String
            Using stream As New MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    writer.WriteStartArray("requests")
                    Dim n = 0
                    If itemIds IsNot Nothing Then
                        For Each itemId In itemIds
                            If String.IsNullOrWhiteSpace(itemId) Then Continue For
                            If n >= MaxRequests Then Exit For
                            n += 1
                            writer.WriteStartObject()
                            writer.WriteString("id", itemId)
                            writer.WriteString("method", "DELETE")
                            writer.WriteString("url", "/sites/" & siteId & "/lists/" & listId & "/items/" & itemId)
                            writer.WriteEndObject()
                        Next
                    End If
                    writer.WriteEndArray()
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function

        ''' <summary>Item ids whose batch response should be tried again one at a time. 404 is already gone.</summary>
        Public Shared Function RetryIds(responseJson As String) As List(Of String)
            Dim retry As New List(Of String)
            If String.IsNullOrWhiteSpace(responseJson) Then Return retry
            Try
                Using doc = JsonDocument.Parse(responseJson)
                    Dim responses As JsonElement
                    If Not doc.RootElement.TryGetProperty("responses", responses) OrElse responses.ValueKind <> JsonValueKind.Array Then Return retry
                    For Each item In responses.EnumerateArray()
                        Dim status = 0
                        Dim statusElement As JsonElement
                        If item.TryGetProperty("status", statusElement) AndAlso statusElement.ValueKind = JsonValueKind.Number Then
                            status = statusElement.GetInt32()
                        End If
                        If status >= 200 AndAlso status <= 299 Then Continue For
                        If status = 404 Then Continue For
                        Dim idElement As JsonElement
                        If item.TryGetProperty("id", idElement) AndAlso idElement.ValueKind = JsonValueKind.String Then
                            Dim id = idElement.GetString()
                            If Not String.IsNullOrWhiteSpace(id) Then retry.Add(id)
                        End If
                    Next
                End Using
            Catch
            End Try
            Return retry
        End Function
    End Class

End Namespace
