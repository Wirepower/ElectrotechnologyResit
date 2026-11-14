Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text

Namespace ElectrotechnologyResit.Core

    Public Class GraphCall
        Public Property Method As String = "GET"
        Public Property Url As String
        Public Property JsonBody As String
        Public Property AllowNonIndexedFallback As Boolean
        Public Property PreferNonIndexedFallback As Boolean
    End Class

    Public Class GraphResult
        Public Property StatusCode As Integer
        Public Property Body As String
        Public Property RawBytes As Byte()
        Public Property RetryAfter As String

        Public ReadOnly Property IsSuccess As Boolean
            Get
                Return StatusCode >= 200 AndAlso StatusCode <= 299
            End Get
        End Property
    End Class

    Public Class GraphRequestException
        Inherits Exception

        Public Property StatusCode As Integer
        Public Property ResponseBody As String

        Public Sub New(statusCode As Integer, responseBody As String)
            MyBase.New(BuildMessage(statusCode, responseBody))
            Me.StatusCode = statusCode
            Me.ResponseBody = If(responseBody, "")
        End Sub

        Private Shared Function BuildMessage(statusCode As Integer, responseBody As String) As String
            Dim body = If(responseBody, "")
            If body.Length > 600 Then body = body.Substring(0, 600)
            Return "Microsoft Graph error (" & statusCode.ToString() & "): " & body
        End Function
    End Class

    Public Interface IAccessTokenProvider
        Function GetAccessTokenAsync() As Task(Of String)
    End Interface

    Public Interface IGraphTransport
        Function SendAsync(request As GraphCall) As Task(Of GraphResult)
    End Interface

    Public Class GraphExecutor
        Private ReadOnly _transport As IGraphTransport
        Private ReadOnly _delay As Func(Of TimeSpan, Task)

        Public Sub New(transport As IGraphTransport)
            Me.New(transport, Function(wait As TimeSpan) Task.Delay(wait))
        End Sub

        Public Sub New(transport As IGraphTransport, delayAsync As Func(Of TimeSpan, Task))
            If transport Is Nothing Then Throw New ArgumentNullException(NameOf(transport))
            _transport = transport
            _delay = If(delayAsync, Function(wait As TimeSpan) Task.Delay(wait))
        End Sub

        Public Async Function SendAsync(request As GraphCall) As Task(Of GraphResult)
            If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))
            Dim attempt As Integer = 0
            Do
                attempt += 1
                Dim result = Await _transport.SendAsync(request).ConfigureAwait(False)
                If result IsNot Nothing AndAlso result.IsSuccess Then Return result

                Dim status = If(result Is Nothing, 0, result.StatusCode)
                Dim body = If(result Is Nothing, "", result.Body)
                Dim retryAfter = If(result Is Nothing, "", result.RetryAfter)

                If request.AllowNonIndexedFallback AndAlso Not request.PreferNonIndexedFallback AndAlso GraphRetry.IsNonIndexedFilterError(status, body) Then
                    request.PreferNonIndexedFallback = True
                    Continue Do
                End If

                Dim wait = GraphRetry.GetDelay(status, retryAfter, attempt)
                If Not wait.HasValue Then
                    Throw New GraphRequestException(status, body)
                End If
                Await _delay.Invoke(wait.Value).ConfigureAwait(False)
            Loop
        End Function
    End Class

    Public Class GraphHttpTransport
        Implements IGraphTransport

        Private ReadOnly _http As HttpClient
        Private ReadOnly _tokens As IAccessTokenProvider

        Public Sub New(tokens As IAccessTokenProvider)
            Me.New(tokens, Nothing)
        End Sub

        Public Sub New(tokens As IAccessTokenProvider, handler As HttpMessageHandler)
            _tokens = tokens
            If handler Is Nothing Then
                _http = New HttpClient() With {.Timeout = TimeSpan.FromMinutes(2)}
            Else
                _http = New HttpClient(handler) With {.Timeout = TimeSpan.FromMinutes(2)}
            End If
        End Sub

        Public Async Function SendAsync(request As GraphCall) As Task(Of GraphResult) Implements IGraphTransport.SendAsync
            Dim token = ""
            If _tokens IsNot Nothing Then token = Await _tokens.GetAccessTokenAsync().ConfigureAwait(False)
            Using message As New HttpRequestMessage(New HttpMethod(If(request.Method, "GET")), request.Url)
                If Not String.IsNullOrEmpty(token) Then
                    message.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                End If
                message.Headers.Accept.ParseAdd("application/json")
                If request.PreferNonIndexedFallback Then
                    message.Headers.TryAddWithoutValidation(GraphRetry.PreferHeaderName, GraphRetry.PreferNonIndexedValue)
                End If
                If request.JsonBody IsNot Nothing Then
                    message.Content = New StringContent(request.JsonBody, Encoding.UTF8, "application/json")
                End If
                Dim response = Await _http.SendAsync(message).ConfigureAwait(False)
                Dim bytes = Await response.Content.ReadAsByteArrayAsync().ConfigureAwait(False)
                Dim retryAfter = ""
                If response.Headers.RetryAfter IsNot Nothing Then
                    If response.Headers.RetryAfter.Delta.HasValue Then
                        retryAfter = response.Headers.RetryAfter.Delta.Value.TotalSeconds.ToString(Globalization.CultureInfo.InvariantCulture)
                    ElseIf response.Headers.RetryAfter.Date.HasValue Then
                        retryAfter = response.Headers.RetryAfter.Date.Value.ToString("r", Globalization.CultureInfo.InvariantCulture)
                    End If
                End If
                Return New GraphResult With {
                    .StatusCode = CInt(response.StatusCode),
                    .RawBytes = bytes,
                    .Body = Encoding.UTF8.GetString(bytes),
                    .RetryAfter = retryAfter
                }
            End Using
        End Function
    End Class

End Namespace
