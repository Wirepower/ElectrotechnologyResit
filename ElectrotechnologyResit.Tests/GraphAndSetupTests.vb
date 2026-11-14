Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports ElectrotechnologyResit.Core
Imports OfficeOpenXml
Imports Xunit

Public Class GraphAndSetupTests
    <Fact>
    Public Async Function Paging_follows_the_next_link() As Task
        Dim transport As New ScriptedTransport()
        transport.Responses.Enqueue(JsonResult(200, "{""id"":""site-1"",""displayName"":""Electrotechnology""}"))
        transport.Responses.Enqueue(JsonResult(200, "{""value"":[{""id"":""list-1"",""displayName"":""ResitBookings"",""list"":{""template"":""genericList"",""hidden"":false}}]}"))
        transport.Responses.Enqueue(JsonResult(200, "{""value"":[{""id"":""a"",""fields"":{""StudentID"":""9000001""}}],""@odata.nextLink"":""https://graph.example/page2""}"))
        transport.Responses.Enqueue(JsonResult(200, "{""value"":[{""id"":""b"",""fields"":{""StudentID"":""9000002""}}]}"))
        Dim client As New GraphSharePointClient(New GraphExecutor(transport), "https://contoso.sharepoint.com/sites/Electrotechnology")
        Dim items = Await client.QueryItemsAsync("ResitBookings", "fields/StudentID eq '9000001'", True)
        Assert.Equal(2, items.Count)
        Assert.Equal("9000001", items(0).Fields.GetString("StudentID"))
        Assert.Equal("9000002", items(1).Fields.GetString("StudentID"))
        Assert.Contains(transport.Requests, Function(request) request.Url = "https://graph.example/page2")
        Assert.Contains(transport.Requests, Function(request) request.Url.Contains("$filter="))
    End Function

    <Fact>
    Public Async Function Throttling_waits_for_retry_after() As Task
        Dim transport As New ScriptedTransport()
        transport.Responses.Enqueue(New GraphResult With {.StatusCode = 429, .Body = "slow down", .RetryAfter = "3"})
        transport.Responses.Enqueue(JsonResult(200, "{""ok"":true}"))
        Dim waits As New List(Of TimeSpan)
        Dim executor As New GraphExecutor(transport, Function(wait As TimeSpan)
                                                            waits.Add(wait)
                                                            Return Task.CompletedTask
                                                        End Function)
        Dim result = Await executor.SendAsync(New GraphCall With {.Method = "GET", .Url = "https://graph.example/items"})
        Assert.True(result.IsSuccess)
        Assert.Equal(TimeSpan.FromSeconds(3), waits.Single())
        Assert.Equal(2, transport.Requests.Count)
    End Function

    <Fact>
    Public Async Function Service_unavailable_also_retries() As Task
        Dim transport As New ScriptedTransport()
        transport.Responses.Enqueue(New GraphResult With {.StatusCode = 503, .Body = "unavailable", .RetryAfter = "1"})
        transport.Responses.Enqueue(JsonResult(200, "{}"))
        Dim executor As New GraphExecutor(transport, Function(wait As TimeSpan) Task.CompletedTask)
        Dim result = Await executor.SendAsync(New GraphCall With {.Url = "https://graph.example/items"})
        Assert.Equal(200, result.StatusCode)
    End Function

    <Fact>
    Public Async Function Non_indexed_filter_retries_once_with_the_prefer_header() As Task
        Dim transport As New ScriptedTransport()
        transport.Responses.Enqueue(New GraphResult With {.StatusCode = 400, .Body = "{""error"":{""message"":""Field StudentID is not indexed""}}"})
        transport.Responses.Enqueue(JsonResult(200, "{""value"":[]}"))
        Dim executor As New GraphExecutor(transport, Function(wait As TimeSpan) Task.CompletedTask)
        Dim result = Await executor.SendAsync(New GraphCall With {.Url = "https://graph.example/items", .AllowNonIndexedFallback = True})
        Assert.True(result.IsSuccess)
        Assert.False(transport.Requests(0).PreferNonIndexedFallback)
        Assert.True(transport.Requests(1).PreferNonIndexedFallback)
        Assert.Equal(GraphRetry.PreferNonIndexedValue, "HonorNonIndexedQueriesWarningMayFailRandomly")
    End Function

    <Fact>
    Public Async Function Prefer_header_is_not_sent_until_the_fallback() As Task
        Dim handler As New RecordingHandler()
        Dim transport As New GraphHttpTransport(New FixedToken(), handler)
        Await transport.SendAsync(New GraphCall With {.Method = "GET", .Url = "https://graph.example/items", .PreferNonIndexedFallback = False})
        Assert.False(handler.LastRequest.Headers.Contains(GraphRetry.PreferHeaderName))
        Await transport.SendAsync(New GraphCall With {.Method = "GET", .Url = "https://graph.example/items", .PreferNonIndexedFallback = True})
        Assert.Equal(GraphRetry.PreferNonIndexedValue, handler.LastRequest.Headers.GetValues(GraphRetry.PreferHeaderName).Single())
    End Function

    <Fact>
    Public Async Function Setup_is_idempotent_and_never_deletes() As Task
        Dim store As New MemorySchemaStore()
        Dim provisioner As New ListProvisioner(store, SharePointListNames.Default())
        Dim first = Await provisioner.EnsureAsync()
        Assert.True(first.IsSatisfied)
        Assert.DoesNotContain(store.Operations, Function(item) item.StartsWith("delete", StringComparison.OrdinalIgnoreCase))
        Dim afterFirst = store.Operations.Count
        Dim second = Await provisioner.EnsureAsync()
        Assert.True(second.IsSatisfied)
        Assert.Equal(afterFirst, store.Operations.Count)
        Assert.Equal("kept", store.Marker)

        Dim bookings = store.Lists.Single(Function(list) list.DisplayName = "ResitBookings")
        bookings.Columns.RemoveAll(Function(column) column.Name = SharePointColumns.AttemptNo)
        bookings.Columns.Single(Function(column) column.Name = SharePointColumns.ResitDate).Indexed = False
        Dim beforeRepair = store.Operations.Count
        Dim repaired = Await provisioner.EnsureAsync()
        Assert.True(repaired.IsSatisfied)
        Assert.Contains(store.Operations.Skip(beforeRepair), Function(item) item = "add:ResitBookings.AttemptNo")
        Assert.Contains(store.Operations.Skip(beforeRepair), Function(item) item.StartsWith("patch:ResitBookings.ResitDate"))
        Assert.True(bookings.Columns.Single(Function(column) column.Name = SharePointColumns.ResitDate).Indexed)
        Assert.True(bookings.Columns.Single(Function(column) column.Name = SharePointColumns.StudentId).EnforceUniqueValues)
    End Function

    <Fact>
    Public Sub Schema_indexes_student_id_and_resit_date()
        Dim spec = SharePointListSchema.Build(SharePointListNames.Default()).Single(Function(list) list.DisplayName = "ResitBookings")
        Dim student = spec.Columns.Single(Function(column) column.Name = SharePointColumns.StudentId)
        Dim resitDate = spec.Columns.Single(Function(column) column.Name = SharePointColumns.ResitDate)
        Assert.True(student.Indexed)
        Assert.True(student.EnforceUniqueValues)
        Assert.Equal(ColumnKind.DateOnly, resitDate.Kind)
        Assert.True(resitDate.Indexed)
        Assert.False(resitDate.EnforceUniqueValues)
        Dim signature = SharePointListSchema.Build(Nothing).Single(Function(list) list.DisplayName = "EmailSignature")
        Assert.Equal(ListTemplateKind.DocumentLibrary, signature.Template)
    End Sub

    <Fact>
    Public Sub Setup_validation_rejects_a_blank_site_and_duplicate_names()
        Assert.Null(SharePointSetupValidation.Validate("https://contoso.sharepoint.com/sites/Electro", "app-id", SharePointListNames.Default()))
        Assert.NotNull(SharePointSetupValidation.Validate("http://contoso.sharepoint.com/sites/Electro", "app-id", SharePointListNames.Default()))
        Assert.NotNull(SharePointSetupValidation.Validate("https://contoso.sharepoint.com/sites/Electro", "  ", SharePointListNames.Default()))
        Dim names = SharePointListNames.Default()
        names.Students = names.Bookings
        Assert.Equal("Each list needs its own name.", SharePointSetupValidation.Validate("https://contoso.sharepoint.com/sites/Electro", "app-id", names))
    End Sub

    <Fact>
    Public Async Function Csv_and_excel_import_upsert_without_deleting() As Task
        Dim store As New InMemoryResitStore()
        Await store.UpsertAsync(New StudentRecord With {.StudentId = "9000001", .GivenName = "Old", .FamilyName = "Name"})
        Await store.UpsertAsync(New StudentRecord With {.StudentId = "9000004", .GivenName = "Drew", .FamilyName = "Singh"})
        Dim importer As New LookupImportService(store, store, store)
        Dim folder = Path.Combine(Path.GetTempPath(), "resit-import-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(folder)
        Try
            Dim csvPath = Path.Combine(folder, "students.csv")
            File.WriteAllText(csvPath, "Student ID,Given Name,Family Name,Personal Email,Block Group Code" & vbCrLf &
                "9000001,Alex,""Nguyen, Alex"",alex.nguyen@example.com,BG-A" & vbCrLf &
                "9000002,Blair,Okafor,blair.okafor@example.com,BG-B" & vbCrLf)
            Dim parsed = LookupFileReader.ParseCsv("A,""B, C"",D")
            Assert.Equal("A", parsed(0)(0))
            Assert.Equal("B, C", parsed(0)(1))
            Assert.Equal("D", parsed(0)(2))
            Dim tally = Await importer.ImportStudentsAsync(LookupFileReader.ReadStudents(csvPath))
            Assert.Equal(1, tally.Inserted)
            Assert.Equal(1, tally.Updated)
            Assert.Equal("Alex", (Await store.FindByStudentIdAsync("9000001")).GivenName)
            Assert.Equal("Drew", (Await store.FindByStudentIdAsync("9000004")).GivenName)

            Dim xlsx = Path.Combine(folder, "teachers.xlsx")
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Using package As New ExcelPackage()
                Dim sheet = package.Workbook.Worksheets.Add("Teachers")
                sheet.Cells(1, 1).Value = "Full Name"
                sheet.Cells(1, 2).Value = "Email"
                sheet.Cells(2, 1).Value = "Jordan Hale"
                sheet.Cells(2, 2).Value = "jordan.hale@example.com"
                package.SaveAs(New FileInfo(xlsx))
            End Using
            Dim teachers = Await importer.ImportTeachersAsync(LookupFileReader.ReadTeachers(xlsx))
            Assert.Equal(1, teachers.Inserted)
            Assert.Equal("jordan.hale@example.com", Await store.FindEmailAsync("Jordan Hale"))
        Finally
            Directory.Delete(folder, True)
        End Try
    End Function

    Private Shared Function JsonResult(status As Integer, body As String) As GraphResult
        Return New GraphResult With {.StatusCode = status, .Body = body, .RawBytes = Encoding.UTF8.GetBytes(body)}
    End Function
End Class

Public Class ScriptedTransport
    Implements IGraphTransport

    Public Property Responses As New Queue(Of GraphResult)
    Public Property Requests As New List(Of GraphCall)

    Public Function SendAsync(request As GraphCall) As Task(Of GraphResult) Implements IGraphTransport.SendAsync
        Requests.Add(New GraphCall With {
            .Method = request.Method,
            .Url = request.Url,
            .PreferNonIndexedFallback = request.PreferNonIndexedFallback,
            .AllowNonIndexedFallback = request.AllowNonIndexedFallback
        })
        Return Task.FromResult(Responses.Dequeue())
    End Function
End Class

Public Class FixedToken
    Implements IAccessTokenProvider

    Public Function GetAccessTokenAsync() As Task(Of String) Implements IAccessTokenProvider.GetAccessTokenAsync
        Return Task.FromResult("test-token")
    End Function
End Class

Public Class RecordingHandler
    Inherits HttpMessageHandler

    Public Property LastRequest As HttpRequestMessage

    Protected Overrides Function SendAsync(request As HttpRequestMessage, cancellationToken As Threading.CancellationToken) As Task(Of HttpResponseMessage)
        LastRequest = request
        Dim response As New HttpResponseMessage(HttpStatusCode.OK)
        response.Content = New StringContent("{}")
        Return Task.FromResult(response)
    End Function
End Class
