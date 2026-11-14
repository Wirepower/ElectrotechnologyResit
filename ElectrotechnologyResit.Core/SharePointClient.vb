Imports System.Globalization
Imports System.Text
Imports System.Text.Json

Namespace ElectrotechnologyResit.Core

    Public NotInheritable Class GraphUrls
        Public Const Root As String = "https://graph.microsoft.com/v1.0"

        Private Sub New()
        End Sub

        Public Shared Function Site(siteUrl As String) As String
            Dim uri As New Uri(siteUrl.Trim())
            Dim path = uri.AbsolutePath.TrimEnd("/"c)
            If path = "" OrElse path = "/" Then
                Return Root & "/sites/" & uri.Host
            End If
            Return Root & "/sites/" & uri.Host & ":" & path
        End Function

        Public Shared Function Items(siteId As String, listId As String, filter As String, top As Integer) As String
            Dim url = Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) &
                "/items?$expand=fields&$top=" & top.ToString(CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(filter) Then
                url &= "&$filter=" & Uri.EscapeDataString(filter)
            End If
            Return url
        End Function
    End Class

    Public Class ParsedPage
        Public Property Items As New List(Of ListItemRecord)
        Public Property NextLink As String
    End Class

    Public NotInheritable Class GraphItemParser
        Private Sub New()
        End Sub

        Public Shared Function ParsePage(json As String) As ParsedPage
            Dim page As New ParsedPage()
            If String.IsNullOrWhiteSpace(json) Then Return page
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim values As JsonElement
                If root.TryGetProperty("value", values) AndAlso values.ValueKind = JsonValueKind.Array Then
                    For Each item In values.EnumerateArray()
                        page.Items.Add(ParseItem(item))
                    Next
                ElseIf root.TryGetProperty("fields", values) OrElse root.TryGetProperty("id", values) Then
                    page.Items.Add(ParseItem(root))
                End If
                Dim link As JsonElement
                If root.TryGetProperty("@odata.nextLink", link) AndAlso link.ValueKind = JsonValueKind.String Then
                    page.NextLink = link.GetString()
                End If
            End Using
            Return page
        End Function

        Private Shared Function ParseItem(item As JsonElement) As ListItemRecord
            Dim record As New ListItemRecord With {.Fields = New FieldBag()}
            Dim id As JsonElement
            If item.TryGetProperty("id", id) AndAlso id.ValueKind = JsonValueKind.String Then
                record.Id = id.GetString()
            End If
            Dim fields As JsonElement
            If item.TryGetProperty("fields", fields) AndAlso fields.ValueKind = JsonValueKind.Object Then
                For Each propertyValue In fields.EnumerateObject()
                    If propertyValue.Name.StartsWith("@", StringComparison.Ordinal) Then Continue For
                    Select Case propertyValue.Value.ValueKind
                        Case JsonValueKind.True
                            record.Fields.SetBool(propertyValue.Name, True)
                        Case JsonValueKind.False
                            record.Fields.SetBool(propertyValue.Name, False)
                        Case JsonValueKind.String
                            record.Fields.SetString(propertyValue.Name, propertyValue.Value.GetString())
                        Case JsonValueKind.Number
                            record.Fields.SetString(propertyValue.Name, propertyValue.Value.GetRawText())
                        Case JsonValueKind.Null
                            record.Fields.SetString(propertyValue.Name, "")
                    End Select
                Next
            End If
            Return record
        End Function

        Public Shared Function ReadStringProperty(json As String, name As String) As String
            Using doc = JsonDocument.Parse(json)
                Dim value As JsonElement
                If doc.RootElement.TryGetProperty(name, value) AndAlso value.ValueKind = JsonValueKind.String Then
                    Return value.GetString()
                End If
            End Using
            Return ""
        End Function
    End Class

    Public NotInheritable Class SharePointJson
        Private Sub New()
        End Sub

        Public Shared Function FieldsJson(text As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As String
            Using stream As New IO.MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    WriteFields(writer, text, flags)
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function

        Public Shared Function CreateItemJson(text As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As String
            Using stream As New IO.MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    writer.WritePropertyName("fields")
                    writer.WriteStartObject()
                    WriteFields(writer, text, flags)
                    writer.WriteEndObject()
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function

        Private Shared Sub WriteFields(writer As Utf8JsonWriter, text As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean))
            If text IsNot Nothing Then
                For Each pair In text
                    writer.WriteString(pair.Key, If(pair.Value, ""))
                Next
            End If
            If flags IsNot Nothing Then
                For Each pair In flags
                    writer.WriteBoolean(pair.Key, pair.Value)
                Next
            End If
        End Sub

        Public Shared Function ListJson(spec As ListSpec) As String
            Dim template = If(spec.Template = ListTemplateKind.DocumentLibrary, "documentLibrary", "genericList")
            Using stream As New IO.MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    writer.WriteString("displayName", spec.DisplayName)
                    writer.WritePropertyName("list")
                    writer.WriteStartObject()
                    writer.WriteString("template", template)
                    writer.WriteEndObject()
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function

        Public Shared Function ColumnJson(column As ColumnSpec) As String
            Using stream As New IO.MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    writer.WriteString("name", column.Name)
                    writer.WriteString("displayName", column.DisplayName)
                    If column.Indexed OrElse column.EnforceUniqueValues Then writer.WriteBoolean("indexed", True)
                    If column.EnforceUniqueValues Then writer.WriteBoolean("enforceUniqueValues", True)
                    Select Case column.Kind
                        Case ColumnKind.DateOnly
                            writer.WritePropertyName("dateTime")
                            writer.WriteStartObject()
                            writer.WriteString("format", "dateOnly")
                            writer.WriteEndObject()
                        Case ColumnKind.YesNo
                            writer.WritePropertyName("boolean")
                            writer.WriteStartObject()
                            writer.WriteEndObject()
                        Case Else
                            writer.WritePropertyName("text")
                            writer.WriteStartObject()
                            writer.WriteBoolean("allowMultipleLines", False)
                            writer.WriteNumber("maxLength", 255)
                            writer.WriteEndObject()
                    End Select
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function

        Public Shared Function ColumnPatchJson(indexed As Boolean, enforceUniqueValues As Boolean) As String
            Using stream As New IO.MemoryStream()
                Using writer As New Utf8JsonWriter(stream)
                    writer.WriteStartObject()
                    writer.WriteBoolean("indexed", indexed)
                    writer.WriteBoolean("enforceUniqueValues", enforceUniqueValues)
                    writer.WriteEndObject()
                End Using
                Return Encoding.UTF8.GetString(stream.ToArray())
            End Using
        End Function
    End Class

    Public NotInheritable Class BookingFieldMap
        Private Sub New()
        End Sub

        Public Shared Sub Fill(booking As ResitBooking, includeFlags As Boolean, ByRef text As Dictionary(Of String, String), ByRef flags As Dictionary(Of String, Boolean))
            text = New Dictionary(Of String, String) From {
                {"Title", If(booking.StudentId, "")},
                {SharePointColumns.StudentId, If(booking.StudentId, "")},
                {SharePointColumns.StudentFirstname, If(booking.FirstName, "")},
                {SharePointColumns.Surname, If(booking.Surname, "")},
                {SharePointColumns.Email, If(booking.Email, "")},
                {SharePointColumns.Unit, If(booking.UnitCode, "")},
                {SharePointColumns.UnitName, If(booking.UnitName, "")},
                {SharePointColumns.QualificationCode, If(booking.QualificationCode, "")},
                {SharePointColumns.Blockgroup, If(booking.BlockGroup, "")},
                {SharePointColumns.AllocatedTeacher, If(booking.TeacherName, "")},
                {SharePointColumns.AllocatedTeacherEmail, If(booking.TeacherEmail, "")},
                {SharePointColumns.AttemptNo, If(booking.AttemptNo, "")},
                {SharePointColumns.ResitDate, MelbourneCalendar.ToDateOnlyString(booking.ResitDate)}
            }
            flags = New Dictionary(Of String, Boolean)
            If includeFlags Then
                flags(SharePointColumns.EnergyspaceClassCreated) = booking.EnergyspaceCreated
                flags(SharePointColumns.EnergyspaceAssessmentBooked) = booking.EnergyspaceAssessmentBooked
            End If
        End Sub

        Public Shared Function ToBooking(item As ListItemRecord) As ResitBooking
            Dim fields = If(item.Fields, New FieldBag())
            Return New ResitBooking With {
                .SharePointItemId = item.Id,
                .StudentId = fields.GetString(SharePointColumns.StudentId),
                .FirstName = fields.GetString(SharePointColumns.StudentFirstname),
                .Surname = fields.GetString(SharePointColumns.Surname),
                .Email = fields.GetString(SharePointColumns.Email),
                .UnitCode = fields.GetString(SharePointColumns.Unit),
                .UnitName = fields.GetString(SharePointColumns.UnitName),
                .QualificationCode = fields.GetString(SharePointColumns.QualificationCode),
                .BlockGroup = fields.GetString(SharePointColumns.Blockgroup),
                .TeacherName = fields.GetString(SharePointColumns.AllocatedTeacher),
                .TeacherEmail = fields.GetString(SharePointColumns.AllocatedTeacherEmail),
                .AttemptNo = fields.GetString(SharePointColumns.AttemptNo),
                .ResitDate = MelbourneCalendar.ReadCalendarDate(fields.GetString(SharePointColumns.ResitDate)),
                .EnergyspaceCreated = fields.GetBool(SharePointColumns.EnergyspaceClassCreated),
                .EnergyspaceAssessmentBooked = fields.GetBool(SharePointColumns.EnergyspaceAssessmentBooked),
                .Status = fields.GetString(SharePointColumns.Status)
            }
        End Function
    End Class

    Public Interface ISharePointClient
        Function GetSiteAsync() As Task(Of SharePointSite)
        Function ReadListsAsync(displayNames As IEnumerable(Of String)) As Task(Of IReadOnlyList(Of ExistingList))
        Function CreateListAsync(spec As ListSpec) As Task
        Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task
        Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task
        Function QueryItemsAsync(listDisplayName As String, filter As String, allowNonIndexedFallback As Boolean) As Task(Of IReadOnlyList(Of ListItemRecord))
        Function CreateItemAsync(listDisplayName As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task(Of String)
        Function UpdateItemAsync(listDisplayName As String, itemId As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task
        Function DeleteItemAsync(listDisplayName As String, itemId As String) As Task
        Function DeleteItemsAsync(listDisplayName As String, itemIds As IReadOnlyList(Of String)) As Task
        Function DownloadSignatureAsync(libraryDisplayName As String) As Task(Of Byte())
    End Interface

    Public Class GraphSharePointClient
        Implements ISharePointClient

        Private ReadOnly _executor As GraphExecutor
        Private ReadOnly _siteUrl As String
        Private _siteId As String
        Private _listIds As Dictionary(Of String, String)

        Public Sub New(executor As GraphExecutor, siteUrl As String)
            _executor = executor
            _siteUrl = If(siteUrl, "").Trim()
        End Sub

        Public Sub InvalidateListCache()
            _listIds = Nothing
        End Sub

        Public Async Function GetSiteAsync() As Task(Of SharePointSite) Implements ISharePointClient.GetSiteAsync
            If String.IsNullOrWhiteSpace(_siteUrl) Then
                Throw New InvalidOperationException("Set the SharePoint site URL in Settings, then try again.")
            End If
            Dim result = Await SendAsync(New GraphCall With {.Method = "GET", .Url = GraphUrls.Site(_siteUrl)}).ConfigureAwait(False)
            _siteId = GraphItemParser.ReadStringProperty(result.Body, "id")
            Return New SharePointSite With {
                .Id = _siteId,
                .DisplayName = GraphItemParser.ReadStringProperty(result.Body, "displayName")
            }
        End Function

        Public Async Function ReadListsAsync(displayNames As IEnumerable(Of String)) As Task(Of IReadOnlyList(Of ExistingList)) Implements ISharePointClient.ReadListsAsync
            Dim wanted = New HashSet(Of String)(If(displayNames, Enumerable.Empty(Of String)()), StringComparer.OrdinalIgnoreCase)
            Dim metadata = Await ReadListMetadataAsync().ConfigureAwait(False)
            Dim found As New List(Of ExistingList)
            For Each list In metadata
                If Not wanted.Contains(list.DisplayName) Then Continue For
                list.Columns = Await ReadColumnsAsync(list.Id).ConfigureAwait(False)
                found.Add(list)
            Next
            Return found
        End Function

        Public Async Function CreateListAsync(spec As ListSpec) As Task Implements ISharePointClient.CreateListAsync
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Await SendAsync(New GraphCall With {
                .Method = "POST",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists",
                .JsonBody = SharePointJson.ListJson(spec)
            }).ConfigureAwait(False)
            InvalidateListCache()
        End Function

        Public Async Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task Implements ISharePointClient.AddColumnAsync
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Await SendAsync(New GraphCall With {
                .Method = "POST",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/columns",
                .JsonBody = SharePointJson.ColumnJson(column)
            }).ConfigureAwait(False)
        End Function

        Public Async Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task Implements ISharePointClient.UpdateColumnAsync
            Dim lists = Await ReadListsAsync(New String() {listDisplayName}).ConfigureAwait(False)
            Dim list = lists.FirstOrDefault()
            If list Is Nothing Then Throw New InvalidOperationException("SharePoint list '" & listDisplayName & "' was not found.")
            Dim column = list.Columns.FirstOrDefault(Function(item) String.Equals(item.Name, columnName, StringComparison.OrdinalIgnoreCase))
            If column Is Nothing OrElse String.IsNullOrEmpty(column.Id) Then
                Throw New InvalidOperationException("Column '" & columnName & "' was not found on '" & listDisplayName & "'.")
            End If
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Await SendAsync(New GraphCall With {
                .Method = "PATCH",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(list.Id) & "/columns/" & Uri.EscapeDataString(column.Id),
                .JsonBody = SharePointJson.ColumnPatchJson(indexed, enforceUniqueValues)
            }).ConfigureAwait(False)
        End Function

        Public Async Function QueryItemsAsync(listDisplayName As String, filter As String, allowNonIndexedFallback As Boolean) As Task(Of IReadOnlyList(Of ListItemRecord)) Implements ISharePointClient.QueryItemsAsync
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim url = GraphUrls.Items(siteId, listId, filter, 200)
            Dim all As New List(Of ListItemRecord)
            Dim prefer = False
            While Not String.IsNullOrEmpty(url)
                Dim request As New GraphCall With {
                    .Method = "GET",
                    .Url = url,
                    .AllowNonIndexedFallback = allowNonIndexedFallback,
                    .PreferNonIndexedFallback = prefer
                }
                Dim result = Await SendAsync(request).ConfigureAwait(False)
                prefer = request.PreferNonIndexedFallback
                Dim page = GraphItemParser.ParsePage(result.Body)
                all.AddRange(page.Items)
                url = page.NextLink
            End While
            Return all
        End Function

        Public Async Function CreateItemAsync(listDisplayName As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task(Of String) Implements ISharePointClient.CreateItemAsync
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim result = Await SendAsync(New GraphCall With {
                .Method = "POST",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/items",
                .JsonBody = SharePointJson.CreateItemJson(textFields, flags)
            }).ConfigureAwait(False)
            Return GraphItemParser.ReadStringProperty(result.Body, "id")
        End Function

        Public Async Function UpdateItemAsync(listDisplayName As String, itemId As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task Implements ISharePointClient.UpdateItemAsync
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Await SendAsync(New GraphCall With {
                .Method = "PATCH",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/items/" & Uri.EscapeDataString(itemId) & "/fields",
                .JsonBody = SharePointJson.FieldsJson(textFields, flags)
            }).ConfigureAwait(False)
        End Function

        Public Async Function DeleteItemAsync(listDisplayName As String, itemId As String) As Task Implements ISharePointClient.DeleteItemAsync
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Await SendAsync(New GraphCall With {
                .Method = "DELETE",
                .Url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/items/" & Uri.EscapeDataString(itemId)
            }).ConfigureAwait(False)
        End Function

        Public Async Function DeleteItemsAsync(listDisplayName As String, itemIds As IReadOnlyList(Of String)) As Task Implements ISharePointClient.DeleteItemsAsync
            Dim ids = If(itemIds, New List(Of String)()).Where(Function(id) Not String.IsNullOrWhiteSpace(id)).Distinct(StringComparer.Ordinal).ToList()
            If ids.Count = 0 Then Return
            Dim listId = Await ResolveListIdAsync(listDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim index = 0
            While index < ids.Count
                Dim chunk = ids.Skip(index).Take(GraphBatch.MaxRequests).ToList()
                index += chunk.Count
                Dim result = Await SendAsync(New GraphCall With {
                    .Method = "POST",
                    .Url = GraphUrls.Root & "/$batch",
                    .JsonBody = GraphBatch.DeleteBody(siteId, listId, chunk)
                }).ConfigureAwait(False)
                For Each itemId In GraphBatch.RetryIds(result.Body)
                    Await DeleteItemAsync(listDisplayName, itemId).ConfigureAwait(False)
                Next
            End While
        End Function

        Public Async Function DownloadSignatureAsync(libraryDisplayName As String) As Task(Of Byte()) Implements ISharePointClient.DownloadSignatureAsync
            Dim listId = Await ResolveListIdAsync(libraryDisplayName).ConfigureAwait(False)
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/drive/root/children?$top=200"
            Dim newestId As String = Nothing
            Dim newestStamp As DateTimeOffset = DateTimeOffset.MinValue
            While Not String.IsNullOrEmpty(url)
                Dim result As GraphResult
                Try
                    result = Await SendAsync(New GraphCall With {.Method = "GET", .Url = url}).ConfigureAwait(False)
                Catch ex As GraphRequestException When ex.StatusCode = 404
                    Return Nothing
                End Try
                Dim page = ParseDriveChildren(result.Body)
                For Each child In page.Files
                    If child.Modified >= newestStamp Then
                        newestStamp = child.Modified
                        newestId = child.Id
                    End If
                Next
                url = page.NextLink
            End While
            If String.IsNullOrEmpty(newestId) Then Return Nothing
            Dim contentUrl = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/drive/items/" & Uri.EscapeDataString(newestId) & "/content"
            Dim download = Await SendAsync(New GraphCall With {.Method = "GET", .Url = contentUrl}).ConfigureAwait(False)
            Return download.RawBytes
        End Function

        Private Async Function SendAsync(request As GraphCall) As Task(Of GraphResult)
            Return Await _executor.SendAsync(request).ConfigureAwait(False)
        End Function

        Private Async Function EnsureSiteIdAsync() As Task(Of String)
            If Not String.IsNullOrEmpty(_siteId) Then Return _siteId
            Dim site = Await GetSiteAsync().ConfigureAwait(False)
            Return site.Id
        End Function

        Private Async Function ResolveListIdAsync(displayName As String) As Task(Of String)
            If _listIds Is Nothing Then
                _listIds = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each list In Await ReadListMetadataAsync().ConfigureAwait(False)
                    _listIds(list.DisplayName) = list.Id
                Next
            End If
            Dim id As String = Nothing
            If _listIds.TryGetValue(displayName, id) Then Return id
            Throw New InvalidOperationException("SharePoint list '" & displayName & "' was not found. Run SharePoint setup to create it.")
        End Function

        Private Async Function ReadListMetadataAsync() As Task(Of List(Of ExistingList))
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists?$select=id,displayName,list&$top=200"
            Dim lists As New List(Of ExistingList)
            While Not String.IsNullOrEmpty(url)
                Dim result = Await SendAsync(New GraphCall With {.Method = "GET", .Url = url}).ConfigureAwait(False)
                Using doc = JsonDocument.Parse(result.Body)
                    Dim values As JsonElement
                    If doc.RootElement.TryGetProperty("value", values) Then
                        For Each item In values.EnumerateArray()
                            Dim hidden = False
                            Dim listElement As JsonElement
                            If item.TryGetProperty("list", listElement) AndAlso listElement.ValueKind = JsonValueKind.Object Then
                                Dim hiddenElement As JsonElement
                                If listElement.TryGetProperty("hidden", hiddenElement) AndAlso hiddenElement.ValueKind = JsonValueKind.True Then hidden = True
                            End If
                            If hidden Then Continue For
                            Dim template = ListTemplateKind.GenericList
                            If item.TryGetProperty("list", listElement) AndAlso listElement.ValueKind = JsonValueKind.Object Then
                                Dim templateElement As JsonElement
                                If listElement.TryGetProperty("template", templateElement) AndAlso templateElement.ValueKind = JsonValueKind.String AndAlso String.Equals(templateElement.GetString(), "documentLibrary", StringComparison.OrdinalIgnoreCase) Then
                                    template = ListTemplateKind.DocumentLibrary
                                End If
                            End If
                            Dim id = ""
                            Dim name = ""
                            Dim idElement As JsonElement
                            Dim nameElement As JsonElement
                            If item.TryGetProperty("id", idElement) Then id = If(idElement.GetString(), "")
                            If item.TryGetProperty("displayName", nameElement) Then name = If(nameElement.GetString(), "")
                            lists.Add(New ExistingList With {.Id = id, .DisplayName = name, .Template = template, .Columns = New List(Of ExistingColumn)})
                        Next
                    End If
                    Dim link As JsonElement
                    If doc.RootElement.TryGetProperty("@odata.nextLink", link) Then
                        url = link.GetString()
                    Else
                        url = Nothing
                    End If
                End Using
            End While
            Return lists
        End Function

        Private Async Function ReadColumnsAsync(listId As String) As Task(Of List(Of ExistingColumn))
            Dim siteId = Await EnsureSiteIdAsync().ConfigureAwait(False)
            Dim url = GraphUrls.Root & "/sites/" & Uri.EscapeDataString(siteId) & "/lists/" & Uri.EscapeDataString(listId) & "/columns?$select=id,name,indexed,enforceUniqueValues&$top=200"
            Dim columns As New List(Of ExistingColumn)
            While Not String.IsNullOrEmpty(url)
                Dim result = Await SendAsync(New GraphCall With {.Method = "GET", .Url = url}).ConfigureAwait(False)
                Using doc = JsonDocument.Parse(result.Body)
                    Dim values As JsonElement
                    If doc.RootElement.TryGetProperty("value", values) Then
                        For Each item In values.EnumerateArray()
                            Dim column As New ExistingColumn()
                            Dim element As JsonElement
                            If item.TryGetProperty("id", element) Then column.Id = element.GetString()
                            If item.TryGetProperty("name", element) Then column.Name = element.GetString()
                            If item.TryGetProperty("indexed", element) AndAlso element.ValueKind = JsonValueKind.True Then column.Indexed = True
                            If item.TryGetProperty("enforceUniqueValues", element) AndAlso element.ValueKind = JsonValueKind.True Then column.EnforceUniqueValues = True
                            columns.Add(column)
                        Next
                    End If
                    Dim link As JsonElement
                    If doc.RootElement.TryGetProperty("@odata.nextLink", link) Then
                        url = link.GetString()
                    Else
                        url = Nothing
                    End If
                End Using
            End While
            Return columns
        End Function

        Private Class DriveFile
            Public Property Id As String
            Public Property Modified As DateTimeOffset
        End Class

        Private Class DrivePage
            Public Property Files As New List(Of DriveFile)
            Public Property NextLink As String
        End Class

        Private Shared Function ParseDriveChildren(json As String) As DrivePage
            Dim page As New DrivePage()
            Using doc = JsonDocument.Parse(json)
                Dim values As JsonElement
                If doc.RootElement.TryGetProperty("value", values) Then
                    For Each item In values.EnumerateArray()
                        Dim fileElement As JsonElement
                        If Not item.TryGetProperty("file", fileElement) Then Continue For
                        Dim name = ""
                        Dim nameElement As JsonElement
                        If item.TryGetProperty("name", nameElement) Then name = If(nameElement.GetString(), "")
                        Dim extension = IO.Path.GetExtension(name).ToLowerInvariant()
                        If extension <> ".jpg" AndAlso extension <> ".jpeg" AndAlso extension <> ".png" AndAlso extension <> ".gif" AndAlso extension <> ".bmp" Then Continue For
                        Dim id = ""
                        Dim idElement As JsonElement
                        If item.TryGetProperty("id", idElement) Then id = If(idElement.GetString(), "")
                        Dim modified = DateTimeOffset.MinValue
                        Dim modifiedElement As JsonElement
                        If item.TryGetProperty("lastModifiedDateTime", modifiedElement) Then
                            Dim parsed As DateTimeOffset
                            If DateTimeOffset.TryParse(modifiedElement.GetString(), parsed) Then modified = parsed
                        End If
                        page.Files.Add(New DriveFile With {.Id = id, .Modified = modified})
                    Next
                End If
                Dim link As JsonElement
                If doc.RootElement.TryGetProperty("@odata.nextLink", link) Then page.NextLink = link.GetString()
            End Using
            Return page
        End Function
    End Class

    Public Class GraphSchemaStore
        Implements ISharePointSchemaStore

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _names As SharePointListNames

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _names = If(names, SharePointListNames.Default())
        End Sub

        Public Function ReadListsAsync() As Task(Of IReadOnlyList(Of ExistingList)) Implements ISharePointSchemaStore.ReadListsAsync
            Return _client.ReadListsAsync(_names.AllNames())
        End Function

        Public Function CreateListAsync(spec As ListSpec) As Task Implements ISharePointSchemaStore.CreateListAsync
            Return _client.CreateListAsync(spec)
        End Function

        Public Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task Implements ISharePointSchemaStore.AddColumnAsync
            Return _client.AddColumnAsync(listDisplayName, column)
        End Function

        Public Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task Implements ISharePointSchemaStore.UpdateColumnAsync
            Return _client.UpdateColumnAsync(listDisplayName, columnName, indexed, enforceUniqueValues)
        End Function
    End Class

End Namespace
