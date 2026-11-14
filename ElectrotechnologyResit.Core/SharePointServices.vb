Namespace ElectrotechnologyResit.Core

    Public Class SharePointBookingService
        Implements IBookingService

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _teachers As ITeacherDirectory
        Private ReadOnly _listName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames, teachers As ITeacherDirectory)
            _client = client
            _teachers = teachers
            _listName = If(names, SharePointListNames.Default()).Bookings
        End Sub

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetAllAsync
            Dim items = Await _client.QueryItemsAsync(_listName, Nothing, False).ConfigureAwait(False)
            Return items.Select(AddressOf BookingFieldMap.ToBooking).ToList()
        End Function

        Public Async Function GetByDateAsync(resitDate As Date) As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetByDateAsync
            Dim filter = ODataFilters.DateEquals(SharePointColumns.ResitDate, resitDate)
            Dim items = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Return items.Select(AddressOf BookingFieldMap.ToBooking).Where(Function(item) item.ResitDate.Date = resitDate.Date).ToList()
        End Function

        Public Async Function UpsertAsync(booking As ResitBooking, preserveEnergyspaceFlags As Boolean) As Task(Of UpsertResult) Implements IBookingService.UpsertAsync
            Dim id = If(booking?.StudentId, "").Trim()
            If id.Length = 0 Then Throw New ArgumentException("A student ID is required.")
            Dim filter = ODataFilters.TextEquals(SharePointColumns.StudentId, id)
            Dim existingItems = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim existing = existingItems.Select(AddressOf BookingFieldMap.ToBooking).FirstOrDefault(Function(item) String.Equals(item.StudentId, id, StringComparison.OrdinalIgnoreCase))
            Dim teachers = If(_teachers Is Nothing, Enumerable.Empty(Of TeacherRecord)(), Await _teachers.GetAllAsync().ConfigureAwait(False))
            Dim merged = BookingUpsert.Merge(existing, booking, preserveEnergyspaceFlags, teachers)
            Dim includeFlags = existing Is Nothing OrElse Not preserveEnergyspaceFlags
            Dim text As Dictionary(Of String, String) = Nothing
            Dim flags As Dictionary(Of String, Boolean) = Nothing
            BookingFieldMap.Fill(merged, includeFlags, text, flags)
            If existing Is Nothing Then
                merged.SharePointItemId = Await _client.CreateItemAsync(_listName, text, flags).ConfigureAwait(False)
                Return New UpsertResult With {.Kind = UpsertKind.Inserted, .Booking = merged}
            End If
            Await _client.UpdateItemAsync(_listName, existing.SharePointItemId, text, flags).ConfigureAwait(False)
            merged.SharePointItemId = existing.SharePointItemId
            Return New UpsertResult With {.Kind = UpsertKind.Updated, .Booking = merged}
        End Function

        Public Async Function DeleteByStudentIdAsync(studentId As String) As Task Implements IBookingService.DeleteByStudentIdAsync
            Dim filter = ODataFilters.TextEquals(SharePointColumns.StudentId, If(studentId, "").Trim())
            Dim items = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            For Each item In items
                Await _client.DeleteItemAsync(_listName, item.Id).ConfigureAwait(False)
            Next
        End Function

        Public Async Function DeleteByDateAsync(resitDate As Date) As Task Implements IBookingService.DeleteByDateAsync
            Dim rows = Await GetByDateAsync(resitDate).ConfigureAwait(False)
            For Each row In rows
                Await _client.DeleteItemAsync(_listName, row.SharePointItemId).ConfigureAwait(False)
            Next
        End Function

        Public Async Function DeleteOnOrBeforeAsync(cutoff As Date) As Task(Of Integer) Implements IBookingService.DeleteOnOrBeforeAsync
            Dim filter = ODataFilters.DateOnOrBefore(SharePointColumns.ResitDate, cutoff)
            Dim items = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim ids = items.Select(AddressOf BookingFieldMap.ToBooking).
                Where(Function(item) item.ResitDate.Date <= cutoff.Date AndAlso Not String.IsNullOrEmpty(item.SharePointItemId)).
                Select(Function(item) item.SharePointItemId).
                ToList()
            If ids.Count = 0 Then Return 0
            Await _client.DeleteItemsAsync(_listName, ids).ConfigureAwait(False)
            Return ids.Count
        End Function

        Public Async Function SetEnergyspaceCreatedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceCreatedAsync
            Await SetFlagAsync(resitDate, unitCode, SharePointColumns.EnergyspaceClassCreated, value).ConfigureAwait(False)
        End Function

        Public Async Function SetEnergyspaceAssessmentBookedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceAssessmentBookedAsync
            Await SetFlagAsync(resitDate, unitCode, SharePointColumns.EnergyspaceAssessmentBooked, value).ConfigureAwait(False)
        End Function

        Public Async Function GetUnitFlagsAsync(resitDate As Date) As Task(Of IReadOnlyList(Of UnitNightFlags)) Implements IBookingService.GetUnitFlagsAsync
            Dim rows = Await GetByDateAsync(resitDate).ConfigureAwait(False)
            Return EnergyspaceFlagRules.Summarise(rows, resitDate)
        End Function

        Private Async Function SetFlagAsync(resitDate As Date, unitCode As String, columnName As String, value As Boolean) As Task
            Dim rows = Await GetByDateAsync(resitDate).ConfigureAwait(False)
            Dim flags As New Dictionary(Of String, Boolean) From {{columnName, value}}
            For Each row In rows
                If String.Equals(row.UnitCode, unitCode, StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrEmpty(row.SharePointItemId) Then
                    Await _client.UpdateItemAsync(_listName, row.SharePointItemId, New Dictionary(Of String, String), flags).ConfigureAwait(False)
                End If
            Next
        End Function
    End Class

    Public Class SharePointStudentDirectory
        Implements IStudentDirectory

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _listName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _listName = If(names, SharePointListNames.Default()).Students
        End Sub

        Public Async Function FindByStudentIdAsync(studentId As String) As Task(Of StudentRecord) Implements IStudentDirectory.FindByStudentIdAsync
            Dim id = If(studentId, "").Trim()
            If id.Length = 0 Then Return Nothing
            Dim filter = ODataFilters.TextEquals(SharePointColumns.StudentId, id)
            Dim items = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim item = items.FirstOrDefault(Function(row) String.Equals(row.Fields.GetString(SharePointColumns.StudentId), id, StringComparison.OrdinalIgnoreCase))
            If item Is Nothing Then Return Nothing
            Return Map(item)
        End Function

        Public Async Function UpsertAsync(student As StudentRecord) As Task(Of UpsertKind) Implements IStudentDirectory.UpsertAsync
            Dim id = If(student?.StudentId, "").Trim()
            If id.Length = 0 Then Throw New ArgumentException("A student ID is required.")
            Dim filter = ODataFilters.TextEquals(SharePointColumns.StudentId, id)
            Dim existing = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim text = StudentFields(student)
            Dim match = existing.FirstOrDefault()
            If match Is Nothing Then
                Await _client.CreateItemAsync(_listName, text, Nothing).ConfigureAwait(False)
                Return UpsertKind.Inserted
            End If
            Await _client.UpdateItemAsync(_listName, match.Id, text, Nothing).ConfigureAwait(False)
            Return UpsertKind.Updated
        End Function

        Private Shared Function Map(item As ListItemRecord) As StudentRecord
            Return New StudentRecord With {
                .StudentId = item.Fields.GetString(SharePointColumns.StudentId),
                .GivenName = item.Fields.GetString(SharePointColumns.StudentGivenName),
                .FamilyName = item.Fields.GetString(SharePointColumns.StudentFamilyName),
                .PersonalEmail = item.Fields.GetString(SharePointColumns.StudentPersonalEmail),
                .BlockGroupCode = item.Fields.GetString(SharePointColumns.BlockGroupCode)
            }
        End Function

        Private Shared Function StudentFields(student As StudentRecord) As Dictionary(Of String, String)
            Return New Dictionary(Of String, String) From {
                {"Title", student.StudentId},
                {SharePointColumns.StudentId, student.StudentId},
                {SharePointColumns.StudentGivenName, If(student.GivenName, "")},
                {SharePointColumns.StudentFamilyName, If(student.FamilyName, "")},
                {SharePointColumns.StudentPersonalEmail, If(student.PersonalEmail, "")},
                {SharePointColumns.BlockGroupCode, If(student.BlockGroupCode, "")}
            }
        End Function
    End Class

    Public Class SharePointTeacherDirectory
        Implements ITeacherDirectory

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _listName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _listName = If(names, SharePointListNames.Default()).Teachers
        End Sub

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of TeacherRecord)) Implements ITeacherDirectory.GetAllAsync
            Dim items = Await _client.QueryItemsAsync(_listName, Nothing, False).ConfigureAwait(False)
            Return items.Select(Function(item) New TeacherRecord With {
                .FullName = item.Fields.GetString(SharePointColumns.TeacherFullName),
                .Email = item.Fields.GetString(SharePointColumns.Email)
            }).OrderBy(Function(item) item.FullName, StringComparer.OrdinalIgnoreCase).ToList()
        End Function

        Public Async Function FindEmailAsync(fullName As String) As Task(Of String) Implements ITeacherDirectory.FindEmailAsync
            Dim teachers = Await GetAllAsync().ConfigureAwait(False)
            Return TeacherEmailResolver.EmailFor(teachers, fullName)
        End Function

        Public Async Function UpsertAsync(teacher As TeacherRecord) As Task(Of UpsertKind) Implements ITeacherDirectory.UpsertAsync
            Dim name = If(teacher?.FullName, "").Trim()
            If name.Length = 0 Then Throw New ArgumentException("A teacher name is required.")
            Dim filter = ODataFilters.TextEquals(SharePointColumns.TeacherFullName, name)
            Dim existing = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim text As New Dictionary(Of String, String) From {
                {"Title", name},
                {SharePointColumns.TeacherFullName, name},
                {SharePointColumns.Email, If(teacher.Email, "")}
            }
            Dim match = existing.FirstOrDefault(Function(item) String.Equals(item.Fields.GetString(SharePointColumns.TeacherFullName), name, StringComparison.OrdinalIgnoreCase))
            If match Is Nothing Then
                Await _client.CreateItemAsync(_listName, text, Nothing).ConfigureAwait(False)
                Return UpsertKind.Inserted
            End If
            Await _client.UpdateItemAsync(_listName, match.Id, text, Nothing).ConfigureAwait(False)
            Return UpsertKind.Updated
        End Function
    End Class

    Public Class SharePointUnitDirectory
        Implements IUnitDirectory

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _listName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _listName = If(names, SharePointListNames.Default()).Units
        End Sub

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of UnitRecord)) Implements IUnitDirectory.GetAllAsync
            Dim items = Await _client.QueryItemsAsync(_listName, Nothing, False).ConfigureAwait(False)
            Return items.Select(Function(item) New UnitRecord With {
                .UnitCode = item.Fields.GetString(SharePointColumns.UnitCode),
                .UnitTitle = item.Fields.GetString(SharePointColumns.UnitTitle),
                .QualificationCode = item.Fields.GetString(SharePointColumns.QualificationCode)
            }).ToList()
        End Function

        Public Async Function FindTitleAsync(unitCode As String) As Task(Of String) Implements IUnitDirectory.FindTitleAsync
            Dim units = Await GetAllAsync().ConfigureAwait(False)
            For Each unit In units
                If String.Equals(unit.UnitCode, If(unitCode, "").Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Return If(unit.UnitTitle, "")
                End If
            Next
            Return ""
        End Function

        Public Async Function UpsertAsync(unit As UnitRecord) As Task(Of UpsertKind) Implements IUnitDirectory.UpsertAsync
            Dim code = If(unit?.UnitCode, "").Trim()
            If code.Length = 0 Then Throw New ArgumentException("A unit code is required.")
            Dim package = If(unit.QualificationCode, "").Trim()
            Dim filter = ODataFilters.TextEquals(SharePointColumns.UnitCode, code)
            Dim existing = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim text As New Dictionary(Of String, String) From {
                {"Title", code},
                {SharePointColumns.UnitCode, code},
                {SharePointColumns.UnitTitle, If(unit.UnitTitle, "")},
                {SharePointColumns.QualificationCode, package}
            }
            Dim match = existing.FirstOrDefault(Function(item) String.Equals(If(item.Fields, New FieldBag()).GetString(SharePointColumns.QualificationCode), package, StringComparison.OrdinalIgnoreCase))
            If match Is Nothing Then
                Await _client.CreateItemAsync(_listName, text, Nothing).ConfigureAwait(False)
                Return UpsertKind.Inserted
            End If
            Await _client.UpdateItemAsync(_listName, match.Id, text, Nothing).ConfigureAwait(False)
            Return UpsertKind.Updated
        End Function
    End Class

    Public Class SharePointSignatureStore
        Implements ISignatureImageStore

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _libraryName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _libraryName = If(names, SharePointListNames.Default()).SignatureLibrary
        End Sub

        Public Function GetCurrentImageAsync() As Task(Of Byte()) Implements ISignatureImageStore.GetCurrentImageAsync
            Return _client.DownloadSignatureAsync(_libraryName)
        End Function
    End Class

    Public Class SharePointQualificationDirectory
        Implements IQualificationDirectory

        Private ReadOnly _client As ISharePointClient
        Private ReadOnly _listName As String

        Public Sub New(client As ISharePointClient, names As SharePointListNames)
            _client = client
            _listName = If(names, SharePointListNames.Default()).Qualifications
        End Sub

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of QualificationRecord)) Implements IQualificationDirectory.GetAllAsync
            Dim items = Await _client.QueryItemsAsync(_listName, Nothing, False).ConfigureAwait(False)
            Return items.Select(AddressOf ToRecord).Where(Function(item) item IsNot Nothing).ToList()
        End Function

        Public Async Function FindAsync(qualificationCode As String) As Task(Of QualificationRecord) Implements IQualificationDirectory.FindAsync
            Dim code = If(qualificationCode, "").Trim()
            If code.Length = 0 Then Return Nothing
            Dim filter = ODataFilters.TextEquals(SharePointColumns.QualificationCode, code)
            Dim items = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Return items.Select(AddressOf ToRecord).FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(item.QualificationCode, code, StringComparison.OrdinalIgnoreCase))
        End Function

        Public Async Function UpsertAsync(qualification As QualificationRecord) As Task(Of UpsertKind) Implements IQualificationDirectory.UpsertAsync
            Dim code = If(qualification?.QualificationCode, "").Trim()
            If code.Length = 0 Then Throw New ArgumentException("A qualification code is required.")
            Dim rules = AbbreviationRules.Format(qualification.Rules)
            If rules.Length > 255 Then Throw New InvalidOperationException("Abbreviation rules must be 255 characters or fewer.")
            Dim filter = ODataFilters.TextEquals(SharePointColumns.QualificationCode, code)
            Dim existing = Await _client.QueryItemsAsync(_listName, filter, True).ConfigureAwait(False)
            Dim text As New Dictionary(Of String, String) From {
                {"Title", code},
                {SharePointColumns.QualificationCode, code},
                {SharePointColumns.QualificationTitle, If(qualification.Title, "")},
                {SharePointColumns.Version, If(qualification.Version, "")},
                {SharePointColumns.Status, QualificationStatus.Normalise(qualification.Status)},
                {SharePointColumns.ClassGroupSuffix, If(qualification.ClassGroupSuffix, "")},
                {SharePointColumns.AbbreviationRules, rules}
            }
            Dim match = existing.FirstOrDefault(Function(item) String.Equals(If(item.Fields, New FieldBag()).GetString(SharePointColumns.QualificationCode), code, StringComparison.OrdinalIgnoreCase))
            If match Is Nothing Then
                Await _client.CreateItemAsync(_listName, text, Nothing).ConfigureAwait(False)
                Return UpsertKind.Inserted
            End If
            Await _client.UpdateItemAsync(_listName, match.Id, text, Nothing).ConfigureAwait(False)
            Return UpsertKind.Updated
        End Function

        Private Shared Function ToRecord(item As ListItemRecord) As QualificationRecord
            If item Is Nothing OrElse item.Fields Is Nothing Then Return Nothing
            Dim code = item.Fields.GetString(SharePointColumns.QualificationCode)
            If String.IsNullOrWhiteSpace(code) Then Return Nothing
            Return New QualificationRecord With {
                .QualificationCode = code,
                .Title = item.Fields.GetString(SharePointColumns.QualificationTitle),
                .Version = item.Fields.GetString(SharePointColumns.Version),
                .Status = QualificationStatus.Normalise(item.Fields.GetString(SharePointColumns.Status)),
                .ClassGroupSuffix = item.Fields.GetString(SharePointColumns.ClassGroupSuffix),
                .Rules = AbbreviationRules.Parse(item.Fields.GetString(SharePointColumns.AbbreviationRules))
            }
        End Function
    End Class

End Namespace
