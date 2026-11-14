Imports System.Data

Namespace ElectrotechnologyResit.Core

    Public Interface IBookingService
        Function GetAllAsync() As Task(Of IReadOnlyList(Of ResitBooking))
        Function GetByDateAsync(resitDate As Date) As Task(Of IReadOnlyList(Of ResitBooking))
        Function UpsertAsync(booking As ResitBooking, preserveEnergyspaceFlags As Boolean) As Task(Of UpsertResult)
        Function DeleteByStudentIdAsync(studentId As String) As Task
        Function DeleteByDateAsync(resitDate As Date) As Task
        Function DeleteOnOrBeforeAsync(cutoff As Date) As Task
        Function SetEnergyspaceCreatedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task
        Function SetEnergyspaceAssessmentBookedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task
        Function GetUnitFlagsAsync(resitDate As Date) As Task(Of IReadOnlyList(Of UnitNightFlags))
    End Interface

    Public Interface IStudentDirectory
        Function FindByStudentIdAsync(studentId As String) As Task(Of StudentRecord)
        Function UpsertAsync(student As StudentRecord) As Task(Of UpsertKind)
    End Interface

    Public Interface ITeacherDirectory
        Function GetAllAsync() As Task(Of IReadOnlyList(Of TeacherRecord))
        Function FindEmailAsync(fullName As String) As Task(Of String)
        Function UpsertAsync(teacher As TeacherRecord) As Task(Of UpsertKind)
    End Interface

    Public Interface IUnitDirectory
        Function GetAllAsync() As Task(Of IReadOnlyList(Of UnitRecord))
        Function FindTitleAsync(unitCode As String) As Task(Of String)
        Function UpsertAsync(unit As UnitRecord) As Task(Of UpsertKind)
    End Interface

    Public Interface ISignatureImageStore
        Function GetCurrentImageAsync() As Task(Of Byte())
    End Interface

    ''' <summary>
    ''' Optional Microsoft Bookings import. Live Bookings is owned by VU Timetable and needs Frank Offer's sign-off.
    ''' </summary>
    Public Interface IBookingImporter
        Function PullAsync(resitDate As Date) As Task(Of String)
    End Interface

    Public NotInheritable Class BookingTable
        Private Sub New()
        End Sub

        Public Shared ReadOnly Headers As String() = {
            "Student ID",
            "Unit",
            "Resit date",
            "Blockgroup",
            "AllocatedTeacher",
            "status",
            "AllocatedTeacher email",
            "EnergyspaceClassCreated",
            "EnergyspaceAssessmentBooked",
            "AttemptNo",
            "Student Firstname",
            "Surname",
            "Email",
            "Unit Name"
        }

        Public Shared Function FromBookings(rows As IEnumerable(Of ResitBooking)) As DataTable
            Dim table As New DataTable()
            table.Columns.Add("Student ID", GetType(String))
            table.Columns.Add("Unit", GetType(String))
            table.Columns.Add("Resit date", GetType(Date))
            table.Columns.Add("Blockgroup", GetType(String))
            table.Columns.Add("AllocatedTeacher", GetType(String))
            table.Columns.Add("status", GetType(String))
            table.Columns.Add("AllocatedTeacher email", GetType(String))
            table.Columns.Add("EnergyspaceClassCreated", GetType(Boolean))
            table.Columns.Add("EnergyspaceAssessmentBooked", GetType(Boolean))
            table.Columns.Add("AttemptNo", GetType(String))
            table.Columns.Add("Student Firstname", GetType(String))
            table.Columns.Add("Surname", GetType(String))
            table.Columns.Add("Email", GetType(String))
            table.Columns.Add("Unit Name", GetType(String))
            If rows Is Nothing Then Return table
            For Each booking In rows
                table.Rows.Add(
                    booking.StudentId,
                    booking.UnitCode,
                    booking.ResitDate.Date,
                    booking.BlockGroup,
                    booking.TeacherName,
                    If(booking.Status, ""),
                    booking.TeacherEmail,
                    booking.EnergyspaceCreated,
                    booking.EnergyspaceAssessmentBooked,
                    booking.AttemptNo,
                    booking.FirstName,
                    booking.Surname,
                    booking.Email,
                    booking.UnitName)
            Next
            Return table
        End Function
    End Class

    Public Class InMemoryResitStore
        Implements IBookingService
        Implements IStudentDirectory
        Implements ITeacherDirectory
        Implements IUnitDirectory
        Implements ISignatureImageStore

        Private ReadOnly _bookings As New Dictionary(Of String, ResitBooking)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _students As New Dictionary(Of String, StudentRecord)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _teachers As New Dictionary(Of String, TeacherRecord)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _units As New Dictionary(Of String, UnitRecord)(StringComparer.OrdinalIgnoreCase)
        Public Property SignatureImage As Byte()
        Public Property StudentLookupCount As Integer

        Public Function GetBookingsAsync() As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetAllAsync
            Return Task.FromResult(Of IReadOnlyList(Of ResitBooking))(ClonedBookings(_bookings.Values))
        End Function

        Public Function GetByDateAsync(resitDate As Date) As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetByDateAsync
            Dim rows = _bookings.Values.Where(Function(item) item.ResitDate.Date = resitDate.Date)
            Return Task.FromResult(Of IReadOnlyList(Of ResitBooking))(ClonedBookings(rows))
        End Function

        Public Function UpsertAsync(booking As ResitBooking, preserveEnergyspaceFlags As Boolean) As Task(Of UpsertResult) Implements IBookingService.UpsertAsync
            Dim id = If(booking?.StudentId, "").Trim()
            If id.Length = 0 Then Throw New ArgumentException("A student ID is required.")
            Dim existing As ResitBooking = Nothing
            _bookings.TryGetValue(id, existing)
            Dim merged = BookingUpsert.Merge(existing, booking, preserveEnergyspaceFlags, _teachers.Values)
            If String.IsNullOrEmpty(merged.SharePointItemId) Then merged.SharePointItemId = Guid.NewGuid().ToString("N")
            _bookings(id) = merged
            Return Task.FromResult(New UpsertResult With {
                .Kind = If(existing Is Nothing, UpsertKind.Inserted, UpsertKind.Updated),
                .Booking = merged.Clone()
            })
        End Function

        Public Function DeleteByStudentIdAsync(studentId As String) As Task Implements IBookingService.DeleteByStudentIdAsync
            _bookings.Remove(If(studentId, "").Trim())
            Return Task.CompletedTask
        End Function

        Public Function DeleteByDateAsync(resitDate As Date) As Task Implements IBookingService.DeleteByDateAsync
            RemoveWhere(Function(item) item.ResitDate.Date = resitDate.Date)
            Return Task.CompletedTask
        End Function

        Public Function DeleteOnOrBeforeAsync(cutoff As Date) As Task Implements IBookingService.DeleteOnOrBeforeAsync
            RemoveWhere(Function(item) item.ResitDate.Date <= cutoff.Date)
            Return Task.CompletedTask
        End Function

        Public Function SetEnergyspaceCreatedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceCreatedAsync
            For Each booking In Matching(resitDate, unitCode)
                booking.EnergyspaceCreated = value
            Next
            Return Task.CompletedTask
        End Function

        Public Function SetEnergyspaceAssessmentBookedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceAssessmentBookedAsync
            For Each booking In Matching(resitDate, unitCode)
                booking.EnergyspaceAssessmentBooked = value
            Next
            Return Task.CompletedTask
        End Function

        Public Function GetUnitFlagsAsync(resitDate As Date) As Task(Of IReadOnlyList(Of UnitNightFlags)) Implements IBookingService.GetUnitFlagsAsync
            Return Task.FromResult(EnergyspaceFlagRules.Summarise(_bookings.Values, resitDate))
        End Function

        Public Function FindByStudentIdAsync(studentId As String) As Task(Of StudentRecord) Implements IStudentDirectory.FindByStudentIdAsync
            StudentLookupCount += 1
            Dim found As StudentRecord = Nothing
            If _students.TryGetValue(If(studentId, "").Trim(), found) Then
                Return Task.FromResult(found.Clone())
            End If
            Return Task.FromResult(Of StudentRecord)(Nothing)
        End Function

        Public Function UpsertAsync(student As StudentRecord) As Task(Of UpsertKind) Implements IStudentDirectory.UpsertAsync
            Dim id = If(student?.StudentId, "").Trim()
            If id.Length = 0 Then Throw New ArgumentException("A student ID is required.")
            Dim kind = If(_students.ContainsKey(id), UpsertKind.Updated, UpsertKind.Inserted)
            _students(id) = student.Clone()
            _students(id).StudentId = id
            Return Task.FromResult(kind)
        End Function

        Public Function GetTeachersAsync() As Task(Of IReadOnlyList(Of TeacherRecord)) Implements ITeacherDirectory.GetAllAsync
            Dim rows = _teachers.Values.Select(Function(item) item.Clone()).OrderBy(Function(item) item.FullName, StringComparer.OrdinalIgnoreCase).ToList()
            Return Task.FromResult(Of IReadOnlyList(Of TeacherRecord))(rows)
        End Function

        Public Function FindEmailAsync(fullName As String) As Task(Of String) Implements ITeacherDirectory.FindEmailAsync
            Return Task.FromResult(TeacherEmailResolver.EmailFor(_teachers.Values, fullName))
        End Function

        Public Function UpsertAsync(teacher As TeacherRecord) As Task(Of UpsertKind) Implements ITeacherDirectory.UpsertAsync
            Dim name = If(teacher?.FullName, "").Trim()
            If name.Length = 0 Then Throw New ArgumentException("A teacher name is required.")
            Dim kind = If(_teachers.ContainsKey(name), UpsertKind.Updated, UpsertKind.Inserted)
            _teachers(name) = New TeacherRecord With {.FullName = name, .Email = If(teacher.Email, "").Trim()}
            Return Task.FromResult(kind)
        End Function

        Public Function GetUnitsAsync() As Task(Of IReadOnlyList(Of UnitRecord)) Implements IUnitDirectory.GetAllAsync
            Dim rows = _units.Values.Select(Function(item) item.Clone()).ToList()
            Return Task.FromResult(Of IReadOnlyList(Of UnitRecord))(rows)
        End Function

        Public Function FindTitleAsync(unitCode As String) As Task(Of String) Implements IUnitDirectory.FindTitleAsync
            Dim found As UnitRecord = Nothing
            If _units.TryGetValue(If(unitCode, "").Trim(), found) Then Return Task.FromResult(If(found.UnitTitle, ""))
            Return Task.FromResult("")
        End Function

        Public Function UpsertAsync(unit As UnitRecord) As Task(Of UpsertKind) Implements IUnitDirectory.UpsertAsync
            Dim code = If(unit?.UnitCode, "").Trim()
            If code.Length = 0 Then Throw New ArgumentException("A unit code is required.")
            Dim kind = If(_units.ContainsKey(code), UpsertKind.Updated, UpsertKind.Inserted)
            _units(code) = New UnitRecord With {.UnitCode = code, .UnitTitle = If(unit.UnitTitle, "").Trim()}
            Return Task.FromResult(kind)
        End Function

        Public Function GetCurrentImageAsync() As Task(Of Byte()) Implements ISignatureImageStore.GetCurrentImageAsync
            Return Task.FromResult(SignatureImage)
        End Function

        Private Function Matching(resitDate As Date, unitCode As String) As IEnumerable(Of ResitBooking)
            Return _bookings.Values.Where(Function(item) item.ResitDate.Date = resitDate.Date AndAlso String.Equals(item.UnitCode, unitCode, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Sub RemoveWhere(predicate As Func(Of ResitBooking, Boolean))
            For Each id In _bookings.Where(Function(pair) predicate(pair.Value)).Select(Function(pair) pair.Key).ToList()
                _bookings.Remove(id)
            Next
        End Sub

        Private Shared Function ClonedBookings(rows As IEnumerable(Of ResitBooking)) As IReadOnlyList(Of ResitBooking)
            Return rows.Select(Function(item) item.Clone()).ToList()
        End Function
    End Class

    Public Class CachingStudentDirectory
        Implements IStudentDirectory

        Private ReadOnly _inner As IStudentDirectory
        Private ReadOnly _cache As New Dictionary(Of String, StudentRecord)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _missing As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Public Sub New(inner As IStudentDirectory)
            _inner = inner
        End Sub

        Public Async Function FindByStudentIdAsync(studentId As String) As Task(Of StudentRecord) Implements IStudentDirectory.FindByStudentIdAsync
            Dim id = If(studentId, "").Trim()
            If id.Length = 0 Then Return Nothing
            Dim cached As StudentRecord = Nothing
            If _cache.TryGetValue(id, cached) Then Return cached.Clone()
            If _missing.Contains(id) Then Return Nothing
            Dim found = Await _inner.FindByStudentIdAsync(id).ConfigureAwait(False)
            If found Is Nothing Then
                _missing.Add(id)
                Return Nothing
            End If
            _cache(id) = found
            Return found.Clone()
        End Function

        Public Async Function UpsertAsync(student As StudentRecord) As Task(Of UpsertKind) Implements IStudentDirectory.UpsertAsync
            Dim kind = Await _inner.UpsertAsync(student).ConfigureAwait(False)
            Dim id = If(student?.StudentId, "").Trim()
            If id.Length > 0 Then
                _missing.Remove(id)
                _cache(id) = student.Clone()
            End If
            Return kind
        End Function
    End Class

    Public Class CachingTeacherDirectory
        Implements ITeacherDirectory

        Private ReadOnly _inner As ITeacherDirectory
        Private _all As IReadOnlyList(Of TeacherRecord)

        Public Sub New(inner As ITeacherDirectory)
            _inner = inner
        End Sub

        Private _loadCount As Integer

        Public ReadOnly Property LoadCount As Integer
            Get
                Return _loadCount
            End Get
        End Property

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of TeacherRecord)) Implements ITeacherDirectory.GetAllAsync
            If _all Is Nothing Then
                _loadCount += 1
                _all = Await _inner.GetAllAsync().ConfigureAwait(False)
            End If
            Return _all.Select(Function(item) item.Clone()).ToList()
        End Function

        Public Async Function FindEmailAsync(fullName As String) As Task(Of String) Implements ITeacherDirectory.FindEmailAsync
            Dim teachers = Await GetAllAsync().ConfigureAwait(False)
            Return TeacherEmailResolver.EmailFor(teachers, fullName)
        End Function

        Public Async Function UpsertAsync(teacher As TeacherRecord) As Task(Of UpsertKind) Implements ITeacherDirectory.UpsertAsync
            Dim kind = Await _inner.UpsertAsync(teacher).ConfigureAwait(False)
            _all = Nothing
            Return kind
        End Function
    End Class

    Public Class CachingUnitDirectory
        Implements IUnitDirectory

        Private ReadOnly _inner As IUnitDirectory
        Private _all As IReadOnlyList(Of UnitRecord)

        Public Sub New(inner As IUnitDirectory)
            _inner = inner
        End Sub

        Private _loadCount As Integer

        Public ReadOnly Property LoadCount As Integer
            Get
                Return _loadCount
            End Get
        End Property

        Public Async Function GetAllAsync() As Task(Of IReadOnlyList(Of UnitRecord)) Implements IUnitDirectory.GetAllAsync
            If _all Is Nothing Then
                _loadCount += 1
                _all = Await _inner.GetAllAsync().ConfigureAwait(False)
            End If
            Return _all.Select(Function(item) item.Clone()).ToList()
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
            Dim kind = Await _inner.UpsertAsync(unit).ConfigureAwait(False)
            _all = Nothing
            Return kind
        End Function
    End Class

End Namespace
