Imports System.IO
Imports System.Text.Json

Namespace ElectrotechnologyResit.Core

    ''' <summary>Fictional seed for the local sample store. No real names or university addresses.</summary>
    Public NotInheritable Class SampleDataCatalog
        Private Sub New()
        End Sub

        Public Shared Function Build(today As Date) As SampleSnapshot
            Dim night = today.Date
            Dim recent = night.AddDays(-2)
            Dim old = night.AddDays(-16)
            Dim teachers = New List(Of TeacherRecord) From {
                Teacher("Jordan Hale", "jordan.hale@example.com"),
                Teacher("Morgan Ellis", "morgan.ellis@example.com"),
                Teacher("Riley Chen", "riley.chen@example.com"),
                Teacher("Avery Quinn", "avery.quinn@example.com")
            }
            Dim units = New List(Of UnitRecord) From {
                Unit("UEE30820", "UEECD0019", "Fabricate, assemble and dismantle utilities industry components"),
                Unit("UEE30820", "UEECD0020", "Fix and secure electrotechnology equipment"),
                Unit("UEE30820", "UEECD0046", "Solve problems in single path circuits"),
                Unit("UEE30820", "UEEEL0005", "Develop and connect electrical control circuits"),
                Unit("UEE30820", "UEEEL0012", "Install low voltage wiring, appliances, switchgear and associated accessories"),
                Unit("UEE30820", "UEEEL0021", "Solve problems in magnetic and electromagnetic devices"),
                Unit("UEE30820", "UEEEL0025", "Test and connect transformers")
            }
            Dim qualifications = New List(Of QualificationRecord) From {
                New QualificationRecord With {
                    .QualificationCode = "UEE30820",
                    .Title = "Certificate III in Electrotechnology Electrician",
                    .Version = "1",
                    .Status = QualificationStatus.Active,
                    .ClassGroupSuffix = "CERTIII",
                    .Rules = AbbreviationRules.Uee30820Rules()
                }
            }
            Dim people = New List(Of StudentRecord) From {
                Student("9000001", "Alex", "Nguyen", "BG-A"),
                Student("9000002", "Blair", "Okafor", "BG-B"),
                Student("9000003", "Casey", "Patel", "BG-A"),
                Student("9000004", "Drew", "Singh", "BG-C"),
                Student("9000005", "Eden", "Walsh", "BG-B"),
                Student("9000006", "Frankie", "Rossi", "BG-A"),
                Student("9000007", "Grey", "Ahmed", "BG-C"),
                Student("9000008", "Harper", "Cole", "BG-B"),
                Student("9000009", "Indigo", "Brooks", "BG-A"),
                Student("9000010", "Jules", "Marin", "BG-C"),
                Student("9000011", "Kai", "Okonkwo", "BG-B"),
                Student("9000012", "Lane", "Fischer", "BG-A"),
                Student("9000013", "Micah", "Adeyemi", "BG-C"),
                Student("9000014", "Noor", "Rahman", "BG-B")
            }
            Dim bookings = New List(Of ResitBooking) From {
                Booking(people(0), units(0), teachers(0), night, "1", True, True),
                Booking(people(1), units(1), teachers(0), night, "2", True, False),
                Booking(people(2), units(2), teachers(1), night, "3", False, False),
                Booking(people(3), units(3), teachers(1), night, "1", True, True),
                Booking(people(4), units(4), teachers(2), night, "4", True, False),
                Booking(people(5), units(5), teachers(3), night, "2", False, True),
                Booking(people(6), units(6), teachers(2), recent, "1", True, True),
                Booking(people(7), units(0), teachers(2), recent, "6", False, False),
                Booking(people(8), units(1), teachers(0), recent, "1", True, True),
                Booking(people(9), units(2), teachers(3), recent, "3", False, True),
                Booking(people(10), units(3), teachers(1), recent, "2", True, False),
                Booking(people(11), units(4), teachers(0), old, "5", False, False),
                Booking(people(12), units(5), teachers(2), old, "1", True, True),
                Booking(people(13), units(6), teachers(3), old, "2", True, False)
            }
            Return New SampleSnapshot With {
                .Bookings = bookings,
                .Students = people,
                .Teachers = teachers,
                .Units = units,
                .Qualifications = qualifications
            }
        End Function

        Private Shared Function Teacher(name As String, email As String) As TeacherRecord
            Return New TeacherRecord With {.FullName = name, .Email = email}
        End Function

        Private Shared Function Unit(qualificationCode As String, code As String, title As String) As UnitRecord
            Return New UnitRecord With {.QualificationCode = qualificationCode, .UnitCode = code, .UnitTitle = title}
        End Function

        Private Shared Function Student(id As String, given As String, family As String, block As String) As StudentRecord
            Return New StudentRecord With {
                .StudentId = id,
                .GivenName = given,
                .FamilyName = family,
                .PersonalEmail = given.ToLowerInvariant() & "." & family.ToLowerInvariant() & "@example.com",
                .BlockGroupCode = block
            }
        End Function

        Private Shared Function Booking(person As StudentRecord, unit As UnitRecord, teacher As TeacherRecord, night As Date, attempt As String, created As Boolean, booked As Boolean) As ResitBooking
            Return New ResitBooking With {
                .StudentId = person.StudentId,
                .FirstName = person.GivenName,
                .Surname = person.FamilyName,
                .Email = person.PersonalEmail,
                .UnitCode = unit.UnitCode,
                .UnitName = unit.UnitTitle,
                .QualificationCode = unit.QualificationCode,
                .BlockGroup = person.BlockGroupCode,
                .TeacherName = teacher.FullName,
                .TeacherEmail = teacher.Email,
                .AttemptNo = attempt,
                .ResitDate = night.Date,
                .EnergyspaceCreated = created,
                .EnergyspaceAssessmentBooked = booked,
                .Status = ""
            }
        End Function
    End Class

    Public Class SampleSnapshot
        Public Property Bookings As List(Of ResitBooking)
        Public Property Students As List(Of StudentRecord)
        Public Property Teachers As List(Of TeacherRecord)
        Public Property Units As List(Of UnitRecord)
        Public Property Qualifications As List(Of QualificationRecord)
    End Class

    ''' <summary>Local JSON store used when the data source is sample data. Not SharePoint.</summary>
    Public Class SampleDataStore
        Implements IBookingService
        Implements IStudentDirectory
        Implements ITeacherDirectory
        Implements IUnitDirectory
        Implements IQualificationDirectory
        Implements ISignatureImageStore

        Private ReadOnly _memory As New InMemoryResitStore()
        Private ReadOnly _folder As String
        Private _today As Date
        Private Shared ReadOnly JsonOptions As New JsonSerializerOptions With {.WriteIndented = True, .PropertyNameCaseInsensitive = True}

        ' 1x1 JPEG placeholder so the teacher-email image path can be tried. Not a person's signature.
        Private Shared ReadOnly PlaceholderJpeg As Byte() = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=")

        Public Sub New(folder As String, today As Date)
            _folder = folder
            _today = today.Date
            Directory.CreateDirectory(_folder)
            If File.Exists(BookingsPath()) Then
                Load()
            Else
                Seed()
            End If
            If Not File.Exists(QualificationsPath()) Then
                QualificationCatalog.EnsureSeedAsync(Me, Me).GetAwaiter().GetResult()
            End If
        End Sub

        Public Shared Function DefaultFolder() As String
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ElectrotechnologyResit", "SampleData")
        End Function

        Public Sub Reset(today As Date)
            _today = today.Date
            For Each name In New String() {"bookings.json", "students.json", "teachers.json", "units.json", "qualifications.json", "signature.jpg"}
                Dim filePath = IO.Path.Combine(_folder, name)
                If File.Exists(filePath) Then File.Delete(filePath)
            Next
            _memory.ClearAll()
            Seed()
        End Sub

        Public Function GetBookingsAsync() As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetAllAsync
            Return _memory.GetBookingsAsync()
        End Function

        Public Function GetByDateAsync(resitDate As Date) As Task(Of IReadOnlyList(Of ResitBooking)) Implements IBookingService.GetByDateAsync
            Return _memory.GetByDateAsync(resitDate)
        End Function

        Public Async Function UpsertAsync(booking As ResitBooking, preserveEnergyspaceFlags As Boolean) As Task(Of UpsertResult) Implements IBookingService.UpsertAsync
            Dim result = Await _memory.UpsertAsync(booking, preserveEnergyspaceFlags).ConfigureAwait(False)
            SaveBookings()
            Return result
        End Function

        Public Function DeleteByStudentIdAsync(studentId As String) As Task Implements IBookingService.DeleteByStudentIdAsync
            Dim task = _memory.DeleteByStudentIdAsync(studentId)
            SaveBookings()
            Return task
        End Function

        Public Function DeleteByDateAsync(resitDate As Date) As Task Implements IBookingService.DeleteByDateAsync
            Dim task = _memory.DeleteByDateAsync(resitDate)
            SaveBookings()
            Return task
        End Function

        Public Async Function DeleteOnOrBeforeAsync(cutoff As Date) As Task(Of Integer) Implements IBookingService.DeleteOnOrBeforeAsync
            Dim removed = Await _memory.DeleteOnOrBeforeAsync(cutoff).ConfigureAwait(False)
            SaveBookings()
            Return removed
        End Function

        Public Function SetEnergyspaceCreatedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceCreatedAsync
            Dim task = _memory.SetEnergyspaceCreatedAsync(resitDate, unitCode, value)
            SaveBookings()
            Return task
        End Function

        Public Function SetEnergyspaceAssessmentBookedAsync(resitDate As Date, unitCode As String, value As Boolean) As Task Implements IBookingService.SetEnergyspaceAssessmentBookedAsync
            Dim task = _memory.SetEnergyspaceAssessmentBookedAsync(resitDate, unitCode, value)
            SaveBookings()
            Return task
        End Function

        Public Function GetUnitFlagsAsync(resitDate As Date) As Task(Of IReadOnlyList(Of UnitNightFlags)) Implements IBookingService.GetUnitFlagsAsync
            Return _memory.GetUnitFlagsAsync(resitDate)
        End Function

        Public Function FindByStudentIdAsync(studentId As String) As Task(Of StudentRecord) Implements IStudentDirectory.FindByStudentIdAsync
            Return _memory.FindByStudentIdAsync(studentId)
        End Function

        Public Async Function UpsertAsync(student As StudentRecord) As Task(Of UpsertKind) Implements IStudentDirectory.UpsertAsync
            Dim kind = Await _memory.UpsertAsync(student).ConfigureAwait(False)
            SaveStudents()
            Return kind
        End Function

        Public Function GetTeachersAsync() As Task(Of IReadOnlyList(Of TeacherRecord)) Implements ITeacherDirectory.GetAllAsync
            Return _memory.GetTeachersAsync()
        End Function

        Public Function FindEmailAsync(fullName As String) As Task(Of String) Implements ITeacherDirectory.FindEmailAsync
            Return _memory.FindEmailAsync(fullName)
        End Function

        Public Async Function UpsertAsync(teacher As TeacherRecord) As Task(Of UpsertKind) Implements ITeacherDirectory.UpsertAsync
            Dim kind = Await _memory.UpsertAsync(teacher).ConfigureAwait(False)
            SaveTeachers()
            Return kind
        End Function

        Public Function GetUnitsAsync() As Task(Of IReadOnlyList(Of UnitRecord)) Implements IUnitDirectory.GetAllAsync
            Return _memory.GetUnitsAsync()
        End Function

        Public Function FindTitleAsync(unitCode As String) As Task(Of String) Implements IUnitDirectory.FindTitleAsync
            Return _memory.FindTitleAsync(unitCode)
        End Function

        Public Async Function UpsertAsync(unit As UnitRecord) As Task(Of UpsertKind) Implements IUnitDirectory.UpsertAsync
            Dim kind = Await _memory.UpsertAsync(unit).ConfigureAwait(False)
            SaveUnits()
            Return kind
        End Function

        Public Function GetAllAsync() As Task(Of IReadOnlyList(Of QualificationRecord)) Implements IQualificationDirectory.GetAllAsync
            Return _memory.GetQualificationsAsync()
        End Function

        Public Function FindAsync(qualificationCode As String) As Task(Of QualificationRecord) Implements IQualificationDirectory.FindAsync
            Return _memory.FindQualificationAsync(qualificationCode)
        End Function

        Public Async Function UpsertAsync(qualification As QualificationRecord) As Task(Of UpsertKind) Implements IQualificationDirectory.UpsertAsync
            Dim kind = Await _memory.UpsertAsync(qualification).ConfigureAwait(False)
            SaveQualifications()
            Return kind
        End Function

        Public Function GetCurrentImageAsync() As Task(Of Byte()) Implements ISignatureImageStore.GetCurrentImageAsync
            Return _memory.GetCurrentImageAsync()
        End Function

        Private Sub Seed()
            Dim snapshot = SampleDataCatalog.Build(_today)
            For Each qualification In snapshot.Qualifications
                _memory.UpsertAsync(qualification).GetAwaiter().GetResult()
            Next
            For Each teacher In snapshot.Teachers
                _memory.UpsertAsync(teacher).GetAwaiter().GetResult()
            Next
            For Each unit In snapshot.Units
                _memory.UpsertAsync(unit).GetAwaiter().GetResult()
            Next
            For Each student In snapshot.Students
                _memory.UpsertAsync(student).GetAwaiter().GetResult()
            Next
            For Each booking In snapshot.Bookings
                _memory.UpsertAsync(booking, False).GetAwaiter().GetResult()
            Next
            _memory.SignatureImage = PlaceholderJpeg
            SaveAll()
        End Sub

        Private Sub Load()
            Dim qualifications = ReadList(Of QualificationRecord)(QualificationsPath())
            Dim teachers = ReadList(Of TeacherRecord)(TeachersPath())
            Dim units = ReadList(Of UnitRecord)(UnitsPath())
            Dim students = ReadList(Of StudentRecord)(StudentsPath())
            Dim bookings = ReadList(Of ResitBooking)(BookingsPath())
            For Each qualification In qualifications
                _memory.UpsertAsync(qualification).GetAwaiter().GetResult()
            Next
            For Each teacher In teachers
                _memory.UpsertAsync(teacher).GetAwaiter().GetResult()
            Next
            For Each unit In units
                _memory.UpsertAsync(unit).GetAwaiter().GetResult()
            Next
            For Each student In students
                _memory.UpsertAsync(student).GetAwaiter().GetResult()
            Next
            For Each booking In bookings
                _memory.UpsertAsync(booking, False).GetAwaiter().GetResult()
            Next
            Dim imagePath = Path.Combine(_folder, "signature.jpg")
            If File.Exists(imagePath) Then _memory.SignatureImage = File.ReadAllBytes(imagePath)
        End Sub

        Private Sub SaveAll()
            SaveQualifications()
            SaveTeachers()
            SaveUnits()
            SaveStudents()
            SaveBookings()
            File.WriteAllBytes(Path.Combine(_folder, "signature.jpg"), If(_memory.SignatureImage, PlaceholderJpeg))
        End Sub

        Private Sub SaveBookings()
            WriteList(BookingsPath(), _memory.GetBookingsAsync().GetAwaiter().GetResult())
        End Sub

        Private Sub SaveStudents()
            WriteList(StudentsPath(), _memory.ExportStudents())
        End Sub

        Private Sub SaveTeachers()
            WriteList(TeachersPath(), _memory.GetTeachersAsync().GetAwaiter().GetResult())
        End Sub

        Private Sub SaveUnits()
            WriteList(UnitsPath(), _memory.GetUnitsAsync().GetAwaiter().GetResult())
        End Sub

        Private Sub SaveQualifications()
            WriteList(QualificationsPath(), _memory.GetQualificationsAsync().GetAwaiter().GetResult())
        End Sub

        Private Function BookingsPath() As String
            Return Path.Combine(_folder, "bookings.json")
        End Function

        Private Function StudentsPath() As String
            Return Path.Combine(_folder, "students.json")
        End Function

        Private Function TeachersPath() As String
            Return Path.Combine(_folder, "teachers.json")
        End Function

        Private Function UnitsPath() As String
            Return Path.Combine(_folder, "units.json")
        End Function

        Private Function QualificationsPath() As String
            Return Path.Combine(_folder, "qualifications.json")
        End Function

        Private Shared Function ReadList(Of T)(path As String) As List(Of T)
            If Not File.Exists(path) Then Return New List(Of T)
            Dim json = File.ReadAllText(path)
            Dim rows = JsonSerializer.Deserialize(Of List(Of T))(json, JsonOptions)
            If rows Is Nothing Then Return New List(Of T)
            Return rows
        End Function

        Private Shared Sub WriteList(Of T)(path As String, rows As IEnumerable(Of T))
            File.WriteAllText(path, JsonSerializer.Serialize(rows, JsonOptions))
        End Sub
    End Class

End Namespace
