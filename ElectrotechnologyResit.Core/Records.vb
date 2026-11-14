Namespace ElectrotechnologyResit.Core

    Public Class ResitBooking
        Public Property SharePointItemId As String
        Public Property StudentId As String
        Public Property FirstName As String
        Public Property Surname As String
        Public Property Email As String
        Public Property UnitCode As String
        Public Property UnitName As String
        Public Property QualificationCode As String
        Public Property BlockGroup As String
        Public Property TeacherName As String
        Public Property TeacherEmail As String
        Public Property AttemptNo As String
        Public Property ResitDate As Date
        Public Property EnergyspaceCreated As Boolean
        Public Property EnergyspaceAssessmentBooked As Boolean
        Public Property Status As String

        Public Function Clone() As ResitBooking
            Return New ResitBooking With {
                .SharePointItemId = SharePointItemId,
                .StudentId = StudentId,
                .FirstName = FirstName,
                .Surname = Surname,
                .Email = Email,
                .UnitCode = UnitCode,
                .UnitName = UnitName,
                .QualificationCode = QualificationCode,
                .BlockGroup = BlockGroup,
                .TeacherName = TeacherName,
                .TeacherEmail = TeacherEmail,
                .AttemptNo = AttemptNo,
                .ResitDate = ResitDate,
                .EnergyspaceCreated = EnergyspaceCreated,
                .EnergyspaceAssessmentBooked = EnergyspaceAssessmentBooked,
                .Status = Status
            }
        End Function
    End Class

    Public Class StudentRecord
        Public Property StudentId As String
        Public Property GivenName As String
        Public Property FamilyName As String
        Public Property PersonalEmail As String
        Public Property BlockGroupCode As String

        Public Function Clone() As StudentRecord
            Return New StudentRecord With {
                .StudentId = StudentId,
                .GivenName = GivenName,
                .FamilyName = FamilyName,
                .PersonalEmail = PersonalEmail,
                .BlockGroupCode = BlockGroupCode
            }
        End Function
    End Class

    Public Class TeacherRecord
        Public Property FullName As String
        Public Property Email As String

        Public Function Clone() As TeacherRecord
            Return New TeacherRecord With {.FullName = FullName, .Email = Email}
        End Function
    End Class

    Public Class UnitRecord
        Public Property UnitCode As String
        Public Property UnitTitle As String
        Public Property QualificationCode As String

        Public Function Clone() As UnitRecord
            Return New UnitRecord With {.UnitCode = UnitCode, .UnitTitle = UnitTitle, .QualificationCode = QualificationCode}
        End Function
    End Class

    Public Enum UpsertKind
        Inserted = 1
        Updated = 2
    End Enum

    Public Class UpsertResult
        Public Property Kind As UpsertKind
        Public Property Booking As ResitBooking
    End Class

    ''' <summary>
    ''' Saved Energyspace checkpoints for one unit on one night.
    ''' A tick is shown only when every booking for that unit and date is true.
    ''' </summary>
    Public Class UnitNightFlags
        Public Property UnitCode As String
        Public Property ClassCreated As Boolean
        Public Property AssessmentBooked As Boolean
    End Class

    Public Class ImportTally
        Public Property Inserted As Integer
        Public Property Updated As Integer
        Public Property Skipped As Integer
    End Class

    Public Class SharePointListNames
        Public Property Bookings As String = "ResitBookings"
        Public Property Students As String = "Students"
        Public Property Teachers As String = "Teachers"
        Public Property Units As String = "Units"
        Public Property SignatureLibrary As String = "EmailSignature"
        Public Property Qualifications As String = "Qualifications"

        Public Shared Function FromSettings(bookings As String, students As String, teachers As String, units As String, signatureLibrary As String) As SharePointListNames
            Return New SharePointListNames With {
                .Bookings = OrDefault(bookings, "ResitBookings"),
                .Students = OrDefault(students, "Students"),
                .Teachers = OrDefault(teachers, "Teachers"),
                .Units = OrDefault(units, "Units"),
                .SignatureLibrary = OrDefault(signatureLibrary, "EmailSignature"),
                .Qualifications = "Qualifications"
            }
        End Function

        Public Shared Function FromSettings(bookings As String, students As String, teachers As String, units As String, signatureLibrary As String, qualifications As String) As SharePointListNames
            Dim names = FromSettings(bookings, students, teachers, units, signatureLibrary)
            names.Qualifications = OrDefault(qualifications, "Qualifications")
            Return names
        End Function

        Public Shared Function [Default]() As SharePointListNames
            Return FromSettings("", "", "", "", "")
        End Function

        Private Shared Function OrDefault(value As String, fallback As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return fallback
            Return value.Trim()
        End Function

        Public Function AllNames() As IEnumerable(Of String)
            Return New String() {Bookings, Students, Teachers, Units, Qualifications, SignatureLibrary}
        End Function
    End Class

    Public Enum ColumnKind
        Text
        DateOnly
        YesNo
    End Enum

    Public Enum ListTemplateKind
        GenericList
        DocumentLibrary
    End Enum

    Public Class ColumnSpec
        Public Property Name As String
        Public Property DisplayName As String
        Public Property Kind As ColumnKind
        Public Property Indexed As Boolean
        Public Property EnforceUniqueValues As Boolean
    End Class

    Public Class ListSpec
        Public Property DisplayName As String
        Public Property Template As ListTemplateKind
        Public Property Columns As IReadOnlyList(Of ColumnSpec)
    End Class

    Public Class ExistingColumn
        Public Property Name As String
        Public Property Id As String
        Public Property Indexed As Boolean
        Public Property EnforceUniqueValues As Boolean
    End Class

    Public Class ExistingList
        Public Property DisplayName As String
        Public Property Id As String
        Public Property Template As ListTemplateKind
        Public Property Columns As List(Of ExistingColumn)
    End Class

    Public Class SchemaGap
        Public Property ListName As String
        Public Property ColumnName As String

        Public Overrides Function ToString() As String
            Return ListName & "." & ColumnName
        End Function
    End Class

    Public Class SchemaReport
        Public Property MissingLists As New List(Of String)
        Public Property MissingColumns As New List(Of SchemaGap)
        Public Property IndexGaps As New List(Of SchemaGap)
        Public Property UniqueGaps As New List(Of SchemaGap)
        Public Property TemplateMismatches As New List(Of String)

        Public ReadOnly Property IsSatisfied As Boolean
            Get
                Return MissingLists.Count = 0 AndAlso MissingColumns.Count = 0 AndAlso IndexGaps.Count = 0 AndAlso UniqueGaps.Count = 0 AndAlso TemplateMismatches.Count = 0
            End Get
        End Property

        Public Function Summary() As String
            If IsSatisfied Then Return "Every required list and column is in place."
            Dim lines As New List(Of String)
            If MissingLists.Count > 0 Then lines.Add("Missing lists: " & String.Join(", ", MissingLists))
            If MissingColumns.Count > 0 Then lines.Add("Missing columns: " & String.Join(", ", MissingColumns.Select(Function(g) g.ToString())))
            If IndexGaps.Count > 0 Then lines.Add("Columns that need an index: " & String.Join(", ", IndexGaps.Select(Function(g) g.ToString())))
            If UniqueGaps.Count > 0 Then lines.Add("Columns that need unique values: " & String.Join(", ", UniqueGaps.Select(Function(g) g.ToString())))
            If TemplateMismatches.Count > 0 Then lines.Add("Wrong list type (left untouched): " & String.Join("; ", TemplateMismatches))
            Return String.Join(vbCrLf, lines)
        End Function
    End Class

    Public Class SharePointSite
        Public Property Id As String
        Public Property DisplayName As String
    End Class

    Public Class ListItemRecord
        Public Property Id As String
        Public Property Fields As FieldBag
    End Class

    Public Class FieldBag
        Private ReadOnly _text As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _bools As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)

        Public Sub SetString(name As String, value As String)
            _text(name) = If(value, "")
        End Sub

        Public Sub SetBool(name As String, value As Boolean)
            _bools(name) = value
            _text(name) = If(value, "true", "false")
        End Sub

        Public Function Has(name As String) As Boolean
            Return _text.ContainsKey(name) OrElse _bools.ContainsKey(name)
        End Function

        Public Function GetString(name As String) As String
            Dim value As String = Nothing
            If _text.TryGetValue(name, value) Then Return If(value, "")
            Return ""
        End Function

        Public Function GetBool(name As String) As Boolean
            Dim value As Boolean
            If _bools.TryGetValue(name, value) Then Return value
            Return YesNoBoolean.Parse(GetString(name))
        End Function
    End Class

End Namespace
