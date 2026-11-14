Namespace ElectrotechnologyResit.Core

    Public Class AbbreviationRule
        Public Property IsOverride As Boolean
        Public Property Match As String
        Public Property Replacement As String

        Public Function Clone() As AbbreviationRule
            Return New AbbreviationRule With {.IsOverride = IsOverride, .Match = Match, .Replacement = Replacement}
        End Function
    End Class

    Public Class QualificationRecord
        Public Property QualificationCode As String
        Public Property Title As String
        Public Property Version As String
        Public Property Status As String
        Public Property ClassGroupSuffix As String
        Public Property Rules As List(Of AbbreviationRule)

        Public Function Clone() As QualificationRecord
            Dim copy As New QualificationRecord With {
                .QualificationCode = QualificationCode,
                .Title = Title,
                .Version = Version,
                .Status = Status,
                .ClassGroupSuffix = ClassGroupSuffix,
                .Rules = New List(Of AbbreviationRule)
            }
            If Rules IsNot Nothing Then
                For Each rule In Rules
                    copy.Rules.Add(rule.Clone())
                Next
            End If
            Return copy
        End Function
    End Class

    Public Class ClassGroupEntry
        Public Property UnitCode As String
        Public Property ClassGroupSuffix As String
        Public Property Rules As IList(Of AbbreviationRule)
    End Class

    ''' <summary>Draft, Active and Superseded, matching a versioned training-package list. Only Active is offered for a new booking.</summary>
    Public NotInheritable Class QualificationStatus
        Public Const Draft As String = "Draft"
        Public Const Active As String = "Active"
        Public Const Superseded As String = "Superseded"

        Private Sub New()
        End Sub

        Public Shared Function Normalise(value As String) As String
            Dim text = If(value, "").Trim()
            If text.Equals(Draft, StringComparison.OrdinalIgnoreCase) Then Return Draft
            If text.Equals(Superseded, StringComparison.OrdinalIgnoreCase) Then Return Superseded
            If text.Equals(Active, StringComparison.OrdinalIgnoreCase) Then Return Active
            Return text
        End Function

        Public Shared Function OfferedForNewBookings(status As String) As Boolean
            Return String.Equals(Normalise(status), Active, StringComparison.Ordinal)
        End Function
    End Class

    Public NotInheritable Class AbbreviationRules
        Private Sub New()
        End Sub

        Public Shared Function Parse(text As String) As List(Of AbbreviationRule)
            Dim rules As New List(Of AbbreviationRule)
            If String.IsNullOrWhiteSpace(text) Then Return rules
            For Each raw In text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf)
                Dim line = raw.Trim()
                If line.Length = 0 OrElse line.StartsWith("#") Then Continue For
                Dim isOverride = False
                If line.StartsWith("override ", StringComparison.OrdinalIgnoreCase) Then
                    isOverride = True
                    line = line.Substring(9).Trim()
                ElseIf line.StartsWith("prefix ", StringComparison.OrdinalIgnoreCase) Then
                    line = line.Substring(7).Trim()
                End If
                Dim eq = line.IndexOf("="c)
                If eq <= 0 Then Continue For
                rules.Add(New AbbreviationRule With {
                    .IsOverride = isOverride,
                    .Match = line.Substring(0, eq).Trim(),
                    .Replacement = line.Substring(eq + 1).Trim()
                })
            Next
            Return rules
        End Function

        Public Shared Function Format(rules As IEnumerable(Of AbbreviationRule)) As String
            If rules Is Nothing Then Return ""
            Dim lines As New List(Of String)
            For Each rule In rules
                Dim kind = If(rule.IsOverride, "override ", "prefix ")
                lines.Add(kind & rule.Match & "=" & rule.Replacement)
            Next
            Return String.Join(vbCrLf, lines)
        End Function

        ''' <summary>Override first, then the longest matching prefix. No match leaves the code unchanged.</summary>
        Public Shared Function Abbreviate(unitCode As String, rules As IEnumerable(Of AbbreviationRule)) As String
            If String.IsNullOrWhiteSpace(unitCode) Then Return unitCode
            Dim code = unitCode.Trim()
            If rules IsNot Nothing Then
                Dim match As AbbreviationRule = Nothing
                For Each rule In rules
                    If rule Is Nothing OrElse String.IsNullOrWhiteSpace(rule.Match) Then Continue For
                    If rule.IsOverride AndAlso code.Equals(rule.Match.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        Return If(rule.Replacement, "")
                    End If
                Next
                For Each rule In rules
                    If rule Is Nothing OrElse rule.IsOverride OrElse String.IsNullOrWhiteSpace(rule.Match) Then Continue For
                    Dim prefix = rule.Match.Trim()
                    If code.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                        If match Is Nothing OrElse prefix.Length > match.Match.Trim().Length Then match = rule
                    End If
                Next
                If match IsNot Nothing Then Return If(match.Replacement, "") & code.Substring(match.Match.Trim().Length)
            End If
            Return code
        End Function

        Public Shared Function Uee30820Rules() As List(Of AbbreviationRule)
            Return Parse("prefix UEECD=CD" & vbLf & "prefix UEEEL=L" & vbLf & "prefix UEECO=CO" & vbLf & "prefix UEERE=RE" & vbLf & "prefix UEEDV=V")
        End Function
    End Class

    Public NotInheritable Class QualificationCatalog
        Private Sub New()
        End Sub

        Public Shared Function DefaultPackage() As QualificationRecord
            Return New QualificationRecord With {
                .QualificationCode = "UEE30820",
                .Title = "Certificate III in Electrotechnology Electrician",
                .Version = "1",
                .Status = QualificationStatus.Active,
                .ClassGroupSuffix = "CERTIII",
                .Rules = AbbreviationRules.Uee30820Rules()
            }
        End Function

        Public Shared Function DefaultUnits() As IReadOnlyList(Of UnitRecord)
            Dim code = "UEE30820"
            Return New List(Of UnitRecord) From {
                PackagedUnit(code, "UEECD0019", "Fabricate, assemble and dismantle utilities industry components"),
                PackagedUnit(code, "UEECD0020", "Fix and secure electrotechnology equipment"),
                PackagedUnit(code, "UEECD0046", "Solve problems in single path circuits"),
                PackagedUnit(code, "UEEEL0005", "Develop and connect electrical control circuits"),
                PackagedUnit(code, "UEEEL0012", "Install low voltage wiring, appliances, switchgear and associated accessories"),
                PackagedUnit(code, "UEEEL0021", "Solve problems in magnetic and electromagnetic devices"),
                PackagedUnit(code, "UEEEL0025", "Test and connect transformers")
            }
        End Function

        ''' <summary>Inserts UEE30820 only when the store has no qualifications yet. Existing rows are left in place.</summary>
        Public Shared Async Function EnsureSeedAsync(qualifications As IQualificationDirectory, units As IUnitDirectory) As Task
            If qualifications Is Nothing OrElse units Is Nothing Then Return
            Dim existing = Await qualifications.GetAllAsync().ConfigureAwait(False)
            If existing IsNot Nothing AndAlso existing.Count > 0 Then Return
            Await qualifications.UpsertAsync(DefaultPackage()).ConfigureAwait(False)
            Dim current = Await units.GetAllAsync().ConfigureAwait(False)
            For Each packaged In DefaultUnits()
                Dim match As UnitRecord = Nothing
                If current IsNot Nothing Then
                    match = current.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(item.UnitCode, packaged.UnitCode, StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(If(item.QualificationCode, ""), packaged.QualificationCode, StringComparison.OrdinalIgnoreCase))
                    If match Is Nothing Then
                        match = current.FirstOrDefault(Function(item) item IsNot Nothing AndAlso String.Equals(item.UnitCode, packaged.UnitCode, StringComparison.OrdinalIgnoreCase) AndAlso String.IsNullOrWhiteSpace(item.QualificationCode))
                    End If
                End If
                If match Is Nothing Then
                    Await units.UpsertAsync(packaged).ConfigureAwait(False)
                ElseIf String.IsNullOrWhiteSpace(match.QualificationCode) Then
                    Await units.UpsertAsync(New UnitRecord With {
                        .UnitCode = match.UnitCode,
                        .UnitTitle = If(String.IsNullOrWhiteSpace(match.UnitTitle), packaged.UnitTitle, match.UnitTitle),
                        .QualificationCode = packaged.QualificationCode
                    }).ConfigureAwait(False)
                End If
            Next
        End Function

        ''' <summary>Blank bookings are stored as this qualification. An Active UEE30820 wins, otherwise the first Active package, otherwise UEE30820.</summary>
        Public Shared Function CodeForBlankBookings(qualifications As IEnumerable(Of QualificationRecord)) As String
            Dim firstActive = ""
            If qualifications IsNot Nothing Then
                For Each qualification In qualifications
                    If qualification Is Nothing OrElse Not QualificationStatus.OfferedForNewBookings(qualification.Status) Then Continue For
                    Dim code = If(qualification.QualificationCode, "").Trim()
                    If code.Length = 0 Then Continue For
                    If code.Equals("UEE30820", StringComparison.OrdinalIgnoreCase) Then Return code
                    If firstActive.Length = 0 Then firstActive = code
                Next
            End If
            If firstActive.Length > 0 Then Return firstActive
            Return "UEE30820"
        End Function

        Private Shared Function PackagedUnit(qualificationCode As String, code As String, title As String) As UnitRecord
            Return New UnitRecord With {.QualificationCode = qualificationCode, .UnitCode = code, .UnitTitle = title}
        End Function
    End Class

    Public NotInheritable Class QualificationSelection
        Private Sub New()
        End Sub

        Public Shared Function ForNewBooking(unitCode As String, units As IEnumerable(Of UnitRecord), qualifications As IEnumerable(Of QualificationRecord)) As String
            Dim code = If(unitCode, "").Trim()
            If code.Length = 0 OrElse units Is Nothing Then Return ""
            Dim active As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If qualifications IsNot Nothing Then
                For Each qualification In qualifications
                    If qualification IsNot Nothing AndAlso QualificationStatus.OfferedForNewBookings(qualification.Status) Then
                        active.Add(If(qualification.QualificationCode, "").Trim())
                    End If
                Next
            End If
            For Each unit In units
                If unit Is Nothing Then Continue For
                If Not code.Equals(If(unit.UnitCode, "").Trim(), StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim package = If(unit.QualificationCode, "").Trim()
                If active.Contains(package) Then Return package
            Next
            Return ""
        End Function
    End Class

End Namespace
