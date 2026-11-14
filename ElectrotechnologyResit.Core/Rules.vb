Imports System.Globalization

Namespace ElectrotechnologyResit.Core

    ''' <summary>
    ''' Parses Energyspace and similar Yes/No values.
    ''' Isolated so the rule can change without hunting through the forms.
    ''' Accepts Yes, No, True and False (any capitalisation). Anything else does not parse.
    ''' </summary>
    Public NotInheritable Class YesNoBoolean
        Private Sub New()
        End Sub

        Public Shared Function TryParse(text As String, ByRef value As Boolean) As Boolean
            value = False
            If text Is Nothing Then Return False
            Dim trimmed = text.Trim()
            If trimmed.Equals("yes", StringComparison.OrdinalIgnoreCase) OrElse trimmed.Equals("true", StringComparison.OrdinalIgnoreCase) Then
                value = True
                Return True
            End If
            If trimmed.Equals("no", StringComparison.OrdinalIgnoreCase) OrElse trimmed.Equals("false", StringComparison.OrdinalIgnoreCase) Then
                value = False
                Return True
            End If
            Return False
        End Function

        Public Shared Function Parse(text As String) As Boolean
            Dim value As Boolean
            If TryParse(text, value) Then Return value
            Return False
        End Function

        ''' <summary>Text shown in the Edit Student Energyspace combos. Yes is ticked, No is not.</summary>
        Public Shared Function ToYesNo(value As Boolean) As String
            Return If(value, "Yes", "No")
        End Function
    End Class

    Public NotInheritable Class EnglishOrdinal
        Private Sub New()
        End Sub

        Public Shared Function Suffix(day As Integer) As String
            Dim mod100 = day Mod 100
            If mod100 >= 11 AndAlso mod100 <= 13 Then Return "th"
            Select Case day Mod 10
                Case 1
                    Return "st"
                Case 2
                    Return "nd"
                Case 3
                    Return "rd"
                Case Else
                    Return "th"
            End Select
        End Function

        ''' <summary>Example: 3rd November, 2026. 11th, 12th and 13th use "th".</summary>
        Public Shared Function FormatResitNight(day As Date) As String
            Dim culture = CultureInfo.GetCultureInfo("en-AU")
            Return day.Day.ToString(culture) & Suffix(day.Day) & " " & day.ToString("MMMM, yyyy", culture)
        End Function
    End Class

    Public NotInheritable Class MelbourneCalendar
        Private Sub New()
        End Sub

        Public Shared Function Zone() As TimeZoneInfo
            For Each id In New String() {"Australia/Melbourne", "AUS Eastern Standard Time"}
                Try
                    Return TimeZoneInfo.FindSystemTimeZoneById(id)
                Catch ex As TimeZoneNotFoundException
                Catch ex As InvalidTimeZoneException
                End Try
            Next
            Throw New InvalidOperationException("The Australia/Melbourne time zone was not found on this computer.")
        End Function

        Public Shared Function Today() As Date
            Return TimeZoneInfo.ConvertTimeFromUtc(Date.UtcNow, Zone()).Date
        End Function

        ''' <summary>
        ''' Reads a SharePoint date-only value as a calendar date.
        ''' The first yyyy-MM-dd in the stored text is the date. It is not shifted through UTC.
        ''' </summary>
        Public Shared Function ReadCalendarDate(stored As String) As Date
            If String.IsNullOrWhiteSpace(stored) Then Return Date.MinValue
            Dim text = stored.Trim().Trim(""""c)
            If text.Length >= 10 Then
                Dim parsed As Date
                If Date.TryParseExact(text.Substring(0, 10), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
                    Return parsed
                End If
            End If
            Return Date.MinValue
        End Function

        Public Shared Function ToDateOnlyString(day As Date) As String
            Return day.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        End Function
    End Class

    ''' <summary>Exactly one email mode is on. The two choices behave like radio buttons.</summary>
    Public NotInheritable Class EmailSendMode
        Private Sub New()
        End Sub

        Public Shared Function AfterDisplayChanged(displayChecked As Boolean) As (DisplayManually As Boolean, SendImmediately As Boolean)
            If displayChecked Then Return (True, False)
            Return (False, True)
        End Function

        Public Shared Function AfterSendChanged(sendChecked As Boolean) As (DisplayManually As Boolean, SendImmediately As Boolean)
            If sendChecked Then Return (False, True)
            Return (True, False)
        End Function
    End Class

    Public NotInheritable Class TeacherEmailResolver
        Private Sub New()
        End Sub

        ''' <summary>Email that belongs to the selected teacher. Empty when the name is unknown. Never a stale label.</summary>
        Public Shared Function EmailFor(teachers As IEnumerable(Of TeacherRecord), selectedName As String) As String
            If teachers Is Nothing OrElse String.IsNullOrWhiteSpace(selectedName) Then Return ""
            Dim wanted = selectedName.Trim()
            For Each teacher In teachers
                If teacher Is Nothing Then Continue For
                If String.Equals(If(teacher.FullName, "").Trim(), wanted, StringComparison.OrdinalIgnoreCase) Then
                    Return If(teacher.Email, "").Trim()
                End If
            Next
            Return ""
        End Function
    End Class

    Public NotInheritable Class BookingUpsert
        Private Sub New()
        End Sub

        ''' <summary>
        ''' One student ID is one booking. When preserveFlags is true and a row already exists,
        ''' Energyspace checkpoints and status stay as they were. A new row uses the incoming flags.
        ''' </summary>
        Public Shared Function Merge(existing As ResitBooking, incoming As ResitBooking, preserveFlags As Boolean, teachers As IEnumerable(Of TeacherRecord)) As ResitBooking
            If incoming Is Nothing Then Throw New ArgumentNullException(NameOf(incoming))
            Dim merged = incoming.Clone()
            merged.StudentId = If(incoming.StudentId, "").Trim()
            merged.ResitDate = incoming.ResitDate.Date
            If existing IsNot Nothing AndAlso preserveFlags Then
                merged.EnergyspaceCreated = existing.EnergyspaceCreated
                merged.EnergyspaceAssessmentBooked = existing.EnergyspaceAssessmentBooked
                merged.Status = existing.Status
                merged.SharePointItemId = existing.SharePointItemId
            ElseIf existing IsNot Nothing Then
                merged.SharePointItemId = existing.SharePointItemId
                If String.IsNullOrEmpty(merged.Status) Then merged.Status = existing.Status
            End If
            merged.TeacherName = If(merged.TeacherName, "").Trim()
            merged.TeacherEmail = TeacherEmailResolver.EmailFor(teachers, merged.TeacherName)
            If existing IsNot Nothing Then
                merged.QualificationCode = If(existing.QualificationCode, "")
            Else
                merged.QualificationCode = If(incoming.QualificationCode, "").Trim()
            End If
            Return merged
        End Function
    End Class

    Public NotInheritable Class EnergyspaceFlagRules
        Private Sub New()
        End Sub

        ''' <summary>Ticked only when every booking for that unit on that night is true.</summary>
        Public Shared Function Summarise(bookings As IEnumerable(Of ResitBooking), night As Date) As IReadOnlyList(Of UnitNightFlags)
            Dim groups As New Dictionary(Of String, List(Of ResitBooking))(StringComparer.OrdinalIgnoreCase)
            Dim order As New List(Of String)
            If bookings IsNot Nothing Then
                For Each booking In bookings
                    If booking Is Nothing Then Continue For
                    If booking.ResitDate.Date <> night.Date Then Continue For
                    Dim unit = If(booking.UnitCode, "").Trim()
                    If unit.Length = 0 Then Continue For
                    If Not groups.ContainsKey(unit) Then
                        groups(unit) = New List(Of ResitBooking)
                        order.Add(unit)
                    End If
                    groups(unit).Add(booking)
                Next
            End If
            order.Sort(StringComparer.OrdinalIgnoreCase)
            Dim result As New List(Of UnitNightFlags)
            For Each unit In order
                Dim rows = groups(unit)
                result.Add(New UnitNightFlags With {
                    .UnitCode = unit,
                    .ClassCreated = rows.Count > 0 AndAlso rows.All(Function(b) b.EnergyspaceCreated),
                    .AssessmentBooked = rows.Count > 0 AndAlso rows.All(Function(b) b.EnergyspaceAssessmentBooked)
                })
            Next
            Return result
        End Function
    End Class

    Public NotInheritable Class ODataFilters
        Private Sub New()
        End Sub

        Public Shared Function EscapeLiteral(value As String) As String
            Return If(value, "").Replace("'", "''")
        End Function

        Public Shared Function TextEquals(column As String, value As String) As String
            Return "fields/" & column & " eq '" & EscapeLiteral(value) & "'"
        End Function

        Public Shared Function DateEquals(column As String, day As Date) As String
            Dim start = MelbourneCalendar.ToDateOnlyString(day.Date)
            Dim finish = MelbourneCalendar.ToDateOnlyString(day.Date.AddDays(1))
            Return "fields/" & column & " ge '" & start & "' and fields/" & column & " lt '" & finish & "'"
        End Function

        Public Shared Function DateOnOrBefore(column As String, day As Date) As String
            Dim finish = MelbourneCalendar.ToDateOnlyString(day.Date.AddDays(1))
            Return "fields/" & column & " lt '" & finish & "'"
        End Function
    End Class

    Public NotInheritable Class GraphRetry
        Public Const MaxAttempts As Integer = 5
        Public Const PreferHeaderName As String = "Prefer"
        Public Const PreferNonIndexedValue As String = "HonorNonIndexedQueriesWarningMayFailRandomly"

        Private Sub New()
        End Sub

        Public Shared Function IsNonIndexedFilterError(statusCode As Integer, body As String) As Boolean
            If statusCode <> 400 AndAlso statusCode <> 500 Then Return False
            If String.IsNullOrEmpty(body) Then Return False
            Dim text = body.ToLowerInvariant()
            Return text.Contains("index") OrElse text.Contains("honornonindexedqueries")
        End Function

        Public Shared Function GetDelay(statusCode As Integer, retryAfter As String, attempt As Integer) As TimeSpan?
            If statusCode <> 429 AndAlso statusCode <> 503 Then Return Nothing
            If attempt >= MaxAttempts Then Return Nothing
            Dim seconds As Double
            If Double.TryParse(If(retryAfter, "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, seconds) AndAlso seconds >= 0 Then
                Return TimeSpan.FromSeconds(Math.Min(seconds, 120))
            End If
            Dim retryAt As DateTimeOffset
            If DateTimeOffset.TryParse(retryAfter, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, retryAt) Then
                Dim wait = retryAt - DateTimeOffset.UtcNow
                If wait < TimeSpan.Zero Then wait = TimeSpan.Zero
                If wait > TimeSpan.FromMinutes(2) Then wait = TimeSpan.FromMinutes(2)
                Return wait
            End If
            Dim fallback = Math.Pow(2, Math.Max(attempt - 1, 0))
            Return TimeSpan.FromSeconds(Math.Min(fallback, 30))
        End Function
    End Class

    ''' <summary>
    ''' Delegated Graph scopes for this app.
    ''' Creating a list or column needs Sites.Manage.All. Reading and writing rows needs Sites.ReadWrite.All.
    ''' Sites.Selected replaces both, and the site grant itself must allow list management.
    ''' </summary>
    Public NotInheritable Class GraphPermissionScopes
        Public Const ManageAll As String = "https://graph.microsoft.com/Sites.Manage.All"
        Public Const ReadWriteAll As String = "https://graph.microsoft.com/Sites.ReadWrite.All"
        Public Const Selected As String = "https://graph.microsoft.com/Sites.Selected"
        Public Const UserRead As String = "https://graph.microsoft.com/User.Read"

        Private Sub New()
        End Sub

        Public Shared Function ForSharePoint(useSitesSelected As Boolean) As String()
            If useSitesSelected Then
                Return New String() {Selected, UserRead}
            End If
            Return New String() {ManageAll, ReadWriteAll, UserRead}
        End Function
    End Class

    Public NotInheritable Class SharePointSetupValidation
        Private Sub New()
        End Sub

        ''' <summary>Returns an error message, or Nothing when the setup values can be saved.</summary>
        Public Shared Function Validate(siteUrl As String, clientId As String, names As SharePointListNames) As String
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(If(siteUrl, "").Trim(), UriKind.Absolute, uri) OrElse Not uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then
                Return "Enter the SharePoint site address, starting with https://."
            End If
            If String.IsNullOrWhiteSpace(clientId) Then
                Return "Enter the Azure AD application (client) ID."
            End If
            If names Is Nothing Then Return "Enter a name for each list."
            Dim listNames = names.AllNames().ToList()
            For Each listName In listNames
                Dim problem = ValidateListName(listName)
                If problem IsNot Nothing Then Return problem
            Next
            Dim distinct = New HashSet(Of String)(listNames, StringComparer.OrdinalIgnoreCase)
            If distinct.Count <> listNames.Count Then
                Return "Each list needs its own name."
            End If
            Return Nothing
        End Function

        Private Shared Function ValidateListName(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return "Enter a name for each list."
            If name.Length > 80 Then Return "List names must be 80 characters or fewer."
            Dim banned = "\/:*?""<>|#%"
            For Each ch In name
                If banned.Contains(ch) Then
                    Return "List names cannot contain \ / : * ? "" < > | # %."
                End If
            Next
            Return Nothing
        End Function
    End Class

End Namespace
