Imports System.Globalization
Imports System.Linq
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports ElectrotechnologyResit.Core

' TODO: Live Microsoft Bookings is owned by VU Timetable (a separate app) and needs Frank Offer's sign-off.
' This import stays hidden behind ShowPullFromBookingsButton. Do not extend it into a supported live pull.

''' <summary>Maps Microsoft Bookings appointments onto resit bookings. Writes go through IBookingService and keep existing Energyspace flags.</summary>
Friend NotInheritable Class BookingsDataImporter
    Implements IBookingImporter

    Private ReadOnly _owner As Form
    Private ReadOnly _bookings As IBookingService

    Public Sub New(owner As Form, bookings As IBookingService)
        _owner = owner
        _bookings = bookings
    End Sub

    Public Function PullAsync(resitDate As Date) As Task(Of String) Implements IBookingImporter.PullAsync
        Return PullFromBookingsAsync(_owner, resitDate, _bookings)
    End Function
    Private Shared ReadOnly IdRegex As New Regex("\b\d{6,10}\b", RegexOptions.Compiled)
    ''' <summary>First token like UEECO0023 from "UEECO0023 Long unit title…"</summary>
    Private Shared ReadOnly UnitCodeRegex As New Regex("^\s*([A-Za-z]{2,}\d+[A-Za-z0-9]*)\b", RegexOptions.Compiled)

    Private Structure VuBookingAnswers
        Public StudentId As String
        Public Blockgroup As String
        Public Teacher As String
        Public UnitAnswer As String
        Public AttemptNo As String
        Public EnergySpaceEmail As String
    End Structure

    Friend Shared Async Function PullFromBookingsAsync(owner As Form, resitDate As Date, bookings As IBookingService) As Task(Of String)
        Dim clientId = My.Settings.BookingsAzureClientId
        If String.IsNullOrWhiteSpace(clientId) Then
            Return "Set the Azure AD application (client) ID in Settings (Microsoft Bookings / Graph section), then try again."
        End If

        Dim tenant = My.Settings.BookingsAzureTenantId
        Dim pca = BookingsGraphClient.CreatePublicClientApplication(clientId, tenant)
        Dim auth = Await BookingsGraphClient.AcquireTokenAsync(pca, owner).ConfigureAwait(True)

        Dim businessId = My.Settings.BookingsBusinessId?.Trim()
        If String.IsNullOrEmpty(businessId) Then
            Dim businesses = Await BookingsGraphClient.ListBookingBusinessIdsAsync(auth.AccessToken).ConfigureAwait(True)
            If businesses Is Nothing OrElse businesses.Count = 0 Then
                Return "No Microsoft Bookings businesses were returned. Check permissions (Bookings.Read.All) and that your account can access Bookings."
            End If
            businessId = businesses(0).Item1
            Dim dn = businesses(0).Item2
            My.Settings.BookingsBusinessId = businessId
            My.Settings.Save()
            ' Inform which business was bound on first use
            System.Diagnostics.Debug.WriteLine("Bookings business: " & dn & " / " & businessId)
        End If

        Dim staff = Await BookingsGraphClient.ListStaffMembersAsync(auth.AccessToken, businessId).ConfigureAwait(True)
        Dim questionLabels = Await BookingsGraphClient.ListCustomQuestionLabelsAsync(auth.AccessToken, businessId).ConfigureAwait(True)

        Dim startIso As String = Nothing
        Dim endIso As String = Nothing
        BookingsGraphClient.GetLocalDayRangeIso(resitDate.Date, startIso, endIso)

        Dim rawList = Await BookingsGraphClient.ListCalendarAppointmentsRawAsync(auth.AccessToken, businessId, startIso, endIso).ConfigureAwait(True)

        Dim resitDateStr = resitDate.ToString("yyyy-MM-dd")
        Dim inserted = 0
        Dim updated = 0
        Dim skipped = 0
        Dim errors As New List(Of String)()

        For Each raw In rawList
            Try
                Using doc = JsonDocument.Parse(raw)
                    Dim appt = doc.RootElement
                    If Not AppointmentOnLocalDate(appt, resitDate.Date) Then
                        skipped += 1
                        Continue For
                    End If

                    Dim vu = ParseVuCustomAnswers(appt, questionLabels)

                    Dim studentId = If(vu.StudentId, "").Trim()
                    If String.IsNullOrWhiteSpace(studentId) Then studentId = ExtractStudentId(appt)
                    If String.IsNullOrWhiteSpace(studentId) Then
                        skipped += 1
                        errors.Add("Skipped an appointment (no student ID found): " & GetAppointmentLabel(appt))
                        Continue For
                    End If

                    Dim email = If(vu.EnergySpaceEmail, "").Trim()
                    If String.IsNullOrEmpty(email) Then email = GetCustomerEmail(appt)

                    Dim custName = GetCustomerName(appt)
                    Dim firstName = ""
                    Dim lastName = ""
                    SplitCustomerName(custName, firstName, lastName)

                    Dim unitCode = ""
                    Dim unitName = ""
                    SplitUnitAnswer(vu.UnitAnswer, unitCode, unitName)
                    If String.IsNullOrEmpty(unitCode) Then
                        Dim unEl As JsonElement
                        If appt.TryGetProperty("serviceName", unEl) Then
                            Dim svc = If(unEl.GetString(), "").Trim()
                            unitCode = svc
                            unitName = svc
                        End If
                    End If

                    Dim teacher = If(vu.Teacher, "").Trim()
                    Dim teacherEmail = ""
                    Dim idsEl As JsonElement
                    If appt.TryGetProperty("staffMemberIds", idsEl) AndAlso idsEl.ValueKind = JsonValueKind.Array Then
                        For Each sid In idsEl.EnumerateArray()
                            Dim s = sid.GetString()
                            If Not String.IsNullOrEmpty(s) AndAlso staff.ContainsKey(s) Then
                                If String.IsNullOrEmpty(teacher) Then teacher = staff(s).Item1
                                teacherEmail = staff(s).Item2
                                Exit For
                            End If
                        Next
                    End If

                    Dim attemptNo = If(vu.AttemptNo, "").Trim()
                    If String.IsNullOrEmpty(attemptNo) Then attemptNo = "1"

                    Dim blockgroup = If(vu.Blockgroup, "").Trim()

                    Dim result = Await UpsertResitRow(bookings, studentId, firstName, lastName, email, teacher, teacherEmail, unitCode, unitName, attemptNo, blockgroup, resitDateStr).ConfigureAwait(True)
                    If result = 1 Then inserted += 1
                    If result = 2 Then updated += 1
                End Using
            Catch ex As Exception
                errors.Add(ex.Message)
            End Try
        Next

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Bookings import for " & resitDateStr & " (local calendar day).")
        sb.AppendLine("Appointments fetched: " & rawList.Count.ToString())
        sb.AppendLine("Inserted: " & inserted.ToString() & ", updated: " & updated.ToString() & ", skipped: " & skipped.ToString() & ".")
        If errors.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Notes:")
            For Each e In errors.Take(8)
                sb.AppendLine(" - " & e)
            Next
            If errors.Count > 8 Then sb.AppendLine(" - … (" & (errors.Count - 8).ToString() & " more)")
        End If
        Return sb.ToString().TrimEnd()
    End Function

    Private Shared Function GetAppointmentLabel(appt As JsonElement) As String
        Dim svc = ""
        Dim n = ""
        Dim s As JsonElement
        If appt.TryGetProperty("serviceName", s) Then svc = If(s.GetString(), "")
        If appt.TryGetProperty("customerName", s) Then n = If(s.GetString(), "")
        Return Trim(svc & " — " & n).Trim(" "c, "—"c)
    End Function

    ''' <summary>True if appointment start falls on the given local calendar date.</summary>
    Private Shared Function AppointmentOnLocalDate(appt As JsonElement, localDate As Date) As Boolean
        Dim startObj As JsonElement
        If Not appt.TryGetProperty("start", startObj) Then Return True
        Dim dtEl As JsonElement
        If Not startObj.TryGetProperty("dateTime", dtEl) Then Return True
        Dim dtStr = If(dtEl.GetString(), "")
        If String.IsNullOrEmpty(dtStr) Then Return True
        Dim parsed As DateTime
        If DateTime.TryParse(dtStr, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, parsed) Then
            Dim local = parsed
            If parsed.Kind = DateTimeKind.Utc Then
                local = TimeZoneInfo.ConvertTimeFromUtc(parsed, TimeZoneInfo.Local)
            ElseIf parsed.Kind = DateTimeKind.Unspecified Then
                local = DateTime.SpecifyKind(parsed, DateTimeKind.Local)
            End If
            Return local.Date = localDate.Date
        End If
        Return True
    End Function

    Private Shared Function GetCustomerEmail(appt As JsonElement) As String
        Dim e As JsonElement
        If appt.TryGetProperty("customerEmailAddress", e) Then Return If(e.GetString(), "").Trim()
        Dim cust As JsonElement
        If appt.TryGetProperty("customers", cust) AndAlso cust.ValueKind = JsonValueKind.Array Then
            For Each c In cust.EnumerateArray()
                If c.TryGetProperty("emailAddress", e) Then Return If(e.GetString(), "").Trim()
            Next
        End If
        Return ""
    End Function

    Private Shared Function GetCustomerName(appt As JsonElement) As String
        Dim e As JsonElement
        If appt.TryGetProperty("customerName", e) Then Return If(e.GetString(), "").Trim()
        If appt.TryGetProperty("customers", e) AndAlso e.ValueKind = JsonValueKind.Array Then
            For Each c In e.EnumerateArray()
                Dim n As JsonElement
                If c.TryGetProperty("name", n) Then Return If(n.GetString(), "").Trim()
            Next
        End If
        Return ""
    End Function

    Private Shared Sub SplitCustomerName(raw As String, ByRef firstName As String, ByRef lastName As String)
        firstName = ""
        lastName = ""
        If String.IsNullOrWhiteSpace(raw) Then Return
        raw = raw.Trim()
        If raw.Contains(","c) Then
            Dim parts = raw.Split(","c, 2)
            lastName = parts(0).Trim()
            firstName = If(parts.Length > 1, parts(1).Trim(), "")
        Else
            firstName = raw
        End If
    End Sub

    ''' <summary>Maps VU Electrotechnology Assessment Resit Class custom questions (by display text) into DB fields.</summary>
    Private Shared Function ParseVuCustomAnswers(appt As JsonElement, questionIdToLabel As Dictionary(Of String, String)) As VuBookingAnswers
        Dim r As New VuBookingAnswers()
        Dim e As JsonElement
        If Not appt.TryGetProperty("customers", e) OrElse e.ValueKind <> JsonValueKind.Array Then Return r
        For Each c In e.EnumerateArray()
            Dim qa As JsonElement
            If Not c.TryGetProperty("customQuestionAnswers", qa) OrElse qa.ValueKind <> JsonValueKind.Array Then Continue For
            For Each q In qa.EnumerateArray()
                Dim qidEl As JsonElement
                Dim ansEl As JsonElement
                If Not q.TryGetProperty("questionId", qidEl) OrElse Not q.TryGetProperty("answer", ansEl) Then Continue For
                Dim qid = If(qidEl.GetString(), "")
                Dim answer = If(ansEl.GetString(), "").Trim()
                Dim label = ""
                If Not String.IsNullOrEmpty(qid) AndAlso questionIdToLabel IsNot Nothing AndAlso questionIdToLabel.ContainsKey(qid) Then
                    label = questionIdToLabel(qid)
                End If
                Dim low = label.ToLowerInvariant()

                If low.Contains("energys") AndAlso low.Contains("email") Then
                    r.EnergySpaceEmail = answer
                    Continue For
                End If
                If low.Contains("attempt") AndAlso low.Contains("number") AndAlso Not low.Contains("previous") Then
                    r.AttemptNo = answer
                    Continue For
                End If
                If low.Contains("student id") OrElse (low.Contains("student") AndAlso low.Contains("number")) Then
                    Dim m = IdRegex.Match(answer)
                    If m.Success Then r.StudentId = m.Value
                    Continue For
                End If
                If low.Contains("unit") AndAlso (low.Contains("booking") OrElse low.Contains("which unit")) Then
                    r.UnitAnswer = answer
                    Continue For
                End If
                If low.Contains("teacher") Then
                    r.Teacher = answer
                    Continue For
                End If
                If (low.Contains("class") AndAlso low.Contains("group")) AndAlso Not low.Contains("teacher") Then
                    r.Blockgroup = answer
                    Continue For
                End If
            Next
        Next
        Return r
    End Function

    Private Shared Sub SplitUnitAnswer(full As String, ByRef unitCode As String, ByRef unitName As String)
        unitCode = ""
        unitName = ""
        If String.IsNullOrWhiteSpace(full) Then Return
        full = full.Trim()
        Dim m = UnitCodeRegex.Match(full)
        If m.Success Then
            unitCode = m.Groups(1).Value
            Dim after = full.Substring(m.Groups(1).Index + m.Groups(1).Length).Trim()
            unitName = If(String.IsNullOrEmpty(after), unitCode, after)
        Else
            unitName = full
            unitCode = full
        End If
    End Sub

    Private Shared Function ExtractStudentId(appt As JsonElement) As String
        Dim hay As New List(Of String)()
        Dim e As JsonElement
        If appt.TryGetProperty("customerNotes", e) Then hay.Add(If(e.GetString(), ""))
        If appt.TryGetProperty("serviceNotes", e) Then hay.Add(If(e.GetString(), ""))
        Dim em = GetCustomerEmail(appt)
        If Not String.IsNullOrEmpty(em) Then
            hay.Add(em)
            Dim at = em.IndexOf("@"c)
            If at > 0 Then hay.Add(em.Substring(0, at))
        End If
        If appt.TryGetProperty("customers", e) AndAlso e.ValueKind = JsonValueKind.Array Then
            For Each c In e.EnumerateArray()
                Dim qa As JsonElement
                If c.TryGetProperty("customQuestionAnswers", qa) AndAlso qa.ValueKind = JsonValueKind.Array Then
                    For Each q In qa.EnumerateArray()
                        Dim a As JsonElement
                        If q.TryGetProperty("answer", a) Then hay.Add(If(a.GetString(), ""))
                    Next
                End If
            Next
        End If

        For Each h In hay
            If String.IsNullOrEmpty(h) Then Continue For
            Dim m = IdRegex.Match(h)
            If m.Success Then Return m.Value
        Next
        Return ""
    End Function

    ''' <summary>Returns 1 = insert, 2 = update. Existing Energyspace flags are kept.</summary>
    Private Shared Async Function UpsertResitRow(
        bookings As IBookingService,
        studentId As String,
        firstName As String,
        lastName As String,
        email As String,
        teacher As String,
        teacherEmail As String,
        unit As String,
        unitName As String,
        attemptNo As String,
        blockgroup As String,
        resitDateStr As String) As Task(Of Integer)

        Dim night As Date
        If Not Date.TryParseExact(resitDateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, night) Then
            night = Date.Today
        End If
        Dim booking As New ResitBooking With {
            .StudentId = studentId,
            .FirstName = firstName,
            .Surname = lastName,
            .Email = email,
            .TeacherName = teacher,
            .TeacherEmail = teacherEmail,
            .UnitCode = unit,
            .UnitName = unitName,
            .AttemptNo = attemptNo,
            .BlockGroup = blockgroup,
            .ResitDate = night,
            .EnergyspaceCreated = False,
            .EnergyspaceAssessmentBooked = False
        }
        Dim result = Await bookings.UpsertAsync(booking, preserveEnergyspaceFlags:=True).ConfigureAwait(False)
        If result.Kind = UpsertKind.Inserted Then Return 1
        Return 2
    End Function
End Class
