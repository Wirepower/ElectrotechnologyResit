Imports System.IO
Imports ElectrotechnologyResit.Core
Imports OfficeOpenXml
Imports Xunit

Public Class BugFixTests
    Private Shared Function SampleBooking(studentId As String, unitCode As String, night As Date, Optional created As Boolean = False, Optional booked As Boolean = False) As ResitBooking
        Return New ResitBooking With {
            .StudentId = studentId,
            .FirstName = "Alex",
            .Surname = "Nguyen",
            .Email = "alex.nguyen@example.com",
            .UnitCode = unitCode,
            .UnitName = "Sample unit",
            .BlockGroup = "BG-A",
            .TeacherName = "Jordan Hale",
            .TeacherEmail = "jordan.hale@example.com",
            .AttemptNo = "1",
            .ResitDate = night,
            .EnergyspaceCreated = created,
            .EnergyspaceAssessmentBooked = booked
        }
    End Function

    <Fact>
    Public Async Function One_student_id_is_one_booking_row() As Task
        Dim store As New InMemoryResitStore()
        Dim night = New Date(2026, 11, 3)
        Await store.UpsertAsync(SampleBooking("9000001", "UEECD0044", night), False)
        Await store.UpsertAsync(SampleBooking("9000001", "UEEEL0021", night.AddDays(1)), False)
        Dim rows = Await store.GetBookingsAsync()
        Assert.Single(rows)
        Assert.Equal("UEEEL0021", rows(0).UnitCode)
        Assert.Equal(night.AddDays(1), rows(0).ResitDate)
    End Function

    <Fact>
    Public Async Function Bookings_upsert_keeps_existing_energyspace_flags() As Task
        Dim store As New InMemoryResitStore()
        Dim night = New Date(2026, 11, 3)
        Await store.UpsertAsync(SampleBooking("9000002", "UEECD0044", night, created:=True, booked:=True), False)
        Dim incoming = SampleBooking("9000002", "HLTAID009", night.AddDays(2), created:=False, booked:=False)
        incoming.TeacherName = "Morgan Ellis"
        Dim result = Await store.UpsertAsync(incoming, preserveEnergyspaceFlags:=True)
        Assert.Equal(UpsertKind.Updated, result.Kind)
        Assert.True(result.Booking.EnergyspaceCreated)
        Assert.True(result.Booking.EnergyspaceAssessmentBooked)
        Assert.Equal("HLTAID009", result.Booking.UnitCode)
        Assert.Equal("Morgan Ellis", result.Booking.TeacherName)
    End Function

    <Fact>
    Public Async Function New_booking_with_preserve_still_starts_unticked() As Task
        Dim store As New InMemoryResitStore()
        Dim result = Await store.UpsertAsync(SampleBooking("9000003", "UEECD0044", New Date(2026, 11, 3)), True)
        Assert.Equal(UpsertKind.Inserted, result.Kind)
        Assert.False(result.Booking.EnergyspaceCreated)
        Assert.False(result.Booking.EnergyspaceAssessmentBooked)
    End Function

    <Fact>
    Public Async Function Untick_clears_only_that_unit_and_night_and_lists_show_saved_state() As Task
        Dim store As New InMemoryResitStore()
        Dim night = New Date(2026, 11, 3)
        Dim otherNight = New Date(2026, 11, 4)
        Await store.UpsertAsync(SampleBooking("9000001", "UEECD0044", night, True, True), False)
        Await store.UpsertAsync(SampleBooking("9000002", "UEECD0044", night, True, True), False)
        Await store.UpsertAsync(SampleBooking("9000003", "UEEEL0021", night, True, False), False)
        Await store.UpsertAsync(SampleBooking("9000004", "UEECD0044", otherNight, True, True), False)

        Dim flags = Await store.GetUnitFlagsAsync(night)
        Assert.True(flags.Single(Function(item) item.UnitCode = "UEECD0044").ClassCreated)
        Assert.True(flags.Single(Function(item) item.UnitCode = "UEECD0044").AssessmentBooked)
        Assert.True(flags.Single(Function(item) item.UnitCode = "UEEEL0021").ClassCreated)
        Assert.False(flags.Single(Function(item) item.UnitCode = "UEEEL0021").AssessmentBooked)

        Await store.SetEnergyspaceCreatedAsync(night, "UEECD0044", False)
        Dim after = Await store.GetUnitFlagsAsync(night)
        Assert.False(after.Single(Function(item) item.UnitCode = "UEECD0044").ClassCreated)
        Assert.True(after.Single(Function(item) item.UnitCode = "UEECD0044").AssessmentBooked)
        Assert.True(after.Single(Function(item) item.UnitCode = "UEEEL0021").ClassCreated)
        Dim other = Await store.GetUnitFlagsAsync(otherNight)
        Assert.True(other.Single().ClassCreated)

        Await store.SetEnergyspaceAssessmentBookedAsync(night, "UEEEL0021", True)
        Dim ticked = Await store.GetUnitFlagsAsync(night)
        Assert.True(ticked.Single(Function(item) item.UnitCode = "UEEEL0021").AssessmentBooked)
    End Function

    <Fact>
    Public Sub Yes_no_true_false_parse_in_one_function()
        Assert.True(YesNoBoolean.Parse("Yes"))
        Assert.True(YesNoBoolean.Parse("yes"))
        Assert.True(YesNoBoolean.Parse("TRUE"))
        Assert.False(YesNoBoolean.Parse("No"))
        Assert.False(YesNoBoolean.Parse("no"))
        Assert.False(YesNoBoolean.Parse("False"))
        Dim parsed As Boolean
        Assert.True(YesNoBoolean.TryParse(" Yes ", parsed))
        Assert.True(parsed)
        Assert.False(YesNoBoolean.TryParse("1", parsed))
        Assert.False(parsed)
    End Sub

    <Fact>
    Public Sub Teacher_email_comes_from_the_selected_teacher()
        Dim teachers = New List(Of TeacherRecord) From {
            New TeacherRecord With {.FullName = "Jordan Hale", .Email = "jordan.hale@example.com"},
            New TeacherRecord With {.FullName = "Morgan Ellis", .Email = "morgan.ellis@example.com"}
        }
        Dim staleLabel = "jordan.hale@example.com"
        Dim selected = TeacherEmailResolver.EmailFor(teachers, "Morgan Ellis")
        Assert.Equal("morgan.ellis@example.com", selected)
        Assert.NotEqual(staleLabel, selected)
        Assert.Equal("", TeacherEmailResolver.EmailFor(teachers, "Nobody"))
    End Sub

    <Theory>
    <InlineData(1, "st")>
    <InlineData(2, "nd")>
    <InlineData(3, "rd")>
    <InlineData(4, "th")>
    <InlineData(11, "th")>
    <InlineData(12, "th")>
    <InlineData(13, "th")>
    <InlineData(21, "st")>
    <InlineData(22, "nd")>
    <InlineData(23, "rd")>
    <InlineData(31, "st")>
    Public Sub Ordinals_follow_normal_english(day As Integer, suffix As String)
        Assert.Equal(suffix, EnglishOrdinal.Suffix(day))
        Dim month = If(day >= 31, 1, 11)
        Dim monthName = If(day >= 31, "January", "November")
        Dim text = EnglishOrdinal.FormatResitNight(New Date(2026, month, day))
        Assert.Equal(day.ToString() & suffix & " " & monthName & ", 2026", text)
    End Sub

    <Fact>
    Public Sub Email_modes_are_mutually_exclusive()
        Dim displayOn = EmailSendMode.AfterDisplayChanged(True)
        Assert.True(displayOn.DisplayManually)
        Assert.False(displayOn.SendImmediately)
        Dim displayOff = EmailSendMode.AfterDisplayChanged(False)
        Assert.False(displayOff.DisplayManually)
        Assert.True(displayOff.SendImmediately)
        Dim sendOn = EmailSendMode.AfterSendChanged(True)
        Assert.False(sendOn.DisplayManually)
        Assert.True(sendOn.SendImmediately)
        Dim sendOff = EmailSendMode.AfterSendChanged(False)
        Assert.True(sendOff.DisplayManually)
        Assert.False(sendOff.SendImmediately)
    End Sub

    <Fact>
    Public Async Function Deletes_are_scoped_to_student_night_and_fourteen_days() As Task
        Dim store As New InMemoryResitStore()
        Await store.UpsertAsync(SampleBooking("9000001", "UEECD0044", New Date(2026, 10, 30)), False)
        Await store.UpsertAsync(SampleBooking("9000002", "UEECD0044", New Date(2026, 10, 31)), False)
        Await store.UpsertAsync(SampleBooking("9000003", "UEECD0044", New Date(2026, 11, 1)), False)
        Await store.UpsertAsync(SampleBooking("9000004", "UEEEL0021", New Date(2026, 11, 3)), False)
        Await store.DeleteByStudentIdAsync("9000004")
        Await store.DeleteOnOrBeforeAsync(New Date(2026, 10, 31))
        Dim left = Await store.GetBookingsAsync()
        Assert.Single(left)
        Assert.Equal("9000003", left(0).StudentId)
        Await store.DeleteByDateAsync(New Date(2026, 11, 1))
        Assert.Empty(Await store.GetBookingsAsync())
    End Function

    <Fact>
    Public Sub Melbourne_date_keeps_the_calendar_day()
        Assert.Equal(New Date(2026, 11, 14), MelbourneCalendar.ReadCalendarDate("2026-11-14T00:00:00Z"))
        Assert.Equal(New Date(2026, 11, 3), MelbourneCalendar.ReadCalendarDate("2026-11-03"))
        Assert.Equal("2026-11-03", MelbourneCalendar.ToDateOnlyString(New Date(2026, 11, 3)))
    End Sub

    <Fact>
    Public Sub Preserve_flags_omits_them_from_the_sharepoint_patch()
        Dim booking = SampleBooking("9000001", "UEECD0044", New Date(2026, 11, 3), True, True)
        Dim text As Dictionary(Of String, String) = Nothing
        Dim flags As Dictionary(Of String, Boolean) = Nothing
        BookingFieldMap.Fill(booking, includeFlags:=False, text, flags)
        Assert.Empty(flags)
        Assert.Equal("UEECD0044", text(SharePointColumns.Unit))
        Assert.Equal("2026-11-03", text(SharePointColumns.ResitDate))
        Assert.DoesNotContain(SharePointColumns.Status, text.Keys)
    End Sub
End Class

Public Class LookupCacheTests
    <Fact>
    Public Async Function Student_lookup_is_cached_and_queries_by_id() As Task
        Dim inner As New InMemoryResitStore()
        Await inner.UpsertAsync(New StudentRecord With {
            .StudentId = "9000001",
            .GivenName = "Alex",
            .FamilyName = "Nguyen",
            .PersonalEmail = "alex.nguyen@example.com",
            .BlockGroupCode = "BG-A"
        })
        Await inner.UpsertAsync(New StudentRecord With {.StudentId = "9000002", .GivenName = "Blair", .FamilyName = "Okafor"})
        Dim cache As New CachingStudentDirectory(inner)
        Dim first = Await cache.FindByStudentIdAsync("9000001")
        Dim second = Await cache.FindByStudentIdAsync("9000001")
        Assert.Equal("Alex", first.GivenName)
        Assert.Equal("Alex", second.GivenName)
        Assert.Equal(1, inner.StudentLookupCount)
        Await cache.FindByStudentIdAsync("9000004")
        Await cache.FindByStudentIdAsync("9000004")
        Assert.Equal(2, inner.StudentLookupCount)

        Dim client As New RecordingSharePointClient()
        client.Items.Add(Item("9000001", "Casey", "Patel"))
        Dim directory As New SharePointStudentDirectory(client, SharePointListNames.Default())
        Dim found = Await directory.FindByStudentIdAsync("9000001")
        Assert.Equal("Casey", found.GivenName)
        Assert.Equal("Students", client.LastList)
        Assert.Equal("fields/StudentID eq '9000001'", client.LastFilter)
        Assert.True(client.LastAllowFallback)
        Assert.Equal(1, client.QueryCount)
    End Function

    <Fact>
    Public Async Function Teacher_and_unit_lists_are_cached_for_the_session() As Task
        Dim inner As New InMemoryResitStore()
        Await inner.UpsertAsync(New TeacherRecord With {.FullName = "Riley Chen", .Email = "riley.chen@example.com"})
        Await inner.UpsertAsync(New UnitRecord With {.UnitCode = "UEECD0044", .UnitTitle = "Sample circuits"})
        Dim teachers As New CachingTeacherDirectory(inner)
        Assert.Equal("riley.chen@example.com", Await teachers.FindEmailAsync("Riley Chen"))
        Assert.Equal("riley.chen@example.com", Await teachers.FindEmailAsync("Riley Chen"))
        Assert.Equal(1, teachers.LoadCount)
        Dim units As New CachingUnitDirectory(inner)
        Assert.Equal("Sample circuits", Await units.FindTitleAsync("UEECD0044"))
        Assert.Equal("", Await units.FindTitleAsync("NO-SUCH"))
        Assert.Equal(1, units.LoadCount)
    End Function

    <Fact>
    Public Async Function SharePoint_booking_upsert_keeps_flags_by_not_sending_them() As Task
        Dim client As New RecordingSharePointClient()
        Dim existing As New ListItemRecord With {.Id = "item-1", .Fields = New FieldBag()}
        existing.Fields.SetString(SharePointColumns.StudentId, "9000004")
        existing.Fields.SetBool(SharePointColumns.EnergyspaceCreated, True)
        existing.Fields.SetBool(SharePointColumns.EnergyspaceAssessmentBooked, True)
        existing.Fields.SetString(SharePointColumns.Unit, "UEECD0044")
        existing.Fields.SetString(SharePointColumns.ResitDate, "2026-11-03")
        client.Items.Add(existing)
        Dim service As New SharePointBookingService(client, SharePointListNames.Default())
        Dim incoming = New ResitBooking With {
            .StudentId = "9000004",
            .UnitCode = "UETDRRF004",
            .ResitDate = New Date(2026, 11, 10),
            .EnergyspaceCreated = False,
            .EnergyspaceAssessmentBooked = False,
            .TeacherName = "Jordan Hale"
        }
        Dim result = Await service.UpsertAsync(incoming, preserveEnergyspaceFlags:=True)
        Assert.True(result.Booking.EnergyspaceCreated)
        Assert.True(result.Booking.EnergyspaceAssessmentBooked)
        Assert.Equal("UETDRRF004", result.Booking.UnitCode)
        Assert.Empty(client.LastFlags)
        Assert.Equal("UETDRRF004", client.LastText(SharePointColumns.Unit))
    End Function

    Private Shared Function Item(studentId As String, given As String, family As String) As ListItemRecord
        Dim record As New ListItemRecord With {.Id = "s-" & studentId, .Fields = New FieldBag()}
        record.Fields.SetString(SharePointColumns.StudentId, studentId)
        record.Fields.SetString(SharePointColumns.StudentGivenName, given)
        record.Fields.SetString(SharePointColumns.StudentFamilyName, family)
        Return record
    End Function
End Class

Public Class RecordingSharePointClient
    Implements ISharePointClient

    Public Property LastFilter As String
    Public Property LastList As String
    Public Property LastAllowFallback As Boolean
    Public Property QueryCount As Integer
    Public Property LastText As IDictionary(Of String, String)
    Public Property LastFlags As IDictionary(Of String, Boolean)
    Public Property Items As New List(Of ListItemRecord)

    Public Function GetSiteAsync() As Task(Of SharePointSite) Implements ISharePointClient.GetSiteAsync
        Return Task.FromResult(New SharePointSite With {.Id = "site", .DisplayName = "Test"})
    End Function

    Public Function ReadListsAsync(displayNames As IEnumerable(Of String)) As Task(Of IReadOnlyList(Of ExistingList)) Implements ISharePointClient.ReadListsAsync
        Return Task.FromResult(Of IReadOnlyList(Of ExistingList))(New List(Of ExistingList))
    End Function

    Public Function CreateListAsync(spec As ListSpec) As Task Implements ISharePointClient.CreateListAsync
        Return Task.CompletedTask
    End Function

    Public Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task Implements ISharePointClient.AddColumnAsync
        Return Task.CompletedTask
    End Function

    Public Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task Implements ISharePointClient.UpdateColumnAsync
        Return Task.CompletedTask
    End Function

    Public Function QueryItemsAsync(listDisplayName As String, filter As String, allowNonIndexedFallback As Boolean) As Task(Of IReadOnlyList(Of ListItemRecord)) Implements ISharePointClient.QueryItemsAsync
        QueryCount += 1
        LastList = listDisplayName
        LastFilter = filter
        LastAllowFallback = allowNonIndexedFallback
        If String.Equals(listDisplayName, "Students", StringComparison.OrdinalIgnoreCase) AndAlso String.IsNullOrWhiteSpace(filter) Then
            Throw New InvalidOperationException("Student lookup must use a Graph $filter.")
        End If
        Return Task.FromResult(Of IReadOnlyList(Of ListItemRecord))(Items.ToList())
    End Function

    Public Function CreateItemAsync(listDisplayName As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task(Of String) Implements ISharePointClient.CreateItemAsync
        LastText = textFields
        LastFlags = flags
        Return Task.FromResult("new-item")
    End Function

    Public Function UpdateItemAsync(listDisplayName As String, itemId As String, textFields As IDictionary(Of String, String), flags As IDictionary(Of String, Boolean)) As Task Implements ISharePointClient.UpdateItemAsync
        LastText = textFields
        LastFlags = If(flags, New Dictionary(Of String, Boolean))
        Return Task.CompletedTask
    End Function

    Public Function DeleteItemAsync(listDisplayName As String, itemId As String) As Task Implements ISharePointClient.DeleteItemAsync
        Return Task.CompletedTask
    End Function

    Public Function DownloadSignatureAsync(libraryDisplayName As String) As Task(Of Byte()) Implements ISharePointClient.DownloadSignatureAsync
        Return Task.FromResult(Of Byte())(Nothing)
    End Function
End Class
