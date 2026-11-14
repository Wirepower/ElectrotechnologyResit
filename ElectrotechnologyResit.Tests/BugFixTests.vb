Imports System.IO
Imports System.Threading
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
    Public Sub Edit_yes_saves_as_ticked_and_a_stored_flag_loads_as_yes_or_no()
        Assert.True(YesNoBoolean.Parse("Yes"))
        Assert.False(YesNoBoolean.Parse("No"))
        Assert.Equal("Yes", YesNoBoolean.ToYesNo(True))
        Assert.Equal("No", YesNoBoolean.ToYesNo(False))
        Assert.Equal("Yes", YesNoBoolean.ToYesNo(YesNoBoolean.Parse("True")))
        Assert.Equal("No", YesNoBoolean.ToYesNo(YesNoBoolean.Parse("False")))
        Assert.True(YesNoBoolean.Parse(YesNoBoolean.ToYesNo(True)))
        Assert.False(YesNoBoolean.Parse(YesNoBoolean.ToYesNo(False)))
    End Sub

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
        Dim removed = Await store.DeleteOnOrBeforeAsync(New Date(2026, 10, 31))
        Assert.Equal(2, removed)
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
        existing.Fields.SetBool(SharePointColumns.EnergyspaceClassCreated, True)
        existing.Fields.SetBool(SharePointColumns.EnergyspaceAssessmentBooked, True)
        existing.Fields.SetString(SharePointColumns.Unit, "UEECD0044")
        existing.Fields.SetString(SharePointColumns.ResitDate, "2026-11-03")
        client.Items.Add(existing)
        Dim teachers As New InMemoryResitStore()
        Await teachers.UpsertAsync(New TeacherRecord With {.FullName = "Jordan Hale", .Email = "jordan.hale@example.com"})
        Dim service As New SharePointBookingService(client, SharePointListNames.Default(), teachers)
        Dim incoming = New ResitBooking With {
            .StudentId = "9000004",
            .UnitCode = "UETDRRF004",
            .ResitDate = New Date(2026, 11, 10),
            .EnergyspaceCreated = False,
            .EnergyspaceAssessmentBooked = False,
            .TeacherName = "Jordan Hale",
            .TeacherEmail = "stale.previous.teacher@example.com"
        }
        Dim result = Await service.UpsertAsync(incoming, preserveEnergyspaceFlags:=True)
        Assert.True(result.Booking.EnergyspaceCreated)
        Assert.True(result.Booking.EnergyspaceAssessmentBooked)
        Assert.Equal("UETDRRF004", result.Booking.UnitCode)
        Assert.Equal("jordan.hale@example.com", result.Booking.TeacherEmail)
        Assert.Equal("jordan.hale@example.com", client.LastText(SharePointColumns.AllocatedTeacherEmail))
        Assert.Empty(client.LastFlags)
        Assert.Equal("UETDRRF004", client.LastText(SharePointColumns.Unit))
    End Function

    <Fact>
    Public Async Function Changing_the_teacher_replaces_the_email_from_the_teachers_list() As Task
        Dim store As New InMemoryResitStore()
        Await store.UpsertAsync(New TeacherRecord With {.FullName = "Jordan Hale", .Email = "jordan.hale@example.com"})
        Await store.UpsertAsync(New TeacherRecord With {.FullName = "Morgan Ellis", .Email = "morgan.ellis@example.com"})
        Dim night = New Date(2026, 11, 3)
        Await store.UpsertAsync(Booking("9000001", "UEECD0044", night), False)

        Dim changed = Booking("9000001", "UEECD0044", night)
        changed.TeacherName = "Morgan Ellis"
        changed.TeacherEmail = "jordan.hale@example.com"
        Dim updated = Await store.UpsertAsync(changed, False)
        Assert.Equal("Morgan Ellis", updated.Booking.TeacherName)
        Assert.Equal("morgan.ellis@example.com", updated.Booking.TeacherEmail)

        changed.TeacherName = "Jordan Hale"
        changed.TeacherEmail = "typed-by-hand@example.com"
        updated = Await store.UpsertAsync(changed, False)
        Assert.Equal("jordan.hale@example.com", updated.Booking.TeacherEmail)

        changed.TeacherName = "Nobody Known"
        changed.TeacherEmail = "morgan.ellis@example.com"
        updated = Await store.UpsertAsync(changed, False)
        Assert.Equal("", updated.Booking.TeacherEmail)
    End Function

    <Fact>
    Public Sub Grid_columns_match_the_live_resit_screen()
        Dim table = BookingTable.FromBookings(Nothing)
        Assert.Equal(BookingTable.Headers.Length, table.Columns.Count)
        For index = 0 To BookingTable.Headers.Length - 1
            Assert.Equal(BookingTable.Headers(index), table.Columns(index).ColumnName)
        Next
        Dim spec = SharePointListSchema.Build(SharePointListNames.Default()).Single(Function(list) list.DisplayName = "ResitBookings")
        Assert.Equal(BookingTable.Headers.Length, spec.Columns.Count)
        For Each header In BookingTable.Headers
            Assert.Contains(spec.Columns, Function(column) column.DisplayName = header)
        Next
        Assert.DoesNotContain(spec.Columns, Function(column) column.Name = "TeacherEmailSent" OrElse column.DisplayName = "TeacherEmailSent")
    End Sub

    <Fact>
    Public Sub Word_class_groups_come_from_class_created_ticks_only()
        Dim rules = AbbreviationRules.Uee30820Rules()
        Dim entries As New List(Of ClassGroupEntry) From {
            New ClassGroupEntry With {.UnitCode = "UEEEL0021", .ClassGroupSuffix = "CERTIII", .Rules = rules},
            New ClassGroupEntry With {.UnitCode = "UEECD0044", .ClassGroupSuffix = "CERTIII", .Rules = rules}
        }
        Dim block = ResitClassGroups.Build(entries, 2026)
        Assert.Equal("CD0044-RESIT-CERTIII-2026" & vbCrLf & "L0021-RESIT-CERTIII-2026", block)
        Assert.Equal(ResitClassGroups.EmptySelectionText, ResitClassGroups.Build(Array.Empty(Of ClassGroupEntry)(), 2026))
        Assert.Contains("Energyspace Class Created", ResitClassGroups.EmptySelectionText)
        Assert.False(ResitOutputVisibility.ShowLoginSheet(False))
        Assert.True(ResitOutputVisibility.ShowLoginSheet(True))
        Assert.True(ResitOutputVisibility.ShowExcel(False, True))
        Assert.True(ResitOutputVisibility.ShowExcel(True, False))
        Assert.False(ResitOutputVisibility.ShowExcel(False, False))
    End Sub

    <Theory>
    <InlineData("UEECD0019", "CD0019")>
    <InlineData("UEECD0020", "CD0020")>
    <InlineData("UEECD0046", "CD0046")>
    <InlineData("UEEEL0005", "L0005")>
    <InlineData("UEEEL0012", "L0012")>
    <InlineData("UEEEL0021", "L0021")>
    <InlineData("UEEEL0025", "L0025")>
    Public Sub Cert_iii_resit_units_abbreviate_for_the_login_sheet(code As String, abbreviated As String)
        Dim rules = AbbreviationRules.Uee30820Rules()
        Assert.Equal(abbreviated, AbbreviationRules.Abbreviate(code, rules))
        Dim entries As New List(Of ClassGroupEntry) From {
            New ClassGroupEntry With {.UnitCode = code, .ClassGroupSuffix = "CERTIII", .Rules = rules}
        }
        Dim block = ResitClassGroups.Build(entries, 2026)
        Assert.Equal(abbreviated & "-RESIT-CERTIII-2026", block)
    End Sub

    <Fact>
    Public Sub A_new_package_uses_its_own_suffix_and_does_not_need_a_hard_coded_unit_list()
        Dim rules = AbbreviationRules.Parse("prefix UEECD=CD")
        Dim entries As New List(Of ClassGroupEntry) From {
            New ClassGroupEntry With {.UnitCode = "UEECD0099", .ClassGroupSuffix = "CERTIV", .Rules = rules}
        }
        Dim block = ResitClassGroups.Build(entries, 2027)
        Assert.Equal("CD0099-RESIT-CERTIV-2027", block)
        Assert.Equal("UEEZZ0001", AbbreviationRules.Abbreviate("UEEZZ0001", rules))
    End Sub

    <Fact>
    Public Async Function A_booking_keeps_the_qualification_it_was_created_with() As Task
        Dim store As New InMemoryResitStore()
        Await store.UpsertAsync(New TeacherRecord With {.FullName = "Jordan Hale", .Email = "jordan.hale@example.com"})
        Dim created = Await store.UpsertAsync(New ResitBooking With {
            .StudentId = "9000001",
            .TeacherName = "Jordan Hale",
            .UnitCode = "UEECD0019",
            .QualificationCode = "UEE30820",
            .ResitDate = New Date(2026, 11, 14)
        }, False)
        Assert.Equal("UEE30820", created.Booking.QualificationCode)
        Dim changed = created.Booking.Clone()
        changed.QualificationCode = "UEE99999"
        changed.UnitCode = "UEECD0020"
        Dim updated = Await store.UpsertAsync(changed, False)
        Assert.Equal("UEE30820", updated.Booking.QualificationCode)
        Assert.Equal("UEECD0020", updated.Booking.UnitCode)
    End Function

    <Fact>
    Public Sub Only_an_active_qualification_is_offered_for_a_new_booking()
        Dim qualifications = New QualificationRecord() {
            New QualificationRecord With {.QualificationCode = "UEE30820", .Status = QualificationStatus.Superseded},
            New QualificationRecord With {.QualificationCode = "UEE40000", .Status = QualificationStatus.Draft},
            New QualificationRecord With {.QualificationCode = "UEE50000", .Status = QualificationStatus.Active}
        }
        Dim units = New UnitRecord() {
            New UnitRecord With {.UnitCode = "UEECD0019", .QualificationCode = "UEE30820"},
            New UnitRecord With {.UnitCode = "UEECD0019", .QualificationCode = "UEE50000"}
        }
        Assert.Equal("UEE50000", QualificationSelection.ForNewBooking("UEECD0019", units, qualifications))
        Assert.Equal("", QualificationSelection.ForNewBooking("UEECD0019", New UnitRecord() {units(0)}, qualifications))
    End Sub

    <Fact>
    Public Sub Office_wait_times_out_and_can_be_cancelled_without_blocking()
        Dim done As New ManualResetEventSlim(False)
        Assert.Equal(OfficeWaitKind.TimedOut, OfficeCallWait.Wait(done, TimeSpan.FromMilliseconds(30), CancellationToken.None, TimeSpan.FromMilliseconds(5)))
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Assert.Equal(OfficeWaitKind.Cancelled, OfficeCallWait.Wait(done, TimeSpan.FromSeconds(5), cancelled.Token, TimeSpan.FromMilliseconds(5)))
        End Using
        done.Set()
        Assert.Equal(OfficeWaitKind.Completed, OfficeCallWait.Wait(done, TimeSpan.FromSeconds(5), CancellationToken.None, TimeSpan.FromMilliseconds(5)))
    End Sub

    <Fact>
    Public Sub Office_step_log_drops_addresses_and_keeps_a_short_step()
        Assert.Equal("Email: message 1 of 4", OfficeStepLog.Sanitise("Email: message 1 of 4"))
        Dim cleaned = OfficeStepLog.Sanitise("Email message alex@example.com 9000001")
        Assert.DoesNotContain("@", cleaned)
        Assert.DoesNotContain("9000001", cleaned)
    End Sub

    <Fact>
    Public Sub Graph_batch_delete_names_item_ids_only()
        Dim body = GraphBatch.DeleteBody("site", "list", New String() {"11", "22"})
        Assert.Contains("DELETE", body)
        Assert.Contains("/items/11", body)
        Assert.DoesNotContain("9000001", body)
        Dim retry = GraphBatch.RetryIds("{""responses"":[{""id"":""11"",""status"":204},{""id"":""22"",""status"":429},{""id"":""33"",""status"":404}]}")
        Assert.Equal("22", Assert.Single(retry))
    End Sub

    <Fact>
    Public Sub Student_forms_drop_the_spare_width_when_lookup_labels_are_hidden()
        Assert.Equal(530, StudentFormWidth.ForLookupLabels(True, 530, 482, 16))
        Assert.Equal(498, StudentFormWidth.ForLookupLabels(False, 530, 482, 16))
        Assert.Equal(530, StudentFormWidth.ForLookupLabels(False, 530, 600, 16))
    End Sub

    <Fact>
    Public Sub Energyspace_lists_grow_to_every_unit_and_fold_into_two_columns_when_there_are_many()
        Assert.Equal(1, EnergyspacePanelLayout.ColumnCount(7))
        Assert.Equal(7, EnergyspacePanelLayout.RowCount(7))
        Assert.Equal(24 + (7 * 20), EnergyspacePanelLayout.ListHeight(7, 20, 24))
        Assert.Equal(2, EnergyspacePanelLayout.ColumnCount(9))
        Assert.Equal(5, EnergyspacePanelLayout.RowCount(9))
        Assert.Equal(24 + (7 * 20), EnergyspacePanelLayout.SharedListHeight(7, 3, 20, 24))
    End Sub

    <Fact>
    Public Async Function Blank_qualifications_are_filled_with_the_active_package_and_set_codes_stay() As Task
        Dim store As New InMemoryResitStore()
        Await store.UpsertAsync(New QualificationRecord With {.QualificationCode = "UEE30820", .Status = QualificationStatus.Active})
        Await store.UpsertAsync(New ResitBooking With {.StudentId = "9000001", .ResitDate = New Date(2026, 11, 1)}, False)
        Await store.UpsertAsync(New ResitBooking With {.StudentId = "9000002", .ResitDate = New Date(2026, 11, 1), .QualificationCode = "UEE50000"}, False)
        Dim code = QualificationCatalog.CodeForBlankBookings(Await store.GetQualificationsAsync())
        Assert.Equal("UEE30820", code)
        Assert.Equal(1, Await store.BackfillQualificationsAsync(code))
        Dim rows = Await store.GetBookingsAsync()
        Assert.Equal("UEE30820", rows.Single(Function(item) item.StudentId = "9000001").QualificationCode)
        Assert.Equal("UEE50000", rows.Single(Function(item) item.StudentId = "9000002").QualificationCode)
        Assert.Equal(0, Await store.BackfillQualificationsAsync(code))
    End Function

    <Fact>
    Public Async Function SharePoint_backfill_writes_only_the_qualification_on_blank_rows() As Task
        Dim client As New RecordingSharePointClient()
        client.Items.Add(BookingItem("11", "2026-11-01"))
        Dim kept = BookingItem("22", "2026-11-01")
        kept.Fields.SetString(SharePointColumns.QualificationCode, "UEE50000")
        client.Items.Add(kept)
        Dim service As New SharePointBookingService(client, SharePointListNames.Default(), New InMemoryResitStore())
        Assert.Equal(1, Await service.BackfillQualificationsAsync("UEE30820"))
        Assert.Equal("UEE30820", client.LastText(SharePointColumns.QualificationCode))
        Assert.False(client.LastText.ContainsKey(SharePointColumns.StudentId))
    End Function

    <Fact>
    Public Async Function SharePoint_purge_filters_the_indexed_resit_date_and_returns_the_count() As Task
        Dim client As New RecordingSharePointClient()
        client.Items.Add(BookingItem("11", "2026-10-01"))
        client.Items.Add(BookingItem("22", "2026-11-14"))
        Dim service As New SharePointBookingService(client, SharePointListNames.Default(), New InMemoryResitStore())
        Dim removed = Await service.DeleteOnOrBeforeAsync(New Date(2026, 10, 31))
        Assert.Equal(1, removed)
        Assert.Contains("ResitDate", client.LastFilter)
        Assert.Contains("11", client.DeletedIds)
        Assert.DoesNotContain("22", client.DeletedIds)
    End Function

    Private Shared Function Booking(studentId As String, unitCode As String, night As Date) As ResitBooking
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
            .ResitDate = night
        }
    End Function

    Private Shared Function BookingItem(itemId As String, resitDate As String) As ListItemRecord
        Dim record As New ListItemRecord With {.Id = itemId, .Fields = New FieldBag()}
        record.Fields.SetString(SharePointColumns.ResitDate, resitDate)
        record.Fields.SetString(SharePointColumns.StudentId, "9000001")
        Return record
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
    Public Property DeletedIds As New List(Of String)

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
        DeletedIds.Add(itemId)
        Return Task.CompletedTask
    End Function

    Public Function DeleteItemsAsync(listDisplayName As String, itemIds As IReadOnlyList(Of String)) As Task Implements ISharePointClient.DeleteItemsAsync
        If itemIds IsNot Nothing Then DeletedIds.AddRange(itemIds)
        Return Task.CompletedTask
    End Function

    Public Function DownloadSignatureAsync(libraryDisplayName As String) As Task(Of Byte()) Implements ISharePointClient.DownloadSignatureAsync
        Return Task.FromResult(Of Byte())(Nothing)
    End Function
End Class
