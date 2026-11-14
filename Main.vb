'Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports ElectrotechnologyResit.Core
Imports Microsoft.VisualBasic
Imports OfficeOpenXml
Imports OfficeOpenXml.Style



Public Class Main
    Private _ready As Boolean
    Private _bindingGrid As Boolean
    Private _loadingUnitLists As Boolean
    Private _emailModeGuard As Boolean
    Sub New()
        ' Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        InitializeComponent()
        Label5.Text = AppVersion.DisplayText
        Text = "Electrotechnology Resit — " & AppVersion.DisplayText
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NewStudent.Show()
        NewStudent.DateTimePicker1.Value = DateTimePicker1.Value
        'UpdateDataGridView()
    End Sub

    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim editStudentForm As New EditStudent()
        editStudentForm.MainFormDataGridView = DataGridView1
        editStudentForm.Show()
        Await UpdateDataGridView()
    End Sub

    Private Async Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If AppServices.NeedsWizard() Then
            Using wizard As New SetupWizard()
                If wizard.ShowDialog(Me) <> DialogResult.OK Then
                    Settings.Show()
                    Me.Hide()
                    Return
                End If
            End Using
        End If

        Try
            AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
        Catch ex As Exception
            MessageBox.Show("Could not sign in or open the SharePoint site. " & ex.Message & vbCrLf & vbCrLf & "Open Settings?", "SharePoint sign-in", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Settings.Show()
            Me.Hide()
            Return
        End Try

        Dim access = Await AppServices.CheckSiteAccessAsync()
        If Not access.Ok Then
            Dim openSettings = MessageBox.Show(access.Message & vbCrLf & vbCrLf & "Open Settings?", "SharePoint sign-in", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If openSettings = DialogResult.Yes Then
                Settings.Show()
                Me.Hide()
                Return
            End If
        Else
            Await OfferSchemaRepairAsync()
        End If

        Try
            Me.Cursor = Cursors.WaitCursor
            Await UpdateDataGridView()
            FilterData()
            ApplyPullFromBookingsButtonVisibility()
            _ready = True
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Async Function OfferSchemaRepairAsync() As Task
        Try
            Dim report = Await AppServices.Provisioner.VerifyAsync()
            If report.IsSatisfied Then Return
            Dim fix = MessageBox.Show(
                "The SharePoint site is missing something this app needs:" & vbCrLf & vbCrLf & report.Summary() & vbCrLf & vbCrLf &
                "Repair it now? Lists and columns that are already there are left as they are. Nothing is deleted.",
                "SharePoint setup",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)
            If fix <> DialogResult.Yes Then Return
            report = Await AppServices.Provisioner.EnsureAsync()
            If Not report.IsSatisfied Then
                MessageBox.Show("Some items are still missing:" & vbCrLf & vbCrLf & report.Summary(), "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show("Could not check the SharePoint lists. " & ex.Message, "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Public Async Function UpdateDataGridView() As Task
        _bindingGrid = True
        Try
            Dim rows = Await AppServices.Bookings.GetAllAsync()
            DataGridView1.DataSource = BookingTable.FromBookings(rows)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            _bindingGrid = False
        End Try
        Await RefreshUnitListsAsync()
    End Function

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If _emailModeGuard Then Return
        _emailModeGuard = True
        Dim mode = EmailSendMode.AfterSendChanged(CheckBox2.Checked)
        CheckBox1.Checked = mode.DisplayManually
        CheckBox2.Checked = mode.SendImmediately
        _emailModeGuard = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If _emailModeGuard Then Return
        _emailModeGuard = True
        Dim mode = EmailSendMode.AfterDisplayChanged(CheckBox1.Checked)
        CheckBox1.Checked = mode.DisplayManually
        CheckBox2.Checked = mode.SendImmediately
        _emailModeGuard = False
    End Sub

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a row to delete.")
            Return
        End If
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim studentID As String = selectedRow.Cells("Student ID").Value.ToString()
        If MessageBox.Show("Delete this student record permanently?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If
        Try
            Await AppServices.Bookings.DeleteByStudentIdAsync(studentID)
            DataGridView1.Rows.Remove(selectedRow)
            Await RefreshUnitListsAsync()
        Catch ex As Exception
            MessageBox.Show("Could not delete from SharePoint: " & ex.Message, "Delete failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim selectedDate As Date = DateTimePicker1.Value.Date
        If MessageBox.Show("Delete ALL resit bookings for " & selectedDate.ToString("D") & "? This cannot be undone.", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If
        Try
            Await AppServices.Bookings.DeleteByDateAsync(selectedDate)
        Catch ex As Exception
            MessageBox.Show("Could not delete from SharePoint: " & ex.Message, "Delete failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        MessageBox.Show("All data for " & selectedDate.ToString("D") & " has been deleted.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Await UpdateDataGridView()
    End Sub

    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim imageData As Byte() = Nothing
        Try
            imageData = Await AppServices.Signature.GetCurrentImageAsync()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            Return
        End Try

        Dim outlookApp As Object = OutlookInterop.TryCreateOutlookApplication()
        If outlookApp Is Nothing Then
            MessageBox.Show("Microsoft Outlook could not be started. Install Outlook (desktop) and try again.", "Outlook required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim bookings = Await AppServices.Bookings.GetByDateAsync(DateTimePicker1.Value.Date)
            For Each booking In bookings
                Dim formattedResitDate = EnglishOrdinal.FormatResitNight(booking.ResitDate)
                Dim body As String = "Dear " & booking.TeacherName & "," & vbCrLf & vbCrLf & "<br>" &
                "This is a reminder about an upcoming resit on " & formattedResitDate & ".<br>" &
                vbCrLf & vbCrLf &
                "Please see below for your applicable student:" & vbCrLf & vbCrLf & "<br><br>" &
                "Student ID: " & booking.StudentId & vbCrLf & "<br>" &
                "Student Firstname: " & booking.FirstName & vbCrLf & "<br>" &
                "Student Surname: " & booking.Surname & vbCrLf & "<br>" &
                "Blockgroup: " & booking.BlockGroup & vbCrLf & "<br>" &
                "Allocated Teacher: " & booking.TeacherName & vbCrLf & "<br>" &
                "Unit: " & booking.UnitCode & "- " & booking.UnitName & vbCrLf & vbCrLf & "<br><br>" &
                "Please monitor this student and update their results upon resit completion,<br><br>" & vbCrLf & vbCrLf &
                "Thank you," & vbCrLf & "<br>" &
                "Electrotechnology.Admin Team"

                Dim imgHtml As String = ""
                If imageData IsNot Nothing AndAlso imageData.Length > 0 Then
                    imgHtml = "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'>"
                End If
                Dim html As String = "<html><body><font color='black'>" & body & "</font></body>" & imgHtml & "</html>"
                Dim subject As String = "One of your students has been booked into a resit night on " & formattedResitDate
                OutlookInterop.CreateDisplayOrSendMail(outlookApp, booking.TeacherEmail, subject, html, "Electrotechnology.admin@vu.edu.au", CheckBox1.Checked, CheckBox2.Checked)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            Marshal.FinalReleaseComObject(outlookApp)
        End Try
    End Sub

    Private Async Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Not _ready Then Return
        FilterData()
        Await RefreshUnitListsAsync()
    End Sub
    Public Sub FilterData()
        Dim selectedDate As Date = DateTimePicker1.Value.Date
        Dim dataTable = TryCast(DataGridView1.DataSource, DataTable)
        If dataTable Is Nothing Then Return
        Dim startText = selectedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        Dim endText = selectedDate.AddDays(1).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        dataTable.DefaultView.RowFilter = "[Resit date] >= #" & startText & "# AND [Resit date] < #" & endText & "#"
    End Sub

    Private Async Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Show all data in the DataGridView
        Await UpdateDataGridView()
    End Sub
    Private Sub ExportToExcel(dataGridView As DataGridView, exportDate As DateTime, columnsToHide As List(Of String), columnWidths As Dictionary(Of String, Integer))
        Dim dateString As String = exportDate.ToString("dd-MM-yyyy")
        Dim fileName As String = $"ElectricalResit_{dateString}.xlsx"

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        saveFileDialog.Title = "Save Excel File"
        saveFileDialog.FileName = fileName

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Using package As New ExcelPackage()
                Dim worksheet = package.Workbook.Worksheets.Add("Sheet1")

                ' Create a mapping between column headers and their respective indices
                Dim headerIndexMap As New Dictionary(Of String, Integer)()
                For j As Integer = 0 To dataGridView.Columns.Count - 1
                    headerIndexMap.Add(dataGridView.Columns(j).HeaderText, j)
                Next

                ' Define the desired order of column headers
                Dim desiredOrderHeaders As String() = {"Student ID", "Student Firstname", "Student Surname", "Student Email", "Unit", "Unit Name", "Blockgroup", "AllocatedTeacher", "AllocatedTeacherEmail", "AttemptNo", "Resit date", "EnergyspaceCreated", "EnergyspaceAssessmentBooked", "status", "Attendance", "Mark", "Pass/Fail"} ' Adjust with your desired order

                ' Sort the mapping based on the desired order of column headers
                Dim sortedMap = headerIndexMap.OrderBy(Function(x) Array.IndexOf(desiredOrderHeaders, x.Key))

                ' Add headers to the worksheet based on the sorted mapping
                Dim columnIndex As Integer = 1
                For Each kvp In sortedMap
                    If Not columnsToHide.Contains(kvp.Key) Then ' Check if the column should not be hidden
                        worksheet.Cells(1, columnIndex).Value = kvp.Key
                        If columnWidths.ContainsKey(kvp.Key) Then ' Check if width is specified for this column
                            worksheet.Column(columnIndex).Width = columnWidths(kvp.Key)
                        Else
                            worksheet.Column(columnIndex).Width = 20 ' Set default width to 20
                        End If
                        ' Apply formatting to the headers (thick border)
                        worksheet.Cells(1, columnIndex).Style.Border.BorderAround(ExcelBorderStyle.Thick)
                        worksheet.Cells(1, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(1, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        columnIndex += 1
                    End If
                Next

                ' Add the three additional columns to the worksheet
                For Each additionalHeader In {"Attendance", "Mark %", "Pass/Fail"}
                    If Not columnsToHide.Contains(additionalHeader) Then ' Check if the column should not be hidden
                        worksheet.Cells(1, columnIndex).Value = additionalHeader
                        worksheet.Column(columnIndex).Width = 12 ' Set default width to 20
                        ' Apply formatting to the added header
                        worksheet.Cells(1, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(1, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(1, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.LightBlue)

                        ' Apply formatting to the headers (thick border)
                        worksheet.Cells(1, columnIndex).Style.Border.BorderAround(ExcelBorderStyle.Thick)
                        worksheet.Cells(1, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(1, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        columnIndex += 1
                    End If
                Next

                ' Populate the worksheet with DataGridView data, skipping hidden columns
                For i As Integer = 0 To dataGridView.Rows.Count - 2
                    columnIndex = 1
                    For Each kvp In sortedMap
                        If Not columnsToHide.Contains(kvp.Key) Then ' Check if the column should not be hidden
                            Dim cellValue = dataGridView.Rows(i).Cells(kvp.Value).Value
                            ' Check if the column is "Resit Date" and format it as a long date
                            If kvp.Key = "Resit date" AndAlso TypeOf cellValue Is DateTime Then
                                worksheet.Cells(i + 2, columnIndex).Value = CType(cellValue, DateTime).ToString("D")
                            Else
                                worksheet.Cells(i + 2, columnIndex).Value = cellValue
                            End If

                            ' Set normal borders for the cell
                            worksheet.Cells(i + 2, columnIndex).Style.Border.Top.Style = ExcelBorderStyle.Thin
                            worksheet.Cells(i + 2, columnIndex).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            worksheet.Cells(i + 2, columnIndex).Style.Border.Left.Style = ExcelBorderStyle.Thin
                            worksheet.Cells(i + 2, columnIndex).Style.Border.Right.Style = ExcelBorderStyle.Thin
                            ' Set thin borders for column L (12th column)
                            Dim columnLRange = worksheet.Cells(2, 12, dataGridView.Rows.Count, 12)
                            columnLRange.Style.Border.Left.Style = ExcelBorderStyle.Thin
                            columnLRange.Style.Border.Right.Style = ExcelBorderStyle.Thin
                            columnLRange.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            columnLRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin

                            ' Set thin borders for column M (13th column)
                            Dim columnMRange = worksheet.Cells(2, 13, dataGridView.Rows.Count, 13)
                            columnMRange.Style.Border.Left.Style = ExcelBorderStyle.Thin
                            columnMRange.Style.Border.Right.Style = ExcelBorderStyle.Thin
                            columnMRange.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            columnMRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin

                            ' Set thin borders for column N (14th column)
                            Dim columnNRange = worksheet.Cells(2, 14, dataGridView.Rows.Count, 14)
                            columnNRange.Style.Border.Left.Style = ExcelBorderStyle.Thin
                            columnNRange.Style.Border.Right.Style = ExcelBorderStyle.Thin
                            columnNRange.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            columnNRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin

                            columnIndex += 1
                        End If
                    Next

                Next
                ' Format the first row with a bigger size, light blue background, and bold black text
                For j As Integer = 1 To columnIndex - 1
                    worksheet.Cells(1, j).Style.Font.Bold = True
                    worksheet.Cells(1, j).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(1, j).Style.Fill.BackgroundColor.SetColor(Color.LightBlue)

                Next
                worksheet.Cells(31, 1).Value = "NOTES:"
                ' Apply formatting to cell A31
                worksheet.Cells("A31:N31").Style.Font.Bold = True
                worksheet.Cells("A31:N31").Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells("A31:N31").Style.Fill.BackgroundColor.SetColor(Color.LightBlue)
                worksheet.Cells("A31:N31").Style.Border.Bottom.Style = ExcelBorderStyle.Thick

                Dim lastColumnIndex As Integer = worksheet.Dimension.End.Column
                Dim bottomRowRange = worksheet.Cells(40, 1, 40, lastColumnIndex)
                bottomRowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thick
                Dim lastRowIndex As Integer = worksheet.Dimension.End.Row
                Dim rightColumnRange = worksheet.Cells(1, lastColumnIndex, lastRowIndex, lastColumnIndex)
                rightColumnRange.Style.Border.Right.Style = ExcelBorderStyle.Thick
                Dim leftColumnRange = worksheet.Cells(1, 1, lastRowIndex, 1)
                leftColumnRange.Style.Border.Left.Style = ExcelBorderStyle.Thick
                Dim topRowRange = worksheet.Cells(1, 1, 1, lastColumnIndex)
                topRowRange.Style.Border.Top.Style = ExcelBorderStyle.Thick

                ' Hide columns after column N
                Dim lastColumnIndexToHide As Integer = 14 ' Assuming column N is column 14
                For columnIndexToHide As Integer = lastColumnIndexToHide + 1 To worksheet.Dimension.End.Column
                    worksheet.Column(columnIndexToHide).Hidden = True
                Next

                ' Hide rows after row 40
                Dim lastRowIndexToHide As Integer = 40
                For rowIndexToHide As Integer = lastRowIndexToHide + 1 To worksheet.Dimension.End.Row
                    worksheet.Row(rowIndexToHide).Hidden = True
                Next

                Dim rangeWithThinBorders = worksheet.Cells("A2:N30")
                rangeWithThinBorders.Style.Border.Top.Style = ExcelBorderStyle.Thin
                rangeWithThinBorders.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                rangeWithThinBorders.Style.Border.Left.Style = ExcelBorderStyle.Thin
                rangeWithThinBorders.Style.Border.Right.Style = ExcelBorderStyle.Thin

                ' Define the range from A1 to N30
                Dim rangeWithThinBorders1 = worksheet.Cells("A1:N30")
                ' Set text size to 12 for rows 2 to 41
                worksheet.Cells("A2:N41").Style.Font.Size = 12
                ' Set horizontal and vertical alignment to center for each cell in the range
                For Each cell In rangeWithThinBorders1
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next
                ' Shift existing headers down by one row
                worksheet.InsertRow(1, 1)
                ' Set the background color to a darker blue
                Dim BackGroundcolor As Color = Color.FromArgb(0, 102, 204) ' You can adjust the RGB values as needed
                worksheet.Cells("A1:N1").Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells("A1:N1").Style.Fill.BackgroundColor.SetColor(BackGroundcolor)
                Dim TextColor As Color = Color.FromArgb(204, 0, 0) ' You can adjust the RGB values as needed
                worksheet.Cells("A1").Style.Font.Color.SetColor(TextColor)
                worksheet.Cells("A1:N1").Merge = True
                worksheet.Cells("A1").Value = "RESIT NIGHT  -  " & DateTimePicker1.Text
                worksheet.Cells("A1").Style.Font.Bold = True
                worksheet.Cells("A1").Style.Font.Size = 48
                worksheet.Cells("A1").Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells("A1").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                ' Save the Excel package to the selected file path
                package.SaveAs(New FileInfo(saveFileDialog.FileName))
            End Using
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim columnsToHide As New List(Of String) From {"status", "EnergyspaceCreated", "EnergyspaceAssessmentBooked"}
        Dim columnWidths As New Dictionary(Of String, Integer) From {
            {"Student ID", 16},
            {"Student Firstname", 22},
            {"Student Surname", 22},
            {"Student Email", 30},
            {"Unit", 26},
            {"Unit Name", 75},
            {"Blockgroup", 28},
            {"AllocatedTeacher", 20},
            {"AllocatedTeacherEmail", 30},
            {"AttemptNo", 14},
            {"Resit date", 26},
            {"Attendance", 14},
            {"Mark %", 14},
            {"Pass/Fail", 26}
        }
        ' Add more columns as needed


        ' Export to Excel
        ExportToExcel(DataGridView1, DateTimePicker1.Value, columnsToHide, columnWidths)
    End Sub

    Private Async Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If MessageBox.Show("Delete all resit records older than 14 days? This cannot be undone.", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If
        Try
            Await AppServices.Bookings.DeleteOnOrBeforeAsync(MelbourneCalendar.Today().AddDays(-14))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            Return
        End Try
        Await UpdateDataGridView()
        MessageBox.Show("All resit records older than 14 days have been deleted.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Async Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Reset All reloads the grid and both unit lists and does not reapply the date filter.
        ' Frank asked to keep this until he confirms a change.
        Await UpdateDataGridView()
    End Sub
    Private Async Function RefreshUnitListsAsync() As Task
        If AppServices.Bookings Is Nothing Then Return
        Dim flags As IReadOnlyList(Of UnitNightFlags)
        Try
            flags = Await AppServices.Bookings.GetUnitFlagsAsync(DateTimePicker1.Value.Date)
        Catch ex As Exception
            MessageBox.Show("Error retrieving data from SharePoint: " & ex.Message)
            Return
        End Try
        _loadingUnitLists = True
        Try
            ListView1.Items.Clear()
            ListView2.Items.Clear()
            For Each unit In flags
                Dim created As New ListViewItem(unit.UnitCode)
                Dim booked As New ListViewItem(unit.UnitCode)
                ListView1.Items.Add(created)
                ListView2.Items.Add(booked)
                created.Checked = unit.ClassCreated
                booked.Checked = unit.AssessmentBooked
            Next
        Finally
            _loadingUnitLists = False
        End Try
        UpdateResitExportButtonsVisibility()
    End Function

    Private Async Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        If _loadingUnitLists OrElse _bindingGrid Then Return
        Dim unit = e.Item.Text
        Dim night = DateTimePicker1.Value.Date
        Try
            Await AppServices.Bookings.SetEnergyspaceCreatedAsync(night, unit, e.Item.Checked)
            UpdateGridFlag(unit, "EnergyspaceCreated", e.Item.Checked)
        Catch ex As Exception
            RevertCheck(e.Item)
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
        UpdateResitExportButtonsVisibility()
    End Sub

    Private Function ListViewHasAnyChecked(listView As ListView) As Boolean
        For Each item As ListViewItem In listView.Items
            If item.Checked Then Return True
        Next
        Return False
    End Function

    ''' <summary>Show Export to Excel and Resit Login Details when ListView1 or ListView2 has a ticked row.</summary>
    Private Sub UpdateResitExportButtonsVisibility()
        Dim show = ListViewHasAnyChecked(ListView1) OrElse ListViewHasAnyChecked(ListView2)
        Button9.Visible = show
        Button12.Visible = show
    End Sub

    Private Async Sub ListView2_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView2.ItemChecked
        If _loadingUnitLists OrElse _bindingGrid Then Return
        Dim unit = e.Item.Text
        Dim night = DateTimePicker1.Value.Date
        Try
            Await AppServices.Bookings.SetEnergyspaceAssessmentBookedAsync(night, unit, e.Item.Checked)
            UpdateGridFlag(unit, "EnergyspaceAssessmentBooked", e.Item.Checked)
        Catch ex As Exception
            RevertCheck(e.Item)
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
        UpdateResitExportButtonsVisibility()
    End Sub

    Private Sub RevertCheck(item As ListViewItem)
        _loadingUnitLists = True
        Try
            item.Checked = Not item.Checked
        Finally
            _loadingUnitLists = False
        End Try
    End Sub

    Private Sub UpdateGridFlag(unitCode As String, columnName As String, value As Boolean)
        Dim table = TryCast(DataGridView1.DataSource, DataTable)
        If table Is Nothing Then Return
        Dim night = DateTimePicker1.Value.Date
        For Each row As DataRow In table.Rows
            If row.RowState = DataRowState.Deleted Then Continue For
            Dim unit = Convert.ToString(row("Unit"))
            If row.IsNull("Resit date") Then Continue For
            Dim resitDay = Convert.ToDateTime(row("Resit date")).Date
            If resitDay = night AndAlso String.Equals(unit, unitCode, StringComparison.OrdinalIgnoreCase) Then
                row(columnName) = value
            End If
        Next
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Settings.Show()
    End Sub

    Private Sub Main_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        ApplyPullFromBookingsButtonVisibility()
    End Sub

    Private Sub ApplyPullFromBookingsButtonVisibility()
        Button13.Visible = My.Settings.ShowPullFromBookingsButton
    End Sub

    Private Sub TryOpenSavedResitLoginDocument(filePath As String)
        Try
            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show("Could not open the file: " & ex.Message, "Resit login sheet", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>Builds lines like CD0044-RESIT-CERTIII-2026 from ListView2 — only items with the checkbox ticked.</summary>
    Private Function BuildResitClassGroupLines(listView As ListView, certYear As Integer) As String
        Dim lines As New List(Of String)
        For Each item As ListViewItem In listView.Items
            If Not item.Checked Then Continue For
            Dim u = item.Text.Trim()
            If u.Length > 0 Then
                Dim abbrev = UnitCodeAbbreviation.ToAbbreviatedClassCode(u)
                lines.Add(abbrev & "-RESIT-CERTIII-" & certYear.ToString())
            End If
        Next
        lines.Sort(StringComparer.OrdinalIgnoreCase)
        If lines.Count = 0 Then
            Return "(no class groups selected — tick the units under Energyspace Assessment Booked to include them here)"
        End If
        Return String.Join(vbCrLf, lines)
    End Function

    ''' <summary>
    ''' Tokens: {{RESIT_NIGHT_DATE}} {{CLASS_GROUPS}} {{DAILY_ENERGYSPACE_PASSWORD}}
    ''' {{PASSWORD_LINE}} / {{PC_LOGIN_PASSWORD}} = PC monthly (Settings). {{MONTHLY_PASSWORD}} removed (replaced empty if present in old templates).
    ''' {{USERNAME_LINE}} optional.
    ''' </summary>
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim dailyPwd As String = Interaction.InputBox(
            "Today's daily Energyspace password (changes each day).",
            "Daily Energyspace password",
            "")
        If String.IsNullOrWhiteSpace(dailyPwd) Then Return

        Dim pcLoginMonthly As String = My.Settings.PcLoginMonthlyPassword

        Dim templatePath As String = My.Settings.ResitLoginWordTemplatePath
        If Not File.Exists(templatePath) Then
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Word documents (*.docx)|*.docx|All files|*.*"
                ofd.Title = "Select the Resit login Word template"
                If ofd.ShowDialog(Me) <> DialogResult.OK Then Return
                templatePath = ofd.FileName
                My.Settings.ResitLoginWordTemplatePath = templatePath
                My.Settings.Save()
            End Using
        End If

        Dim sfd As New SaveFileDialog()
        sfd.Filter = "Word document (*.docx)|*.docx"
        sfd.FileName = "Resit login details - " & DateTimePicker1.Value.ToString("yyyy-MM-dd") & ".docx"
        sfd.Title = "Save generated Resit login sheet"
        If sfd.ShowDialog(Me) <> DialogResult.OK Then Return

        Dim certYear = DateTimePicker1.Value.Year
        Dim classGroupsBlock = BuildResitClassGroupLines(ListView2, certYear)
        Dim resitNightLong As String = DateTimePicker1.Value.ToString("D")

        Dim wordApp As Object = WordDocInterop.TryCreateWordApplication()
        If wordApp Is Nothing Then
            MessageBox.Show("Microsoft Word is required. Install desktop Word and try again.", "Word required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        wordApp.Visible = False
        Dim doc As Object = Nothing
        Try
            doc = wordApp.Documents.Open(templatePath)
            WordDocInterop.ReplaceAllInDocument(doc, "{{RESIT_NIGHT_DATE}}", resitNightLong)
            WordDocInterop.ReplaceAllInDocument(doc, "{{MONTHLY_PASSWORD}}", "")
            WordDocInterop.ReplaceAllInDocument(doc, "{{CLASS_GROUPS}}", classGroupsBlock)
            WordDocInterop.ReplaceAllInDocument(doc, "{{DAILY_ENERGYSPACE_PASSWORD}}", dailyPwd)
            WordDocInterop.ReplaceAllInDocument(doc, "{{USERNAME_LINE}}", "Username: student")
            WordDocInterop.ReplaceAllInDocument(doc, "{{PASSWORD_LINE}}", "Password: " & pcLoginMonthly)
            WordDocInterop.ReplaceAllInDocument(doc, "{{PC_LOGIN_PASSWORD}}", pcLoginMonthly)
            ApplyLeaVuStamp(doc)

            Try
                doc.SaveAs2(sfd.FileName)
            Catch
                doc.SaveAs(sfd.FileName)
            End Try

            doc.Close(0)
            doc = Nothing
            Dim savedPath = sfd.FileName
            If My.Settings.ResitLoginOpenDocumentAfterSave Then
                MessageBox.Show("Document saved:" & vbCrLf & savedPath, "Resit login sheet", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TryOpenSavedResitLoginDocument(savedPath)
            Else
                Dim openNow = MessageBox.Show(
                    "Document saved:" & vbCrLf & savedPath & vbCrLf & vbCrLf & "Open it now?",
                    "Resit login sheet",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
                If openNow = DialogResult.Yes Then
                    TryOpenSavedResitLoginDocument(savedPath)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Word could not complete the document: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If doc IsNot Nothing Then
                Try
                    doc.Close(0)
                Catch
                End Try
                Marshal.FinalReleaseComObject(doc)
            End If
            WordDocInterop.QuitWord(wordApp)
        End Try
    End Sub

    ''' <summary>
    ''' Fills legacy LEA stamp placeholders/bookmarks so older templates still render correctly.
    ''' </summary>
    Private Sub ApplyLeaVuStamp(doc As Object)
        Const vuStampText As String = "VU Stamp"
        WordDocInterop.TryReplaceBookmarkText(doc, "stamp1", vuStampText)
        WordDocInterop.TryReplaceBookmarkText(doc, "stamp2", vuStampText)

        ' Legacy text tokens seen in earlier LEA templates.
        WordDocInterop.ReplaceAllInDocument(doc, "{{stamp1}}", vuStampText)
        WordDocInterop.ReplaceAllInDocument(doc, "{{stamp2}}", vuStampText)
        WordDocInterop.ReplaceAllInDocument(doc, "{{STAMP1}}", vuStampText)
        WordDocInterop.ReplaceAllInDocument(doc, "{{STAMP2}}", vuStampText)
    End Sub

    Private Async Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Cursor = Cursors.WaitCursor
        Button13.Enabled = False
        Try
            ' TODO: Live Microsoft Bookings belongs to VU Timetable and needs Frank Offer's sign-off.
            Dim importer As IBookingImporter = New BookingsDataImporter(Me, AppServices.Bookings)
            Dim msg = Await importer.PullAsync(DateTimePicker1.Value.Date)
            MessageBox.Show(msg, "Pull from Bookings", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Await UpdateDataGridView()
            FilterData()
            UpdateResitExportButtonsVisibility()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Pull from Bookings", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Button13.Enabled = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class
