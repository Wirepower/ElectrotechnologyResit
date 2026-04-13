'Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports Microsoft.Win32
Imports Microsoft.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports OfficeOpenXml
Imports OfficeOpenXml.Style



Public Class Main
    Dim dataTable As New DataTable()
    Private connection As SqlConnection
    Dim adapter As New SqlDataAdapter()
    Sub New()
        ' Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        InitializeComponent()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NewStudent.Show()
        NewStudent.DateTimePicker1.Value = DateTimePicker1.Value
        'UpdateDataGridView()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Create an instance of EditStudent form
        Dim editStudentForm As New EditStudent()

        ' Set the MainFormDataGridView property of the EditStudent form to the DataGridView in the MainForm
        editStudentForm.MainFormDataGridView = DataGridView1

        ' Show the EditStudent form
        editStudentForm.Show()
        UpdateDataGridView()
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SQLCon.InitializeStatusLabel(statusLabel)
        'connection = SQLCon.GetConnection()
        'SQLCon.OpenConnection(connection)
        Dim result As DialogResult = MessageBox.Show("Are you Connected to the VPN", "VPN Connection", MessageBoxButtons.YesNo)
        Dim SQLConnection As DialogResult = MessageBox.Show("Are you having SQL connection issues", "SQL Connection", MessageBoxButtons.YesNo)
        If SQLConnection = vbYes Then
            Settings.Show()
            Me.Hide()
            Exit Sub
        End If
        If result = vbYes Then


            ' Connection string to your SQL Server database
            Dim connectionString As String = SQLCon.connectionString

            ' SQL query to retrieve data from the table
            Dim query As String = "SELECT * FROM ElectrotechnologyReports.dbo.ElectricalResit"

            ' Create a SqlConnection
            Using connection As New SqlConnection(connectionString)
                Try
                    ' Open the connection
                    connection.Open()

                    ' Create a SqlCommand to execute the query
                    Using command As New SqlCommand(query, connection)
                        ' Create a DataTable to hold the data
                        Dim dataTable As New DataTable()

                        ' Create a SqlDataAdapter to fill the DataTable
                        Using adapter As New SqlDataAdapter(command)
                            ' Fill the DataTable
                            adapter.Fill(dataTable)
                        End Using

                        ' Bind the DataTable to the DataGridView
                        DataGridView1.DataSource = dataTable
                    End Using
                Catch ex As System.Exception
                    ' Handle any exceptions
                    MessageBox.Show("Error: " & ex.Message)
                End Try
            End Using
            ' Add event handlers to DataGridView events
            AddHandler DataGridView1.RowsAdded, AddressOf DataGridView1_RowsAdded
            AddHandler DataGridView1.RowsRemoved, AddressOf DataGridView1_RowsRemoved
            AddHandler DataGridView1.CellValueChanged, AddressOf DataGridView1_CellValueChanged
            UpdateListView(DateTimePicker1.Value)
            UpdateListView1(DateTimePicker1.Value)
            FilterData()
        Else
            MessageBox.Show("Closing Application, Connect to VPN and re-open Application", "Exiting Application")
            Settings.Show()
            'System.Environment.Exit(0)
        End If
    End Sub

    Public Sub UpdateDataGridView()
        ' Construct the SQL query to fetch data from the ElectricalResit table
        Dim query As String = "SELECT * FROM ElectrotechnologyReports.dbo.ElectricalResit"

        ' Create a DataTable to store the retrieved data
        Dim dataTable As New DataTable()

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlDataAdapter to fetch the data
                Using adapter As New SqlDataAdapter(query, connection)
                    ' Fill the DataTable with data from the database
                    adapter.Fill(dataTable)
                End Using
            Catch ex As System.Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using

        ' Bind the DataTable to the DataGridView
        DataGridView1.DataSource = dataTable
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If Not CheckBox2.Checked Then
            CheckBox1.Checked = True
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Not CheckBox1.Checked Then
            CheckBox2.Checked = True
        End If
    End Sub
    Private Sub DeleteRecordFromDatabase(studentID As String)
        ' Assuming you have a SqlConnection object named 'connection' already defined
        Using connection As New SqlConnection(SQLCon.connectionString)
            ' Open the connection
            connection.Open()

            ' Define the SQL command to delete the record
            Dim sql As String = "DELETE FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Student ID] = @StudentID"

            ' Create a SqlCommand object
            Using command As New SqlCommand(sql, connection)
                ' Add parameter for the student ID
                command.Parameters.AddWithValue("@StudentID", studentID)

                ' Execute the command
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
            DeleteRecordFromDatabase(studentID)
            DataGridView1.Rows.Remove(selectedRow)
        Catch ex As Exception
            MessageBox.Show("Could not delete from the database: " & ex.Message, "Delete failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim selectedDate As Date = DateTimePicker1.Value.Date
        If MessageBox.Show("Delete ALL resit bookings for " & selectedDate.ToString("D") & "? This cannot be undone.", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If
        DeleteRecordsFromDatabase(selectedDate)
        MessageBox.Show("All data for " & selectedDate.ToString("D") & " has been deleted.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
        UpdateDataGridView()
    End Sub
    Private Sub DeleteRecordsFromDatabase(selectedDate As Date)
        Using connection As New SqlConnection(SQLCon.connectionString)
            connection.Open()

            Dim sql As String = "DELETE FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE CAST([Resit date] AS date) = CAST(@SelectedDate AS date)"

            Using command As New SqlCommand(sql, connection)
                command.Parameters.Add("@SelectedDate", SqlDbType.Date).Value = selectedDate
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Function RetrieveImageDataFromDatabase() As Byte()
        Dim imageData As Byte() = Nothing

        ' Your SQL query to retrieve the image data
        Dim query As String = "SELECT TOP 1 [Email Signature Image] FROM ElectrotechnologyReports.dbo.EmailSettings"


        ' Define your SQL connection string
        Dim connectionString As String = SQLCon.connectionString

        ' Create a SqlConnection object
        Using connection As New SqlConnection(connectionString)
            ' Open the connection
            connection.Open()

            ' Create a SqlCommand object with your query and connection
            Using command As New SqlCommand(query, connection)
                Dim scalar = command.ExecuteScalar()
                If scalar IsNot Nothing AndAlso Not Convert.IsDBNull(scalar) Then
                    imageData = DirectCast(scalar, Byte())
                End If
            End Using
        End Using

        Return imageData
    End Function
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim selectedDate As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        Dim imageData As Byte() = RetrieveImageDataFromDatabase()
        Dim query As String = "SELECT * FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE CONVERT(date, [Resit date]) = @SelectedDate"

        Dim outlookApp As Object = OutlookInterop.TryCreateOutlookApplication()
        If outlookApp Is Nothing Then
            MessageBox.Show("Microsoft Outlook could not be started. Install Outlook (desktop) and try again.", "Outlook required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using connection As New SqlConnection(SQLCon.connectionString)
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@SelectedDate", selectedDate)

                    connection.Open()

                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                        Dim emailAddress As String = reader("Student Email").ToString()
                            Dim Blockgroup As String = reader("Blockgroup").ToString()
                            Dim resitDate As Date = CType(reader("Resit date"), Date)
                            Dim StudentFirstname As String = reader("Student Firstname").ToString()
                            Dim StudentSurname As String = reader("Student Surname").ToString()
                            Dim Student_ID As String = reader("Student ID").ToString()
                            Dim Unit As String = reader("Unit").ToString()
                            Dim UnitName As String = reader("Unit Name").ToString()
                            Dim AllocatedTeacher As String = reader("AllocatedTeacher").ToString()
                            Dim AllocatedTeacherEmail As String = reader("AllocatedTeacherEmail").ToString()
                            Dim formattedResitDate As String = resitDate.ToString("dd")

                            Select Case resitDate.Day
                                Case 1, 21, 31
                                    formattedResitDate &= "st"
                                Case 2, 22
                                    formattedResitDate &= "nd"
                                Case 3, 23
                                    formattedResitDate &= "rd"
                                Case Else
                                    formattedResitDate &= "th"
                            End Select

                            formattedResitDate &= " " & resitDate.ToString("MMMM, yyyy")

                            Dim body As String = "Dear " & AllocatedTeacher & "," & vbCrLf & vbCrLf & "<br>" &
                            "This is a reminder about an upcoming resit on " & formattedResitDate & ".<br>" &
                            vbCrLf & vbCrLf &
                            "Please see below for your applicable student:" & vbCrLf & vbCrLf & "<br><br>" &
                            "Student ID: " & Student_ID & vbCrLf & "<br>" &
                            "Student Firstname: " & StudentFirstname & vbCrLf & "<br>" &
                            "Student Surname: " & StudentSurname & vbCrLf & "<br>" &
                            "Blockgroup: " & Blockgroup & vbCrLf & "<br>" &
                            "Allocated Teacher: " & AllocatedTeacher & vbCrLf & "<br>" &
                            "Unit: " & Unit & "- " & UnitName & vbCrLf & vbCrLf & "<br><br>" &
                            "Please monitor this student and update their results upon resit completion,<br><br>" & vbCrLf & vbCrLf &
                            "Thank you," & vbCrLf & "<br>" &
                            "Electrotechnology.Admin Team"

                            Dim imgHtml As String = ""
                            If imageData IsNot Nothing AndAlso imageData.Length > 0 Then
                                imgHtml = "<br><img src='data:image/jpeg;base64," & Convert.ToBase64String(imageData) & "' width='90%'>"
                            End If
                            Dim html As String = "<html><body><font color='black'>" & body & "</font></body>" & imgHtml & "</html>"
                            Dim subject As String = "One of your students has been booked into a resit night on " & formattedResitDate
                            OutlookInterop.CreateDisplayOrSendMail(outlookApp, AllocatedTeacherEmail, subject, html, "Electrotechnology.admin@vu.edu.au", CheckBox1.Checked, CheckBox2.Checked)
                        End While
                    End Using
                End Using
            End Using
        Finally
            Marshal.FinalReleaseComObject(outlookApp)
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        FilterData()
        UpdateListView(DateTimePicker1.Value)
        UpdateListView1(DateTimePicker1.Value)
    End Sub
    Public Sub FilterData()
        Dim selectedDate As DateTime = DateTimePicker1.Value
        Dim dataTable = TryCast(DataGridView1.DataSource, DataTable)
        If dataTable Is Nothing Then Return
        dataTable.DefaultView.RowFilter = "[Resit date] = '" & selectedDate.ToString("yyyy-MM-dd") & "'"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Show all data in the DataGridView
        UpdateDataGridView()
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

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If MessageBox.Show("Delete all resit records older than 14 days? This cannot be undone.", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If
        Dim query As String = "DELETE FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Resit date] <= DATEADD(day, -14, GETDATE())"

        Using connection As New SqlConnection(SQLCon.connectionString)
            Using command As New SqlCommand(query, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        End Using
        UpdateDataGridView()
        MessageBox.Show("All resit records older than 14 days have been deleted.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        UpdateDataGridView()
        UpdateListView(DateTimePicker1.Value)
        UpdateListView1(DateTimePicker1.Value)
    End Sub
    Private Sub UpdateListView(selectedDate As Date)
        ' Clear the ListView before updating
        ListView1.Items.Clear()

        ' Format the selectedDate to match the 'yyyy-MM-dd' format
        Dim formattedDate As String = selectedDate.ToString("yyyy-MM-dd")

        ' Connection string for your SQL Server database
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve unique values from the "Unit" column
        Dim query As String = "SELECT DISTINCT Unit FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Resit date] = @SelectedDate"

        Try
            ' Open a connection to the database
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                ' Execute the SQL query
                Using command As New SqlCommand(query, connection)
                    ' Add parameter for the formatted date
                    command.Parameters.AddWithValue("@SelectedDate", formattedDate)

                    ' Execute the query and read the results
                    Using reader As SqlDataReader = command.ExecuteReader()
                        ' Iterate through the result set and add unique values to the ListView
                        While reader.Read()
                            ' Add items to the ListView without checking the checkboxes
                            ListView1.Items.Add(reader("Unit").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As System.Exception
            ' Handle any errors that occur during database access
            MessageBox.Show("Error retrieving data from the database: " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        UpdateListView(DateTimePicker1.Value)
        UpdateListView1(DateTimePicker1.Value)
    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        UpdateListView(DateTimePicker1.Value)
        UpdateListView1(DateTimePicker1.Value)
    End Sub

    Private Sub DataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        UpdateListView(DateTimePicker1.Value)
        UpdateListView1(DateTimePicker1.Value)
    End Sub
    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        ' Get the checked item in the ListView
        Dim checkedItem = e.Item

        ' Get the text of the checked item
        Dim unit = checkedItem.Text

        ' Get the date from the DatePicker
        Dim resitDate = DateTimePicker1.Value.Date

        ' Check if the item is checked
        If checkedItem.Checked Then
            ' Update SQL table
            UpdateDataInSQL(resitDate, unit)
        End If
        'UpdateDataGridView()
    End Sub

    Private Sub UpdateDataInSQL(resitDate As Date, unit As String)
        ' Connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.ElectricalResit SET [EnergyspaceCreated] = @Checked WHERE [Resit date] = @ResitDate AND Unit = @Unit"

        Try
            Using connection As New SqlConnection(connectionString)
                ' Open the connection
                connection.Open()

                ' Create command
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@ResitDate", resitDate)
                    command.Parameters.AddWithValue("@Unit", unit)
                    ' You should replace "YourColumnName" with the name of the column you want to update in your SQL table
                    command.Parameters.AddWithValue("@Checked", True) ' Or whatever value you want to set when the checkbox is checked

                    ' Execute the query
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As System.Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub
    '---------------------------------------
    Private Sub UpdateListView1(selectedDate As Date)
        ' Clear the ListView before updating
        ListView2.Items.Clear()

        ' Format the selectedDate to match the 'yyyy-MM-dd' format
        Dim formattedDate As String = selectedDate.ToString("yyyy-MM-dd")

        ' Connection string for your SQL Server database
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query to retrieve unique values from the "Unit" column
        Dim query As String = "SELECT DISTINCT Unit FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Resit date] = @SelectedDate"

        Try
            ' Open a connection to the database
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                ' Execute the SQL query
                Using command As New SqlCommand(query, connection)
                    ' Add parameter for the formatted date
                    command.Parameters.AddWithValue("@SelectedDate", formattedDate)

                    ' Execute the query and read the results
                    Using reader As SqlDataReader = command.ExecuteReader()
                        ' Iterate through the result set and add unique values to the ListView
                        While reader.Read()
                            ' Add items to the ListView without checking the checkboxes
                            ListView2.Items.Add(reader("Unit").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As System.Exception
            ' Handle any errors that occur during database access
            MessageBox.Show("Error retrieving data from the database: " & ex.Message)
        End Try
        UpdateResitExportButtonsVisibility()
    End Sub

    Private Function ListView2HasAnyChecked() As Boolean
        For Each item As ListViewItem In ListView2.Items
            If item.Checked Then Return True
        Next
        Return False
    End Function

    ''' <summary>Show Export to Excel and Resit Login Details only when at least one ListView2 row is ticked.</summary>
    Private Sub UpdateResitExportButtonsVisibility()
        Dim show = ListView2HasAnyChecked()
        Button9.Visible = show
        Button12.Visible = show
    End Sub

    Private Sub ListView2_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView2.ItemChecked
        ' Get the checked item in the ListView
        Dim checkedItem = e.Item

        ' Get the text of the checked item
        Dim unit = checkedItem.Text

        ' Get the date from the DatePicker
        Dim resitDate = DateTimePicker1.Value.Date

        ' Check if the item is checked
        If checkedItem.Checked Then
            ' Update SQL table
            UpdateDataInSQL2(resitDate, unit)
        End If
        UpdateResitExportButtonsVisibility()
        'UpdateDataGridView()
    End Sub
    Private Sub UpdateDataInSQL2(resitDate As Date, unit As String)
        ' Connection string
        Dim connectionString As String = SQLCon.connectionString

        ' SQL query
        Dim query As String = "UPDATE ElectrotechnologyReports.dbo.ElectricalResit SET [EnergyspaceAssessmentBooked] = @Checked WHERE [Resit date] = @ResitDate AND Unit = @Unit"

        Try
            Using connection As New SqlConnection(connectionString)
                ' Open the connection
                connection.Open()

                ' Create command
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@ResitDate", resitDate)
                    command.Parameters.AddWithValue("@Unit", unit)
                    ' You should replace "YourColumnName" with the name of the column you want to update in your SQL table
                    command.Parameters.AddWithValue("@Checked", True) ' Or whatever value you want to set when the checkbox is checked

                    ' Execute the query
                    command.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As System.Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Settings.Show()
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

End Class
