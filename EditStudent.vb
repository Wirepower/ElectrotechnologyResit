Imports ElectrotechnologyResit.Core

Public Class EditStudent
    Public Property MainFormDataGridView As DataGridView
    Private currentRowIndex As Integer = -1
    Private _storedQualification As String = ""
    Private _units As IReadOnlyList(Of UnitRecord) = Array.Empty(Of UnitRecord)()
    Private _qualifications As IReadOnlyList(Of QualificationRecord) = Array.Empty(Of QualificationRecord)()
    Private ReadOnly _qualificationText As New TextBox()
    Private _avatar As InitialAvatar
    Private _headerId As Label
    Private _headerName As Label
    Private _headerEmail As Label
    Private _recordLabel As Label
    Private _classCreated As ToggleSwitch
    Private _assessmentBooked As ToggleSwitch

    Public Sub New()
        InitializeComponent()
        BuildLayout()
    End Sub

    Private Sub BuildLayout()
        StudentDialogLayout.Prepare(Me)
        AcceptButton = Button1
        CancelButton = Button2
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDown
        ComboBox2.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox3.DropDownStyle = ComboBoxStyle.DropDownList
        _qualificationText.ReadOnly = True
        _qualificationText.TabStop = False
        _qualificationText.BackColor = SystemColors.Control
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dddd, d MMMM yyyy"
        lblTeacherEmail.Text = ""
        lblUnit.Text = ""

        _avatar = New InitialAvatar()
        _headerId = New Label With {.AutoSize = True, .ForeColor = Color.FromArgb(0, 90, 180), .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold), .Text = "Student ID"}
        _headerName = New Label With {.AutoSize = True, .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold), .Text = "Student name"}
        _headerEmail = New Label With {.AutoSize = True, .ForeColor = Color.DimGray, .Text = ""}
        _recordLabel = New Label With {.AutoSize = True, .Text = "Record 0 of 0", .Margin = New Padding(0, 10, 12, 0), .TextAlign = ContentAlignment.MiddleRight}
        Dim sampleBadge = StudentDialogLayout.Badge("SAMPLE DATA", Color.DarkRed, Color.White)
        sampleBadge.Visible = AppServices.IsSampleMode
        Button3.AutoSize = True
        Button4.AutoSize = True
        Dim nav As New FlowLayoutPanel With {.FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .AutoSize = True, .Dock = DockStyle.Right}
        nav.Controls.Add(sampleBadge)
        nav.Controls.Add(Button3)
        nav.Controls.Add(Button4)
        nav.Controls.Add(_recordLabel)

        Dim identity As New FlowLayoutPanel With {.FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoSize = True, .Margin = New Padding(12, 6, 0, 0)}
        identity.Controls.Add(_headerId)
        identity.Controls.Add(_headerName)
        identity.Controls.Add(_headerEmail)
        Dim header As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 3, .RowCount = 1, .BackColor = Color.FromArgb(245, 247, 250), .Padding = New Padding(8)}
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 72))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        header.Controls.Add(_avatar, 0, 0)
        header.Controls.Add(identity, 1, 0)
        header.Controls.Add(nav, 2, 0)

        Dim idRow As New TableLayoutPanel With {.ColumnCount = 2, .Dock = DockStyle.Top, .AutoSize = True, .Width = 300}
        idRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        idRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160))
        TextBox2.Dock = DockStyle.Top
        Button5.Dock = DockStyle.Top
        Button5.Height = 27
        idRow.Controls.Add(TextBox2, 0, 0)
        idRow.Controls.Add(Button5, 1, 0)
        Dim studentColumn As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .Padding = New Padding(8)}
        studentColumn.Controls.Add(SectionTitle("STUDENT"))
        studentColumn.Controls.Add(Labeled("Student ID", idRow))
        studentColumn.Controls.Add(StudentDialogLayout.Field("First name", TextBox3, Nothing))
        studentColumn.Controls.Add(StudentDialogLayout.Field("Last name", TextBox4, Nothing))
        studentColumn.Controls.Add(StudentDialogLayout.Field("Student email", TextBox5, Nothing))
        For Each control As Control In studentColumn.Controls
            control.Width = 300
        Next

        Dim bookingColumn As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .Padding = New Padding(8)}
        bookingColumn.Controls.Add(SectionTitle("RESIT BOOKING"))
        bookingColumn.Controls.Add(StudentDialogLayout.Field("Qualification", _qualificationText, Nothing))
        bookingColumn.Controls.Add(StudentDialogLayout.Field("Unit", ComboBox2, lblUnit))
        bookingColumn.Controls.Add(StudentDialogLayout.Field("Teacher", ComboBox1, lblTeacherEmail))
        Dim pair As New TableLayoutPanel With {.ColumnCount = 2, .Dock = DockStyle.Top, .AutoSize = True, .Width = 320}
        pair.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        pair.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        pair.Controls.Add(StudentDialogLayout.Field("Attempt number", ComboBox3, Nothing), 0, 0)
        pair.Controls.Add(StudentDialogLayout.Field("Blockgroup", txtBlockgroup, Nothing), 1, 0)
        bookingColumn.Controls.Add(pair)
        bookingColumn.Controls.Add(StudentDialogLayout.Field("Booked assessment date", DateTimePicker1, Nothing))
        For Each control As Control In bookingColumn.Controls
            If control.Width < 280 Then control.Width = 320
        Next

        Dim columns As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        columns.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        columns.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        columns.Controls.Add(studentColumn, 0, 0)
        columns.Controls.Add(bookingColumn, 1, 0)

        _classCreated = New ToggleSwitch()
        _assessmentBooked = New ToggleSwitch()
        Dim energy As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .Padding = New Padding(12, 8, 12, 8)}
        energy.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        energy.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        energy.Controls.Add(SwitchRow(_classCreated, "Class created", "Has the Energyspace class been created?"), 0, 0)
        energy.Controls.Add(SwitchRow(_assessmentBooked, "Assessment booked", "Has the assessment been booked on Energyspace?"), 1, 0)
        Dim energyBox As New GroupBox With {.Text = "ENERGYSPACE", .Dock = DockStyle.Fill, .Padding = New Padding(8)}
        energyBox.Controls.Add(energy)

        StudentDialogLayout.StyleClose(Button2)
        StudentDialogLayout.StylePrimary(Button1)
        Dim footer As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft, .Padding = New Padding(0, 8, 0, 0)}
        footer.Controls.Add(Button1)
        footer.Controls.Add(Button2)

        Dim root As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 4, .Padding = New Padding(12)}
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 84))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 96))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
        root.Controls.Add(header, 0, 0)
        root.Controls.Add(columns, 0, 1)
        root.Controls.Add(energyBox, 0, 2)
        root.Controls.Add(footer, 0, 3)
        Controls.Add(root)
    End Sub

    Private Shared Function SectionTitle(text As String) As Label
        Return New Label With {.Text = text, .AutoSize = True, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold), .Margin = New Padding(0, 0, 0, 4)}
    End Function

    Private Shared Function Labeled(caption As String, editor As Control) As Control
        Dim stack As New TableLayoutPanel With {.ColumnCount = 1, .AutoSize = True, .Dock = DockStyle.Top, .Margin = New Padding(0, 0, 0, 4)}
        stack.Controls.Add(New Label With {.Text = caption, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)}, 0, 0)
        editor.Margin = New Padding(0)
        stack.Controls.Add(editor, 0, 1)
        Return stack
    End Function

    Private Shared Function SwitchRow(toggle As ToggleSwitch, title As String, hint As String) As Control
        Dim row As New TableLayoutPanel With {.ColumnCount = 2, .Dock = DockStyle.Fill}
        row.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 60))
        row.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        toggle.Margin = New Padding(0, 8, 8, 0)
        row.Controls.Add(toggle, 0, 0)
        Dim text As New FlowLayoutPanel With {.FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoSize = True}
        text.Controls.Add(New Label With {.Text = title, .AutoSize = True, .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)})
        text.Controls.Add(New Label With {.Text = hint, .AutoSize = True, .ForeColor = Color.DimGray, .MaximumSize = New Size(260, 0)})
        row.Controls.Add(text, 1, 0)
        Return row
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Async Sub EditStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Await PopulateLookups()
        If MainFormDataGridView IsNot Nothing AndAlso MainFormDataGridView.Rows.Count > 0 Then
            currentRowIndex = 0
            Await LoadDataFromDataGridView()
        Else
            Button3.Enabled = False
            Button4.Enabled = False
            _recordLabel.Text = "Record 0 of 0"
        End If
        StudentDialogLayout.ApplyLookupNotes(lblTeacherEmail, lblUnit)
    End Sub

    Private Async Function PopulateLookups() As Task
        Try
            _units = Await AppServices.Units.GetAllAsync()
            _qualifications = Await AppServices.Qualifications.GetAllAsync()
            ComboBox1.Items.Clear()
            For Each teacher In Await AppServices.Teachers.GetAllAsync()
                ComboBox1.Items.Add(teacher.FullName)
            Next
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Private Sub FillUnitsForStoredQualification()
        ComboBox2.Items.Clear()
        For Each unit In _units
            If String.Equals(If(unit.QualificationCode, ""), _storedQualification, StringComparison.OrdinalIgnoreCase) Then
                ComboBox2.Items.Add(New UnitChoice With {.Code = unit.UnitCode, .Title = unit.UnitTitle})
            End If
        Next
    End Sub

    Private Sub SelectUnit(code As String)
        For index = 0 To ComboBox2.Items.Count - 1
            Dim item = TryCast(ComboBox2.Items(index), UnitChoice)
            If item IsNot Nothing AndAlso String.Equals(item.Code, code, StringComparison.OrdinalIgnoreCase) Then
                ComboBox2.SelectedIndex = index
                Return
            End If
        Next
        If Not String.IsNullOrWhiteSpace(code) Then
            ComboBox2.Items.Add(New UnitChoice With {.Code = code, .Title = ""})
            ComboBox2.SelectedIndex = ComboBox2.Items.Count - 1
        End If
    End Sub

    Private Function SelectedUnit() As UnitChoice
        Return TryCast(ComboBox2.SelectedItem, UnitChoice)
    End Function

    Private Async Function LoadDataFromDataGridView() As Task
        If MainFormDataGridView Is Nothing Then Return
        If currentRowIndex < 0 OrElse currentRowIndex >= MainFormDataGridView.Rows.Count Then Return
        Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)
        If row Is Nothing OrElse row.IsNewRow Then Return
        TextBox2.Text = CellText(row, "Student ID")
        TextBox3.Text = CellText(row, "Student Firstname")
        TextBox4.Text = CellText(row, "Surname")
        TextBox5.Text = CellText(row, "Email")
        SelectTeacher(CellText(row, "AllocatedTeacher"))
        _storedQualification = CellText(row, "QualificationCode")
        _qualificationText.Text = QualificationDisplay(_storedQualification)
        FillUnitsForStoredQualification()
        SelectUnit(CellText(row, "Unit"))
        ComboBox3.Text = CellText(row, "AttemptNo")
        _classCreated.Checked = CellBool(row, "EnergyspaceClassCreated")
        _assessmentBooked.Checked = CellBool(row, "EnergyspaceAssessmentBooked")
        txtBlockgroup.Text = CellText(row, "Blockgroup")
        Dim unit = SelectedUnit()
        lblUnit.Text = If(unit Is Nothing OrElse String.IsNullOrEmpty(unit.Title), CellText(row, "Unit Name"), unit.Title)
        Await ShowTeacherEmailAsync(ComboBox1.Text)

        Dim resitDateValue = row.Cells("Resit date").Value
        If resitDateValue IsNot Nothing AndAlso TypeOf resitDateValue IsNot DBNull Then
            DateTimePicker1.Value = Convert.ToDateTime(resitDateValue)
        End If
        Dim count = MainFormDataGridView.Rows.Count
        _recordLabel.Text = "Record " & (currentRowIndex + 1).ToString() & " of " & count.ToString()
        Button3.Enabled = currentRowIndex > 0
        Button4.Enabled = currentRowIndex < count - 1
        RefreshHeader()
    End Function

    Private Sub SelectTeacher(name As String)
        Dim index = ComboBox1.FindStringExact(name)
        If index >= 0 Then
            ComboBox1.SelectedIndex = index
        Else
            ComboBox1.SelectedIndex = -1
            ComboBox1.Text = name
        End If
    End Sub

    Private Function QualificationDisplay(code As String) As String
        If String.IsNullOrWhiteSpace(code) Then Return "(not recorded)"
        For Each qualification In _qualifications
            If String.Equals(qualification.QualificationCode, code, StringComparison.OrdinalIgnoreCase) Then
                Return New PackageChoice With {.Code = qualification.QualificationCode, .Title = qualification.Title}.ToString()
            End If
        Next
        Return code
    End Function

    Private Shared Function CellText(row As DataGridViewRow, columnName As String) As String
        Dim value = row.Cells(columnName).Value
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return ""
        Return value.ToString()
    End Function

    Private Shared Function CellBool(row As DataGridViewRow, columnName As String) As Boolean
        Dim value = row.Cells(columnName).Value
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return False
        If TypeOf value Is Boolean Then Return CBool(value)
        Dim parsed As Boolean
        If Boolean.TryParse(value.ToString(), parsed) Then Return parsed
        Return False
    End Function

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If currentRowIndex > 0 Then
            currentRowIndex -= 1
            Await LoadDataFromDataGridView()
        End If
    End Sub

    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MainFormDataGridView Is Nothing Then Return
        If currentRowIndex < MainFormDataGridView.Rows.Count - 1 Then
            currentRowIndex += 1
            Await LoadDataFromDataGridView()
        End If
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MainFormDataGridView Is Nothing Then Return
        If currentRowIndex < 0 OrElse currentRowIndex >= MainFormDataGridView.Rows.Count Then Return
        Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)
        If row.IsNewRow Then Return
        Dim unit = SelectedUnit()
        If unit Is Nothing Then
            MessageBox.Show(Me, "Please select a unit.", "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            Dim teacherEmail = Await AppServices.Teachers.FindEmailAsync(ComboBox1.Text)
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(teacherEmail), "Email not found", teacherEmail)
            Dim unitTitle = If(unit.Title, "")
            If String.IsNullOrEmpty(unitTitle) Then unitTitle = lblUnit.Text
            If Not String.IsNullOrEmpty(unitTitle) Then lblUnit.Text = unitTitle

            row.Cells("Student ID").Value = TextBox2.Text
            row.Cells("Student Firstname").Value = TextBox3.Text
            row.Cells("Surname").Value = TextBox4.Text
            row.Cells("Email").Value = TextBox5.Text
            row.Cells("AllocatedTeacher").Value = ComboBox1.Text
            row.Cells("AllocatedTeacher email").Value = teacherEmail
            row.Cells("Unit").Value = unit.Code
            row.Cells("Unit Name").Value = unitTitle
            row.Cells("AttemptNo").Value = ComboBox3.Text
            row.Cells("EnergyspaceClassCreated").Value = _classCreated.Checked
            row.Cells("EnergyspaceAssessmentBooked").Value = _assessmentBooked.Checked
            row.Cells("Resit date").Value = DateTimePicker1.Value.Date
            row.Cells("Blockgroup").Value = txtBlockgroup.Text

            Dim booking As New ResitBooking With {
                .StudentId = TextBox2.Text.Trim(),
                .FirstName = TextBox3.Text,
                .Surname = TextBox4.Text,
                .Email = TextBox5.Text,
                .TeacherName = ComboBox1.Text,
                .TeacherEmail = teacherEmail,
                .UnitCode = unit.Code,
                .UnitName = unitTitle,
                .AttemptNo = ComboBox3.Text,
                .EnergyspaceCreated = _classCreated.Checked,
                .EnergyspaceAssessmentBooked = _assessmentBooked.Checked,
                .ResitDate = DateTimePicker1.Value.Date,
                .BlockGroup = txtBlockgroup.Text
            }
            Await AppServices.Bookings.UpsertAsync(booking, preserveEnergyspaceFlags:=False)
            Dim mainForm As Main = TryCast(Application.OpenForms("Main"), Main)
            If mainForm IsNot Nothing Then Await mainForm.UpdateDataGridView()
            MessageBox.Show(Me, "Data updated successfully.", "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(Me, "Error updating data: " & ex.Message, "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Async Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem Is Nothing AndAlso String.IsNullOrWhiteSpace(ComboBox1.Text) Then
            lblTeacherEmail.Text = ""
            Return
        End If
        Await ShowTeacherEmailAsync(ComboBox1.Text)
    End Sub

    Private Async Function ShowTeacherEmailAsync(teacherFullName As String) As Task
        Try
            Dim email = Await AppServices.Teachers.FindEmailAsync(teacherFullName)
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(email), "Email not found", email)
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim unit = SelectedUnit()
        If unit IsNot Nothing AndAlso Not String.IsNullOrEmpty(unit.Title) Then lblUnit.Text = unit.Title
    End Sub

    Private Sub IdentityChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged
        RefreshHeader()
    End Sub

    Private Sub RefreshHeader()
        If _headerId Is Nothing Then Return
        Dim id = TextBox2.Text.Trim()
        Dim given = TextBox3.Text.Trim()
        Dim family = TextBox4.Text.Trim()
        _headerId.Text = If(id.Length = 0, "Student ID", id)
        Dim name = (given & " " & family).Trim()
        _headerName.Text = If(name.Length = 0, "Student name", name)
        _headerEmail.Text = TextBox5.Text.Trim()
        _avatar.Initials = StudentDialogLayout.Initials(given, family)
    End Sub

    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim studentID As String = TextBox2.Text.Trim()
        If String.IsNullOrWhiteSpace(studentID) Then
            MessageBox.Show(Me, "Please enter Student ID.", "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            Dim student = Await AppServices.Students.FindByStudentIdAsync(studentID)
            If student IsNot Nothing Then
                TextBox3.Text = student.GivenName
                TextBox4.Text = student.FamilyName
                TextBox5.Text = student.PersonalEmail
                txtBlockgroup.Text = student.BlockGroupCode
                _classCreated.Checked = False
                _assessmentBooked.Checked = False
            Else
                MessageBox.Show(Me, "No record found for the provided Student ID.", "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                txtBlockgroup.Text = ""
            End If
            RefreshHeader()
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Edit Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
End Class
