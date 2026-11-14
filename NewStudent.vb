Imports ElectrotechnologyResit.Core

Public Class NewStudent
    Private ReadOnly _qualification As New ComboBox()
    Private _units As IReadOnlyList(Of UnitRecord) = Array.Empty(Of UnitRecord)()
    Private _qualifications As IReadOnlyList(Of QualificationRecord) = Array.Empty(Of QualificationRecord)()
    Private _avatar As InitialAvatar
    Private _headerId As Label
    Private _headerName As Label
    Private _headerEmail As Label
    Private _statusBadge As Label
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
        _qualification.DropDownStyle = ComboBoxStyle.DropDownList
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dddd, d MMMM yyyy"
        lblTeacherEmail.Text = ""
        lblUnit.Text = ""

        _avatar = New InitialAvatar()
        _headerId = New Label With {.AutoSize = True, .ForeColor = Color.FromArgb(0, 90, 180), .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold), .Text = "Student ID"}
        _headerName = New Label With {.AutoSize = True, .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold), .Text = "Student name"}
        _headerEmail = New Label With {.AutoSize = True, .ForeColor = Color.DimGray, .Text = ""}
        _statusBadge = StudentDialogLayout.Badge("New booking", Color.FromArgb(232, 240, 250), Color.FromArgb(0, 90, 180))
        Dim sampleBadge = StudentDialogLayout.Badge("SAMPLE DATA", Color.DarkRed, Color.White)
        sampleBadge.Visible = AppServices.IsSampleMode

        Dim identity As New FlowLayoutPanel With {.FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoSize = True, .Margin = New Padding(12, 6, 0, 0)}
        identity.Controls.Add(_headerId)
        identity.Controls.Add(_headerName)
        identity.Controls.Add(_headerEmail)
        Dim badges As New FlowLayoutPanel With {.FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .AutoSize = True, .Dock = DockStyle.Right}
        badges.Controls.Add(sampleBadge)
        badges.Controls.Add(_statusBadge)
        Dim header As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 3, .RowCount = 1, .BackColor = Color.FromArgb(245, 247, 250), .Padding = New Padding(8)}
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 72))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        header.Controls.Add(_avatar, 0, 0)
        header.Controls.Add(identity, 1, 0)
        header.Controls.Add(badges, 2, 0)

        Dim idRow As New TableLayoutPanel With {.ColumnCount = 2, .Dock = DockStyle.Top, .AutoSize = True}
        idRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        idRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160))
        txtStudentID.Dock = DockStyle.Top
        Button3.Dock = DockStyle.Top
        Button3.Height = 27
        idRow.Controls.Add(txtStudentID, 0, 0)
        idRow.Controls.Add(Button3, 1, 0)
        Dim studentColumn As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .Padding = New Padding(8)}
        studentColumn.Controls.Add(SectionTitle("STUDENT"))
        studentColumn.Controls.Add(Labeled("Student ID", idRow))
        studentColumn.Controls.Add(StudentDialogLayout.Field("First name", txtGivenName, Nothing))
        studentColumn.Controls.Add(StudentDialogLayout.Field("Last name", txtFamilyName, Nothing))
        studentColumn.Controls.Add(StudentDialogLayout.Field("Student email", txtPersonalEmail, Nothing))
        For Each control As Control In studentColumn.Controls
            control.Width = 300
        Next
        idRow.Width = 300

        Dim bookingColumn As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .Padding = New Padding(8)}
        bookingColumn.Controls.Add(SectionTitle("RESIT BOOKING"))
        bookingColumn.Controls.Add(StudentDialogLayout.Field("Qualification", _qualification, Nothing))
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
        AddHandler _qualification.SelectedIndexChanged, Sub() FillUnits()
        RefreshHeader()
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

    Private Async Sub NewStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Await PopulateTeacherComboBox()
        Await PopulateUnitComboBox()
        StudentDialogLayout.ApplyLookupNotes(lblTeacherEmail, lblUnit)
        RefreshHeader()
    End Sub

    Private Async Function PopulateUnitComboBox() As Task
        Try
            _units = Await AppServices.Units.GetAllAsync()
            _qualifications = Await AppServices.Qualifications.GetAllAsync()
            _qualification.Items.Clear()
            For Each qualification In _qualifications
                If QualificationStatus.OfferedForNewBookings(qualification.Status) Then
                    _qualification.Items.Add(New PackageChoice With {.Code = qualification.QualificationCode, .Title = qualification.Title})
                End If
            Next
            If _qualification.Items.Count > 0 Then _qualification.SelectedIndex = 0 Else FillUnits()
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "New Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Private Sub FillUnits()
        Dim package = SelectedQualificationCode()
        ComboBox2.Items.Clear()
        For Each unit In _units
            If String.Equals(If(unit.QualificationCode, ""), package, StringComparison.OrdinalIgnoreCase) Then
                ComboBox2.Items.Add(New UnitChoice With {.Code = unit.UnitCode, .Title = unit.UnitTitle})
            End If
        Next
        lblUnit.Text = ""
    End Sub

    Private Function SelectedQualificationCode() As String
        Dim item = TryCast(_qualification.SelectedItem, PackageChoice)
        If item Is Nothing Then Return ""
        Return If(item.Code, "")
    End Function

    Private Function SelectedUnit() As UnitChoice
        Return TryCast(ComboBox2.SelectedItem, UnitChoice)
    End Function

    Private Async Function PopulateTeacherComboBox() As Task
        Try
            ComboBox1.Items.Clear()
            For Each teacher In Await AppServices.Teachers.GetAllAsync()
                ComboBox1.Items.Add(teacher.FullName)
            Next
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "New Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim studentID As String = txtStudentID.Text.Trim()
        If String.IsNullOrWhiteSpace(studentID) Then
            MessageBox.Show(Me, "Please enter Student ID.", "New Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Try
            Dim student = Await AppServices.Students.FindByStudentIdAsync(studentID)
            If student IsNot Nothing Then
                txtGivenName.Text = student.GivenName
                txtFamilyName.Text = student.FamilyName
                txtPersonalEmail.Text = student.PersonalEmail
                txtBlockgroup.Text = student.BlockGroupCode
                _statusBadge.Text = "filled from Search Student ID"
            Else
                MessageBox.Show(Me, "No record found for the provided Student ID.", "New Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtGivenName.Clear()
                txtFamilyName.Clear()
                txtPersonalEmail.Clear()
                txtBlockgroup.Text = ""
                _statusBadge.Text = "New booking"
            End If
            RefreshHeader()
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "New Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Async Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem Is Nothing Then
            lblTeacherEmail.Text = ""
            Return
        End If
        Try
            Dim email = Await AppServices.Teachers.FindEmailAsync(ComboBox1.SelectedItem.ToString())
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(email), "Email not found", email)
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "New Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim unit = SelectedUnit()
        lblUnit.Text = If(unit Is Nothing OrElse String.IsNullOrEmpty(unit.Title), "", unit.Title)
    End Sub

    Private Sub IdentityChanged(sender As Object, e As EventArgs) Handles txtStudentID.TextChanged, txtGivenName.TextChanged, txtFamilyName.TextChanged, txtPersonalEmail.TextChanged
        RefreshHeader()
    End Sub

    Private Sub RefreshHeader()
        Dim id = txtStudentID.Text.Trim()
        Dim given = txtGivenName.Text.Trim()
        Dim family = txtFamilyName.Text.Trim()
        _headerId.Text = If(id.Length = 0, "Student ID", id)
        Dim name = (given & " " & family).Trim()
        _headerName.Text = If(name.Length = 0, "Student name", name)
        _headerEmail.Text = txtPersonalEmail.Text.Trim()
        _avatar.Initials = StudentDialogLayout.Initials(given, family)
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim studentID As String = txtStudentID.Text.Trim()
        Dim unit = SelectedUnit()
        If String.IsNullOrWhiteSpace(studentID) OrElse ComboBox1.SelectedItem Is Nothing OrElse unit Is Nothing OrElse ComboBox3.SelectedItem Is Nothing Then
            MessageBox.Show(Me, "Please enter a Student ID, and select teacher, unit, and attempt number.", "New Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim teacherName = ComboBox1.SelectedItem.ToString()
        Try
            Dim teacherEmail = Await AppServices.Teachers.FindEmailAsync(teacherName)
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(teacherEmail), "Email not found", teacherEmail)
            Dim booking As New ResitBooking With {
                .StudentId = studentID,
                .FirstName = txtGivenName.Text.Trim(),
                .Surname = txtFamilyName.Text.Trim(),
                .Email = txtPersonalEmail.Text.Trim(),
                .TeacherName = teacherName,
                .TeacherEmail = teacherEmail,
                .UnitCode = unit.Code,
                .UnitName = If(unit.Title, ""),
                .QualificationCode = SelectedQualificationCode(),
                .AttemptNo = ComboBox3.SelectedItem.ToString(),
                .EnergyspaceCreated = _classCreated.Checked,
                .EnergyspaceAssessmentBooked = _assessmentBooked.Checked,
                .BlockGroup = txtBlockgroup.Text,
                .ResitDate = DateTimePicker1.Value.Date
            }
            Dim result = Await AppServices.Bookings.UpsertAsync(booking, preserveEnergyspaceFlags:=False)
            If result.Kind = UpsertKind.Updated Then
                MessageBox.Show(Me, "Data updated successfully.", "New Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show(Me, "Data inserted successfully.", "New Student", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Dim mainForm As Main = CType(Application.OpenForms("Main"), Main)
            If mainForm IsNot Nothing Then Await mainForm.UpdateDataGridView()
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "New Student", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try
        txtStudentID.Text = ""
        txtGivenName.Text = ""
        txtFamilyName.Text = ""
        txtPersonalEmail.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        lblTeacherEmail.Text = ""
        lblUnit.Text = ""
        _statusBadge.Text = "New booking"
        RefreshHeader()
    End Sub
End Class
