Imports ElectrotechnologyResit.Core

Public Class NewStudent
    Private ReadOnly _qualification As New ComboBox()
    Private _units As IReadOnlyList(Of UnitRecord) = Array.Empty(Of UnitRecord)()
    Private _designedWidth As Integer

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Async Sub NewStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If _designedWidth = 0 Then _designedWidth = ClientSize.Width
        _qualification.DropDownStyle = ComboBoxStyle.DropDownList
        _qualification.Location = New Point(182, 92)
        _qualification.Size = New Size(157, 23)
        Controls.Add(New Label With {.AutoSize = True, .Location = New Point(96, 95), .Text = "Qualification"})
        Controls.Add(_qualification)
        AddHandler _qualification.SelectedIndexChanged, AddressOf QualificationChanged
        Await PopulateTeacherComboBox()
        Await PopulateUnitComboBox()
        StudentFormLookupLabels.Apply(Me, lblTeacherEmail, lblUnit, _designedWidth)
    End Sub

    Private Async Function PopulateUnitComboBox() As Task
        Try
            _units = Await AppServices.Units.GetAllAsync()
            Dim qualifications = Await AppServices.Qualifications.GetAllAsync()
            _qualification.Items.Clear()
            For Each qualification In qualifications
                If QualificationStatus.OfferedForNewBookings(qualification.Status) Then
                    _qualification.Items.Add(qualification.QualificationCode)
                End If
            Next
            If _qualification.Items.Count > 0 Then _qualification.SelectedIndex = 0 Else FillUnits()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Function

    Private Sub QualificationChanged(sender As Object, e As EventArgs)
        FillUnits()
    End Sub

    Private Sub FillUnits()
        Dim package = If(_qualification.SelectedItem, "").ToString()
        ComboBox2.Items.Clear()
        For Each unit In _units
            If String.Equals(If(unit.QualificationCode, ""), package, StringComparison.OrdinalIgnoreCase) Then
                ComboBox2.Items.Add(unit.UnitCode)
            End If
        Next
        lblUnit.Text = ""
    End Sub

    Private Async Function PopulateTeacherComboBox() As Task
        Try
            Dim teachers = Await AppServices.Teachers.GetAllAsync()
            ComboBox1.Items.Clear()
            For Each teacher In teachers
                ComboBox1.Items.Add(teacher.FullName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Function

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim studentID As String = txtStudentID.Text.Trim()
        If String.IsNullOrWhiteSpace(studentID) Then
            MessageBox.Show("Please enter Student ID.")
            Return
        End If
        Try
            Dim student = Await AppServices.Students.FindByStudentIdAsync(studentID)
            If student IsNot Nothing Then
                txtGivenName.Text = student.GivenName
                txtFamilyName.Text = student.FamilyName
                txtPersonalEmail.Text = student.PersonalEmail
                txtBlockgroup.Text = student.BlockGroupCode
            Else
                MessageBox.Show("No record found for the provided Student ID.")
                txtGivenName.Clear()
                txtFamilyName.Clear()
                txtPersonalEmail.Clear()
                txtBlockgroup.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Async Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            Await ShowTeacherEmailAsync(ComboBox1.SelectedItem.ToString())
        End If
    End Sub

    Private Async Function ShowTeacherEmailAsync(teacherFullName As String) As Task
        Try
            Dim email = Await AppServices.Teachers.FindEmailAsync(teacherFullName)
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(email), "Email not found", email)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem Is Nothing Then Return
        Dim title = TitleForSelectedUnit()
        lblUnit.Text = If(String.IsNullOrEmpty(title), "Unit Title not found", title)
    End Sub

    Private Function TitleForSelectedUnit() As String
        Dim package = If(_qualification.SelectedItem, "").ToString()
        Dim code = If(ComboBox2.SelectedItem, "").ToString()
        For Each unit In _units
            If String.Equals(unit.UnitCode, code, StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(If(unit.QualificationCode, ""), package, StringComparison.OrdinalIgnoreCase) Then
                Return If(unit.UnitTitle, "")
            End If
        Next
        Return ""
    End Function

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim studentID As String = txtStudentID.Text.Trim()
        If String.IsNullOrWhiteSpace(studentID) OrElse ComboBox1.SelectedItem Is Nothing OrElse ComboBox2.SelectedItem Is Nothing OrElse ComboBox3.SelectedItem Is Nothing OrElse ComboBox4.SelectedItem Is Nothing OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Please enter a Student ID, and select teacher, unit, attempt number, and both Energyspace options.")
            Return
        End If
        Dim teacherName = ComboBox1.SelectedItem.ToString()
        Dim teacherEmail = ""
        Dim unitTitle = lblUnit.Text
        Try
            teacherEmail = Await AppServices.Teachers.FindEmailAsync(teacherName)
            Dim lookedUpTitle = TitleForSelectedUnit()
            If Not String.IsNullOrEmpty(lookedUpTitle) Then unitTitle = lookedUpTitle
            If unitTitle = "Unit Title not found" Then unitTitle = ""
            Dim booking As New ResitBooking With {
                .StudentId = studentID,
                .FirstName = txtGivenName.Text.Trim(),
                .Surname = txtFamilyName.Text.Trim(),
                .Email = txtPersonalEmail.Text.Trim(),
                .TeacherName = teacherName,
                .TeacherEmail = teacherEmail,
                .UnitCode = ComboBox2.SelectedItem.ToString(),
                .UnitName = unitTitle,
                .QualificationCode = If(_qualification.SelectedItem, "").ToString(),
                .AttemptNo = ComboBox3.SelectedItem.ToString(),
                .EnergyspaceCreated = YesNoBoolean.Parse(ComboBox5.SelectedItem.ToString()),
                .EnergyspaceAssessmentBooked = YesNoBoolean.Parse(ComboBox4.SelectedItem.ToString()),
                .BlockGroup = txtBlockgroup.Text,
                .ResitDate = DateTimePicker1.Value.Date
            }
            Dim result = Await AppServices.Bookings.UpsertAsync(booking, preserveEnergyspaceFlags:=False)
            If result.Kind = UpsertKind.Updated Then
                MessageBox.Show("Data updated successfully.")
            Else
                MessageBox.Show("Data inserted successfully.")
            End If
            Dim mainForm As Main = CType(Application.OpenForms("Main"), Main)
            If mainForm IsNot Nothing Then Await mainForm.UpdateDataGridView()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            Return
        End Try
        txtStudentID.Text = ""
        txtGivenName.Text = ""
        txtFamilyName.Text = ""
        txtPersonalEmail.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
    End Sub

    Private Sub lblUnit_Click(sender As Object, e As EventArgs) Handles lblUnit.Click
    End Sub
End Class
