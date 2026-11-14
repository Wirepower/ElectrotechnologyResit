Imports ElectrotechnologyResit.Core

Public Class EditStudent
    Public Property MainFormDataGridView As DataGridView
    Private currentRowIndex As Integer = -1
    Private ReadOnly _qualification As New Label()
    Private _storedQualification As String = ""
    Private _units As IReadOnlyList(Of UnitRecord) = Array.Empty(Of UnitRecord)()
    Private _designedWidth As Integer

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Async Sub EditStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If _designedWidth = 0 Then _designedWidth = ClientSize.Width
        _qualification.AutoSize = True
        _qualification.Location = New Point(184, 96)
        Controls.Add(New Label With {.AutoSize = True, .Location = New Point(96, 96), .Text = "Qualification"})
        Controls.Add(_qualification)
        Await PopulateUnitComboBox()
        Await PopulateTeacherComboBox()
        If MainFormDataGridView IsNot Nothing AndAlso MainFormDataGridView.Rows.Count > 0 Then
            currentRowIndex = 0
            Await LoadDataFromDataGridView()
        Else
            Button3.Enabled = False
            Button4.Enabled = False
        End If
        StudentFormLookupLabels.Apply(Me, lblTeacherEmail, lblUnit, _designedWidth)
    End Sub

    Private Async Function PopulateUnitComboBox() As Task
        Try
            _units = Await AppServices.Units.GetAllAsync()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Function

    Private Sub FillUnitsForStoredQualification()
        ComboBox2.Items.Clear()
        For Each unit In _units
            If String.Equals(If(unit.QualificationCode, ""), _storedQualification, StringComparison.OrdinalIgnoreCase) Then
                ComboBox2.Items.Add(unit.UnitCode)
            End If
        Next
    End Sub

    Private Async Function PopulateTeacherComboBox() As Task
        Try
            ComboBox1.Items.Clear()
            For Each teacher In Await AppServices.Teachers.GetAllAsync()
                ComboBox1.Items.Add(teacher.FullName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Function

    Private Async Function LoadDataFromDataGridView() As Task
        If currentRowIndex >= 0 AndAlso currentRowIndex < MainFormDataGridView.Rows.Count Then
            Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)
            If row IsNot Nothing AndAlso Not row.IsNewRow Then
                TextBox2.Text = CellText(row, "Student ID")
                TextBox3.Text = CellText(row, "Student Firstname")
                TextBox4.Text = CellText(row, "Surname")
                TextBox5.Text = CellText(row, "Email")
                ComboBox1.Text = CellText(row, "AllocatedTeacher")
                _storedQualification = CellText(row, "QualificationCode")
                _qualification.Text = If(String.IsNullOrWhiteSpace(_storedQualification), "(not recorded)", _storedQualification)
                FillUnitsForStoredQualification()
                ComboBox2.Text = CellText(row, "Unit")
                ComboBox3.Text = CellText(row, "AttemptNo")
                ComboBox5.Text = YesNoBoolean.ToYesNo(YesNoBoolean.Parse(CellText(row, "EnergyspaceClassCreated")))
                ComboBox4.Text = YesNoBoolean.ToYesNo(YesNoBoolean.Parse(CellText(row, "EnergyspaceAssessmentBooked")))
                txtBlockgroup.Text = CellText(row, "Blockgroup")
                lblUnit.Text = CellText(row, "Unit Name")
                Await ShowTeacherEmailAsync(ComboBox1.Text)

                Dim resitDateValue = row.Cells("Resit date").Value
                If resitDateValue IsNot Nothing AndAlso TypeOf resitDateValue IsNot DBNull Then
                    DateTimePicker1.Value = Convert.ToDateTime(resitDateValue)
                Else
                    DateTimePicker1.Value = DateTimePicker1.MinDate
                End If

                Button3.Enabled = currentRowIndex > 0
                Button4.Enabled = currentRowIndex < MainFormDataGridView.Rows.Count - 1
            End If
        End If
    End Function

    Private Shared Function CellText(row As DataGridViewRow, columnName As String) As String
        Dim value = row.Cells(columnName).Value
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return ""
        Return value.ToString()
    End Function

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If currentRowIndex > 0 Then
            currentRowIndex -= 1
            Await LoadDataFromDataGridView()
        End If
    End Sub

    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If currentRowIndex < MainFormDataGridView.Rows.Count - 1 Then
            currentRowIndex += 1
            Await LoadDataFromDataGridView()
            Button4.Enabled = currentRowIndex < MainFormDataGridView.Rows.Count - 1
        End If
    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim energySpaceCreated = YesNoBoolean.Parse(ComboBox5.Text)
        Dim energySpaceAssessmentBooked = YesNoBoolean.Parse(ComboBox4.Text)
        If currentRowIndex < 0 OrElse currentRowIndex >= MainFormDataGridView.Rows.Count Then Return
        Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)
        If row.IsNewRow Then Return
        Try
            Dim teacherEmail = Await AppServices.Teachers.FindEmailAsync(ComboBox1.Text)
            Dim unitTitle = Await AppServices.Units.FindTitleAsync(ComboBox2.Text)
            If String.IsNullOrEmpty(unitTitle) Then unitTitle = lblUnit.Text
            If unitTitle = "Unit Title not found" Then unitTitle = ""
            lblTeacherEmail.Text = If(String.IsNullOrEmpty(teacherEmail), "Email not found", teacherEmail)
            If Not String.IsNullOrEmpty(unitTitle) Then lblUnit.Text = unitTitle

            row.Cells("Student ID").Value = TextBox2.Text
            row.Cells("Student Firstname").Value = TextBox3.Text
            row.Cells("Surname").Value = TextBox4.Text
            row.Cells("Email").Value = TextBox5.Text
            row.Cells("AllocatedTeacher").Value = ComboBox1.Text
            row.Cells("AllocatedTeacher email").Value = teacherEmail
            row.Cells("Unit").Value = ComboBox2.Text
            row.Cells("Unit Name").Value = unitTitle
            row.Cells("AttemptNo").Value = ComboBox3.Text
            row.Cells("EnergyspaceClassCreated").Value = energySpaceCreated
            row.Cells("EnergyspaceAssessmentBooked").Value = energySpaceAssessmentBooked
            row.Cells("Resit date").Value = DateTimePicker1.Value.Date
            row.Cells("Blockgroup").Value = txtBlockgroup.Text

            Dim booking As New ResitBooking With {
                .StudentId = TextBox2.Text.Trim(),
                .FirstName = TextBox3.Text,
                .Surname = TextBox4.Text,
                .Email = TextBox5.Text,
                .TeacherName = ComboBox1.Text,
                .TeacherEmail = teacherEmail,
                .UnitCode = ComboBox2.Text,
                .UnitName = unitTitle,
                .AttemptNo = ComboBox3.Text,
                .EnergyspaceCreated = energySpaceCreated,
                .EnergyspaceAssessmentBooked = energySpaceAssessmentBooked,
                .ResitDate = DateTimePicker1.Value.Date,
                .BlockGroup = txtBlockgroup.Text
            }
            Await AppServices.Bookings.UpsertAsync(booking, preserveEnergyspaceFlags:=False)
            Dim mainForm As Main = TryCast(Application.OpenForms("Main"), Main)
            If mainForm IsNot Nothing Then Await mainForm.UpdateDataGridView()
            MessageBox.Show("Data updated successfully.")
        Catch ex As Exception
            MessageBox.Show("Error updating data: " & ex.Message)
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

    Private Async Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem Is Nothing Then Return
        Try
            Dim title = Await AppServices.Units.FindTitleAsync(ComboBox2.SelectedItem.ToString())
            lblUnit.Text = If(String.IsNullOrEmpty(title), "Unit Title not found", title)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim studentID As String = TextBox2.Text.Trim()
        If String.IsNullOrWhiteSpace(studentID) Then
            MessageBox.Show("Please enter Student ID.")
            Return
        End If
        Try
            Dim student = Await AppServices.Students.FindByStudentIdAsync(studentID)
            If student IsNot Nothing Then
                TextBox3.Text = student.GivenName
                TextBox4.Text = student.FamilyName
                TextBox5.Text = student.PersonalEmail
                txtBlockgroup.Text = student.BlockGroupCode
                ComboBox5.Text = YesNoBoolean.ToYesNo(False)
                ComboBox4.Text = YesNoBoolean.ToYesNo(False)
            Else
                MessageBox.Show("No record found for the provided Student ID.")
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                txtBlockgroup.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
End Class
