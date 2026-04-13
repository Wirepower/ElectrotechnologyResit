Public Class EditStudent
    Public Property MainFormDataGridView As DataGridView
    ' Define a variable to keep track of the current row index
    Private currentRowIndex As Integer = -1

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub EditStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateUnitComboBox()
        PopulateTeacherComboBox()
        ' Initialize currentRowIndex to 0 if there are rows in the DataGridView
        If MainFormDataGridView.Rows.Count > 0 Then
            currentRowIndex = 0
            LoadDataFromDataGridView()
        Else
            ' If there are no rows, disable navigation buttons
            Button3.Enabled = False
            Button4.Enabled = False
        End If
    End Sub
    Private Sub PopulateUnitComboBox()
        ' Construct the SQL query to retrieve Unit_Code column from UEE30820units table
        Dim query As String = "SELECT Unit_Code FROM ElectrotechnologyReports.dbo.UEE30820units"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Execute the query
                    Using reader As SqlDataReader = command.ExecuteReader()
                        ' Clear the ComboBox
                        'ComboBox2.Items.Clear()
                        ComboBox2.Items.Clear()
                        ' Populate the ComboBox with data from the query result
                        While reader.Read()
                            'ComboBox2.Items.Add(reader("Unit_Code").ToString())
                            ComboBox2.Items.Add(reader("Unit_Code").ToString())
                        End While
                    End Using
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
    Private Sub PopulateTeacherComboBox()
        ' Construct the SQL query to retrieve Teacher_Full_Name column from TeacherList table
        Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList ORDER BY Teacher_Full_Name ASC"


        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Execute the query
                    Using reader As SqlDataReader = command.ExecuteReader()
                        ' Clear the ComboBox
                        ComboBox1.Items.Clear()
                        ' Populate the ComboBox with data from the query result
                        While reader.Read()
                            ComboBox1.Items.Add(reader("Teacher_Full_Name").ToString())
                        End While
                    End Using
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
    Private Sub LoadDataFromDataGridView()
        If currentRowIndex >= 0 AndAlso currentRowIndex < MainFormDataGridView.Rows.Count Then
            Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)

            If row IsNot Nothing Then
                TextBox2.Text = If(row.Cells("Student ID").Value IsNot Nothing, row.Cells("Student ID").Value.ToString(), "")
                TextBox3.Text = If(row.Cells("Student Firstname").Value IsNot Nothing, row.Cells("Student Firstname").Value.ToString(), "")
                TextBox4.Text = If(row.Cells("Student Surname").Value IsNot Nothing, row.Cells("Student Surname").Value.ToString(), "")
                TextBox5.Text = If(row.Cells("Student Email").Value IsNot Nothing, row.Cells("Student Email").Value.ToString(), "")
                ComboBox1.Text = If(row.Cells("AllocatedTeacher").Value IsNot Nothing, row.Cells("AllocatedTeacher").Value.ToString(), "")
                ComboBox2.Text = If(row.Cells("Unit").Value IsNot Nothing, row.Cells("Unit").Value.ToString(), "")
                ComboBox3.Text = If(row.Cells("AttemptNo").Value IsNot Nothing, row.Cells("AttemptNo").Value.ToString(), "")
                ComboBox5.Text = If(row.Cells("EnergyspaceCreated").Value IsNot Nothing, row.Cells("EnergyspaceCreated").Value.ToString(), "")
                ComboBox4.Text = If(row.Cells("EnergyspaceAssessmentBooked").Value IsNot Nothing, row.Cells("EnergyspaceAssessmentBooked").Value.ToString(), "")
                txtBlockgroup.Text = If(row.Cells("Blockgroup").Value IsNot Nothing, row.Cells("Blockgroup").Value.ToString(), "")

                ' Handle DateTimePicker value separately
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
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If currentRowIndex > 0 Then
            currentRowIndex -= 1
            LoadDataFromDataGridView()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If currentRowIndex < MainFormDataGridView.Rows.Count - 1 Then
            currentRowIndex += 1
            LoadDataFromDataGridView()
            ' Disable the Next button if there are no more rows after navigating
            Button4.Enabled = currentRowIndex < MainFormDataGridView.Rows.Count - 1
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim energySpaceCreated As Boolean
        Dim energySpaceAssessmentBooked As Boolean
        If Not Boolean.TryParse(ComboBox5.Text, energySpaceCreated) Then energySpaceCreated = False
        If Not Boolean.TryParse(ComboBox4.Text, energySpaceAssessmentBooked) Then energySpaceAssessmentBooked = False

        ' Update data in the current row of the DataGridView
        If currentRowIndex >= 0 AndAlso currentRowIndex < MainFormDataGridView.Rows.Count Then
            Dim row As DataGridViewRow = MainFormDataGridView.Rows(currentRowIndex)
            row.Cells("Student ID").Value = TextBox2.Text
            row.Cells("Student Firstname").Value = TextBox3.Text
            row.Cells("Student Surname").Value = TextBox4.Text
            row.Cells("Student Email").Value = TextBox5.Text
            row.Cells("AllocatedTeacher").Value = ComboBox1.Text
            row.Cells("AllocatedTeacherEmail").Value = lblTeacherEmail.Text
            row.Cells("Unit").Value = ComboBox2.Text
            row.Cells("Unit Name").Value = lblUnit.Text
            row.Cells("AttemptNo").Value = ComboBox3.Text
            row.Cells("EnergyspaceCreated").Value = ComboBox5.Text
            row.Cells("EnergyspaceAssessmentBooked").Value = ComboBox4.Text
            row.Cells("Resit date").Value = DateTimePicker1.Value.ToString("yyyy-MM-dd")
            row.Cells("Blockgroup").Value = txtBlockgroup.Text

            ' Update the corresponding row in the SQL table
            Try
                Using connection As New SqlConnection(SQLCon.connectionString)
                    connection.Open()
                    Dim updateCommand As String = "UPDATE ElectrotechnologyReports.dbo.ElectricalResit SET [Student Firstname] = @StudentFirstname, " &
                                              "[Student Surname] = @StudentSurname, " &
                                              "[Student Email] = @StudentEmail, " &
                                              "[AllocatedTeacher] = @AllocatedTeacher, " &
                                              "[AllocatedTeacherEmail]=@AllocatedTeacherEmail, " &
                                              "Unit = @Unit, " &
                                              "[Unit Name] = @UnitName, " &
                                              "AttemptNo = @AttemptNo, " &
                                              "EnergyspaceCreated = @EnergyspaceCreated, " &
                                              "EnergyspaceAssessmentBooked = @EnergyspaceAssessmentBooked, " &
                                              "[Resit date] = @ResitDate, " &
                                              "Blockgroup = @Blockgroup " &
                                              "WHERE [Student ID] = @StudentID"

                    Using command As New SqlCommand(updateCommand, connection)
                        ' Add parameters
                        command.Parameters.AddWithValue("@StudentID", TextBox2.Text)
                        command.Parameters.AddWithValue("@StudentFirstname", TextBox3.Text)
                        command.Parameters.AddWithValue("@StudentSurname", TextBox4.Text)
                        command.Parameters.AddWithValue("@StudentEmail", TextBox5.Text)
                        command.Parameters.AddWithValue("@AllocatedTeacher", ComboBox1.Text)
                        command.Parameters.AddWithValue("@AllocatedTeacherEmail", lblTeacherEmail.Text)
                        command.Parameters.AddWithValue("@Unit", ComboBox2.Text)
                        command.Parameters.AddWithValue("@UnitName", lblUnit.Text)
                        command.Parameters.AddWithValue("@AttemptNo", ComboBox3.Text)
                        command.Parameters.AddWithValue("@EnergyspaceCreated", energySpaceCreated)
                        command.Parameters.AddWithValue("@EnergyspaceAssessmentBooked", energySpaceAssessmentBooked)
                        command.Parameters.AddWithValue("@ResitDate", DateTimePicker1.Value)
                        command.Parameters.AddWithValue("@Blockgroup", txtBlockgroup.Text)

                        ' Execute the command
                        command.ExecuteNonQuery()
                    End Using
                End Using
                MessageBox.Show("Data updated successfully.")
            Catch ex As Exception
                MessageBox.Show("Error updating data: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Update the label with the email of the selected item in the ComboBox
        If ComboBox1.SelectedItem IsNot Nothing Then
            UpdateEmailLabel(ComboBox1.SelectedItem.ToString())
        End If
    End Sub
    Private Sub UpdateEmailLabel(teacherFullName As String)
        ' Construct the SQL query to retrieve the Email column based on the selected Teacher_Full_Name
        Dim query As String = "SELECT Email FROM ElectrotechnologyReports.dbo.TeacherList WHERE Teacher_Full_Name = @TeacherFullName"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@TeacherFullName", teacherFullName)

                    ' Execute the query
                    Dim email As Object = command.ExecuteScalar()

                    ' Update the label with the retrieved email
                    If email IsNot Nothing Then
                        lblTeacherEmail.Text = email.ToString()
                    Else
                        lblTeacherEmail.Text = "Email not found"
                    End If
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
    Private Sub UpdateUnitLabel(unitCode As String)
        ' Construct the SQL query with parameter
        Dim query As String = "SELECT Unit_Title FROM ElectrotechnologyReports.dbo.UEE30820units WHERE Unit_Code = @UnitCode"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Add parameter
                    command.Parameters.AddWithValue("@UnitCode", unitCode)

                    ' Execute the query
                    Dim unitTitle As Object = command.ExecuteScalar()

                    ' Update lblUnit with the retrieved Unit_Title
                    If unitTitle IsNot Nothing Then
                        lblUnit.Text = unitTitle.ToString()
                    Else
                        lblUnit.Text = "Unit Title not found"
                    End If
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem IsNot Nothing Then
            UpdateUnitLabel(ComboBox2.SelectedItem.ToString())
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Get the Student ID from the textbox
        Dim studentID As String = TextBox2.Text.Trim()

        ' Check if Student ID is provided
        If String.IsNullOrWhiteSpace(studentID) Then
            MessageBox.Show("Please enter Student ID.")
            Return
        End If

        ' Construct the SQL query to retrieve other details based on Student ID
        Dim query As String = "SELECT [Student Given Name], [Student Family Name], [Student Personal Email], [Block Group Code] FROM ElectrotechnologyReports.dbo.AgreementsDetails WHERE [Student ID] = @StudentID"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@StudentID", studentID)

                    ' Execute the query
                    Dim reader As SqlDataReader = command.ExecuteReader()

                    ' Check if any records are returned
                    If reader.HasRows Then
                        ' Read the first record
                        reader.Read()

                        ' Populate the text boxes with the retrieved values
                        TextBox3.Text = reader("Student Given Name").ToString()
                        TextBox4.Text = reader("Student Family Name").ToString()
                        TextBox5.Text = reader("Student Personal Email").ToString()
                        txtBlockgroup.Text = reader("Block Group Code").ToString()
                        ComboBox5.Text = "False"
                        ComboBox4.Text = "False"
                    Else
                        ' If no record found, display a message
                        MessageBox.Show("No record found for the provided Student ID.")
                        ' Clear text boxes
                        TextBox3.Clear()
                        TextBox4.Clear()
                        TextBox5.Clear()
                        txtBlockgroup.Text = ""
                    End If

                    ' Close the data reader
                    reader.Close()
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
End Class