Public Class NewStudent
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub NewStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate the ComboBox with Teacher_Full_Name column from TeacherList table
        PopulateTeacherComboBox()
        PopulateUnitComboBox()


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
                        ComboBox2.Items.Clear()
                        'EditStudent.ComboBox2.Items.Clear()
                        ' Populate the ComboBox with data from the query result
                        While reader.Read()
                            ComboBox2.Items.Add(reader("Unit_Code").ToString())
                            'EditStudent.ComboBox2.Items.Add(reader("Unit_Code").ToString())
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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Get the Student ID from the textbox
        Dim studentID As String = txtStudentID.Text.Trim()

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
                        txtGivenName.Text = reader("Student Given Name").ToString()
                        txtFamilyName.Text = reader("Student Family Name").ToString()
                        txtPersonalEmail.Text = reader("Student Personal Email").ToString()
                        txtBlockgroup.Text = reader("Block Group Code").ToString()
                    Else
                        ' If no record found, display a message
                        MessageBox.Show("No record found for the provided Student ID.")
                        ' Clear text boxes
                        txtGivenName.Clear()
                        txtFamilyName.Clear()
                        txtPersonalEmail.Clear()
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
        ' Construct the SQL query to retrieve Unit_Title based on the selected Unit_Code
        Dim query As String = "SELECT Unit_Title FROM ElectrotechnologyReports.dbo.UEE30820units WHERE Unit_Code = @UnitCode"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get values from controls
        Dim studentID As String = txtStudentID.Text.Trim()
        Dim givenName As String = txtGivenName.Text.Trim()
        Dim familyName As String = txtFamilyName.Text.Trim()
        Dim personalEmail As String = txtPersonalEmail.Text.Trim()
        If ComboBox1.SelectedItem Is Nothing OrElse ComboBox2.SelectedItem Is Nothing OrElse ComboBox3.SelectedItem Is Nothing OrElse ComboBox4.SelectedItem Is Nothing OrElse ComboBox5.SelectedItem Is Nothing Then
            MessageBox.Show("Please select teacher, unit, attempt number, and both Energyspace options.")
            Return
        End If
        Dim allocatedTeacher As String = ComboBox1.SelectedItem.ToString()
        Dim allocatedTeacherEmail As String = lblTeacherEmail.Text
        Dim unit As String = ComboBox2.SelectedItem.ToString()
        Dim UnitName As String = lblUnit.Text
        Dim attemptNo As String = ComboBox3.SelectedItem.ToString()
        Dim energySpaceCreatedValue As Boolean = If(ComboBox5.SelectedItem.ToString() = "Yes", True, False)
        Dim energySpaceAssessmentBookedValue As Boolean = If(ComboBox4.SelectedItem.ToString() = "Yes", True, False)
        Dim blockgroup As String = txtBlockgroup.Text
        Dim resitDate As String = DateTimePicker1.Value.ToString("yyyy-MM-dd") ' Format date as "yyyy-MM-dd"

        ' Construct the SQL query to check if the student ID exists
        Dim queryCheck As String = "SELECT COUNT(*) FROM ElectrotechnologyReports.dbo.ElectricalResit WHERE [Student ID] = @StudentID"

        ' Create a SqlConnection
        Using connection As New SqlConnection(SQLCon.connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand to check if the student ID exists
                Using commandCheck As New SqlCommand(queryCheck, connection)
                    ' Add parameter for the student ID
                    commandCheck.Parameters.AddWithValue("@StudentID", studentID)
                    ' Execute the query to count rows with the given student ID
                    Dim count As Integer = CInt(commandCheck.ExecuteScalar())

                    ' Check if the student ID exists
                    If count > 0 Then
                        ' Student ID exists, so update the existing row
                        Dim updateQuery As String = "UPDATE ElectrotechnologyReports.dbo.ElectricalResit SET [Student Firstname] = @GivenName, [Student Surname] = @FamilyName, [Student Email] = @PersonalEmail, AllocatedTeacher = @AllocatedTeacher, AllocatedTeacherEmail = @AllocatedTeacherEmail, Unit = @Unit, [Unit Name] = @UnitName, AttemptNo = @AttemptNo, EnergyspaceCreated = @EnergySpaceCreated, EnergyspaceAssessmentBooked = @EnergySpaceAssessmentBooked, Blockgroup = @Blockgroup, [Resit date] = @ResitDate WHERE [Student ID] = @StudentID"

                        ' Create a SqlCommand for the update query
                        Using updateCommand As New SqlCommand(updateQuery, connection)
                            ' Add parameters for the update command
                            updateCommand.Parameters.AddWithValue("@GivenName", givenName)
                            updateCommand.Parameters.AddWithValue("@FamilyName", familyName)
                            updateCommand.Parameters.AddWithValue("@PersonalEmail", personalEmail)
                            updateCommand.Parameters.AddWithValue("@AllocatedTeacher", allocatedTeacher)
                            updateCommand.Parameters.AddWithValue("@AllocatedTeacherEmail", allocatedTeacherEmail)
                            updateCommand.Parameters.AddWithValue("@Unit", unit)
                            updateCommand.Parameters.AddWithValue("@UnitName", UnitName)
                            updateCommand.Parameters.AddWithValue("@AttemptNo", attemptNo)
                            updateCommand.Parameters.AddWithValue("@EnergySpaceCreated", energySpaceCreatedValue)
                            updateCommand.Parameters.AddWithValue("@EnergySpaceAssessmentBooked", energySpaceAssessmentBookedValue)
                            updateCommand.Parameters.AddWithValue("@Blockgroup", blockgroup)
                            updateCommand.Parameters.AddWithValue("@ResitDate", resitDate)
                            updateCommand.Parameters.AddWithValue("@StudentID", studentID)

                            ' Execute the update command
                            Dim rowsAffected As Integer = updateCommand.ExecuteNonQuery()

                            ' Check if rows were affected
                            If rowsAffected > 0 Then
                                MessageBox.Show("Data updated successfully.")
                            Else
                                MessageBox.Show("Failed to update data.")
                            End If
                        End Using
                    Else
                        ' Student ID does not exist, so insert a new row
                        Dim insertQuery As String = "INSERT INTO ElectrotechnologyReports.dbo.ElectricalResit ([Student ID], [Student Firstname], [Student Surname], [Student Email], AllocatedTeacher, AllocatedTeacherEmail, Unit, [Unit Name], AttemptNo, EnergyspaceCreated, EnergyspaceAssessmentBooked, Blockgroup, [Resit date]) VALUES (@StudentID, @GivenName, @FamilyName, @PersonalEmail, @AllocatedTeacher, @AllocatedTeacherEmail, @Unit, @UnitName, @AttemptNo, @EnergySpaceCreated, @EnergySpaceAssessmentBooked, @Blockgroup, @ResitDate)"

                        ' Create a SqlCommand for the insert query
                        Using insertCommand As New SqlCommand(insertQuery, connection)
                            ' Add parameters for the insert command
                            insertCommand.Parameters.AddWithValue("@StudentID", studentID)
                            insertCommand.Parameters.AddWithValue("@GivenName", givenName)
                            insertCommand.Parameters.AddWithValue("@FamilyName", familyName)
                            insertCommand.Parameters.AddWithValue("@PersonalEmail", personalEmail)
                            insertCommand.Parameters.AddWithValue("@AllocatedTeacher", allocatedTeacher)
                            insertCommand.Parameters.AddWithValue("@AllocatedTeacherEmail", allocatedTeacherEmail)
                            insertCommand.Parameters.AddWithValue("@Unit", unit)
                            insertCommand.Parameters.AddWithValue("@UnitName", UnitName)
                            insertCommand.Parameters.AddWithValue("@AttemptNo", attemptNo)
                            insertCommand.Parameters.AddWithValue("@EnergySpaceCreated", energySpaceCreatedValue)
                            insertCommand.Parameters.AddWithValue("@EnergySpaceAssessmentBooked", energySpaceAssessmentBookedValue)
                            insertCommand.Parameters.AddWithValue("@Blockgroup", blockgroup)
                            insertCommand.Parameters.AddWithValue("@ResitDate", resitDate)

                            ' Execute the insert command
                            Dim rowsAffected As Integer = insertCommand.ExecuteNonQuery()

                            ' Check if rows were affected
                            If rowsAffected > 0 Then
                                MessageBox.Show("Data inserted successfully.")
                            Else
                                MessageBox.Show("Failed to insert data.")
                            End If
                        End Using
                    End If
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
        Dim mainForm As Main = CType(Application.OpenForms("Main"), Main)
        mainForm.UpdateDataGridView()
        ' Clear input fields
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