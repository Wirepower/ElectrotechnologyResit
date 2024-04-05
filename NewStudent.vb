Public Class NewStudent
    Private connectionString As String = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;TrustServerCertificate=True;Multi Subnet Failover=False;"
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
        Using connection As New SqlConnection(connectionString)
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
        Dim query As String = "SELECT Teacher_Full_Name FROM ElectrotechnologyReports.dbo.TeacherList"

        ' Create a SqlConnection
        Using connection As New SqlConnection(connectionString)
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
        Using connection As New SqlConnection(connectionString)
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
                        lblBlockgroup.Text = reader("Block Group Code").ToString()
                    Else
                        ' If no record found, display a message
                        MessageBox.Show("No record found for the provided Student ID.")
                        ' Clear text boxes
                        txtGivenName.Clear()
                        txtFamilyName.Clear()
                        txtPersonalEmail.Clear()
                        lblBlockgroup.Text = ""
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
        Using connection As New SqlConnection(connectionString)
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
        Using connection As New SqlConnection(connectionString)
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
        Dim allocatedTeacher As String = ComboBox1.SelectedItem.ToString()
        Dim allocatedTeacherEmail As String = lblTeacherEmail.Text
        Dim unit As String = ComboBox2.SelectedItem.ToString()
        Dim UnitName As String = lblUnit.Text
        Dim attemptNo As String = ComboBox3.SelectedItem.ToString()
        Dim energySpaceCreatedValue As Boolean = If(ComboBox5.SelectedItem.ToString() = "Yes", True, False)
        Dim energySpaceAssessmentBookedValue As Boolean = If(ComboBox4.SelectedItem.ToString() = "Yes", True, False)
        Dim blockgroup As String = lblBlockgroup.Text
        Dim resitDate As String = DateTimePicker1.Value.ToString("yyyy-MM-dd") ' Format date as "yyyy-MM-dd"

        ' Construct the SQL query
        Dim query As String = "INSERT INTO ElectrotechnologyReports.dbo.ElectricalResit ([Student ID], [Student Firstname], [Student Surname], [Student Email], AllocatedTeacher, AllocatedTeacherEmail, Unit, [Unit Name], AttemptNo, EnergyspaceCreated, EnergyspaceAssessmentBooked, Blockgroup, [Resit date]) VALUES (@StudentID, @GivenName, @FamilyName, @PersonalEmail, @AllocatedTeacher, @AllocatedTeacherEmail, @Unit, @UnitName, @AttemptNo, @EnergySpaceCreated, @EnergySpaceAssessmentBooked, @Blockgroup, @ResitDate)"

        ' Create a SqlConnection
        Using connection As New SqlConnection(connectionString)
            Try
                ' Open the connection
                connection.Open()

                ' Create a SqlCommand
                Using command As New SqlCommand(query, connection)
                    ' Add parameters
                    command.Parameters.AddWithValue("@StudentID", studentID)
                    command.Parameters.AddWithValue("@GivenName", givenName)
                    command.Parameters.AddWithValue("@FamilyName", familyName)
                    command.Parameters.AddWithValue("@PersonalEmail", personalEmail)
                    command.Parameters.AddWithValue("@AllocatedTeacher", allocatedTeacher)
                    command.Parameters.AddWithValue("@AllocatedTeacherEmail", allocatedTeacherEmail)
                    command.Parameters.AddWithValue("@Unit", unit)
                    command.Parameters.AddWithValue("@UnitName", UnitName)
                    command.Parameters.AddWithValue("@AttemptNo", attemptNo)
                    command.Parameters.AddWithValue("@EnergySpaceCreated", energySpaceCreatedValue)
                    command.Parameters.AddWithValue("@EnergySpaceAssessmentBooked", energySpaceAssessmentBookedValue)
                    command.Parameters.AddWithValue("@Blockgroup", blockgroup)
                    command.Parameters.AddWithValue("@ResitDate", resitDate)

                    ' Execute the query
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()

                    ' Check if rows were affected
                    If rowsAffected > 0 Then
                        MessageBox.Show("Data saved successfully.")
                    Else
                        MessageBox.Show("Failed to save data.")
                    End If
                End Using
            Catch ex As Exception
                ' Handle any exceptions
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
        Dim mainForm As Main = CType(Application.OpenForms("Main"), Main)
        mainForm.UpdateDataGridView()
        txtStudentID.Text = ""
        txtGivenName.Text = ""
        txtFamilyName.Text = ""
        txtPersonalEmail.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
    End Sub


End Class