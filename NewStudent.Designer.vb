<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewStudent
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Button1 = New Button()
        Button2 = New Button()
        txtStudentID = New TextBox()
        txtGivenName = New TextBox()
        txtFamilyName = New TextBox()
        txtPersonalEmail = New TextBox()
        ComboBox1 = New ComboBox()
        ComboBox2 = New ComboBox()
        ComboBox3 = New ComboBox()
        ComboBox4 = New ComboBox()
        ComboBox5 = New ComboBox()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        Label4 = New Label()
        Label5 = New Label()
        Label6 = New Label()
        Label7 = New Label()
        Label8 = New Label()
        Label9 = New Label()
        Label10 = New Label()
        Label11 = New Label()
        lblBlockgroup = New Label()
        Label15 = New Label()
        Button3 = New Button()
        lblTeacherEmail = New Label()
        lblUnit = New Label()
        DateTimePicker1 = New DateTimePicker()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(62, 12)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 32)
        Button1.TabIndex = 0
        Button1.Text = "Save"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(380, 12)
        Button2.Name = "Button2"
        Button2.Size = New Size(75, 32)
        Button2.TabIndex = 1
        Button2.Text = "Close"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' txtStudentID
        ' 
        txtStudentID.Location = New Point(182, 127)
        txtStudentID.Name = "txtStudentID"
        txtStudentID.Size = New Size(157, 23)
        txtStudentID.TabIndex = 3
        ' 
        ' txtGivenName
        ' 
        txtGivenName.Location = New Point(182, 161)
        txtGivenName.Name = "txtGivenName"
        txtGivenName.Size = New Size(157, 23)
        txtGivenName.TabIndex = 4
        ' 
        ' txtFamilyName
        ' 
        txtFamilyName.Location = New Point(182, 199)
        txtFamilyName.Name = "txtFamilyName"
        txtFamilyName.Size = New Size(157, 23)
        txtFamilyName.TabIndex = 5
        ' 
        ' txtPersonalEmail
        ' 
        txtPersonalEmail.Location = New Point(182, 234)
        txtPersonalEmail.Name = "txtPersonalEmail"
        txtPersonalEmail.Size = New Size(157, 23)
        txtPersonalEmail.TabIndex = 6
        ' 
        ' ComboBox1
        ' 
        ComboBox1.FormattingEnabled = True
        ComboBox1.Location = New Point(182, 272)
        ComboBox1.Name = "ComboBox1"
        ComboBox1.Size = New Size(157, 23)
        ComboBox1.TabIndex = 7
        ' 
        ' ComboBox2
        ' 
        ComboBox2.FormattingEnabled = True
        ComboBox2.Location = New Point(182, 311)
        ComboBox2.Name = "ComboBox2"
        ComboBox2.Size = New Size(157, 23)
        ComboBox2.TabIndex = 8
        ' 
        ' ComboBox3
        ' 
        ComboBox3.AutoCompleteCustomSource.AddRange(New String() {"1", "2", "3", "4", "5", "6"})
        ComboBox3.FormattingEnabled = True
        ComboBox3.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        ComboBox3.Location = New Point(182, 340)
        ComboBox3.Name = "ComboBox3"
        ComboBox3.Size = New Size(157, 23)
        ComboBox3.TabIndex = 9
        ' 
        ' ComboBox4
        ' 
        ComboBox4.AutoCompleteCustomSource.AddRange(New String() {"Yes", "No"})
        ComboBox4.FormattingEnabled = True
        ComboBox4.Items.AddRange(New Object() {"Yes", "No"})
        ComboBox4.Location = New Point(309, 395)
        ComboBox4.Name = "ComboBox4"
        ComboBox4.Size = New Size(157, 23)
        ComboBox4.TabIndex = 10
        ComboBox4.Text = "No"
        ' 
        ' ComboBox5
        ' 
        ComboBox5.AutoCompleteCustomSource.AddRange(New String() {"Yes", "No"})
        ComboBox5.FormattingEnabled = True
        ComboBox5.Items.AddRange(New Object() {"Yes", "No"})
        ComboBox5.Location = New Point(43, 395)
        ComboBox5.Name = "ComboBox5"
        ComboBox5.Size = New Size(157, 23)
        ComboBox5.TabIndex = 11
        ComboBox5.Text = "No"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(111, 127)
        Label1.Name = "Label1"
        Label1.Size = New Size(65, 15)
        Label1.TabIndex = 13
        Label1.Text = "Student ID:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(109, 161)
        Label2.Name = "Label2"
        Label2.Size = New Size(67, 15)
        Label2.TabIndex = 14
        Label2.Text = "First Name:"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(111, 199)
        Label3.Name = "Label3"
        Label3.Size = New Size(66, 15)
        Label3.TabIndex = 15
        Label3.Text = "Last Name:"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(94, 237)
        Label4.Name = "Label4"
        Label4.Size = New Size(83, 15)
        Label4.TabIndex = 16
        Label4.Text = "Student Email:"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(127, 275)
        Label5.Name = "Label5"
        Label5.Size = New Size(50, 15)
        Label5.TabIndex = 17
        Label5.Text = "Teacher:"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(145, 314)
        Label6.Name = "Label6"
        Label6.Size = New Size(32, 15)
        Label6.TabIndex = 18
        Label6.Text = "Unit:"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(76, 343)
        Label7.Name = "Label7"
        Label7.Size = New Size(101, 15)
        Label7.TabIndex = 19
        Label7.Text = "Attempt Number:"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(26, 377)
        Label8.Name = "Label8"
        Label8.Size = New Size(200, 15)
        Label8.TabIndex = 20
        Label8.Text = "Has Energyspace class been created?"
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Location = New Point(258, 377)
        Label9.Name = "Label9"
        Label9.Size = New Size(255, 15)
        Label9.TabIndex = 21
        Label9.Text = "Has Assessment been booked on Energyspace?"
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Location = New Point(224, 422)
        Label10.Name = "Label10"
        Label10.Size = New Size(69, 15)
        Label10.TabIndex = 22
        Label10.Text = "BlockGroup"
        ' 
        ' Label11
        ' 
        Label11.AutoSize = True
        Label11.Location = New Point(182, 483)
        Label11.Name = "Label11"
        Label11.Size = New Size(139, 15)
        Label11.TabIndex = 23
        Label11.Text = "Booked Assessment Date"
        ' 
        ' lblBlockgroup
        ' 
        lblBlockgroup.BackColor = SystemColors.ButtonHighlight
        lblBlockgroup.BorderStyle = BorderStyle.FixedSingle
        lblBlockgroup.Location = New Point(182, 437)
        lblBlockgroup.Name = "lblBlockgroup"
        lblBlockgroup.Size = New Size(157, 23)
        lblBlockgroup.TabIndex = 25
        lblBlockgroup.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' Label15
        ' 
        Label15.AutoSize = True
        Label15.Font = New Font("Segoe UI", 18F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label15.ForeColor = Color.Blue
        Label15.Location = New Point(145, 47)
        Label15.Name = "Label15"
        Label15.Size = New Size(234, 32)
        Label15.TabIndex = 27
        Label15.Text = "Add a New Student"
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(347, 126)
        Button3.Name = "Button3"
        Button3.Size = New Size(135, 24)
        Button3.TabIndex = 28
        Button3.Text = "Search Student ID"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' lblTeacherEmail
        ' 
        lblTeacherEmail.Location = New Point(347, 275)
        lblTeacherEmail.Name = "lblTeacherEmail"
        lblTeacherEmail.Size = New Size(171, 15)
        lblTeacherEmail.TabIndex = 29
        ' 
        ' lblUnit
        ' 
        lblUnit.Font = New Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        lblUnit.Location = New Point(347, 303)
        lblUnit.Name = "lblUnit"
        lblUnit.Size = New Size(181, 55)
        lblUnit.TabIndex = 30
        ' 
        ' DateTimePicker1
        ' 
        DateTimePicker1.Location = New Point(111, 501)
        DateTimePicker1.Name = "DateTimePicker1"
        DateTimePicker1.Size = New Size(283, 23)
        DateTimePicker1.TabIndex = 31
        ' 
        ' NewStudent
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(530, 569)
        Controls.Add(DateTimePicker1)
        Controls.Add(lblUnit)
        Controls.Add(lblTeacherEmail)
        Controls.Add(Button3)
        Controls.Add(Label15)
        Controls.Add(lblBlockgroup)
        Controls.Add(Label11)
        Controls.Add(Label10)
        Controls.Add(Label9)
        Controls.Add(Label8)
        Controls.Add(Label7)
        Controls.Add(Label6)
        Controls.Add(Label5)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(ComboBox5)
        Controls.Add(ComboBox4)
        Controls.Add(ComboBox3)
        Controls.Add(ComboBox2)
        Controls.Add(ComboBox1)
        Controls.Add(txtPersonalEmail)
        Controls.Add(txtFamilyName)
        Controls.Add(txtGivenName)
        Controls.Add(txtStudentID)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Name = "NewStudent"
        Text = "NewStudent"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents txtStudentID As TextBox
    Friend WithEvents txtGivenName As TextBox
    Friend WithEvents txtFamilyName As TextBox
    Friend WithEvents txtPersonalEmail As TextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents ComboBox5 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents lblBlockgroup As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents lblTeacherEmail As Label
    Friend WithEvents lblUnit As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
End Class
