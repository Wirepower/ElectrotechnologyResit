<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        Button4 = New Button()
        Button5 = New Button()
        DataGridView1 = New DataGridView()
        Button6 = New Button()
        Button7 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        Label4 = New Label()
        Label5 = New Label()
        CheckBox1 = New CheckBox()
        CheckBox2 = New CheckBox()
        statusLabel = New Label()
        DateTimePicker1 = New DateTimePicker()
        SqlCommand1 = New SqlCommand()
        Button8 = New Button()
        Button9 = New Button()
        Button10 = New Button()
        ListView1 = New ListView()
        Label6 = New Label()
        ListView2 = New ListView()
        Label7 = New Label()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(32, 107)
        Button1.Name = "Button1"
        Button1.Size = New Size(121, 64)
        Button1.TabIndex = 0
        Button1.Text = "New Student"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(175, 107)
        Button2.Name = "Button2"
        Button2.Size = New Size(117, 63)
        Button2.TabIndex = 1
        Button2.Text = "Edit Student"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(311, 107)
        Button3.Name = "Button3"
        Button3.Size = New Size(117, 62)
        Button3.TabIndex = 2
        Button3.Text = "Delete Student"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Button4
        ' 
        Button4.Location = New Point(583, 143)
        Button4.Name = "Button4"
        Button4.Size = New Size(355, 29)
        Button4.TabIndex = 3
        Button4.Text = "Delete All Entries for selected resit date"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(1123, 107)
        Button5.Name = "Button5"
        Button5.Size = New Size(364, 65)
        Button5.TabIndex = 4
        Button5.Text = "Email Teachers of their Students for selected resit date"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' DataGridView1
        ' 
        DataGridView1.AllowUserToOrderColumns = True
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(32, 199)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.ReadOnly = True
        DataGridView1.Size = New Size(1455, 391)
        DataGridView1.TabIndex = 6
        ' 
        ' Button6
        ' 
        Button6.Location = New Point(32, 684)
        Button6.Name = "Button6"
        Button6.Size = New Size(127, 66)
        Button6.TabIndex = 7
        Button6.Text = "Reset All"
        Button6.UseVisualStyleBackColor = True
        ' 
        ' Button7
        ' 
        Button7.Location = New Point(1360, 684)
        Button7.Name = "Button7"
        Button7.Size = New Size(127, 66)
        Button7.TabIndex = 8
        Button7.Text = "Exit"
        Button7.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(721, 88)
        Label1.Name = "Label1"
        Label1.Size = New Size(140, 15)
        Label1.TabIndex = 9
        Label1.Text = "Select Date of Resit Night"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI Black", 20.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.ForeColor = Color.Blue
        Label2.Location = New Point(325, 25)
        Label2.Name = "Label2"
        Label2.Size = New Size(781, 37)
        Label2.TabIndex = 10
        Label2.Text = "RESIT BOOKING/DATABASE TEACHER EMAILING SYSTEM"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(680, 695)
        Label3.Name = "Label3"
        Label3.Size = New Size(185, 15)
        Label3.TabIndex = 11
        Label3.Text = "Created by Frank Offer (E5112471)"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(634, 710)
        Label4.Name = "Label4"
        Label4.Size = New Size(261, 15)
        Label4.TabIndex = 12
        Label4.Text = "Contact me via email on Frank.Offer@vu.edu.au"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(706, 725)
        Label5.Name = "Label5"
        Label5.Size = New Size(127, 15)
        Label5.TabIndex = 13
        Label5.Text = "Version 1.0 SQL edition"
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Checked = True
        CheckBox1.CheckState = CheckState.Checked
        CheckBox1.Location = New Point(1339, 174)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(148, 19)
        CheckBox1.TabIndex = 14
        CheckBox1.Text = "Display & Send Manually"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' CheckBox2
        ' 
        CheckBox2.AutoSize = True
        CheckBox2.Location = New Point(1150, 174)
        CheckBox2.Name = "CheckBox2"
        CheckBox2.Size = New Size(174, 19)
        CheckBox2.TabIndex = 15
        CheckBox2.Text = "Send Email - Do Not Display"
        CheckBox2.UseVisualStyleBackColor = True
        ' 
        ' statusLabel
        ' 
        statusLabel.Location = New Point(12, 9)
        statusLabel.Name = "statusLabel"
        statusLabel.Size = New Size(100, 23)
        statusLabel.TabIndex = 16
        ' 
        ' DateTimePicker1
        ' 
        DateTimePicker1.Location = New Point(583, 107)
        DateTimePicker1.Name = "DateTimePicker1"
        DateTimePicker1.Size = New Size(355, 23)
        DateTimePicker1.TabIndex = 17
        ' 
        ' SqlCommand1
        ' 
        SqlCommand1.CommandTimeout = 30
        SqlCommand1.EnableOptimizedParameterBinding = False
        ' 
        ' Button8
        ' 
        Button8.Location = New Point(944, 107)
        Button8.Name = "Button8"
        Button8.Size = New Size(173, 23)
        Button8.TabIndex = 19
        Button8.Text = "Show All"
        Button8.UseVisualStyleBackColor = True
        ' 
        ' Button9
        ' 
        Button9.Location = New Point(229, 684)
        Button9.Name = "Button9"
        Button9.Size = New Size(127, 66)
        Button9.TabIndex = 20
        Button9.Text = "Export to Excel"
        Button9.UseVisualStyleBackColor = True
        ' 
        ' Button10
        ' 
        Button10.Location = New Point(944, 143)
        Button10.Name = "Button10"
        Button10.Size = New Size(173, 29)
        Button10.TabIndex = 21
        Button10.Text = "Delete all 14 Day Old resits"
        Button10.UseVisualStyleBackColor = True
        ' 
        ' ListView1
        ' 
        ListView1.CheckBoxes = True
        ListView1.GridLines = True
        ListView1.LabelWrap = False
        ListView1.Location = New Point(32, 623)
        ListView1.Margin = New Padding(30)
        ListView1.Name = "ListView1"
        ListView1.Size = New Size(720, 55)
        ListView1.TabIndex = 22
        ListView1.UseCompatibleStateImageBehavior = False
        ListView1.View = View.SmallIcon
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(32, 606)
        Label6.Name = "Label6"
        Label6.Size = New Size(150, 15)
        Label6.TabIndex = 23
        Label6.Text = "Energyspace Class Created:"
        ' 
        ' ListView2
        ' 
        ListView2.CheckBoxes = True
        ListView2.GridLines = True
        ListView2.LabelWrap = False
        ListView2.Location = New Point(763, 623)
        ListView2.Margin = New Padding(30)
        ListView2.Name = "ListView2"
        ListView2.Size = New Size(724, 55)
        ListView2.TabIndex = 24
        ListView2.UseCompatibleStateImageBehavior = False
        ListView2.View = View.SmallIcon
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(763, 606)
        Label7.Name = "Label7"
        Label7.Size = New Size(184, 15)
        Label7.TabIndex = 25
        Label7.Text = "Energyspace Assessment Booked:"
        ' 
        ' Main
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1517, 762)
        Controls.Add(Label7)
        Controls.Add(ListView2)
        Controls.Add(Label6)
        Controls.Add(ListView1)
        Controls.Add(Button10)
        Controls.Add(Button9)
        Controls.Add(Button8)
        Controls.Add(DateTimePicker1)
        Controls.Add(statusLabel)
        Controls.Add(CheckBox2)
        Controls.Add(CheckBox1)
        Controls.Add(Label5)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(DataGridView1)
        Controls.Add(Button5)
        Controls.Add(Button4)
        Controls.Add(Button3)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Name = "Main"
        Text = "Form1"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents statusLabel As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents SqlCommand1 As SqlCommand
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents Label6 As Label
    Friend WithEvents ListView2 As ListView
    Friend WithEvents Label7 As Label

End Class
