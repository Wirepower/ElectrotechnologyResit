<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
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
        TextBox1 = New TextBox()
        Label1 = New Label()
        Button1 = New Button()
        Button2 = New Button()
        Label2 = New Label()
        TextBox2 = New TextBox()
        Label3 = New Label()
        TextBox3 = New TextBox()
        Button3 = New Button()
        CheckBoxResitLoginOpenAfterSave = New CheckBox()
        SuspendLayout()
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(65, 75)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(982, 23)
        TextBox1.TabIndex = 0
        TextBox1.Text = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;TrustServerCertificate=True;Multi Subnet Failover=False;"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 18F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(457, 35)
        Label1.Name = "Label1"
        Label1.Size = New Size(255, 32)
        Label1.TabIndex = 1
        Label1.Text = "SQL Connection String"
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(586, 365)
        Button1.Name = "Button1"
        Button1.Size = New Size(150, 37)
        Button1.TabIndex = 2
        Button1.Text = "Save and Exit"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(381, 365)
        Button2.Name = "Button2"
        Button2.Size = New Size(150, 37)
        Button2.TabIndex = 3
        Button2.Text = "Close without Saving"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(65, 118)
        Label2.Name = "Label2"
        Label2.Size = New Size(420, 15)
        Label2.TabIndex = 4
        Label2.Text = "PC / lab login password (monthly) — ""Password:"" line in the Word sheet"
        ' 
        ' TextBox2
        ' 
        TextBox2.Location = New Point(65, 138)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(982, 23)
        TextBox2.TabIndex = 5
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(65, 184)
        Label3.Name = "Label3"
        Label3.Size = New Size(420, 15)
        Label3.TabIndex = 6
        Label3.Text = "Word template (.docx). Default is created in the app Templates folder on first run."
        ' 
        ' TextBox3
        ' 
        TextBox3.Location = New Point(65, 204)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(860, 23)
        TextBox3.TabIndex = 7
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(940, 200)
        Button3.Name = "Button3"
        Button3.Size = New Size(107, 27)
        Button3.TabIndex = 8
        Button3.Text = "Browse…"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' CheckBoxResitLoginOpenAfterSave
        ' 
        CheckBoxResitLoginOpenAfterSave.AutoSize = True
        CheckBoxResitLoginOpenAfterSave.Location = New Point(65, 242)
        CheckBoxResitLoginOpenAfterSave.Name = "CheckBoxResitLoginOpenAfterSave"
        CheckBoxResitLoginOpenAfterSave.Size = New Size(520, 19)
        CheckBoxResitLoginOpenAfterSave.TabIndex = 9
        CheckBoxResitLoginOpenAfterSave.Text = "Open the Resit login Word document automatically after it is saved (skip 'Open it now?')"
        ' 
        ' Settings
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1123, 430)
        Controls.Add(CheckBoxResitLoginOpenAfterSave)
        Controls.Add(Button3)
        Controls.Add(TextBox3)
        Controls.Add(Label3)
        Controls.Add(TextBox2)
        Controls.Add(Label2)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(Label1)
        Controls.Add(TextBox1)
        Name = "Settings"
        Text = "Settings"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents CheckBoxResitLoginOpenAfterSave As CheckBox
End Class
