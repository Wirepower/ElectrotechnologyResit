<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EditStudent
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        Button4 = New Button()
        Button5 = New Button()
        TextBox2 = New TextBox()
        TextBox3 = New TextBox()
        TextBox4 = New TextBox()
        TextBox5 = New TextBox()
        txtBlockgroup = New TextBox()
        ComboBox1 = New ComboBox()
        ComboBox2 = New ComboBox()
        ComboBox3 = New ComboBox()
        DateTimePicker1 = New DateTimePicker()
        lblTeacherEmail = New Label()
        lblUnit = New Label()
        SuspendLayout()
        Button1.Text = "Update"
        Button2.Text = "Close"
        Button3.Text = "Previous"
        Button4.Text = "Next"
        Button5.Text = "Search Student ID"
        ComboBox3.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(720, 620)
        MinimumSize = New Size(680, 560)
        Name = "EditStudent"
        StartPosition = FormStartPosition.CenterParent
        Text = "Edit Student"
        ResumeLayout(False)
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents txtBlockgroup As TextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents lblTeacherEmail As Label
    Friend WithEvents lblUnit As Label
End Class
