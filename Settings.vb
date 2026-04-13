Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Settings.SQLConnection
        TextBox2.Text = My.Settings.PcLoginMonthlyPassword
        TextBox3.Text = My.Settings.ResitLoginWordTemplatePath
        CheckBoxResitLoginOpenAfterSave.Checked = My.Settings.ResitLoginOpenDocumentAfterSave
        TextBoxBookingsClientId.Text = My.Settings.BookingsAzureClientId
        TextBoxBookingsTenant.Text = My.Settings.BookingsAzureTenantId
        TextBoxBookingsBusinessId.Text = My.Settings.BookingsBusinessId
        CheckBoxShowPullFromBookingsButton.Checked = My.Settings.ShowPullFromBookingsButton
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Word documents (*.docx)|*.docx|All files|*.*"
            ofd.Title = "Resit login Word template"
            If ofd.ShowDialog(Me) = DialogResult.OK Then
                TextBox3.Text = ofd.FileName
            End If
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.SQLConnection = TextBox1.Text
        My.Settings.PcLoginMonthlyPassword = TextBox2.Text.Trim()
        My.Settings.ResitLoginWordTemplatePath = TextBox3.Text.Trim()
        My.Settings.ResitLoginOpenDocumentAfterSave = CheckBoxResitLoginOpenAfterSave.Checked
        My.Settings.BookingsAzureClientId = TextBoxBookingsClientId.Text.Trim()
        My.Settings.BookingsAzureTenantId = TextBoxBookingsTenant.Text.Trim()
        My.Settings.BookingsBusinessId = TextBoxBookingsBusinessId.Text.Trim()
        My.Settings.ShowPullFromBookingsButton = CheckBoxShowPullFromBookingsButton.Checked
        My.Settings.Save()
        Me.Hide()
        MessageBox.Show("Settings Saved, Please Restart Application", "Application Restart", MessageBoxButtons.OK)
        Application.Exit()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class