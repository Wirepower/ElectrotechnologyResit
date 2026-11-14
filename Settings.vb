Imports ElectrotechnologyResit.Core

Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadFields()
    End Sub

    Private Sub LoadFields()
        TextBoxSiteUrl.Text = My.Settings.SharePointSiteUrl
        TextBoxGraphClientId.Text = My.Settings.GraphClientId
        TextBoxGraphTenant.Text = If(String.IsNullOrWhiteSpace(My.Settings.GraphTenantId), "common", My.Settings.GraphTenantId)
        CheckBoxSitesSelected.Checked = My.Settings.GraphUseSitesSelected
        TextBoxBookingsList.Text = SharePointListNames.FromSettings(My.Settings.ResitBookingsListName, "", "", "", "").Bookings
        TextBoxStudentsList.Text = SharePointListNames.FromSettings("", My.Settings.StudentsListName, "", "", "").Students
        TextBoxTeachersList.Text = SharePointListNames.FromSettings("", "", My.Settings.TeachersListName, "", "").Teachers
        TextBoxUnitsList.Text = SharePointListNames.FromSettings("", "", "", My.Settings.UnitsListName, "").Units
        TextBoxSignatureLibrary.Text = SharePointListNames.FromSettings("", "", "", "", My.Settings.SignatureImageLocation).SignatureLibrary
        TextBox2.Text = My.Settings.PcLoginMonthlyPassword
        TextBox3.Text = My.Settings.ResitLoginWordTemplatePath
        CheckBoxResitLoginOpenAfterSave.Checked = My.Settings.ResitLoginOpenDocumentAfterSave
        TextBoxBookingsClientId.Text = My.Settings.BookingsAzureClientId
        TextBoxBookingsTenant.Text = My.Settings.BookingsAzureTenantId
        TextBoxBookingsBusinessId.Text = My.Settings.BookingsBusinessId
        CheckBoxShowPullFromBookingsButton.Checked = My.Settings.ShowPullFromBookingsButton
    End Sub

    Private Sub ApplySharePointFields()
        My.Settings.SharePointSiteUrl = TextBoxSiteUrl.Text.Trim()
        My.Settings.GraphClientId = TextBoxGraphClientId.Text.Trim()
        My.Settings.GraphTenantId = If(String.IsNullOrWhiteSpace(TextBoxGraphTenant.Text), "common", TextBoxGraphTenant.Text.Trim())
        My.Settings.GraphUseSitesSelected = CheckBoxSitesSelected.Checked
        Dim names = SharePointListNames.FromSettings(TextBoxBookingsList.Text, TextBoxStudentsList.Text, TextBoxTeachersList.Text, TextBoxUnitsList.Text, TextBoxSignatureLibrary.Text)
        My.Settings.ResitBookingsListName = names.Bookings
        My.Settings.StudentsListName = names.Students
        My.Settings.TeachersListName = names.Teachers
        My.Settings.UnitsListName = names.Units
        My.Settings.SignatureImageLocation = names.SignatureLibrary
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

    Private Sub ButtonWizard_Click(sender As Object, e As EventArgs) Handles ButtonWizard.Click
        ApplySharePointFields()
        My.Settings.Save()
        Using wizard As New SetupWizard()
            wizard.ShowDialog(Me)
        End Using
        LoadFields()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ApplySharePointFields()
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
