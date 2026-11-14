Imports ElectrotechnologyResit.Core

Public Class Settings
    Private ComboBoxDataSource As ComboBox
    Private ButtonResetSample As Button
    Private ButtonTestConnection As Button
    Private ButtonImportCsv As Button

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BuildDataSourceControls()
        LoadFields()
    End Sub

    Private Sub BuildDataSourceControls()
        If ComboBoxDataSource IsNot Nothing Then Return
        Dim sourceGroup As New GroupBox With {
            .Text = "Data source",
            .Location = New Point(12, 12),
            .Size = New Size(1088, 72),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        sourceGroup.Controls.Add(New Label With {.AutoSize = True, .Location = New Point(12, 32), .Text = "Use"})
        ComboBoxDataSource = New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(48, 28),
            .Width = 260
        }
        ComboBoxDataSource.Items.Add("Sample data (testing)")
        ComboBoxDataSource.Items.Add("SharePoint")
        ButtonResetSample = New Button With {
            .Text = "Reset sample data",
            .Location = New Point(324, 26),
            .Size = New Size(160, 27)
        }
        sourceGroup.Controls.Add(ComboBoxDataSource)
        sourceGroup.Controls.Add(ButtonResetSample)
        Controls.Add(sourceGroup)
        AddHandler ComboBoxDataSource.SelectedIndexChanged, AddressOf DataSourceChanged
        AddHandler ButtonResetSample.Click, AddressOf ResetSample_Click

        GroupBoxSharePoint.Top = 92
        GroupBoxSharePoint.Height = 408
        ButtonWizard.Text = "Run first-time setup"
        ButtonWizard.Location = New Point(230, 352)
        ButtonWizard.Size = New Size(180, 28)
        ButtonTestConnection = New Button With {
            .Text = "Sign in / Test connection",
            .Location = New Point(12, 352),
            .Size = New Size(206, 28)
        }
        ButtonImportCsv = New Button With {
            .Text = "Import CSV/Excel",
            .Location = New Point(422, 352),
            .Size = New Size(160, 28)
        }
        GroupBoxSharePoint.Controls.Add(ButtonTestConnection)
        GroupBoxSharePoint.Controls.Add(ButtonImportCsv)
        AddHandler ButtonTestConnection.Click, AddressOf TestConnection_Click
        AddHandler ButtonImportCsv.Click, AddressOf ImportCsv_Click

        GroupBoxResitDoc.Top = 508
        GroupBoxBookings.Top = 684
        Button1.Top = 948
        Button2.Top = 948
        If ClientSize.Height < 1000 Then ClientSize = New Size(ClientSize.Width, 1004)
    End Sub

    Private Sub DataSourceChanged(sender As Object, e As EventArgs)
        If ButtonResetSample Is Nothing OrElse ComboBoxDataSource Is Nothing Then Return
        ButtonResetSample.Visible = ComboBoxDataSource.SelectedIndex = 0
    End Sub

    Private Sub LoadFields()
        ComboBoxDataSource.SelectedIndex = If(String.Equals(My.Settings.DataSource, "SharePoint", StringComparison.OrdinalIgnoreCase), 1, 0)
        ButtonResetSample.Visible = ComboBoxDataSource.SelectedIndex = 0
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
        My.Settings.DataSource = If(ComboBoxDataSource IsNot Nothing AndAlso ComboBoxDataSource.SelectedIndex = 1, "SharePoint", "Sample")
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
        ComboBoxDataSource.SelectedIndex = 1
        ApplySharePointFields()
        My.Settings.Save()
        AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
        Using wizard As New SetupWizard()
            wizard.ShowDialog(Me)
        End Using
        LoadFields()
        MessageBox.Show("If setup finished, Save and Exit, then open the app again to use SharePoint.", "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Async Sub TestConnection_Click(sender As Object, e As EventArgs)
        ApplySharePointFields()
        My.Settings.Save()
        ButtonTestConnection.Enabled = False
        UseWaitCursor = True
        Try
            Dim result = Await AppServices.TestSharePointConnectionAsync(Function() CType(Me, IWin32Window))
            MessageBox.Show(result.Message, "SharePoint sign-in", MessageBoxButtons.OK, If(result.Ok, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Finally
            UseWaitCursor = False
            ButtonTestConnection.Enabled = True
        End Try
    End Sub

    Private Async Sub ImportCsv_Click(sender As Object, e As EventArgs)
        ApplySharePointFields()
        My.Settings.Save()
        Try
            AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try
        Dim studentsFile = BrowseImport("Students CSV or Excel")
        Dim teachersFile = BrowseImport("Teachers CSV or Excel")
        Dim unitsFile = BrowseImport("Units CSV or Excel")
        If studentsFile = "" AndAlso teachersFile = "" AndAlso unitsFile = "" Then Return
        UseWaitCursor = True
        Try
            Dim lines As New List(Of String)
            If AppServices.IsSampleMode Then lines.Add("Imported into sample data (fictional store).")
            If studentsFile <> "" Then
                Dim tally = Await AppServices.LookupImport.ImportStudentsAsync(LookupFileReader.ReadStudents(studentsFile))
                lines.Add("Students: " & tally.Inserted.ToString() & " added, " & tally.Updated.ToString() & " updated, " & tally.Skipped.ToString() & " skipped.")
            End If
            If teachersFile <> "" Then
                Dim tally = Await AppServices.LookupImport.ImportTeachersAsync(LookupFileReader.ReadTeachers(teachersFile))
                lines.Add("Teachers: " & tally.Inserted.ToString() & " added, " & tally.Updated.ToString() & " updated, " & tally.Skipped.ToString() & " skipped.")
            End If
            If unitsFile <> "" Then
                Dim tally = Await AppServices.LookupImport.ImportUnitsAsync(LookupFileReader.ReadUnits(unitsFile))
                lines.Add("Units: " & tally.Inserted.ToString() & " added, " & tally.Updated.ToString() & " updated, " & tally.Skipped.ToString() & " skipped.")
            End If
            Dim mainForm = TryCast(Application.OpenForms("Main"), Main)
            If mainForm IsNot Nothing AndAlso mainForm.Visible Then Await mainForm.UpdateDataGridView()
            MessageBox.Show(String.Join(vbCrLf, lines), "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            UseWaitCursor = False
        End Try
    End Sub

    Private Function BrowseImport(title As String) As String
        Using dialog As New OpenFileDialog()
            dialog.Title = title
            dialog.Filter = "CSV or Excel (*.csv;*.xlsx)|*.csv;*.xlsx"
            If dialog.ShowDialog(Me) = DialogResult.OK Then Return dialog.FileName
        End Using
        Return ""
    End Function

    Private Async Sub ResetSample_Click(sender As Object, e As EventArgs)
        If ComboBoxDataSource.SelectedIndex <> 0 Then Return
        If MessageBox.Show("Replace the local sample bookings with the fictional seed? This does not touch SharePoint.", "Reset sample data", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
        ApplySharePointFields()
        My.Settings.Save()
        AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
        If AppServices.Sample Is Nothing Then Return
        AppServices.Sample.Reset(MelbourneCalendar.Today())
        Dim mainForm = TryCast(Application.OpenForms("Main"), Main)
        If mainForm IsNot Nothing AndAlso mainForm.Visible Then
            mainForm.ApplySampleDataCaption()
            Await mainForm.UpdateDataGridView()
        End If
        MessageBox.Show("Sample data has been reset. Bookings, teachers and units are fictional.", "Reset sample data", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
