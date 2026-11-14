Imports ElectrotechnologyResit.Core

Public Class Settings
    Private ComboBoxDataSource As ComboBox
    Private ButtonResetSample As Button
    Private ButtonTestConnection As Button
    Private ButtonImportCsv As Button
    Private CheckBoxShowLookupLabels As CheckBox
    Private _loading As Boolean
    Private _dirty As Boolean
    Private _saving As Boolean
    Private _wired As Boolean
    Private _loadedRestart As SettingsRestartSnapshot

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BuildDataSourceControls()
        LoadFields()
        WireDirtyTracking()
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
        Dim qualificationsButton As New Button With {
            .Text = "Qualifications and units",
            .Location = New Point(500, 26),
            .Size = New Size(180, 27)
        }
        sourceGroup.Controls.Add(ComboBoxDataSource)
        sourceGroup.Controls.Add(ButtonResetSample)
        CheckBoxShowLookupLabels = New CheckBox With {
            .Text = "Show teacher email and unit name on student forms",
            .AutoSize = True,
            .Location = New Point(700, 30)
        }
        sourceGroup.Controls.Add(qualificationsButton)
        sourceGroup.Controls.Add(CheckBoxShowLookupLabels)
        AddHandler CheckBoxShowLookupLabels.CheckedChanged, AddressOf LookupLabelsChanged
        AddHandler qualificationsButton.Click, Sub()
                                                    ApplySharePointFields()
                                                    My.Settings.Save()
                                                    AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
                                                    Dim form As New QualificationsForm()
                                                    form.ShowDialog(Me)
                                                End Sub
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
        InstallDockedButtons()
    End Sub

    ''' <summary>Buttons stay on the window. Only the settings groups scroll.</summary>
    Private Sub InstallDockedButtons()
        AutoScroll = False
        Dim content As New Panel With {.Dock = DockStyle.Fill, .AutoScroll = True, .Name = "SettingsContent"}
        Dim buttons As New Panel With {.Dock = DockStyle.Bottom, .Height = 56, .Name = "SettingsButtons"}
        Dim moving As New List(Of Control)
        For Each control As Control In Controls
            If control Is Button1 OrElse control Is Button2 Then Continue For
            moving.Add(control)
        Next
        For Each control In moving
            Controls.Remove(control)
            content.Controls.Add(control)
        Next
        Controls.Remove(Button1)
        Controls.Remove(Button2)
        Button1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button2.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        buttons.Controls.Add(Button1)
        buttons.Controls.Add(Button2)
        AddHandler buttons.Resize, Sub() PlaceSettingsButtons(buttons)
        Controls.Add(content)
        Controls.Add(buttons)
        PlaceSettingsButtons(buttons)
    End Sub

    Private Sub PlaceSettingsButtons(panel As Panel)
        Button1.Location = New Point(Math.Max(12, panel.ClientSize.Width - 12 - Button1.Width), 10)
        Button2.Location = New Point(Math.Max(12, Button1.Left - 12 - Button2.Width), 10)
    End Sub

    Private Sub DataSourceChanged(sender As Object, e As EventArgs)
        If ButtonResetSample Is Nothing OrElse ComboBoxDataSource Is Nothing Then Return
        ButtonResetSample.Visible = ComboBoxDataSource.SelectedIndex = 0
    End Sub

    Private Sub LoadFields()
        _loading = True
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
        CheckBoxShowLookupLabels.Checked = My.Settings.ShowTeacherEmailAndUnitName
        _loading = False
        _loadedRestart = CurrentRestartSnapshot()
        If Not _wired Then _dirty = False
    End Sub

    Private Function CurrentRestartSnapshot() As SettingsRestartSnapshot
        Dim source = If(ComboBoxDataSource IsNot Nothing AndAlso ComboBoxDataSource.SelectedIndex = 1, "SharePoint", "Sample")
        Return SettingsRestartSnapshot.Capture(
            source,
            TextBoxSiteUrl.Text,
            TextBoxGraphClientId.Text,
            TextBoxGraphTenant.Text,
            CheckBoxSitesSelected.Checked,
            TextBoxBookingsList.Text,
            TextBoxStudentsList.Text,
            TextBoxTeachersList.Text,
            TextBoxUnitsList.Text,
            TextBoxSignatureLibrary.Text)
    End Function

    Private Sub LookupLabelsChanged(sender As Object, e As EventArgs)
        If _loading OrElse CheckBoxShowLookupLabels Is Nothing Then Return
        My.Settings.ShowTeacherEmailAndUnitName = CheckBoxShowLookupLabels.Checked
        My.Settings.Save()
    End Sub

    Private Sub WireDirtyTracking()
        If _wired Then Return
        _wired = True
        AddHandler ComboBoxDataSource.SelectedIndexChanged, AddressOf MarkDirty
        AddHandler TextBoxSiteUrl.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxGraphClientId.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxGraphTenant.TextChanged, AddressOf MarkDirty
        AddHandler CheckBoxSitesSelected.CheckedChanged, AddressOf MarkDirty
        AddHandler TextBoxBookingsList.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxStudentsList.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxTeachersList.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxUnitsList.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxSignatureLibrary.TextChanged, AddressOf MarkDirty
        AddHandler TextBox2.TextChanged, AddressOf MarkDirty
        AddHandler TextBox3.TextChanged, AddressOf MarkDirty
        AddHandler CheckBoxResitLoginOpenAfterSave.CheckedChanged, AddressOf MarkDirty
        AddHandler TextBoxBookingsClientId.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxBookingsTenant.TextChanged, AddressOf MarkDirty
        AddHandler TextBoxBookingsBusinessId.TextChanged, AddressOf MarkDirty
        AddHandler CheckBoxShowPullFromBookingsButton.CheckedChanged, AddressOf MarkDirty
    End Sub

    Private Sub MarkDirty(sender As Object, e As EventArgs)
        If _loading Then Return
        _dirty = True
    End Sub

    Private Sub Settings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If _saving OrElse Not _dirty Then Return
        Dim answer = MessageBox.Show("Save changes?", "Settings", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        Dim choice = SettingsCloseAnswer.Cancel
        If answer = DialogResult.Yes Then choice = SettingsCloseAnswer.Yes
        If answer = DialogResult.No Then choice = SettingsCloseAnswer.No
        Select Case SettingsClosePrompt.Decide(True, choice)
            Case SettingsCloseAction.StayOpen
                e.Cancel = True
            Case SettingsCloseAction.SaveAndExit
                SaveAndExit()
        End Select
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
        SaveAndExit()
    End Sub

    Private Sub SaveAndExit()
        Dim restart = SettingsRestart.NeedsRestart(_loadedRestart, CurrentRestartSnapshot())
        _dirty = False
        _saving = True
        ApplySharePointFields()
        My.Settings.PcLoginMonthlyPassword = TextBox2.Text.Trim()
        My.Settings.ResitLoginWordTemplatePath = TextBox3.Text.Trim()
        My.Settings.ResitLoginOpenDocumentAfterSave = CheckBoxResitLoginOpenAfterSave.Checked
        My.Settings.BookingsAzureClientId = TextBoxBookingsClientId.Text.Trim()
        My.Settings.BookingsAzureTenantId = TextBoxBookingsTenant.Text.Trim()
        My.Settings.BookingsBusinessId = TextBoxBookingsBusinessId.Text.Trim()
        My.Settings.ShowPullFromBookingsButton = CheckBoxShowPullFromBookingsButton.Checked
        My.Settings.ShowTeacherEmailAndUnitName = CheckBoxShowLookupLabels.Checked
        My.Settings.Save()
        _loadedRestart = CurrentRestartSnapshot()
        Me.Hide()
        If restart Then
            MessageBox.Show("Settings Saved, Please Restart Application", "Application Restart", MessageBoxButtons.OK)
            Application.Exit()
        Else
            MessageBox.Show("Settings saved", "Settings", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
