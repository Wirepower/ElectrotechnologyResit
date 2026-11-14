Imports System.Diagnostics
Imports System.IO
Imports ElectrotechnologyResit.Core

''' <summary>
''' First-run SharePoint setup. Safe to run again from Settings. It never deletes lists, columns, or rows.
''' </summary>
Public Class SetupWizard
    Inherits Form

    Private ReadOnly _panels(4) As Panel
    Private ReadOnly _titles As String() = {
        "Set up SharePoint",
        "Site and list names",
        "Sign in",
        "Create lists",
        "Import lookup lists (optional)"
    }
    Private _step As Integer
    Private _signedIn As Boolean
    Private _listsReady As Boolean
    Private _createStarted As Boolean

    Private lblTitle As Label
    Private btnBack As Button
    Private btnNext As Button
    Private btnCancel As Button

    Private txtSite As TextBox
    Private txtClient As TextBox
    Private txtTenant As TextBox
    Private chkSelected As CheckBox
    Private txtBookings As TextBox
    Private txtStudents As TextBox
    Private txtTeachers As TextBox
    Private txtUnits As TextBox
    Private txtQualifications As TextBox
    Private txtSignature As TextBox

    Private lblSignIn As Label
    Private btnSignIn As Button

    Private txtCreateLog As TextBox
    Private btnRetryCreate As Button

    Private txtStudentsFile As TextBox
    Private txtTeachersFile As TextBox
    Private txtUnitsFile As TextBox
    Private lblImport As Label

    Public Sub New()
        Text = "SharePoint setup"
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(720, 560)
        Font = New Font("Segoe UI", 9.0F)
        BuildChrome()
        BuildPages()
        LoadDefaults()
        ShowStep(0)
    End Sub

    Private Sub BuildChrome()
        lblTitle = New Label With {.AutoSize = False, .Location = New Point(16, 12), .Size = New Size(688, 28), .Font = New Font(Font, FontStyle.Bold), .Text = _titles(0)}
        btnBack = New Button With {.Text = "Back", .Location = New Point(360, 516), .Size = New Size(100, 32)}
        btnNext = New Button With {.Text = "Next", .Location = New Point(468, 516), .Size = New Size(110, 32)}
        btnCancel = New Button With {.Text = "Cancel", .Location = New Point(588, 516), .Size = New Size(116, 32), .DialogResult = DialogResult.Cancel}
        Controls.Add(lblTitle)
        Controls.Add(btnBack)
        Controls.Add(btnNext)
        Controls.Add(btnCancel)
        CancelButton = btnCancel
        AddHandler btnBack.Click, AddressOf Back_Click
        AddHandler btnNext.Click, AddressOf Next_Click
    End Sub

    Private Sub BuildPages()
        For index = 0 To 4
            _panels(index) = New Panel With {.Location = New Point(16, 48), .Size = New Size(688, 456), .Visible = False}
            Controls.Add(_panels(index))
        Next
        BuildWelcome()
        BuildDetails()
        BuildSignIn()
        BuildCreate()
        BuildImport()
    End Sub

    Private Sub BuildWelcome()
        Dim text As New Label With {
            .AutoSize = False,
            .Location = New Point(8, 8),
            .Size = New Size(660, 280),
            .Text = "Electrotechnology Resit keeps the night's bookings in SharePoint, under your own sign-in." & vbCrLf & vbCrLf &
                "This wizard asks for the site address, signs you in, and creates any missing lists and columns:" & vbCrLf &
                "ResitBookings (one row per student ID), Students, Teachers, Units, and a library for the email signature image." & vbCrLf & vbCrLf &
                "You can run it again later from Settings. It only adds what is missing. It does not delete lists, columns, or rows." & vbCrLf & vbCrLf &
                "The monthly PC/lab password is typed in Settings. It is not stored in SharePoint and it is not part of this wizard."
        }
        _panels(0).Controls.Add(text)
    End Sub

    Private Sub BuildDetails()
        Dim page = _panels(1)
        page.Controls.Add(FieldLabel("SharePoint site URL", 8))
        txtSite = FieldBox(28)
        page.Controls.Add(txtSite)
        page.Controls.Add(FieldLabel("Azure AD application (client) ID", 60))
        txtClient = FieldBox(80)
        txtClient.Width = 320
        page.Controls.Add(txtClient)
        page.Controls.Add(New Label With {.AutoSize = True, .Location = New Point(360, 60), .Text = "Directory (tenant) ID"})
        txtTenant = New TextBox With {.Location = New Point(360, 80), .Width = 300}
        page.Controls.Add(txtTenant)
        chkSelected = New CheckBox With {
            .AutoSize = True,
            .Location = New Point(8, 116),
            .MaximumSize = New Size(660, 0),
            .Text = "Use Sites.Selected instead of Sites.ReadWrite.All and Sites.Manage.All. Leave this unticked unless IT has granted this app permission to create lists on the site, not only to edit rows."
        }
        page.Controls.Add(chkSelected)
        page.Controls.Add(FieldLabel("Resit bookings list", 168))
        txtBookings = New TextBox With {.Location = New Point(8, 188), .Width = 320}
        page.Controls.Add(txtBookings)
        page.Controls.Add(New Label With {.AutoSize = True, .Location = New Point(360, 168), .Text = "Students list"})
        txtStudents = New TextBox With {.Location = New Point(360, 188), .Width = 300}
        page.Controls.Add(txtStudents)
        page.Controls.Add(FieldLabel("Teachers list", 220))
        txtTeachers = New TextBox With {.Location = New Point(8, 240), .Width = 320}
        page.Controls.Add(txtTeachers)
        page.Controls.Add(New Label With {.AutoSize = True, .Location = New Point(360, 220), .Text = "Units list"})
        txtUnits = New TextBox With {.Location = New Point(360, 240), .Width = 300}
        page.Controls.Add(txtUnits)
        page.Controls.Add(FieldLabel("Qualifications list", 272))
        txtQualifications = New TextBox With {.Location = New Point(8, 292), .Width = 320}
        page.Controls.Add(txtQualifications)
        page.Controls.Add(FieldLabel("Signature image library", 324))
        txtSignature = FieldBox(344)
        page.Controls.Add(txtSignature)
        AddHandler txtSite.TextChanged, Sub() _signedIn = False
        AddHandler txtClient.TextChanged, Sub() _signedIn = False
    End Sub

    Private Sub BuildSignIn()
        lblSignIn = New Label With {
            .AutoSize = False,
            .Location = New Point(8, 8),
            .Size = New Size(660, 120),
            .Text = "Sign in with the account that can open this SharePoint site. The redirect URI on the app registration must be http://localhost."
        }
        btnSignIn = New Button With {.Text = "Sign in and check site access", .Location = New Point(8, 140), .Size = New Size(260, 32)}
        AddHandler btnSignIn.Click, AddressOf SignIn_Click
        _panels(2).Controls.Add(lblSignIn)
        _panels(2).Controls.Add(btnSignIn)
    End Sub

    Private Sub BuildCreate()
        txtCreateLog = New TextBox With {.Location = New Point(8, 8), .Size = New Size(660, 360), .Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Vertical}
        btnRetryCreate = New Button With {.Text = "Retry", .Location = New Point(8, 380), .Size = New Size(120, 32)}
        AddHandler btnRetryCreate.Click, AddressOf RetryCreate_Click
        _panels(3).Controls.Add(txtCreateLog)
        _panels(3).Controls.Add(btnRetryCreate)
    End Sub

    Private Sub BuildImport()
        Dim intro As New Label With {
            .AutoSize = False,
            .Location = New Point(8, 8),
            .Size = New Size(660, 72),
            .Text = "Optional. Load Students, Teachers and Units from CSV or Excel so the lookup lists can be filled without SQL. Import adds or updates rows. It does not delete rows that are not in the file."
        }
        _panels(4).Controls.Add(intro)
        txtStudentsFile = AddFileRow(_panels(4), "Students", 88)
        txtTeachersFile = AddFileRow(_panels(4), "Teachers", 148)
        txtUnitsFile = AddFileRow(_panels(4), "Units", 208)
        Dim samples As New Button With {.Text = "Open sample files", .Location = New Point(8, 268), .Size = New Size(160, 32)}
        AddHandler samples.Click, AddressOf OpenSamples_Click
        Dim importButton As New Button With {.Text = "Import selected files", .Location = New Point(180, 268), .Size = New Size(180, 32)}
        AddHandler importButton.Click, AddressOf Import_Click
        lblImport = New Label With {.AutoSize = False, .Location = New Point(8, 312), .Size = New Size(660, 120)}
        _panels(4).Controls.Add(samples)
        _panels(4).Controls.Add(importButton)
        _panels(4).Controls.Add(lblImport)
    End Sub

    Private Shared Function FieldLabel(caption As String, y As Integer) As Label
        Return New Label With {.AutoSize = True, .Location = New Point(8, y), .Text = caption}
    End Function

    Private Shared Function FieldBox(y As Integer) As TextBox
        Return New TextBox With {.Location = New Point(8, y), .Width = 652}
    End Function

    Private Function AddFileRow(page As Panel, caption As String, y As Integer) As TextBox
        page.Controls.Add(New Label With {.AutoSize = True, .Location = New Point(8, y), .Text = caption})
        Dim box As New TextBox With {.Location = New Point(8, y + 18), .Width = 520}
        Dim browse As New Button With {.Text = "Browse…", .Location = New Point(540, y + 16), .Size = New Size(120, 27)}
        AddHandler browse.Click,
            Sub()
                Using dialog As New OpenFileDialog()
                    dialog.Filter = "CSV or Excel (*.csv;*.xlsx)|*.csv;*.xlsx"
                    dialog.Title = "Choose the " & caption & " file"
                    If dialog.ShowDialog(Me) = DialogResult.OK Then box.Text = dialog.FileName
                End Using
            End Sub
        page.Controls.Add(box)
        page.Controls.Add(browse)
        Return box
    End Function

    Private Sub LoadDefaults()
        txtSite.Text = My.Settings.SharePointSiteUrl
        txtClient.Text = My.Settings.GraphClientId
        txtTenant.Text = If(String.IsNullOrWhiteSpace(My.Settings.GraphTenantId), "common", My.Settings.GraphTenantId)
        chkSelected.Checked = My.Settings.GraphUseSitesSelected
        Dim names = AppServices.CurrentNames()
        txtBookings.Text = names.Bookings
        txtStudents.Text = names.Students
        txtTeachers.Text = names.Teachers
        txtUnits.Text = names.Units
        txtQualifications.Text = names.Qualifications
        txtSignature.Text = names.SignatureLibrary
    End Sub

    Private Function ReadNames() As SharePointListNames
        Return SharePointListNames.FromSettings(txtBookings.Text, txtStudents.Text, txtTeachers.Text, txtUnits.Text, txtSignature.Text, txtQualifications.Text)
    End Function

    Private Function SaveDetails() As String
        Dim names = ReadNames()
        txtBookings.Text = names.Bookings
        txtStudents.Text = names.Students
        txtTeachers.Text = names.Teachers
        txtUnits.Text = names.Units
        txtQualifications.Text = names.Qualifications
        txtSignature.Text = names.SignatureLibrary
        Dim problem = SharePointSetupValidation.Validate(txtSite.Text, txtClient.Text, names)
        If problem IsNot Nothing Then Return problem
        My.Settings.SharePointSiteUrl = txtSite.Text.Trim()
        My.Settings.GraphClientId = txtClient.Text.Trim()
        My.Settings.GraphTenantId = If(String.IsNullOrWhiteSpace(txtTenant.Text), "common", txtTenant.Text.Trim())
        My.Settings.GraphUseSitesSelected = chkSelected.Checked
        My.Settings.ResitBookingsListName = names.Bookings
        My.Settings.StudentsListName = names.Students
        My.Settings.TeachersListName = names.Teachers
        My.Settings.UnitsListName = names.Units
        My.Settings.QualificationsListName = names.Qualifications
        My.Settings.SignatureImageLocation = names.SignatureLibrary
        My.Settings.Save()
        AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
        Return Nothing
    End Function

    Private Sub ShowStep(index As Integer)
        _step = index
        lblTitle.Text = _titles(index)
        For panelIndex = 0 To 4
            _panels(panelIndex).Visible = panelIndex = index
        Next
        btnBack.Enabled = index > 0
        btnNext.Text = If(index = 4, "Finish", "Next")
        btnNext.Enabled = CanGoNext()
        If index = 3 AndAlso Not _createStarted Then
            _createStarted = True
            RunCreate()
        End If
    End Sub

    Private Function CanGoNext() As Boolean
        If _step = 2 Then Return _signedIn
        If _step = 3 Then Return _listsReady
        Return True
    End Function

    Private Sub Back_Click(sender As Object, e As EventArgs)
        If _step > 0 Then ShowStep(_step - 1)
    End Sub

    Private Sub Next_Click(sender As Object, e As EventArgs)
        If _step = 1 Then
            Dim problem = SaveDetails()
            If problem IsNot Nothing Then
                MessageBox.Show(problem, "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
        End If
        If _step = 4 Then
            If Not _listsReady Then
                MessageBox.Show("Create the SharePoint lists before finishing.", "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            My.Settings.SharePointSetupCompleted = True
            My.Settings.Save()
            DialogResult = DialogResult.OK
            Close()
            Return
        End If
        ShowStep(_step + 1)
    End Sub

    Private Async Sub SignIn_Click(sender As Object, e As EventArgs)
        Dim problem = SaveDetails()
        If problem IsNot Nothing Then
            lblSignIn.Text = problem
            Return
        End If
        btnSignIn.Enabled = False
        UseWaitCursor = True
        Try
            Dim access = Await AppServices.CheckSiteAccessAsync()
            lblSignIn.Text = access.Message
            _signedIn = access.Ok
            If access.Ok Then _createStarted = False
        Finally
            UseWaitCursor = False
            btnSignIn.Enabled = True
            btnNext.Enabled = CanGoNext()
        End Try
    End Sub

    Private Async Sub RetryCreate_Click(sender As Object, e As EventArgs)
        Await CreateListsAsync()
    End Sub

    Private Async Sub RunCreate()
        Await CreateListsAsync()
    End Sub

    Private Async Function CreateListsAsync() As Task
        btnRetryCreate.Enabled = False
        btnNext.Enabled = False
        UseWaitCursor = True
        txtCreateLog.Text = "Checking SharePoint and creating anything that is missing..."
        Try
            Dim problem = SaveDetails()
            If problem IsNot Nothing Then
                txtCreateLog.Text = problem
                _listsReady = False
                Return
            End If
            Dim report = Await AppServices.Provisioner.EnsureAsync()
            txtCreateLog.Text = report.Summary() & vbCrLf & vbCrLf & "Nothing was deleted."
            _listsReady = report.IsSatisfied
            If _listsReady Then
                Await QualificationCatalog.EnsureSeedAsync(AppServices.Qualifications, AppServices.Units)
                txtCreateLog.AppendText(vbCrLf & "The lists are ready. UEE30820 is added when the qualifications list is empty. You can import lookup files on the next step, or skip that and finish.")
            End If
        Catch ex As Exception
            _listsReady = False
            txtCreateLog.Text = "SharePoint setup could not finish. " & ex.Message & vbCrLf & vbCrLf & "Nothing was deleted."
        Finally
            UseWaitCursor = False
            btnRetryCreate.Enabled = True
            btnNext.Enabled = CanGoNext()
        End Try
    End Function

    Private Sub OpenSamples_Click(sender As Object, e As EventArgs)
        Dim folder = Path.Combine(AppContext.BaseDirectory, "samples")
        If Not Directory.Exists(folder) Then
            MessageBox.Show("The sample files folder was not found next to the application.", "SharePoint setup", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Process.Start(New ProcessStartInfo(folder) With {.UseShellExecute = True})
    End Sub

    Private Async Sub Import_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtStudentsFile.Text) AndAlso String.IsNullOrWhiteSpace(txtTeachersFile.Text) AndAlso String.IsNullOrWhiteSpace(txtUnitsFile.Text) Then
            lblImport.Text = "Choose a file, or click Finish to skip the import."
            Return
        End If
        UseWaitCursor = True
        Try
            Dim lines As New List(Of String)
            If Not String.IsNullOrWhiteSpace(txtStudentsFile.Text) Then
                Dim tally = Await AppServices.LookupImport.ImportStudentsAsync(LookupFileReader.ReadStudents(txtStudentsFile.Text))
                lines.Add(Describe("Students", tally))
            End If
            If Not String.IsNullOrWhiteSpace(txtTeachersFile.Text) Then
                Dim tally = Await AppServices.LookupImport.ImportTeachersAsync(LookupFileReader.ReadTeachers(txtTeachersFile.Text))
                lines.Add(Describe("Teachers", tally))
            End If
            If Not String.IsNullOrWhiteSpace(txtUnitsFile.Text) Then
                Dim tally = Await AppServices.LookupImport.ImportUnitsAsync(LookupFileReader.ReadUnits(txtUnitsFile.Text))
                lines.Add(Describe("Units", tally))
            End If
            lblImport.Text = String.Join(vbCrLf, lines)
        Catch ex As Exception
            lblImport.Text = ex.Message
        Finally
            UseWaitCursor = False
        End Try
    End Sub

    Private Shared Function Describe(caption As String, tally As ImportTally) As String
        Return caption & ": " & tally.Inserted.ToString() & " added, " & tally.Updated.ToString() & " updated, " & tally.Skipped.ToString() & " skipped."
    End Function
End Class
