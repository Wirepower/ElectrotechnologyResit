<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
    Inherits System.Windows.Forms.Form

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

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        GroupBoxSharePoint = New GroupBox()
        LabelSharePointHint = New Label()
        LabelSite = New Label()
        TextBoxSiteUrl = New TextBox()
        LabelGraphClient = New Label()
        TextBoxGraphClientId = New TextBox()
        LabelGraphTenant = New Label()
        TextBoxGraphTenant = New TextBox()
        CheckBoxSitesSelected = New CheckBox()
        LabelBookingsList = New Label()
        TextBoxBookingsList = New TextBox()
        LabelStudentsList = New Label()
        TextBoxStudentsList = New TextBox()
        LabelTeachersList = New Label()
        TextBoxTeachersList = New TextBox()
        LabelUnitsList = New Label()
        TextBoxUnitsList = New TextBox()
        LabelSignature = New Label()
        TextBoxSignatureLibrary = New TextBox()
        ButtonWizard = New Button()
        GroupBoxResitDoc = New GroupBox()
        Label2 = New Label()
        TextBox2 = New TextBox()
        Label3 = New Label()
        TextBox3 = New TextBox()
        Button3 = New Button()
        CheckBoxResitLoginOpenAfterSave = New CheckBox()
        GroupBoxBookings = New GroupBox()
        LabelBookingsHint = New Label()
        CheckBoxShowPullFromBookingsButton = New CheckBox()
        LabelBookingsClient = New Label()
        TextBoxBookingsClientId = New TextBox()
        LabelBookingsTenant = New Label()
        TextBoxBookingsTenant = New TextBox()
        LabelBookingsBusiness = New Label()
        TextBoxBookingsBusinessId = New TextBox()
        Button1 = New Button()
        Button2 = New Button()
        GroupBoxSharePoint.SuspendLayout()
        GroupBoxResitDoc.SuspendLayout()
        GroupBoxBookings.SuspendLayout()
        SuspendLayout()
        '
        ' GroupBoxSharePoint
        '
        GroupBoxSharePoint.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GroupBoxSharePoint.Controls.Add(LabelSharePointHint)
        GroupBoxSharePoint.Controls.Add(LabelSite)
        GroupBoxSharePoint.Controls.Add(TextBoxSiteUrl)
        GroupBoxSharePoint.Controls.Add(LabelGraphClient)
        GroupBoxSharePoint.Controls.Add(TextBoxGraphClientId)
        GroupBoxSharePoint.Controls.Add(LabelGraphTenant)
        GroupBoxSharePoint.Controls.Add(TextBoxGraphTenant)
        GroupBoxSharePoint.Controls.Add(CheckBoxSitesSelected)
        GroupBoxSharePoint.Controls.Add(LabelBookingsList)
        GroupBoxSharePoint.Controls.Add(TextBoxBookingsList)
        GroupBoxSharePoint.Controls.Add(LabelStudentsList)
        GroupBoxSharePoint.Controls.Add(TextBoxStudentsList)
        GroupBoxSharePoint.Controls.Add(LabelTeachersList)
        GroupBoxSharePoint.Controls.Add(TextBoxTeachersList)
        GroupBoxSharePoint.Controls.Add(LabelUnitsList)
        GroupBoxSharePoint.Controls.Add(TextBoxUnitsList)
        GroupBoxSharePoint.Controls.Add(LabelSignature)
        GroupBoxSharePoint.Controls.Add(TextBoxSignatureLibrary)
        GroupBoxSharePoint.Controls.Add(ButtonWizard)
        GroupBoxSharePoint.Location = New Point(12, 12)
        GroupBoxSharePoint.Name = "GroupBoxSharePoint"
        GroupBoxSharePoint.Size = New Size(1088, 372)
        GroupBoxSharePoint.TabIndex = 0
        GroupBoxSharePoint.TabStop = False
        GroupBoxSharePoint.Text = "SharePoint"
        '
        ' LabelSharePointHint
        '
        LabelSharePointHint.AutoSize = True
        LabelSharePointHint.ForeColor = SystemColors.GrayText
        LabelSharePointHint.Location = New Point(12, 22)
        LabelSharePointHint.MaximumSize = New Size(1060, 0)
        LabelSharePointHint.Name = "LabelSharePointHint"
        LabelSharePointHint.Size = New Size(900, 30)
        LabelSharePointHint.TabIndex = 0
        LabelSharePointHint.Text = "Bookings are stored in SharePoint lists for the signed-in user. List names are optional and already filled with the defaults. The setup wizard signs in and creates anything that is missing. It does not delete lists, columns, or rows."
        '
        ' LabelSite
        '
        LabelSite.AutoSize = True
        LabelSite.Location = New Point(12, 58)
        LabelSite.Name = "LabelSite"
        LabelSite.Size = New Size(118, 15)
        LabelSite.TabIndex = 1
        LabelSite.Text = "SharePoint site URL"
        '
        ' TextBoxSiteUrl
        '
        TextBoxSiteUrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxSiteUrl.Location = New Point(12, 76)
        TextBoxSiteUrl.Name = "TextBoxSiteUrl"
        TextBoxSiteUrl.PlaceholderText = "https://tenant.sharepoint.com/sites/YourSite"
        TextBoxSiteUrl.Size = New Size(1064, 23)
        TextBoxSiteUrl.TabIndex = 2
        '
        ' LabelGraphClient
        '
        LabelGraphClient.AutoSize = True
        LabelGraphClient.Location = New Point(12, 108)
        LabelGraphClient.Name = "LabelGraphClient"
        LabelGraphClient.Size = New Size(175, 15)
        LabelGraphClient.TabIndex = 3
        LabelGraphClient.Text = "Azure AD application (client) ID"
        '
        ' TextBoxGraphClientId
        '
        TextBoxGraphClientId.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxGraphClientId.Location = New Point(12, 126)
        TextBoxGraphClientId.Name = "TextBoxGraphClientId"
        TextBoxGraphClientId.Size = New Size(520, 23)
        TextBoxGraphClientId.TabIndex = 4
        '
        ' LabelGraphTenant
        '
        LabelGraphTenant.AutoSize = True
        LabelGraphTenant.Location = New Point(548, 108)
        LabelGraphTenant.Name = "LabelGraphTenant"
        LabelGraphTenant.Size = New Size(110, 15)
        LabelGraphTenant.TabIndex = 5
        LabelGraphTenant.Text = "Directory (tenant) ID"
        '
        ' TextBoxGraphTenant
        '
        TextBoxGraphTenant.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        TextBoxGraphTenant.Location = New Point(548, 126)
        TextBoxGraphTenant.Name = "TextBoxGraphTenant"
        TextBoxGraphTenant.Size = New Size(528, 23)
        TextBoxGraphTenant.TabIndex = 6
        '
        ' CheckBoxSitesSelected
        '
        CheckBoxSitesSelected.AutoSize = True
        CheckBoxSitesSelected.Location = New Point(12, 158)
        CheckBoxSitesSelected.MaximumSize = New Size(1060, 0)
        CheckBoxSitesSelected.Name = "CheckBoxSitesSelected"
        CheckBoxSitesSelected.Size = New Size(760, 34)
        CheckBoxSitesSelected.TabIndex = 7
        CheckBoxSitesSelected.Text = "Use Sites.Selected instead of Sites.ReadWrite.All and Sites.Manage.All. The site grant must allow creating lists, not only editing rows. Leave this unticked unless that grant is in place."
        '
        ' LabelBookingsList
        '
        LabelBookingsList.AutoSize = True
        LabelBookingsList.Location = New Point(12, 200)
        LabelBookingsList.Name = "LabelBookingsList"
        LabelBookingsList.Size = New Size(118, 15)
        LabelBookingsList.TabIndex = 8
        LabelBookingsList.Text = "Resit bookings list"
        '
        ' TextBoxBookingsList
        '
        TextBoxBookingsList.Location = New Point(12, 218)
        TextBoxBookingsList.Name = "TextBoxBookingsList"
        TextBoxBookingsList.Size = New Size(520, 23)
        TextBoxBookingsList.TabIndex = 9
        '
        ' LabelStudentsList
        '
        LabelStudentsList.AutoSize = True
        LabelStudentsList.Location = New Point(548, 200)
        LabelStudentsList.Name = "LabelStudentsList"
        LabelStudentsList.Size = New Size(84, 15)
        LabelStudentsList.TabIndex = 10
        LabelStudentsList.Text = "Students list"
        '
        ' TextBoxStudentsList
        '
        TextBoxStudentsList.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        TextBoxStudentsList.Location = New Point(548, 218)
        TextBoxStudentsList.Name = "TextBoxStudentsList"
        TextBoxStudentsList.Size = New Size(528, 23)
        TextBoxStudentsList.TabIndex = 11
        '
        ' LabelTeachersList
        '
        LabelTeachersList.AutoSize = True
        LabelTeachersList.Location = New Point(12, 248)
        LabelTeachersList.Name = "LabelTeachersList"
        LabelTeachersList.Size = New Size(78, 15)
        LabelTeachersList.TabIndex = 12
        LabelTeachersList.Text = "Teachers list"
        '
        ' TextBoxTeachersList
        '
        TextBoxTeachersList.Location = New Point(12, 266)
        TextBoxTeachersList.Name = "TextBoxTeachersList"
        TextBoxTeachersList.Size = New Size(520, 23)
        TextBoxTeachersList.TabIndex = 13
        '
        ' LabelUnitsList
        '
        LabelUnitsList.AutoSize = True
        LabelUnitsList.Location = New Point(548, 248)
        LabelUnitsList.Name = "LabelUnitsList"
        LabelUnitsList.Size = New Size(58, 15)
        LabelUnitsList.TabIndex = 14
        LabelUnitsList.Text = "Units list"
        '
        ' TextBoxUnitsList
        '
        TextBoxUnitsList.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        TextBoxUnitsList.Location = New Point(548, 266)
        TextBoxUnitsList.Name = "TextBoxUnitsList"
        TextBoxUnitsList.Size = New Size(528, 23)
        TextBoxUnitsList.TabIndex = 15
        '
        ' LabelSignature
        '
        LabelSignature.AutoSize = True
        LabelSignature.Location = New Point(12, 296)
        LabelSignature.Name = "LabelSignature"
        LabelSignature.Size = New Size(280, 15)
        LabelSignature.TabIndex = 16
        LabelSignature.Text = "Email signature image library (newest image is used)"
        '
        ' TextBoxSignatureLibrary
        '
        TextBoxSignatureLibrary.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxSignatureLibrary.Location = New Point(12, 314)
        TextBoxSignatureLibrary.Name = "TextBoxSignatureLibrary"
        TextBoxSignatureLibrary.Size = New Size(760, 23)
        TextBoxSignatureLibrary.TabIndex = 17
        '
        ' ButtonWizard
        '
        ButtonWizard.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        ButtonWizard.Location = New Point(788, 312)
        ButtonWizard.Name = "ButtonWizard"
        ButtonWizard.Size = New Size(288, 27)
        ButtonWizard.TabIndex = 18
        ButtonWizard.Text = "SharePoint setup wizard"
        ButtonWizard.UseVisualStyleBackColor = True
        '
        ' GroupBoxResitDoc
        '
        GroupBoxResitDoc.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GroupBoxResitDoc.Controls.Add(Label2)
        GroupBoxResitDoc.Controls.Add(TextBox2)
        GroupBoxResitDoc.Controls.Add(Label3)
        GroupBoxResitDoc.Controls.Add(TextBox3)
        GroupBoxResitDoc.Controls.Add(Button3)
        GroupBoxResitDoc.Controls.Add(CheckBoxResitLoginOpenAfterSave)
        GroupBoxResitDoc.Location = New Point(12, 392)
        GroupBoxResitDoc.Name = "GroupBoxResitDoc"
        GroupBoxResitDoc.Size = New Size(1088, 168)
        GroupBoxResitDoc.TabIndex = 1
        GroupBoxResitDoc.TabStop = False
        GroupBoxResitDoc.Text = "Resit login Word document"
        '
        ' Label2
        '
        Label2.AutoSize = True
        Label2.Location = New Point(12, 24)
        Label2.Name = "Label2"
        Label2.Size = New Size(420, 15)
        Label2.TabIndex = 0
        Label2.Text = "PC / lab password (monthly). Type it here. It is stored only in your Windows user settings."
        '
        ' TextBox2
        '
        TextBox2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBox2.Location = New Point(12, 44)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(1064, 23)
        TextBox2.TabIndex = 1
        '
        ' Label3
        '
        Label3.AutoSize = True
        Label3.Location = New Point(12, 76)
        Label3.Name = "Label3"
        Label3.Size = New Size(312, 15)
        Label3.TabIndex = 2
        Label3.Text = "Word template path (.docx). A default is created on first run."
        '
        ' TextBox3
        '
        TextBox3.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBox3.Location = New Point(12, 96)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(928, 23)
        TextBox3.TabIndex = 3
        '
        ' Button3
        '
        Button3.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button3.Location = New Point(952, 94)
        Button3.Name = "Button3"
        Button3.Size = New Size(124, 27)
        Button3.TabIndex = 4
        Button3.Text = "Browse…"
        Button3.UseVisualStyleBackColor = True
        '
        ' CheckBoxResitLoginOpenAfterSave
        '
        CheckBoxResitLoginOpenAfterSave.AutoSize = True
        CheckBoxResitLoginOpenAfterSave.Location = New Point(12, 132)
        CheckBoxResitLoginOpenAfterSave.Name = "CheckBoxResitLoginOpenAfterSave"
        CheckBoxResitLoginOpenAfterSave.Size = New Size(520, 19)
        CheckBoxResitLoginOpenAfterSave.TabIndex = 5
        CheckBoxResitLoginOpenAfterSave.Text = "Open the Resit login Word document automatically after it is saved (skip 'Open it now?')"
        '
        ' GroupBoxBookings
        '
        GroupBoxBookings.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GroupBoxBookings.Controls.Add(LabelBookingsHint)
        GroupBoxBookings.Controls.Add(CheckBoxShowPullFromBookingsButton)
        GroupBoxBookings.Controls.Add(LabelBookingsClient)
        GroupBoxBookings.Controls.Add(TextBoxBookingsClientId)
        GroupBoxBookings.Controls.Add(LabelBookingsTenant)
        GroupBoxBookings.Controls.Add(TextBoxBookingsTenant)
        GroupBoxBookings.Controls.Add(LabelBookingsBusiness)
        GroupBoxBookings.Controls.Add(TextBoxBookingsBusinessId)
        GroupBoxBookings.Location = New Point(12, 568)
        GroupBoxBookings.Name = "GroupBoxBookings"
        GroupBoxBookings.Size = New Size(1088, 248)
        GroupBoxBookings.TabIndex = 2
        GroupBoxBookings.TabStop = False
        GroupBoxBookings.Text = "Microsoft Bookings (optional, hidden)"
        '
        ' LabelBookingsHint
        '
        LabelBookingsHint.AutoSize = True
        LabelBookingsHint.ForeColor = SystemColors.GrayText
        LabelBookingsHint.Location = New Point(12, 24)
        LabelBookingsHint.MaximumSize = New Size(1060, 0)
        LabelBookingsHint.Name = "LabelBookingsHint"
        LabelBookingsHint.Size = New Size(1034, 30)
        LabelBookingsHint.TabIndex = 0
        LabelBookingsHint.Text = "The Pull from Bookings button stays hidden until this checkbox is on. Live Bookings is owned by VU Timetable and needs Frank Offer's sign-off. Leave the client ID blank until that is agreed."
        '
        ' CheckBoxShowPullFromBookingsButton
        '
        CheckBoxShowPullFromBookingsButton.AutoSize = True
        CheckBoxShowPullFromBookingsButton.Location = New Point(12, 60)
        CheckBoxShowPullFromBookingsButton.Name = "CheckBoxShowPullFromBookingsButton"
        CheckBoxShowPullFromBookingsButton.Size = New Size(420, 19)
        CheckBoxShowPullFromBookingsButton.TabIndex = 1
        CheckBoxShowPullFromBookingsButton.Text = "Show ""Pull from Bookings"" button on the main window"
        '
        ' LabelBookingsClient
        '
        LabelBookingsClient.AutoSize = True
        LabelBookingsClient.Location = New Point(12, 88)
        LabelBookingsClient.Name = "LabelBookingsClient"
        LabelBookingsClient.Size = New Size(220, 15)
        LabelBookingsClient.TabIndex = 2
        LabelBookingsClient.Text = "Bookings Azure AD application (client) ID"
        '
        ' TextBoxBookingsClientId
        '
        TextBoxBookingsClientId.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxBookingsClientId.Location = New Point(12, 108)
        TextBoxBookingsClientId.Name = "TextBoxBookingsClientId"
        TextBoxBookingsClientId.Size = New Size(1064, 23)
        TextBoxBookingsClientId.TabIndex = 3
        '
        ' LabelBookingsTenant
        '
        LabelBookingsTenant.AutoSize = True
        LabelBookingsTenant.Location = New Point(12, 136)
        LabelBookingsTenant.Name = "LabelBookingsTenant"
        LabelBookingsTenant.Size = New Size(260, 15)
        LabelBookingsTenant.TabIndex = 4
        LabelBookingsTenant.Text = "Directory (tenant) ID — often ""common"" or your tenant GUID"
        '
        ' TextBoxBookingsTenant
        '
        TextBoxBookingsTenant.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxBookingsTenant.Location = New Point(12, 154)
        TextBoxBookingsTenant.Name = "TextBoxBookingsTenant"
        TextBoxBookingsTenant.Size = New Size(1064, 23)
        TextBoxBookingsTenant.TabIndex = 5
        '
        ' LabelBookingsBusiness
        '
        LabelBookingsBusiness.AutoSize = True
        LabelBookingsBusiness.Location = New Point(12, 180)
        LabelBookingsBusiness.Name = "LabelBookingsBusiness"
        LabelBookingsBusiness.Size = New Size(380, 15)
        LabelBookingsBusiness.TabIndex = 6
        LabelBookingsBusiness.Text = "Booking business ID (optional — leave blank to use the first calendar you can access)"
        '
        ' TextBoxBookingsBusinessId
        '
        TextBoxBookingsBusinessId.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxBookingsBusinessId.Location = New Point(12, 200)
        TextBoxBookingsBusinessId.Name = "TextBoxBookingsBusinessId"
        TextBoxBookingsBusinessId.Size = New Size(776, 23)
        TextBoxBookingsBusinessId.TabIndex = 7
        '
        ' Button1
        '
        Button1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button1.Location = New Point(586, 10)
        Button1.Name = "Button1"
        Button1.Size = New Size(150, 37)
        Button1.TabIndex = 3
        Button1.Text = "Save and Exit"
        Button1.UseVisualStyleBackColor = True
        '
        ' Button2
        '
        Button2.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Button2.Location = New Point(424, 10)
        Button2.Name = "Button2"
        Button2.Size = New Size(150, 37)
        Button2.TabIndex = 4
        Button2.Text = "Cancel"
        Button2.UseVisualStyleBackColor = True
        '
        ' Settings
        '
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        AutoScroll = False
        ClientSize = New Size(1112, 888)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(GroupBoxBookings)
        Controls.Add(GroupBoxResitDoc)
        Controls.Add(GroupBoxSharePoint)
        MinimumSize = New Size(900, 700)
        Name = "Settings"
        StartPosition = FormStartPosition.CenterParent
        Text = "Settings"
        GroupBoxSharePoint.ResumeLayout(False)
        GroupBoxSharePoint.PerformLayout()
        GroupBoxResitDoc.ResumeLayout(False)
        GroupBoxResitDoc.PerformLayout()
        GroupBoxBookings.ResumeLayout(False)
        GroupBoxBookings.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents GroupBoxSharePoint As GroupBox
    Friend WithEvents LabelSharePointHint As Label
    Friend WithEvents LabelSite As Label
    Friend WithEvents TextBoxSiteUrl As TextBox
    Friend WithEvents LabelGraphClient As Label
    Friend WithEvents TextBoxGraphClientId As TextBox
    Friend WithEvents LabelGraphTenant As Label
    Friend WithEvents TextBoxGraphTenant As TextBox
    Friend WithEvents CheckBoxSitesSelected As CheckBox
    Friend WithEvents LabelBookingsList As Label
    Friend WithEvents TextBoxBookingsList As TextBox
    Friend WithEvents LabelStudentsList As Label
    Friend WithEvents TextBoxStudentsList As TextBox
    Friend WithEvents LabelTeachersList As Label
    Friend WithEvents TextBoxTeachersList As TextBox
    Friend WithEvents LabelUnitsList As Label
    Friend WithEvents TextBoxUnitsList As TextBox
    Friend WithEvents LabelSignature As Label
    Friend WithEvents TextBoxSignatureLibrary As TextBox
    Friend WithEvents ButtonWizard As Button
    Friend WithEvents GroupBoxResitDoc As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents CheckBoxResitLoginOpenAfterSave As CheckBox
    Friend WithEvents GroupBoxBookings As GroupBox
    Friend WithEvents LabelBookingsHint As Label
    Friend WithEvents CheckBoxShowPullFromBookingsButton As CheckBox
    Friend WithEvents LabelBookingsClient As Label
    Friend WithEvents TextBoxBookingsClientId As TextBox
    Friend WithEvents LabelBookingsTenant As Label
    Friend WithEvents TextBoxBookingsTenant As TextBox
    Friend WithEvents LabelBookingsBusiness As Label
    Friend WithEvents TextBoxBookingsBusinessId As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
End Class
