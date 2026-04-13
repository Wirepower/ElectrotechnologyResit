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
        GroupBoxDatabase = New GroupBox()
        LabelSqlCaption = New Label()
        TextBox1 = New TextBox()
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
        GroupBoxDatabase.SuspendLayout()
        GroupBoxResitDoc.SuspendLayout()
        GroupBoxBookings.SuspendLayout()
        SuspendLayout()
        '
        ' GroupBoxDatabase
        '
        GroupBoxDatabase.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GroupBoxDatabase.Controls.Add(LabelSqlCaption)
        GroupBoxDatabase.Controls.Add(TextBox1)
        GroupBoxDatabase.Location = New Point(12, 12)
        GroupBoxDatabase.Name = "GroupBoxDatabase"
        GroupBoxDatabase.Size = New Size(1088, 92)
        GroupBoxDatabase.TabIndex = 0
        GroupBoxDatabase.TabStop = False
        GroupBoxDatabase.Text = "Database"
        '
        ' LabelSqlCaption
        '
        LabelSqlCaption.AutoSize = True
        LabelSqlCaption.Location = New Point(12, 24)
        LabelSqlCaption.Name = "LabelSqlCaption"
        LabelSqlCaption.Size = New Size(99, 15)
        LabelSqlCaption.TabIndex = 0
        LabelSqlCaption.Text = "SQL connection string"
        '
        ' TextBox1
        '
        TextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBox1.Location = New Point(12, 44)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(1064, 23)
        TextBox1.TabIndex = 1
        TextBox1.Text = "Server=DEVSQLCENTRAL.AD.VU.EDU.AU;Integrated Security=True;Connect Timeout=30;Encrypt=True;TrustServerCertificate=True;Multi Subnet Failover=False;"
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
        GroupBoxResitDoc.Location = New Point(12, 114)
        GroupBoxResitDoc.Name = "GroupBoxResitDoc"
        GroupBoxResitDoc.Size = New Size(1088, 198)
        GroupBoxResitDoc.TabIndex = 1
        GroupBoxResitDoc.TabStop = False
        GroupBoxResitDoc.Text = "Resit login Word document"
        '
        ' Label2
        '
        Label2.AutoSize = True
        Label2.Location = New Point(12, 24)
        Label2.Name = "Label2"
        Label2.Size = New Size(298, 15)
        Label2.TabIndex = 0
        Label2.Text = "PC / lab password (monthly) — used in the generated .docx"
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
        GroupBoxBookings.Location = New Point(12, 322)
        GroupBoxBookings.Name = "GroupBoxBookings"
        GroupBoxBookings.Size = New Size(1088, 252)
        GroupBoxBookings.TabIndex = 2
        GroupBoxBookings.TabStop = False
        GroupBoxBookings.Text = "Microsoft Bookings (Microsoft Graph) — optional"
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
        LabelBookingsHint.Text = "Imports appointments for the resit date on the main window. Requires an Azure AD app registration and consent from IT. Leave blank until Graph is configured."
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
        LabelBookingsClient.Size = New Size(140, 15)
        LabelBookingsClient.TabIndex = 2
        LabelBookingsClient.Text = "Azure AD application (client) ID"
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
        LabelBookingsTenant.Location = New Point(12, 140)
        LabelBookingsTenant.Name = "LabelBookingsTenant"
        LabelBookingsTenant.Size = New Size(260, 15)
        LabelBookingsTenant.TabIndex = 4
        LabelBookingsTenant.Text = "Directory (tenant) ID — often ""common"" or your tenant GUID"
        '
        ' TextBoxBookingsTenant
        '
        TextBoxBookingsTenant.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxBookingsTenant.Location = New Point(12, 160)
        TextBoxBookingsTenant.Name = "TextBoxBookingsTenant"
        TextBoxBookingsTenant.Size = New Size(1064, 23)
        TextBoxBookingsTenant.TabIndex = 5
        TextBoxBookingsTenant.Text = "common"
        '
        ' LabelBookingsBusiness
        '
        LabelBookingsBusiness.AutoSize = True
        LabelBookingsBusiness.Location = New Point(12, 192)
        LabelBookingsBusiness.Name = "LabelBookingsBusiness"
        LabelBookingsBusiness.Size = New Size(380, 15)
        LabelBookingsBusiness.TabIndex = 6
        LabelBookingsBusiness.Text = "Booking business ID (optional — leave blank to use the first calendar you can access)"
        '
        ' TextBoxBookingsBusinessId
        '
        TextBoxBookingsBusinessId.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxBookingsBusinessId.Location = New Point(12, 212)
        TextBoxBookingsBusinessId.Name = "TextBoxBookingsBusinessId"
        TextBoxBookingsBusinessId.Size = New Size(1064, 23)
        TextBoxBookingsBusinessId.TabIndex = 7
        '
        ' Button1
        '
        Button1.Anchor = AnchorStyles.Bottom
        Button1.Location = New Point(586, 592)
        Button1.Name = "Button1"
        Button1.Size = New Size(150, 37)
        Button1.TabIndex = 3
        Button1.Text = "Save and Exit"
        Button1.UseVisualStyleBackColor = True
        '
        ' Button2
        '
        Button2.Anchor = AnchorStyles.Bottom
        Button2.Location = New Point(381, 592)
        Button2.Name = "Button2"
        Button2.Size = New Size(150, 37)
        Button2.TabIndex = 4
        Button2.Text = "Close without Saving"
        Button2.UseVisualStyleBackColor = True
        '
        ' Settings
        '
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1112, 648)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(GroupBoxBookings)
        Controls.Add(GroupBoxResitDoc)
        Controls.Add(GroupBoxDatabase)
        MinimumSize = New Size(900, 620)
        Name = "Settings"
        StartPosition = FormStartPosition.CenterParent
        Text = "Settings"
        GroupBoxDatabase.ResumeLayout(False)
        GroupBoxDatabase.PerformLayout()
        GroupBoxResitDoc.ResumeLayout(False)
        GroupBoxResitDoc.PerformLayout()
        GroupBoxBookings.ResumeLayout(False)
        GroupBoxBookings.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents GroupBoxDatabase As GroupBox
    Friend WithEvents LabelSqlCaption As Label
    Friend WithEvents TextBox1 As TextBox
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
