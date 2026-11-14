<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
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
        DataGridView1 = New DataGridView()
        Button7 = New Button()
        Label2 = New Label()
        Label3 = New Label()
        Label5 = New Label()
        CheckBox1 = New RadioButton()
        CheckBox2 = New RadioButton()
        statusLabel = New Label()
        DateTimePicker1 = New DateTimePicker()
        Button8 = New Button()
        Button9 = New Button()
        ListView1 = New ListView()
        Label6 = New Label()
        ListView2 = New ListView()
        Label7 = New Label()
        Button11 = New Button()
        Button12 = New Button()
        Button13 = New Button()
        LabelSampleBadge = New Label()
        LabelStudentCount = New Label()
        LabelClassCount = New Label()
        LabelBookedCount = New Label()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()

        _mainLayout = New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 5, .Padding = New Padding(8)}
        Dim root = _mainLayout
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 128))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 196))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))

        Dim header As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 5, .RowCount = 1}
        header.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
        header.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
        Label2.Text = "Electrotechnology Resit v2.0"
        Label2.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        Label2.AutoSize = True
        Label2.Anchor = AnchorStyles.Left
        Label2.Margin = New Padding(0, 8, 12, 0)
        LabelSampleBadge.Text = "SAMPLE DATA"
        LabelSampleBadge.AutoSize = True
        LabelSampleBadge.Anchor = AnchorStyles.Left
        LabelSampleBadge.BackColor = Color.DarkRed
        LabelSampleBadge.ForeColor = Color.White
        LabelSampleBadge.Padding = New Padding(8, 4, 8, 4)
        LabelSampleBadge.Margin = New Padding(0, 10, 0, 0)
        LabelSampleBadge.Visible = False
        Button11.Text = "Settings"
        Button11.Dock = DockStyle.Fill
        Button11.Margin = New Padding(4, 8, 4, 8)
        Button7.Text = "Exit"
        Button7.Dock = DockStyle.Fill
        Button7.Margin = New Padding(4, 8, 0, 8)
        header.Controls.Add(Label2, 0, 0)
        header.Controls.Add(LabelSampleBadge, 1, 0)
        header.Controls.Add(Button11, 3, 0)
        header.Controls.Add(Button7, 4, 0)
        root.Controls.Add(header, 0, 0)

        Dim toolbar As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .AutoScroll = True
        }
        toolbar.Controls.Add(BuildStudentsGroup())
        toolbar.Controls.Add(BuildNightGroup())
        toolbar.Controls.Add(BuildOutputsGroup())
        toolbar.Controls.Add(BuildEmailGroup())
        root.Controls.Add(toolbar, 0, 1)

        DataGridView1.Dock = DockStyle.Fill
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToOrderColumns = True
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.MultiSelect = False
        DataGridView1.ShowCellToolTips = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        root.Controls.Add(DataGridView1, 0, 2)

        Dim energyspace As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        energyspace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        energyspace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        energyspace.Controls.Add(BuildUnitPanel("ENERGYSPACE CLASS CREATED", ListView1, Label6, LabelClassCount), 0, 0)
        energyspace.Controls.Add(BuildUnitPanel("ENERGYSPACE ASSESSMENT BOOKED", ListView2, Label7, LabelBookedCount), 1, 0)
        root.Controls.Add(energyspace, 0, 3)

        Dim status As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 3, .RowCount = 1}
        status.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40))
        status.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        status.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35))
        statusLabel.Text = ""
        statusLabel.AutoSize = False
        statusLabel.Dock = DockStyle.Fill
        statusLabel.TextAlign = ContentAlignment.MiddleLeft
        LabelStudentCount.Text = ""
        LabelStudentCount.AutoSize = False
        LabelStudentCount.Dock = DockStyle.Fill
        LabelStudentCount.TextAlign = ContentAlignment.MiddleLeft
        Label3.Text = "Created by Frank Offer, Contact me via email on Frank.Offer@vu.edu.au, Version 2.0"
        Label3.AutoSize = False
        Label3.Dock = DockStyle.Fill
        Label3.TextAlign = ContentAlignment.MiddleRight
        Label5.Visible = False
        status.Controls.Add(statusLabel, 0, 0)
        status.Controls.Add(LabelStudentCount, 1, 0)
        status.Controls.Add(Label3, 2, 0)
        root.Controls.Add(status, 0, 4)

        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1280, 800)
        MinimumSize = New Size(1100, 700)
        StartPosition = FormStartPosition.CenterScreen
        Controls.Add(root)
        Name = "Main"
        Text = "Electrotechnology Resit v2.0"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Private Function BuildStudentsGroup() As GroupBox
        Dim box As New GroupBox With {.Text = "STUDENTS", .Size = New Size(250, 112), .Margin = New Padding(0, 0, 8, 0)}
        Button1.Text = "New"
        Button1.Location = New Point(12, 28)
        Button1.Size = New Size(70, 32)
        Button2.Text = "Edit"
        Button2.Location = New Point(88, 28)
        Button2.Size = New Size(70, 32)
        Button3.Text = "Delete"
        Button3.Location = New Point(164, 28)
        Button3.Size = New Size(70, 32)
        Button13.Text = "Pull from Bookings"
        Button13.Location = New Point(12, 68)
        Button13.Size = New Size(222, 28)
        Button13.Visible = False
        box.Controls.Add(Button1)
        box.Controls.Add(Button2)
        box.Controls.Add(Button3)
        box.Controls.Add(Button13)
        Return box
    End Function

    Private Function BuildNightGroup() As GroupBox
        Dim box As New GroupBox With {.Text = "RESIT NIGHT", .Size = New Size(300, 112), .Margin = New Padding(0, 0, 8, 0)}
        DateTimePicker1.Format = DateTimePickerFormat.Short
        DateTimePicker1.Location = New Point(12, 32)
        DateTimePicker1.Size = New Size(130, 27)
        Button8.Text = "Show All"
        Button8.Location = New Point(150, 30)
        Button8.Size = New Size(136, 28)
        Button4.Text = "Delete entries for date"
        Button4.Location = New Point(12, 68)
        Button4.Size = New Size(274, 28)
        Button4.FlatStyle = FlatStyle.Flat
        Button4.BackColor = Color.FromArgb(176, 42, 42)
        Button4.ForeColor = Color.White
        Button4.UseVisualStyleBackColor = False
        box.Controls.Add(DateTimePicker1)
        box.Controls.Add(Button8)
        box.Controls.Add(Button4)
        Return box
    End Function

    Private Function BuildOutputsGroup() As GroupBox
        Dim box As New GroupBox With {.Text = "OUTPUTS", .Size = New Size(180, 112), .Margin = New Padding(0, 0, 8, 0)}
        Button9.Text = "Export to Excel"
        Button9.Location = New Point(12, 28)
        Button9.Size = New Size(156, 32)
        Button9.Visible = False
        Button12.Text = "Resit Login Details"
        Button12.Location = New Point(12, 68)
        Button12.Size = New Size(156, 32)
        Button12.Visible = False
        box.Controls.Add(Button9)
        box.Controls.Add(Button12)
        Return box
    End Function

    Private Function BuildEmailGroup() As GroupBox
        Dim box As New GroupBox With {.Text = "EMAIL TEACHERS", .Size = New Size(360, 112), .Margin = New Padding(0, 0, 8, 0)}
        Button5.Text = "Email teachers of their students"
        Button5.Location = New Point(12, 22)
        Button5.Size = New Size(336, 34)
        Button5.FlatStyle = FlatStyle.Flat
        Button5.BackColor = Color.FromArgb(0, 90, 180)
        Button5.ForeColor = Color.White
        Button5.UseVisualStyleBackColor = False
        CheckBox1.Text = "Display / send manually"
        CheckBox1.Location = New Point(12, 64)
        CheckBox1.AutoSize = True
        CheckBox1.Checked = True
        CheckBox2.Text = "Send - do not display"
        CheckBox2.Location = New Point(180, 64)
        CheckBox2.AutoSize = True
        box.Controls.Add(Button5)
        box.Controls.Add(CheckBox1)
        box.Controls.Add(CheckBox2)
        Return box
    End Function

    Private Function BuildUnitPanel(title As String, list As ListView, heading As Label, count As Label) As Panel
        Dim panel As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 3, .Padding = New Padding(4)}
        panel.RowStyles.Add(New RowStyle(SizeType.Absolute, 22))
        panel.RowStyles.Add(New RowStyle(SizeType.Absolute, 20))
        panel.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        heading.Text = title
        heading.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        heading.Dock = DockStyle.Fill
        count.Text = "0 of 0 units ticked"
        count.Dock = DockStyle.Fill
        list.CheckBoxes = True
        list.View = View.Details
        list.FullRowSelect = True
        list.GridLines = True
        list.HeaderStyle = ColumnHeaderStyle.Nonclickable
        list.Scrollable = False
        list.Dock = DockStyle.Top
        list.Height = 28
        list.Columns.Add("Unit", 120)
        list.Columns.Add("Name", 320)
        panel.Controls.Add(heading, 0, 0)
        panel.Controls.Add(count, 0, 1)
        panel.Controls.Add(list, 0, 2)
        Return panel
    End Function

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button7 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents CheckBox1 As RadioButton
    Friend WithEvents CheckBox2 As RadioButton
    Friend WithEvents statusLabel As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents Label6 As Label
    Friend WithEvents ListView2 As ListView
    Friend WithEvents Label7 As Label
    Friend WithEvents Button11 As Button
    Friend WithEvents Button12 As Button
    Friend WithEvents Button13 As Button
    Friend LabelSampleBadge As Label
    Friend LabelStudentCount As Label
    Friend LabelClassCount As Label
    Friend LabelBookedCount As Label
    Private _mainLayout As TableLayoutPanel
End Class
