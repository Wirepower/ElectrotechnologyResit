Imports ElectrotechnologyResit.Core

''' <summary>Add and edit qualifications and their units. The same form is used for sample data and SharePoint.</summary>
Public Class QualificationsForm
    Inherits Form
    Private ReadOnly _list As New ListBox()
    Private ReadOnly _code As New TextBox()
    Private ReadOnly _title As New TextBox()
    Private ReadOnly _version As New TextBox()
    Private ReadOnly _status As New ComboBox()
    Private ReadOnly _suffix As New TextBox()
    Private ReadOnly _rules As New TextBox()
    Private ReadOnly _units As New ListBox()
    Private ReadOnly _unitCode As New TextBox()
    Private ReadOnly _unitTitle As New TextBox()
    Private _loaded As List(Of QualificationRecord)
    Private _editingExisting As Boolean

    Public Sub New()
        Me.Text = "Qualifications and units"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(860, 560)
        MinimumSize = New Size(760, 520)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)

        Dim root As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1, .Padding = New Padding(8)}
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        _list.Dock = DockStyle.Fill
        root.Controls.Add(_list, 0, 0)

        Dim editor As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 12, .Padding = New Padding(8, 0, 0, 0)}
        editor.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150))
        editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        AddRow(editor, 0, "Code", _code)
        AddRow(editor, 1, "Title", _title)
        AddRow(editor, 2, "Version", _version)
        _status.DropDownStyle = ComboBoxStyle.DropDownList
        _status.Items.AddRange(New Object() {QualificationStatus.Draft, QualificationStatus.Active, QualificationStatus.Superseded})
        AddRow(editor, 3, "Status", _status)
        AddRow(editor, 4, "Class group suffix", _suffix)
        _rules.Multiline = True
        _rules.ScrollBars = ScrollBars.Vertical
        _rules.Height = 80
        AddRow(editor, 5, "Abbreviation rules", _rules)
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 90))

        Dim saveQualification As New Button With {.Text = "Save qualification", .Dock = DockStyle.Left, .Width = 160, .Height = 28}
        Dim addNew As New Button With {.Text = "New", .Dock = DockStyle.Left, .Width = 80, .Height = 28}
        Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.LeftToRight}
        buttons.Controls.Add(saveQualification)
        buttons.Controls.Add(addNew)
        editor.Controls.Add(buttons, 1, 6)
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

        _units.Dock = DockStyle.Fill
        editor.Controls.Add(New Label With {.Text = "Units", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 7)
        editor.Controls.Add(_units, 1, 7)
        editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        AddRow(editor, 8, "Unit code", _unitCode)
        AddRow(editor, 9, "Unit title", _unitTitle)
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
        Dim saveUnit As New Button With {.Text = "Save unit", .Width = 120, .Height = 28}
        Dim importUnits As New Button With {.Text = "Import units...", .Width = 120, .Height = 28}
        Dim unitButtons As New FlowLayoutPanel With {.Dock = DockStyle.Fill}
        unitButtons.Controls.Add(saveUnit)
        unitButtons.Controls.Add(importUnits)
        editor.Controls.Add(unitButtons, 1, 10)
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))
        Dim hint As New Label With {
            .Text = "One line per rule: prefix UEECD=CD or override CODE=ABBR. Only Active qualifications are offered for a new booking. Keep rules under 255 characters.",
            .AutoSize = False,
            .Dock = DockStyle.Fill
        }
        editor.Controls.Add(hint, 1, 11)
        editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
        root.Controls.Add(editor, 1, 0)
        Controls.Add(root)

        AddHandler _list.SelectedIndexChanged, AddressOf SelectedChanged
        AddHandler saveQualification.Click, AddressOf SaveQualificationClick
        AddHandler addNew.Click, Sub() ClearEditor(True)
        AddHandler saveUnit.Click, AddressOf SaveUnitClick
        AddHandler importUnits.Click, AddressOf ImportUnitsClick
        AddHandler MyBase.Load, AddressOf FormLoad
    End Sub

    Private Shared Sub AddRow(table As TableLayoutPanel, row As Integer, caption As String, editor As Control)
        table.Controls.Add(New Label With {.Text = caption, .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, row)
        editor.Dock = DockStyle.Fill
        table.Controls.Add(editor, 1, row)
    End Sub

    Private Async Function ReloadAsync() As Task
        Try
            AppServices.ConfigureFromSettings(Function() CType(Me, IWin32Window))
            _loaded = (Await AppServices.Qualifications.GetAllAsync()).ToList()
            _list.Items.Clear()
            For Each item In _loaded
                _list.Items.Add(item.QualificationCode & "  " & item.Status)
            Next
            If _list.Items.Count > 0 Then _list.SelectedIndex = 0 Else ClearEditor(True)
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Function

    Private Async Sub SelectedChanged(sender As Object, e As EventArgs)
        If _list.SelectedIndex < 0 OrElse _loaded Is Nothing OrElse _list.SelectedIndex >= _loaded.Count Then Return
        Dim item = _loaded(_list.SelectedIndex)
        _editingExisting = True
        _code.ReadOnly = True
        _code.Text = item.QualificationCode
        _title.Text = item.Title
        _version.Text = item.Version
        _status.SelectedItem = QualificationStatus.Normalise(item.Status)
        _suffix.Text = item.ClassGroupSuffix
        _rules.Text = AbbreviationRules.Format(item.Rules)
        Await LoadUnitsAsync(item.QualificationCode)
    End Sub

    Private Async Function LoadUnitsAsync(qualificationCode As String) As Task
        _units.Items.Clear()
        Dim units = Await AppServices.Units.GetAllAsync()
        For Each unit In units
            If String.Equals(If(unit.QualificationCode, ""), qualificationCode, StringComparison.OrdinalIgnoreCase) Then
                _units.Items.Add(unit.UnitCode & "  " & unit.UnitTitle)
            End If
        Next
    End Function

    Private Sub ClearEditor(isNew As Boolean)
        _editingExisting = False
        _code.ReadOnly = False
        _code.Text = ""
        _title.Text = ""
        _version.Text = "1"
        _status.SelectedItem = QualificationStatus.Draft
        _suffix.Text = ""
        _rules.Text = ""
        _units.Items.Clear()
        _unitCode.Text = ""
        _unitTitle.Text = ""
        If isNew Then _list.ClearSelected()
    End Sub

    Private Async Sub FormLoad(sender As Object, e As EventArgs)
        Await ReloadAsync()
    End Sub

    Private Async Sub SaveQualificationClick(sender As Object, e As EventArgs)
        Try
            Dim rules = AbbreviationRules.Parse(_rules.Text)
            If AbbreviationRules.Format(rules).Length > 255 Then
                MessageBox.Show(Me, "Abbreviation rules must be 255 characters or fewer.", "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Await AppServices.Qualifications.UpsertAsync(New QualificationRecord With {
                .QualificationCode = _code.Text.Trim(),
                .Title = _title.Text.Trim(),
                .Version = _version.Text.Trim(),
                .Status = Convert.ToString(_status.SelectedItem),
                .ClassGroupSuffix = _suffix.Text.Trim(),
                .Rules = rules
            })
            Await ReloadAsync()
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Async Sub SaveUnitClick(sender As Object, e As EventArgs)
        If _code.Text.Trim().Length = 0 Then
            MessageBox.Show(Me, "Save the qualification before adding a unit.", "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Await AppServices.Units.UpsertAsync(New UnitRecord With {
                .QualificationCode = _code.Text.Trim(),
                .UnitCode = _unitCode.Text.Trim(),
                .UnitTitle = _unitTitle.Text.Trim()
            })
            _unitCode.Text = ""
            _unitTitle.Text = ""
            Await LoadUnitsAsync(_code.Text.Trim())
        Catch ex As Exception
            MessageBox.Show(Me, ex.Message, "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Async Sub ImportUnitsClick(sender As Object, e As EventArgs)
        If _code.Text.Trim().Length = 0 Then
            MessageBox.Show(Me, "Choose or save a qualification first.", "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Using dialog As New OpenFileDialog()
            dialog.Filter = "CSV or Excel (*.csv;*.xlsx)|*.csv;*.xlsx"
            dialog.Title = "Import units for " & _code.Text.Trim()
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                Dim tally = Await AppServices.LookupImport.ImportUnitsAsync(LookupFileReader.ReadUnits(dialog.FileName, _code.Text.Trim()))
                Await LoadUnitsAsync(_code.Text.Trim())
                MessageBox.Show(Me, "Imported " & tally.Inserted.ToString() & " new and updated " & tally.Updated.ToString() & ".", "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, "Qualifications", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Using
    End Sub
End Class
