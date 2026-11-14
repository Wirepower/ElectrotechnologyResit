Friend NotInheritable Class StudentDialogLayout
    Private Sub New()
    End Sub

    Friend Shared Sub Prepare(form As Form)
        form.Font = New Font("Segoe UI", 9.0F)
        form.BackColor = Color.White
        form.AcceptButton = Nothing
    End Sub

    Friend Shared Function Field(caption As String, editor As Control, note As Label) As Control
        Dim stack As New TableLayoutPanel With {
            .ColumnCount = 1,
            .RowCount = If(note Is Nothing, 2, 3),
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .Margin = New Padding(0, 0, 0, 4)
        }
        stack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        stack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        If note IsNot Nothing Then stack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        stack.Controls.Add(New Label With {.Text = caption, .AutoSize = True, .Margin = New Padding(0, 6, 0, 2)}, 0, 0)
        editor.Dock = DockStyle.Top
        editor.Margin = New Padding(0)
        If editor.Height < 23 Then editor.Height = 27
        stack.Controls.Add(editor, 0, 1)
        If note IsNot Nothing Then
            note.AutoSize = True
            note.ForeColor = Color.DimGray
            note.Margin = New Padding(0, 2, 0, 0)
            note.Visible = False
            stack.Controls.Add(note, 0, 2)
        End If
        Return stack
    End Function

    Friend Shared Function Badge(text As String, back As Color, fore As Color) As Label
        Return New Label With {
            .Text = text,
            .AutoSize = True,
            .BackColor = back,
            .ForeColor = fore,
            .Padding = New Padding(8, 3, 8, 3),
            .Margin = New Padding(8, 8, 0, 0)
        }
    End Function

    Friend Shared Sub StylePrimary(button As Button)
        button.FlatStyle = FlatStyle.Flat
        button.BackColor = Color.FromArgb(0, 90, 180)
        button.ForeColor = Color.White
        button.UseVisualStyleBackColor = False
        button.Size = New Size(110, 34)
    End Sub

    Friend Shared Sub StyleClose(button As Button)
        button.FlatStyle = FlatStyle.Flat
        button.Size = New Size(110, 34)
    End Sub

    Friend Shared Function Initials(given As String, family As String) As String
        Dim first = FirstLetter(given)
        Dim second = FirstLetter(family)
        Dim text = (first & second).ToUpperInvariant()
        If text.Length = 0 Then Return "?"
        Return text
    End Function

    Private Shared Function FirstLetter(value As String) As String
        Dim text = If(value, "").Trim()
        If text.Length = 0 Then Return ""
        Return text.Substring(0, 1)
    End Function

    Friend Shared Sub ApplyLookupNotes(emailLabel As Label, unitLabel As Label)
        Dim show = My.Settings.ShowTeacherEmailAndUnitName
        emailLabel.Visible = show
        unitLabel.Visible = show
    End Sub
End Class
