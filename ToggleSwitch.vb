''' <summary>On/off switch. Checked is the boolean that is saved.</summary>
Friend Class ToggleSwitch
    Inherits CheckBox

    Public Sub New()
        Appearance = Appearance.Normal
        AutoSize = False
        Size = New Size(48, 26)
        Text = ""
        MinimumSize = New Size(48, 26)
        Cursor = Cursors.Hand
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        e.Graphics.Clear(If(Parent Is Nothing, SystemColors.Control, Parent.BackColor))
        Dim track As New Rectangle(0, 2, Width - 1, Height - 5)
        Dim onColor = Color.FromArgb(0, 90, 180)
        Dim offColor = Color.FromArgb(180, 188, 198)
        Using brush As New SolidBrush(If(Checked, onColor, offColor))
            e.Graphics.FillEllipse(brush, track.X, track.Y, track.Height, track.Height)
            e.Graphics.FillEllipse(brush, track.Right - track.Height, track.Y, track.Height, track.Height)
            e.Graphics.FillRectangle(brush, track.X + track.Height \ 2, track.Y, track.Width - track.Height, track.Height)
        End Using
        Dim knob = track.Height - 6
        Dim x = If(Checked, track.Right - knob - 3, track.X + 3)
        Using knobBrush As New SolidBrush(Color.White)
            e.Graphics.FillEllipse(knobBrush, x, track.Y + 3, knob, knob)
        End Using
    End Sub
End Class

Friend Class InitialAvatar
    Inherits Panel

    Public Property Initials As String
        Get
            Return _initials
        End Get
        Set(value As String)
            _initials = If(value, "")
            Invalidate()
        End Set
    End Property

    Private _initials As String = ""

    Public Sub New()
        Size = New Size(56, 56)
        MinimumSize = New Size(56, 56)
        DoubleBuffered = True
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        e.Graphics.Clear(If(Parent Is Nothing, Color.White, Parent.BackColor))
        Using brush As New SolidBrush(Color.FromArgb(0, 90, 180))
            e.Graphics.FillEllipse(brush, 1, 1, Width - 3, Height - 3)
        End Using
        Dim text = If(_initials, "")
        Using font As New Font("Segoe UI", 14.0F, FontStyle.Bold)
            Using textBrush As New SolidBrush(Color.White)
                Dim size = e.Graphics.MeasureString(text, font)
                e.Graphics.DrawString(text, font, textBrush, (Width - size.Width) / 2.0F, (Height - size.Height) / 2.0F)
            End Using
        End Using
    End Sub
End Class

Friend Class PackageChoice
    Public Property Code As String
    Public Property Title As String

    Public Overrides Function ToString() As String
        Dim displayTitle = If(Title, "").Trim()
        If displayTitle.Length = 0 Then Return If(Code, "")
        Return If(Code, "") & " — " & displayTitle
    End Function
End Class

Friend Class UnitChoice
    Public Property Code As String
    Public Property Title As String

    Public Overrides Function ToString() As String
        Dim displayTitle = If(Title, "").Trim()
        If displayTitle.Length = 0 Then Return If(Code, "")
        Return If(Code, "") & " - " & displayTitle
    End Function
End Class
