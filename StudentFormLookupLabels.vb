Imports ElectrotechnologyResit.Core

''' <summary>The teacher-email and unit-name labels sit beside the combos. They are hidden unless Settings asks for them, and the form narrows so the empty space is not left behind.</summary>
Friend NotInheritable Class StudentFormLookupLabels
    Private Sub New()
    End Sub

    Friend Shared Sub Apply(form As Form, emailLabel As Label, unitLabel As Label, designedWidth As Integer)
        Dim show = My.Settings.ShowTeacherEmailAndUnitName
        emailLabel.Visible = show
        unitLabel.Visible = show
        Dim right = 0
        For Each control As Control In form.Controls
            If control Is emailLabel OrElse control Is unitLabel Then Continue For
            right = Math.Max(right, control.Right)
        Next
        form.ClientSize = New Size(StudentFormWidth.ForLookupLabels(show, designedWidth, right, 16), form.ClientSize.Height)
    End Sub
End Class
