''' <summary>Shows or hides the teacher-email and unit-name notes. Hidden notes sit in an auto-sized row, so they leave no gap.</summary>
Friend NotInheritable Class StudentFormLookupLabels
    Private Sub New()
    End Sub

    Friend Shared Sub Apply(emailLabel As Label, unitLabel As Label)
        StudentDialogLayout.ApplyLookupNotes(emailLabel, unitLabel)
    End Sub
End Class
