Namespace ElectrotechnologyResit.Core

    Public Enum SettingsCloseAction
        StayOpen = 0
        CloseWithoutSaving = 1
        SaveAndExit = 2
    End Enum

    Public Enum SettingsCloseAnswer
        Yes = 0
        No = 1
        Cancel = 2
    End Enum

    ''' <summary>What to do when Settings is closed. A display-only tick is saved immediately and does not count as dirty.</summary>
    Public NotInheritable Class SettingsClosePrompt
        Private Sub New()
        End Sub

        Public Shared Function Decide(dirty As Boolean, answer As SettingsCloseAnswer) As SettingsCloseAction
            If Not dirty Then Return SettingsCloseAction.CloseWithoutSaving
            If answer = SettingsCloseAnswer.Yes Then Return SettingsCloseAction.SaveAndExit
            If answer = SettingsCloseAnswer.No Then Return SettingsCloseAction.CloseWithoutSaving
            Return SettingsCloseAction.StayOpen
        End Function
    End Class

End Namespace
