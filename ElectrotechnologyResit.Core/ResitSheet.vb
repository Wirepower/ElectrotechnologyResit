Namespace ElectrotechnologyResit.Core

    Public NotInheritable Class ResitClassGroups
        Public Const EmptySelectionText As String = "(no class groups selected — tick the units under Energyspace Class Created to include them here)"

        Private Sub New()
        End Sub

        ''' <summary>Line shape is fixed: abbreviation, RESIT, the qualification suffix, then the year.</summary>
        Public Shared Function Build(entries As IEnumerable(Of ClassGroupEntry), certYear As Integer) As String
            Dim lines As New List(Of String)
            If entries IsNot Nothing Then
                For Each entry In entries
                    If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.UnitCode) Then Continue For
                    Dim abbreviated = AbbreviationRules.Abbreviate(entry.UnitCode, entry.Rules)
                    Dim suffix = If(entry.ClassGroupSuffix, "").Trim()
                    Dim packagePart = If(suffix.Length = 0, "RESIT", "RESIT-" & suffix)
                    lines.Add(abbreviated & "-" & packagePart & "-" & certYear.ToString())
                Next
            End If
            lines.Sort(StringComparer.OrdinalIgnoreCase)
            If lines.Count = 0 Then Return EmptySelectionText
            Return String.Join(vbCrLf, lines)
        End Function
    End Class

    ''' <summary>Excel appears when either Energyspace list has a tick. The Word sheet appears only for Class Created.</summary>
    Public NotInheritable Class ResitOutputVisibility
        Private Sub New()
        End Sub

        Public Shared Function ShowExcel(classCreatedTicked As Boolean, assessmentBookedTicked As Boolean) As Boolean
            Return classCreatedTicked OrElse assessmentBookedTicked
        End Function

        Public Shared Function ShowLoginSheet(classCreatedTicked As Boolean) As Boolean
            Return classCreatedTicked
        End Function
    End Class

End Namespace
