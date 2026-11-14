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

    ''' <summary>Height of the Energyspace unit lists. One column until there are many units, then two columns of code and name.</summary>
    Public NotInheritable Class EnergyspacePanelLayout
        Public Const ManyUnitThreshold As Integer = 9

        Private Sub New()
        End Sub

        Public Shared Function ColumnCount(itemCount As Integer) As Integer
            If itemCount >= ManyUnitThreshold Then Return 2
            Return 1
        End Function

        Public Shared Function RowCount(itemCount As Integer) As Integer
            If itemCount <= 0 Then Return 0
            Dim columns = ColumnCount(itemCount)
            Return CInt(Math.Ceiling(itemCount / CDbl(columns)))
        End Function

        Public Shared Function ListHeight(itemCount As Integer, rowHeight As Integer, headerHeight As Integer) As Integer
            If rowHeight < 1 Then rowHeight = 1
            If headerHeight < 0 Then headerHeight = 0
            Return headerHeight + (RowCount(itemCount) * rowHeight)
        End Function

        Public Shared Function SharedListHeight(leftCount As Integer, rightCount As Integer, rowHeight As Integer, headerHeight As Integer) As Integer
            Return Math.Max(ListHeight(leftCount, rowHeight, headerHeight), ListHeight(rightCount, rowHeight, headerHeight))
        End Function
    End Class

    ''' <summary>Student forms are narrower when the teacher-email and unit-name labels are hidden.</summary>
    Public NotInheritable Class StudentFormWidth
        Private Sub New()
        End Sub

        Public Shared Function ForLookupLabels(showLabels As Boolean, designedWidth As Integer, contentRightWithoutLabels As Integer, margin As Integer) As Integer
            If designedWidth < 1 Then Return 0
            If showLabels Then Return designedWidth
            Dim fitted = contentRightWithoutLabels + Math.Max(margin, 0)
            If fitted < 1 Then Return designedWidth
            Return Math.Min(designedWidth, fitted)
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
