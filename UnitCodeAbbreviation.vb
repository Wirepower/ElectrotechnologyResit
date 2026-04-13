Imports System.Collections.Generic

''' <summary>Maps full unit codes from the database to abbreviated class codes on resit sheets.</summary>
Friend Module UnitCodeAbbreviation
    Private ReadOnly ExplicitMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"UEECO0023", "CO0023"},
        {"UEECD0007", "CD0007"},
        {"UEECD0019", "CD0019"},
        {"UEECD0020", "CD0020"},
        {"UEECD0051", "CD0051"},
        {"UEECD0046", "CD0046"},
        {"UEECD0044", "CD0044"},
        {"UEEEL0021", "L0021"},
        {"UEEEL0019", "L0019"},
        {"UEERE0001", "RE0001"},
        {"UEEEL0023", "L0023"},
        {"UEEEL0020", "L0020"},
        {"UEEEL0025", "L0025"},
        {"UEEEL0024", "L0024"},
        {"UEEEL0008", "L0008"},
        {"UEEEL0009", "L0009"},
        {"UEEEL0010", "L0010"},
        {"UEEDV0005", "V0005"},
        {"UEEDV0008", "V0008"},
        {"UEEEL0003", "L0003"},
        {"UEEEL0018", "L0018"},
        {"UEEEL0005", "L0005"},
        {"UEECD0016", "CD0016"},
        {"UEEEL0047", "L0047"},
        {"HLTAID009", "HLTAID009"},
        {"UETDRRF004", "UETDRRF004"},
        {"UEEEL0014", "L0014"},
        {"UEEEL0012", "L0012"},
        {"UEEEL0039", "L0039"}
    }

    ''' <summary>Looks up explicit map first, then applies the same prefix rules for new codes added later.</summary>
    Friend Function ToAbbreviatedClassCode(unitCode As String) As String
        If String.IsNullOrWhiteSpace(unitCode) Then Return unitCode
        Dim u = unitCode.Trim()
        Dim mapped As String = Nothing
        If ExplicitMap.TryGetValue(u, mapped) Then Return mapped
        Return AbbreviateByRule(u)
    End Function

    ''' <summary>UEECD*→CD*, UEEEL*→L*, UEECO*→CO*, UEERE*→RE*, UEEDV*→V*; otherwise unchanged.</summary>
    Private Function AbbreviateByRule(u As String) As String
        If u.Length <= 5 Then Return u
        If u.StartsWith("UEECD", StringComparison.OrdinalIgnoreCase) Then Return "CD" & u.Substring(5)
        If u.StartsWith("UEEEL", StringComparison.OrdinalIgnoreCase) Then Return "L" & u.Substring(5)
        If u.StartsWith("UEECO", StringComparison.OrdinalIgnoreCase) Then Return "CO" & u.Substring(5)
        If u.StartsWith("UEERE", StringComparison.OrdinalIgnoreCase) Then Return "RE" & u.Substring(5)
        If u.StartsWith("UEEDV", StringComparison.OrdinalIgnoreCase) Then Return "V" & u.Substring(5)
        Return u
    End Function
End Module
