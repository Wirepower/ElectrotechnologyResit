Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

''' <summary>Late-bound Microsoft Word (no Office PIAs). Requires Word installed.</summary>
Friend Module WordDocInterop
    Private Const WdReplaceAll As Integer = 2

    Friend Function TryCreateWordApplication() As Object
        Dim wordType = Type.GetTypeFromProgID("Word.Application")
        If wordType Is Nothing Then Return Nothing
        Return Activator.CreateInstance(wordType)
    End Function

    Friend Sub ReplaceAllInDocument(doc As Object, findText As String, replaceText As String)
        If findText Is Nothing OrElse replaceText Is Nothing Then Return
        Dim rng = doc.Content
        Dim f = rng.Find
        f.ClearFormatting()
        f.Text = findText
        f.Replacement.ClearFormatting()
        f.Replacement.Text = replaceText
        f.Forward = True
        f.Wrap = 1
        f.Format = False
        f.MatchCase = False
        f.MatchWholeWord = False
        f.MatchWildcards = False
        f.MatchSoundsLike = False
        f.MatchAllWordForms = False
        f.Execute(Replace:=WdReplaceAll)
    End Sub

    Friend Sub QuitWord(wordApp As Object)
        If wordApp Is Nothing Then Return
        Try
            wordApp.Quit(SaveChanges:=0)
        Finally
            Marshal.FinalReleaseComObject(wordApp)
        End Try
    End Sub
End Module
