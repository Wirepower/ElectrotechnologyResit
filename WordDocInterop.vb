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

    ''' <summary>
    ''' Replaces bookmark content and re-adds bookmark (Word drops it after text assignment).
    ''' Returns True when bookmark exists and was updated.
    ''' </summary>
    Friend Function TryReplaceBookmarkText(doc As Object, bookmarkName As String, replacementText As String) As Boolean
        If doc Is Nothing OrElse String.IsNullOrWhiteSpace(bookmarkName) Then Return False
        If replacementText Is Nothing Then replacementText = ""

        Dim bookmarks As Object = Nothing
        Dim bookmark As Object = Nothing
        Dim rng As Object = Nothing
        Try
            bookmarks = doc.Bookmarks
            If bookmarks Is Nothing Then Return False
            If Not CBool(bookmarks.Exists(bookmarkName)) Then Return False

            bookmark = bookmarks.Item(bookmarkName)
            rng = bookmark.Range
            rng.Text = replacementText
            bookmarks.Add(bookmarkName, rng)
            Return True
        Catch
            Return False
        Finally
            ReleaseComIfNeeded(rng)
            ReleaseComIfNeeded(bookmark)
            ReleaseComIfNeeded(bookmarks)
        End Try
    End Function

    Private Sub ReleaseComIfNeeded(obj As Object)
        If obj Is Nothing Then Return
        If Marshal.IsComObject(obj) Then
            Marshal.FinalReleaseComObject(obj)
        End If
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
