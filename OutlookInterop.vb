Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

''' <summary>Late-binds to the installed Outlook COM server so no Office PIAs (office.dll) are required beside the app.</summary>
Friend Module OutlookInterop
    Friend Function TryCreateOutlookApplication() As Object
        Dim outlookType = Type.GetTypeFromProgID("Outlook.Application")
        If outlookType Is Nothing Then Return Nothing
        Return Activator.CreateInstance(outlookType)
    End Function

    Friend Sub CreateDisplayOrSendMail(outlookApp As Object, toAddress As String, subject As String, htmlBody As String, sentOnBehalfOf As String, display As Boolean, send As Boolean)
        Dim mailItem As Object = outlookApp.CreateItem(0) ' olMailItem = 0
        Try
            mailItem.To = toAddress
            mailItem.Subject = subject
            mailItem.HTMLBody = htmlBody
            mailItem.SentOnBehalfOfName = sentOnBehalfOf
            If display Then mailItem.Display()
            If send Then mailItem.Send()
        Finally
            Marshal.FinalReleaseComObject(mailItem)
        End Try
    End Sub
End Module
