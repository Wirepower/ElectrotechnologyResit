Imports System.IO
Imports System.Windows.Forms

Friend Module ResitLoginTemplatePaths
    ''' <summary>Bundled default; bump name when layout changes so startup can point users to the new file.</summary>
    Friend Const BundledTemplateFileName As String = "ResitLoginSheet-v3.docx"

    Friend Function GetTemplatesDirectory() As String
        Return Path.Combine(Application.StartupPath, "Templates")
    End Function

    Friend Function GetBundledTemplateFullPath() As String
        Return Path.Combine(GetTemplatesDirectory(), BundledTemplateFileName)
    End Function
End Module
