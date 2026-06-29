Option Strict On
Option Explicit On

Imports System.Reflection

''' <summary>Displayed app version (from InformationalVersion / AppReleaseVersion in the .vbproj).</summary>
Friend Module AppVersion
    Private ReadOnly _VersionText As String = ResolveVersionText()

    Private Function ResolveVersionText() As String
        Dim attr = Assembly.GetExecutingAssembly().GetCustomAttribute(Of AssemblyInformationalVersionAttribute)()
        If attr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(attr.InformationalVersion) Then
            Dim v = attr.InformationalVersion.Trim()
            Dim plus = v.IndexOf("+"c)
            If plus >= 0 Then v = v.Substring(0, plus).Trim()
            Return v
        End If

        Dim ver = Assembly.GetExecutingAssembly().GetName().Version
        Return ver.Major.ToString() & "." & ver.Minor.ToString()
    End Function

    ''' <summary>e.g. "Version 2.0"</summary>
    Friend ReadOnly Property DisplayText As String
        Get
            Return "Version " & VersionString
        End Get
    End Property

    ''' <summary>e.g. "2.0"</summary>
    Friend ReadOnly Property VersionString As String
        Get
            Return _VersionText
        End Get
    End Property
End Module
