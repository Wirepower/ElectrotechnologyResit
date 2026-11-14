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

    ''' <summary>
    ''' Values that are applied only when the app starts again: data source, SharePoint site and app registration, and the list names.
    ''' The PC password, Word template, open-after-save option and lookup-label option are not included.
    ''' Bookings client, tenant, business id and the pull button are read when they are used, so they are not included either.
    ''' </summary>
    Public Class SettingsRestartSnapshot
        Public Property DataSource As String
        Public Property SiteUrl As String
        Public Property ClientId As String
        Public Property TenantId As String
        Public Property UseSitesSelected As Boolean
        Public Property BookingsList As String
        Public Property StudentsList As String
        Public Property TeachersList As String
        Public Property UnitsList As String
        Public Property SignatureLibrary As String

        Public Shared Function Capture(dataSource As String, siteUrl As String, clientId As String, tenantId As String, useSitesSelected As Boolean, bookingsList As String, studentsList As String, teachersList As String, unitsList As String, signatureLibrary As String) As SettingsRestartSnapshot
            Dim names = SharePointListNames.FromSettings(bookingsList, studentsList, teachersList, unitsList, signatureLibrary)
            Dim source = If(String.Equals(If(dataSource, "").Trim(), "SharePoint", StringComparison.OrdinalIgnoreCase), "SharePoint", "Sample")
            Dim tenant = If(tenantId, "").Trim()
            If tenant.Length = 0 Then tenant = "common"
            Return New SettingsRestartSnapshot With {
                .DataSource = source,
                .SiteUrl = If(siteUrl, "").Trim(),
                .ClientId = If(clientId, "").Trim(),
                .TenantId = tenant,
                .UseSitesSelected = useSitesSelected,
                .BookingsList = names.Bookings,
                .StudentsList = names.Students,
                .TeachersList = names.Teachers,
                .UnitsList = names.Units,
                .SignatureLibrary = names.SignatureLibrary
            }
        End Function
    End Class

    Public NotInheritable Class SettingsRestart
        Private Sub New()
        End Sub

        Public Shared Function NeedsRestart(loaded As SettingsRestartSnapshot, saving As SettingsRestartSnapshot) As Boolean
            If loaded Is Nothing OrElse saving Is Nothing Then Return True
            If Not String.Equals(loaded.DataSource, saving.DataSource, StringComparison.Ordinal) Then Return True
            If Not Same(loaded.SiteUrl, saving.SiteUrl) Then Return True
            If Not Same(loaded.ClientId, saving.ClientId) Then Return True
            If Not Same(loaded.TenantId, saving.TenantId) Then Return True
            If loaded.UseSitesSelected <> saving.UseSitesSelected Then Return True
            If Not Same(loaded.BookingsList, saving.BookingsList) Then Return True
            If Not Same(loaded.StudentsList, saving.StudentsList) Then Return True
            If Not Same(loaded.TeachersList, saving.TeachersList) Then Return True
            If Not Same(loaded.UnitsList, saving.UnitsList) Then Return True
            If Not Same(loaded.SignatureLibrary, saving.SignatureLibrary) Then Return True
            Return False
        End Function

        Private Shared Function Same(left As String, right As String) As Boolean
            Return String.Equals(If(left, ""), If(right, ""), StringComparison.Ordinal)
        End Function
    End Class

End Namespace
