Imports System.Security.Cryptography
Imports ElectrotechnologyResit.Core
Imports Microsoft.Identity.Client

Public Class SiteAccessResult
    Public Property Ok As Boolean
    Public Property Message As String

    Public Shared Function Succeeded(message As String) As SiteAccessResult
        Return New SiteAccessResult With {.Ok = True, .Message = message}
    End Function

    Public Shared Function Failed(message As String) As SiteAccessResult
        Return New SiteAccessResult With {.Ok = False, .Message = message}
    End Function
End Class

Public Class MsalTokenProvider
    Implements IAccessTokenProvider

    Private ReadOnly _pca As IPublicClientApplication
    Private ReadOnly _scopes As String()
    Private ReadOnly _owner As Func(Of IWin32Window)

    Public Sub New(clientId As String, tenantId As String, useSitesSelected As Boolean, owner As Func(Of IWin32Window))
        If String.IsNullOrWhiteSpace(clientId) Then
            Throw New InvalidOperationException("Set the Azure AD application (client) ID in Settings, then try again.")
        End If
        Dim tenant = If(String.IsNullOrWhiteSpace(tenantId), "common", tenantId.Trim())
        _pca = PublicClientApplicationBuilder.Create(clientId.Trim()).
            WithAuthority(AzureCloudInstance.AzurePublic, tenant).
            WithRedirectUri("http://localhost").
            Build()
        BindCache(_pca)
        _scopes = GraphPermissionScopes.ForSharePoint(useSitesSelected)
        _owner = owner
    End Sub

    Public Async Function GetAccessTokenAsync() As Task(Of String) Implements IAccessTokenProvider.GetAccessTokenAsync
        Dim result = Await AcquireAsync().ConfigureAwait(False)
        Return result.AccessToken
    End Function

    Private Async Function AcquireAsync() As Task(Of AuthenticationResult)
        Dim accounts = Await _pca.GetAccountsAsync().ConfigureAwait(False)
        Dim account = accounts.FirstOrDefault()
        Try
            If account IsNot Nothing Then
                Return Await _pca.AcquireTokenSilent(_scopes, account).ExecuteAsync().ConfigureAwait(False)
            End If
        Catch ex As MsalUiRequiredException
        End Try
        Dim builder = _pca.AcquireTokenInteractive(_scopes).WithUseEmbeddedWebView(False)
        Dim window As IWin32Window = Nothing
        If _owner IsNot Nothing Then window = _owner.Invoke()
        If window IsNot Nothing Then builder = builder.WithParentActivityOrWindow(window.Handle)
        Return Await builder.ExecuteAsync().ConfigureAwait(False)
    End Function

    Private Shared Sub BindCache(pca As IPublicClientApplication)
        Dim folder = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ElectrotechnologyResit")
        Dim cachePath = IO.Path.Combine(folder, "msal-sharepoint.bin")
        pca.UserTokenCache.SetBeforeAccess(
            Sub(args)
                Try
                    If IO.File.Exists(cachePath) Then
                        Dim protectedBytes = IO.File.ReadAllBytes(cachePath)
                        args.TokenCache.DeserializeMsalV3(ProtectedData.Unprotect(protectedBytes, Nothing, DataProtectionScope.CurrentUser))
                    End If
                Catch
                End Try
            End Sub)
        pca.UserTokenCache.SetAfterAccess(
            Sub(args)
                Try
                    If args.HasStateChanged Then
                        IO.Directory.CreateDirectory(folder)
                        Dim data = ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), Nothing, DataProtectionScope.CurrentUser)
                        IO.File.WriteAllBytes(cachePath, data)
                    End If
                Catch
                End Try
            End Sub)
    End Sub
End Class

Public Module AppServices
    Public Property Bookings As IBookingService
    Public Property Students As IStudentDirectory
    Public Property Teachers As ITeacherDirectory
    Public Property Units As IUnitDirectory
    Public Property Signature As ISignatureImageStore
    Public Property Provisioner As ISharePointProvisioner
    Public Property LookupImport As LookupImportService
    Public Property Client As ISharePointClient

    Private _configuredKey As String

    Public Function NeedsWizard() As Boolean
        Return String.IsNullOrWhiteSpace(My.Settings.SharePointSiteUrl) OrElse
            String.IsNullOrWhiteSpace(My.Settings.GraphClientId) OrElse
            Not My.Settings.SharePointSetupCompleted
    End Function

    Public Sub ConfigureFromSettings(owner As Func(Of IWin32Window))
        Dim names = CurrentNames()
        Dim key = String.Join("|", My.Settings.SharePointSiteUrl, My.Settings.GraphClientId, My.Settings.GraphTenantId, My.Settings.GraphUseSitesSelected.ToString(), names.Bookings, names.Students, names.Teachers, names.Units, names.SignatureLibrary)
        If key = _configuredKey AndAlso Bookings IsNot Nothing Then Return
        Dim provider As New MsalTokenProvider(My.Settings.GraphClientId, My.Settings.GraphTenantId, My.Settings.GraphUseSitesSelected, owner)
        Dim graph As New GraphSharePointClient(New GraphExecutor(New GraphHttpTransport(provider)), My.Settings.SharePointSiteUrl)
        Client = graph
        Bookings = New SharePointBookingService(graph, names)
        Students = New CachingStudentDirectory(New SharePointStudentDirectory(graph, names))
        Teachers = New CachingTeacherDirectory(New SharePointTeacherDirectory(graph, names))
        Units = New CachingUnitDirectory(New SharePointUnitDirectory(graph, names))
        Signature = New SharePointSignatureStore(graph, names)
        Provisioner = New ListProvisioner(New GraphSchemaStore(graph, names), names)
        LookupImport = New LookupImportService(Students, Teachers, Units)
        _configuredKey = key
    End Sub

    Public Function CurrentNames() As SharePointListNames
        Return SharePointListNames.FromSettings(
            My.Settings.ResitBookingsListName,
            My.Settings.StudentsListName,
            My.Settings.TeachersListName,
            My.Settings.UnitsListName,
            My.Settings.SignatureImageLocation)
    End Function

    Public Async Function CheckSiteAccessAsync() As Task(Of SiteAccessResult)
        Try
            If String.IsNullOrWhiteSpace(My.Settings.GraphClientId) Then
                Return SiteAccessResult.Failed("Set the Azure AD application (client) ID in Settings, then try again.")
            End If
            If String.IsNullOrWhiteSpace(My.Settings.SharePointSiteUrl) Then
                Return SiteAccessResult.Failed("Set the SharePoint site URL in Settings, then try again.")
            End If
            If Client Is Nothing Then
                Return SiteAccessResult.Failed("SharePoint is not configured yet.")
            End If
            Dim site = Await Client.GetSiteAsync().ConfigureAwait(True)
            Dim siteName = If(String.IsNullOrWhiteSpace(site.DisplayName), My.Settings.SharePointSiteUrl, site.DisplayName)
            Return SiteAccessResult.Succeeded("Signed in. SharePoint site: " & siteName)
        Catch ex As Exception
            Return SiteAccessResult.Failed("Could not sign in or open the SharePoint site. " & ex.Message)
        End Try
    End Function
End Module
