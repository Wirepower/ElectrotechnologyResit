Imports System.Globalization
Imports System.Net.Http
Imports System.Text.Json
Imports Microsoft.Identity.Client

' Azure setup: App registration > add delegated Microsoft Graph permissions Bookings.Read.All and User.Read;
' Authentication > Mobile and desktop applications > redirect URI http://localhost (matches MSAL below).
' Grant admin consent for the org if required. Tenant ID can be "common" or your directory GUID.

''' <summary>Microsoft Graph calls for Microsoft Bookings (appointments via calendarView).</summary>
Friend NotInheritable Class BookingsGraphClient
    Private Shared ReadOnly GraphRoot As String = "https://graph.microsoft.com/v1.0"
    Private Shared ReadOnly Http As New HttpClient() With {.Timeout = TimeSpan.FromMinutes(2)}

    Friend Shared ReadOnly Scopes As String() = {
        "https://graph.microsoft.com/Bookings.Read.All",
        "https://graph.microsoft.com/User.Read"
    }

    Friend Shared Function CreatePublicClientApplication(clientId As String, tenantId As String) As IPublicClientApplication
        Dim tid = If(String.IsNullOrWhiteSpace(tenantId), "common", tenantId.Trim())
        Return PublicClientApplicationBuilder.Create(clientId.Trim()).
            WithAuthority(AzureCloudInstance.AzurePublic, tid).
            WithRedirectUri("http://localhost").
            Build()
    End Function

    Friend Shared Async Function AcquireTokenAsync(pca As IPublicClientApplication, parent As IWin32Window) As Task(Of AuthenticationResult)
        Dim accounts = Await pca.GetAccountsAsync().ConfigureAwait(False)
        Dim account = If(accounts.Count > 0, accounts(0), Nothing)
        Try
            If account IsNot Nothing Then
                Return Await pca.AcquireTokenSilent(Scopes, account).ExecuteAsync().ConfigureAwait(False)
            End If
        Catch ex As MsalUiRequiredException
            System.Diagnostics.Debug.WriteLine(ex.ToString())
        End Try

        Dim builder = pca.AcquireTokenInteractive(Scopes)
        If parent IsNot Nothing Then
            builder = builder.WithParentActivityOrWindow(parent.Handle)
        End If
        Return Await builder.ExecuteAsync().ConfigureAwait(False)
    End Function

    Friend Shared Async Function GetJsonAsync(url As String, accessToken As String) As Task(Of String)
        Using req As New HttpRequestMessage(HttpMethod.Get, url)
            req.Headers.Authorization = New Headers.AuthenticationHeaderValue("Bearer", accessToken)
            req.Headers.Accept.ParseAdd("application/json")
            Dim resp = Await Http.SendAsync(req).ConfigureAwait(False)
            Dim body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
            If Not resp.IsSuccessStatusCode Then
                Throw New InvalidOperationException("Microsoft Graph error (" & CInt(resp.StatusCode).ToString() & "): " & body)
            End If
            Return body
        End Using
    End Function

    Friend Shared Async Function ListBookingBusinessIdsAsync(accessToken As String) As Task(Of List(Of (String, String)))
        Dim list As New List(Of (String, String))
        Dim url = GraphRoot & "/solutions/bookingBusinesses?$select=id,displayName"
        While Not String.IsNullOrEmpty(url)
            Dim json = Await GetJsonAsync(url, accessToken).ConfigureAwait(False)
            Dim nextPage As String = Nothing
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim valEl As JsonElement
                If root.TryGetProperty("value", valEl) Then
                    For Each el In valEl.EnumerateArray()
                        Dim id = el.GetProperty("id").GetString()
                        If String.IsNullOrEmpty(id) Then Continue For
                        Dim dn = ""
                        Dim dnEl As JsonElement
                        If el.TryGetProperty("displayName", dnEl) Then dn = If(dnEl.GetString(), "")
                        list.Add((id, dn))
                    Next
                End If
                Dim nl As JsonElement
                If root.TryGetProperty("@odata.nextLink", nl) Then
                    nextPage = nl.GetString()
                End If
            End Using
            url = nextPage
        End While
        Return list
    End Function

    Friend Shared Async Function ListStaffMembersAsync(accessToken As String, businessId As String) As Task(Of Dictionary(Of String, (String, String)))
        Dim map As New Dictionary(Of String, (String, String))(StringComparer.OrdinalIgnoreCase)
        Dim url = GraphRoot & "/solutions/bookingBusinesses/" & Uri.EscapeDataString(businessId) & "/staffMembers?$select=id,displayName,emailAddress"
        While Not String.IsNullOrEmpty(url)
            Dim json = Await GetJsonAsync(url, accessToken).ConfigureAwait(False)
            Dim nextPage As String = Nothing
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim valEl As JsonElement
                If root.TryGetProperty("value", valEl) Then
                    For Each el In valEl.EnumerateArray()
                        Dim sid = el.GetProperty("id").GetString()
                        Dim dn = ""
                        Dim em = ""
                        Dim t As JsonElement
                        If el.TryGetProperty("displayName", t) Then dn = If(t.GetString(), "")
                        If el.TryGetProperty("emailAddress", t) Then em = If(t.GetString(), "")
                        If Not String.IsNullOrEmpty(sid) Then
                            map(sid) = (dn, em)
                        End If
                    Next
                End If
                Dim nl As JsonElement
                If root.TryGetProperty("@odata.nextLink", nl) Then
                    nextPage = nl.GetString()
                End If
            End Using
            url = nextPage
        End While
        Return map
    End Function

    ''' <summary>All appointments in [startIso, endIso) for the booking business calendar.</summary>
    Friend Shared Async Function ListCalendarAppointmentsRawAsync(accessToken As String, businessId As String, startIso As String, endIso As String) As Task(Of List(Of String))
        Dim list As New List(Of String)
        Dim url = GraphRoot & "/solutions/bookingBusinesses/" & Uri.EscapeDataString(businessId) &
            "/calendarView?start=" & Uri.EscapeDataString(startIso) & "&end=" & Uri.EscapeDataString(endIso)
        While Not String.IsNullOrEmpty(url)
            Dim json = Await GetJsonAsync(url, accessToken).ConfigureAwait(False)
            Dim nextPage As String = Nothing
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim valEl As JsonElement
                If root.TryGetProperty("value", valEl) Then
                    For Each el In valEl.EnumerateArray()
                        list.Add(el.GetRawText())
                    Next
                End If
                Dim nl As JsonElement
                If root.TryGetProperty("@odata.nextLink", nl) Then
                    nextPage = nl.GetString()
                End If
            End Using
            url = nextPage
        End While
        Return list
    End Function

    ''' <summary>Local-calendar day as ISO range for Graph calendarView (start inclusive, end exclusive).</summary>
    Friend Shared Sub GetLocalDayRangeIso(d As Date, ByRef startIso As String, ByRef endIso As String)
        Dim tz = TimeZoneInfo.Local
        Dim startLocal = New DateTime(d.Year, d.Month, d.Day, 0, 0, 0, DateTimeKind.Unspecified)
        Dim endLocal = startLocal.AddDays(1)
        Dim startOff = New DateTimeOffset(startLocal, tz.GetUtcOffset(startLocal))
        Dim endOff = New DateTimeOffset(endLocal, tz.GetUtcOffset(endLocal))
        startIso = startOff.ToString("o", CultureInfo.InvariantCulture)
        endIso = endOff.ToString("o", CultureInfo.InvariantCulture)
    End Sub

    ''' <summary>Maps custom question id → display text (for parsing answers on appointments).</summary>
    Friend Shared Async Function ListCustomQuestionLabelsAsync(accessToken As String, businessId As String) As Task(Of Dictionary(Of String, String))
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim url = GraphRoot & "/solutions/bookingBusinesses/" & Uri.EscapeDataString(businessId) & "/customQuestions?$select=id,displayName"
        While Not String.IsNullOrEmpty(url)
            Dim json = Await GetJsonAsync(url, accessToken).ConfigureAwait(False)
            Dim nextPage As String = Nothing
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                Dim valEl As JsonElement
                If root.TryGetProperty("value", valEl) Then
                    For Each el In valEl.EnumerateArray()
                        Dim qid = el.GetProperty("id").GetString()
                        Dim disp = ""
                        Dim dEl As JsonElement
                        If el.TryGetProperty("displayName", dEl) Then disp = If(dEl.GetString(), "")
                        If Not String.IsNullOrEmpty(qid) Then map(qid) = disp
                    Next
                End If
                Dim nl As JsonElement
                If root.TryGetProperty("@odata.nextLink", nl) Then
                    nextPage = nl.GetString()
                End If
            End Using
            url = nextPage
        End While
        Return map
    End Function
End Class
