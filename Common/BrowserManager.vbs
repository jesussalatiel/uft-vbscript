' ==========================================================
' BrowserManager Class
' Manages browser operations for Salesforce automation.
' ==========================================================
Class BrowserManager

    ' Launches Salesforce in the configured browser.
    Public Sub LaunchSalesforcePlatform()
        Dim browserType, browserExe, browserCapabilities, runArguments, browserInstance
        browserType = LCase(G_BROWSER_TYPE)

        Call CloseBrowserInstance(browserType)

        browserExe = GetBrowserExecutable(browserType)
        If browserExe = "" Then
            Reporter.ReportEvent micFail, "LaunchSalesforcePlatform", "Unsupported browser type: '" & browserType & "'"
            Exit Sub
        End If

        browserCapabilities = Capabilities(browserType)

        If browserCapabilities <> "" Then
            runArguments = G_SALESFORCE_BASE_URL & " " & browserCapabilities
        Else
            runArguments = G_SALESFORCE_BASE_URL
        End If

        Reporter.ReportEvent micInfo, "Launching browser", "Executable: '" & browserExe & "' | URL: '" & G_SALESFORCE_BASE_URL & "' | Args: '" & browserCapabilities & "'"

        On Error Resume Next
        SystemUtil.Run browserExe, runArguments
        If Err.Number <> 0 Then
            Reporter.ReportEvent micFail, "Browser launch failed", Err.Description & " (Code: " & Err.Number & ")"
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0

        Set browserInstance = GetBrowserInstance()
        If Not browserInstance Is Nothing Then
            On Error Resume Next
            browserInstance.Sync
            If Err.Number <> 0 Then
                Reporter.ReportEvent micWarning, "Browser sync failed", Err.Description & " (Code: " & Err.Number & ")"
                Err.Clear
            End If
            On Error GoTo 0
            Reporter.ReportEvent micDone, "Browser ready", "Salesforce browser launched and synchronized."
        Else
            Reporter.ReportEvent micFail, "Browser instance not found", "No Salesforce browser found after launch."
        End If
    End Sub

    ' Returns a Salesforce browser instance or the latest browser if not found.
    Public Function GetBrowserInstance()
        Dim desc, browsers, i, timeoutSeconds, elapsed, foundSalesforce
        timeoutSeconds = 10
        elapsed = 0
        foundSalesforce = False

        Set desc = Description.Create
        desc("micclass").Value = "Browser"

        Do While elapsed < timeoutSeconds
            Set browsers = Desktop.ChildObjects(desc)
            For i = 0 To browsers.Count - 1
                If InStr(1, browsers(i).GetROProperty("title"), "Salesforce", vbTextCompare) > 0 Then
                    Set GetBrowserInstance = browsers(i)
                    Reporter.ReportEvent micInfo, "Browser found", "Title: " & browsers(i).GetROProperty("title")
                    foundSalesforce = True
                    Exit Function
                End If
            Next

            Wait 1
            elapsed = elapsed + 1
        Loop

        ' Fallback: last created browser
        Set GetBrowserInstance = Browser("creationtime:=0")
        If Not GetBrowserInstance Is Nothing Then
            Reporter.ReportEvent micWarning, "Fallback browser", "Salesforce not found, returning latest browser instance."
        Else
            Reporter.ReportEvent micFail, "No browser found", "No browser instances available."
        End If
    End Function

    ' Closes the browser process for the given type.
    Public Sub CloseBrowserInstance(browserType)
        Dim exeName
        exeName = GetBrowserExecutable(LCase(browserType))

        If exeName <> "" Then
            Reporter.ReportEvent micInfo, "Closing browser", "Executable: " & exeName
            On Error Resume Next
            SystemUtil.CloseProcessByName exeName
            If Err.Number <> 0 Then
                Reporter.ReportEvent micWarning, "Close failed", Err.Description & " (Code: " & Err.Number & ")"
                Err.Clear
            End If
            On Error GoTo 0
            Reporter.ReportEvent micDone, "Browser closed", "Closed: " & exeName
        Else
            Reporter.ReportEvent micWarning, "Invalid browser type", "Cannot close unknown browser type: '" & browserType & "'"
        End If
    End Sub

    ' Closes all instances of the globally configured browser.
    Public Sub CloseAllBrowserInstances()
        Call CloseBrowserInstance(G_BROWSER_TYPE)
    End Sub

    ' Returns the executable name for a browser type.
    Private Function GetBrowserExecutable(browserType)
        Select Case LCase(browserType)
            Case "edge":   GetBrowserExecutable = "msedge.exe"
            Case "chrome": GetBrowserExecutable = "chrome.exe"
            Case "firefox":GetBrowserExecutable = "firefox.exe"
            Case Else:     GetBrowserExecutable = ""
        End Select
    End Function

    ' Returns additional command-line arguments for each browser type.
    Public Function Capabilities(browserType)
        Select Case LCase(browserType)
            Case "edge", "chrome": Capabilities = "--disable-infobars --start-maximized"
            Case "firefox":        Capabilities = "--width=1280 --height=720"
            Case Else:             Capabilities = ""
        End Select
    End Function

End Class

