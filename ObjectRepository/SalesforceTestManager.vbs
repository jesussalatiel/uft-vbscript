' ==============================================
' SalesforceTestManager Class
' Manages Salesforce login and navigation using POM and BasePage utilities
' ==============================================

Class SalesforceTestManager
    Private base

    ' Initializes the test manager with the browser context
    Public Sub Init(driver)
        Set base = New BasePage
        base.Init driver, "title:=Login \| Salesforce"
    End Sub

    ' Logs into Salesforce using username and password fields
    Public Sub Login(username, password)
        base.SetText "html id:=username", username
        base.SetText "html id:=password", password
        ' Use this if testing a Sandbox login:
        ' base.ClickByIA "button", "Log In to Sandbox"
        base.ClickByIA "button", "Log In"
    End Sub

    ' Opens the App Launcher and selects a specific app
    Public Sub AppLauncher(app)
        Select Case LCase(Trim(app))
            Case "sales"
                base.ClickByIA "button", "categories"
                base.SetAITextField "text_box", "App Launcher", "Sales Center"
            ' Extend with more cases as needed
        End Select
    End Sub

    ' Opens the Settings menu and navigates based on the provided section
    Public Sub Settings(section)
        AIUtil("gear_settings").Click
        AIUtil.FindTextBlock("Setup").Click

        Select Case LCase(Trim(section))
            Case "users"
                NavigateToUsers
            ' Add more cases if needed
        End Select
    End Sub

    ' Navigates to the "Users" section within settings
    Private Sub NavigateToUsers()
        AIUtil.SetContext Browser("creationtime:=1")

        Dim usersBlock
        Set usersBlock = AIUtil.FindTextBlock("Users")
        
        If usersBlock.Exist(5) Then
            usersBlock.Click
        Else
            Reporter.ReportEvent micFail, "Users block", "The first 'Users' element was not found."
            Exit Sub
        End If

        ' Wait dynamically for the second 'Users' element from the bottom
        Dim usersBottom
        Set usersBottom = AIUtil.FindTextBlock("Users", micFromBottom, 1)

        If usersBottom.Exist(5) Then
            usersBottom.Click
        Else
            Reporter.ReportEvent micFail, "Users block (bottom)", "The second 'Users' element (from bottom) was not found."
        End If
    End Sub
End Class

