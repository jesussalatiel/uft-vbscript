' ============================
' BasePage Class
' Common utility methods for page classes
' ============================
Class BasePage
    Public browser
    Private aiContextSet

    ' Initialize browser context
    Public Sub Init(driver)
        Set browser = driver.GetBrowserInstance()
        aiContextSet = False
    End Sub

    ' Set AI context once
    Public Sub InitializeAIContext()
        If Not aiContextSet Then
            AIUtil.SetContext browser
            aiContextSet = True
        End If
    End Sub

    ' Helper to freeze/unfreeze context around an action
    Public Sub RunWithFrozenContext(action)
        InitializeAIContext()
        AIUtil.Context.Freeze
        Execute action
        AIUtil.Context.Unfreeze
    End Sub

    ' Returns a browser page by title
    Public Function GetPageByTitle(titlePattern)
        Set GetPageByTitle = browser.Page("title:=" & titlePattern)
    End Function
End Class
