' ===================================
' BasePage Class
' Common utility methods for page classes
' ===================================
Class BasePage

    ' Publicly accessible browser object
    Public browser

    ' Private flag to ensure AI context is set only once
    Private aiContextSet

    ' Private variable to store the page name
    Private pageName

    ' -----------------------------------
    ' Initializes the browser context
    ' Parameters:
    '   driver: The driver object to get the browser instance from.
    '   titleName: The name of the page.
    ' -----------------------------------
    Public Sub Init(driver, titleName)
        Set browser = driver.GetBrowserInstance()
        aiContextSet = False ' Ensure context is reset for new page instances
        pageName = titleName
    End Sub

    ' -----------------------------------
    ' Sets the AI context if not already set.
    ' This prevents redundant context settings.
    ' -----------------------------------
    Public Sub InitializeAIContext()
        If Not aiContextSet Then
            AIUtil.SetContext browser
            aiContextSet = True
        End If
    End Sub

    ' -----------------------------------
    ' Sets text in an AI-detected element by type and label.
    ' It handles potential element not found scenarios.
    ' Parameters:
    '   controlType: The AI control type (e.g., "WebEdit", "Button").
    '   label: The label or accessibility identifier of the element.
    '   value: The text value to set.
    ' -----------------------------------
    Public Sub SetAITextField(controlType, label, value)
        InitializeAIContext() ' Ensure AI context is initialized

        AIUtil.Context.Freeze ' Freeze AI context for more stable recognition

        Dim element ' Declare element variable

        ' Attempt to find the element using AIUtil
        Set element = AIUtil(controlType, label)

        If Not element Is Nothing Then
            On Error Resume Next ' Enable error handling for interactions
            element.Click ' Click to ensure element is active/focused
            If Err.Number <> 0 Then
                Reporter.ReportEvent micWarning, "SetAITextField", "Failed to click element '" & label & "'. Error: " & Err.Description
                Err.Clear
            End If
            
            ' A small wait can be beneficial for UI stability, adjust as needed
            Wait 0.3 
            
            element.SetText value ' Set the text value
            If Err.Number <> 0 Then
                Reporter.ReportEvent micFail, "SetAITextField", "Failed to set text for '" & label & "'. Error: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0 ' Disable error handling
        Else
            ' Report failure if element is not found
            Reporter.ReportEvent micFail, "SetAITextField", "Element not found: Type - '" & controlType & "', Label - '" & label & "'"
        End If

        AIUtil.Context.Unfreeze ' Unfreeze AI context
    End Sub

    ' -----------------------------------
    ' Clicks an AI-detected element by type and label.
    ' Parameters:
    '   controlType: The AI control type.
    '   label: The label or accessibility identifier of the element.
    ' -----------------------------------
    Public Sub ClickByIA(controlType, label)
        InitializeAIContext() ' Ensure AI context is initialized
        On Error Resume Next ' Enable error handling

        ' Directly attempt to click the element
        AIUtil(controlType, label).Click

        If Err.Number <> 0 Then
            Reporter.ReportEvent micFail, "ClickByIA", "Failed to click element: Type - '" & controlType & "', Label - '" & label & "'. Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0 ' Disable error handling
    End Sub

    ' -----------------------------------
    ' Sets text for a standard WebEdit element using a locator.
    ' Parameters:
    '   locator: The object repository locator for the WebEdit element.
    '   text: The text value to set.
    ' -----------------------------------
    Public Sub SetText(locator, text)
        On Error Resume Next ' Enable error handling
        browser.Page(pageName).WebEdit(locator).Set text
        If Err.Number <> 0 Then
            Reporter.ReportEvent micFail, "SetText", "Failed to set text for locator '" & locator & "' on page '" & pageName & "'. Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0 ' Disable error handling
    End Sub

End Class
