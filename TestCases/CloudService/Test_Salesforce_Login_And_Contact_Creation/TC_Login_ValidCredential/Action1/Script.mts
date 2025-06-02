' ==============================================
' Environment Configuration - Centralized settings
' ==============================================
Dim SALESFORCE_BASE_URL, RICKANDMORTY_API_ENDPOINT, PAGE_LOAD_TIMEOUT, ELEMENT_WAIT_TIMEOUT
SALESFORCE_BASE_URL = Environment.Value("SALESFORCE_URL")
RICKANDMORTY_API_ENDPOINT = Environment.Value("RICKANDMORTY_API")
PAGE_LOAD_TIMEOUT = Environment.Value("PAGE_LOAD_WAIT_TIME")
ELEMENT_WAIT_TIMEOUT = Environment.Value("DEFAULT_WAIT_TIME")

' ==============================================
' Main Salesforce Automation Class
' ==============================================
Class SalesforceAutomation
    ' UI Element Repository - Descriptive names
    Private loginUsernameField, loginPasswordField, loginSubmitButton, contactPhoneField
    Private browserDescriptor, salesforcePageDescriptor

    ' ==============================================
    ' Initialization Methods
    ' ==============================================
    Public Sub InitializeUIElements()
        Set loginUsernameField = CreateWebElementDescriptor("WebEdit", "html id", "username")
        Set loginPasswordField = CreateWebElementDescriptor("WebEdit", "html id", "password")
        Set loginSubmitButton = CreateWebElementDescriptor("WebButton", "html id", "Login")
        Set contactPhoneField = CreateWebElementDescriptor("WebEdit", "name", "Phone")

        Set browserDescriptor = CreateBrowserDescriptor(".*")
        Set salesforcePageDescriptor = CreatePageDescriptor(".*Salesforce.*")
    End Sub

    ' ==============================================
    ' Test Lifecycle Methods
    ' ==============================================
    Public Sub TestSetup()
        InitializeUIElements
        CloseAllBrowserInstances
        LaunchBrowserWithURL SALESFORCE_BASE_URL

        If Not IsElementVisible(loginUsernameField, PAGE_LOAD_TIMEOUT) Then
            ReportTestFailure "Login Page Load", "Username field not visible after " & PAGE_LOAD_TIMEOUT & " seconds"
        End If
    End Sub

    Public Sub TestCleanup()
        CloseAllBrowserInstances
    End Sub

    ' ==============================================
    ' Browser Operations
    ' ==============================================
    Private Sub CloseAllBrowserInstances()
        ' SystemUtil.CloseProcessByName "firefox.exe"
    End Sub

    Private Sub LaunchBrowserWithURL(url)
        SystemUtil.Run "msedge.exe", url
    End Sub

    ' ==============================================
    ' Element Interaction Utilities
    ' ==============================================
    Private Function CreateWebElementDescriptor(elementType, attributeName, attributeValue)
        Dim elementDescription
        Set elementDescription = Description.Create()
        elementDescription("micclass").Value = elementType
        elementDescription(attributeName).Value = attributeValue
        Set CreateWebElementDescriptor = elementDescription
    End Function

    Private Function CreateBrowserDescriptor(namePattern)
        Dim browserDescription
        Set browserDescription = Description.Create()
        browserDescription("micclass").Value = "Browser"
        browserDescription("name").Value = namePattern
        Set CreateBrowserDescriptor = browserDescription
    End Function

    Private Function CreatePageDescriptor(titlePattern)
        Dim pageDescription
        Set pageDescription = Description.Create()
        pageDescription("micclass").Value = "Page"
        pageDescription("title").Value = titlePattern
        Set CreatePageDescriptor = pageDescription
    End Function

    Private Function IsElementVisible(elementDescriptor, timeoutSeconds)
        IsElementVisible = Browser(browserDescriptor).Page(salesforcePageDescriptor).WebEdit(elementDescriptor).Exist(timeoutSeconds)
    End Function

    Private Sub SetFieldValueWithValidation(elementDescriptor, value, fieldName)
        If IsElementVisible(elementDescriptor, ELEMENT_WAIT_TIMEOUT) Then
            Browser(browserDescriptor).Page(salesforcePageDescriptor).WebEdit(elementDescriptor).Set value
        Else
            ReportTestFailure fieldName & " Field Interaction", fieldName & " field not found within " & ELEMENT_WAIT_TIMEOUT & " seconds"
        End If
    End Sub

    Private Sub ClickElementWithValidation(elementDescriptor, elementName)
        If Browser(browserDescriptor).Page(salesforcePageDescriptor).WebButton(elementDescriptor).Exist(ELEMENT_WAIT_TIMEOUT) Then
            Browser(browserDescriptor).Page(salesforcePageDescriptor).WebButton(elementDescriptor).Click
        Else
            ReportTestFailure elementName & " Click", elementName & " not found"
        End If
    End Sub

    ' ==============================================
    ' Reporting Utilities
    ' ==============================================
    Private Sub ReportTestFailure(stepName, message)
        Reporter.ReportEvent micFail, stepName, "❌ FAILURE: " & message
    End Sub

    Private Sub ReportTestWarning(stepName, message)
        Reporter.ReportEvent micWarning, stepName, "⚠️ WARNING: " & message
    End Sub

    Private Sub ReportTestInfo(stepName, message)
        Reporter.ReportEvent micDone, stepName, "ℹ️ INFO: " & message
    End Sub

    ' ==============================================
    ' Business Logic Methods
    ' ==============================================
    Public Sub LoginToSalesforce(username, password)
        SetFieldValueWithValidation loginUsernameField, username, "Username"
        SetFieldValueWithValidation loginPasswordField, password, "Password"
        ClickLoginButton
    End Sub

    Private Sub ClickLoginButton()
        ClickElementWithValidation loginSubmitButton, "Login Button"
    End Sub

    Public Sub NavigateToAppSection(sectionName)
        Select Case sectionName
            Case "Contacts"
                NavigateToContactsSection
            Case "Contracts"
                NavigateToContractsSection
            Case "Home"
                ReportTestInfo "Navigation", "Home section selected"
            Case Else
                ReportTestWarning "Navigation", "Unknown section: " & sectionName
        End Select
    End Sub

    Private Sub NavigateToContactsSection()
        AIUtil.SetContext Browser(browserDescriptor)
        AIUtil.FindTextBlock("Contacts", micFromTop, 1).Click
        AIUtil.FindTextBlock("New", micFromTop, 1).Click
    End Sub

    Private Sub NavigateToContractsSection()
        AIUtil.SetContext Browser(browserDescriptor)
        AIUtil.FindTextBlock("Contracts", micFromTop, 1).Click
    End Sub

    Public Sub CreateNewContactRecord(phoneNumber, salutation, lastName)
        SetFieldValueWithValidation contactPhoneField, phoneNumber, "Phone"
        SelectContactSalutation salutation
        SetLastNameForContact lastName
        SaveContactRecord
    End Sub

    Private Sub SelectContactSalutation(salutation)
        AIUtil("combobox", "Salutation").Click
        AIUtil.FindTextBlock(salutation).Click
    End Sub

    Private Sub SetLastNameForContact(lastName)
        AIUtil("text_box", "Last Name").SetText lastName
    End Sub

    Private Sub SaveContactRecord()
        AIUtil("button", "Save").Click
    End Sub

    ' ==============================================
    ' API Test Utilities
    ' ==============================================
    Public Function ExecuteApiRequest(endpointUrl, httpMethod, requestBody)
        Dim httpRequest
        Set httpRequest = CreateObject("MSXML2.XMLHTTP")

        On Error Resume Next
        httpRequest.Open httpMethod, endpointUrl, False
        httpRequest.setRequestHeader "Content-Type", "application/json"
        httpRequest.Send requestBody

        If Err.Number <> 0 Then
            ReportTestFailure "API Request", "Error executing API call: " & Err.Description
            Set ExecuteApiRequest = Nothing
            Exit Function
        End If

        ProcessApiResponse httpRequest
        Set ExecuteApiRequest = httpRequest
        On Error GoTo 0
    End Function

    Private Sub ProcessApiResponse(httpRequest)
        If httpRequest.Status >= 200 And httpRequest.Status < 300 Then
            Reporter.ReportEvent micDone, "API Response", "✅ SUCCESS: Status " & httpRequest.Status
            LogMessage "API Response: " & httpRequest.responseText
        Else
            Reporter.ReportEvent micWarning, "API Response", "⚠️ UNEXPECTED STATUS: " & httpRequest.Status
            LogMessage "API Error Response: " & httpRequest.responseText
        End If
    End Sub

    Private Sub LogMessage(message)
        MsgBox message
    End Sub
End Class

' ==============================================
' Test Data Utilities
' ==============================================
Function LoadTestDataFromExcel(filePath, requiredColumns, worksheetName)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    If Not fileSystem.FileExists(filePath) Then
        MsgBox "Test data file not found: " & filePath, vbCritical
        Exit Function
    End If

    Dim testData, rowCount, columnIndex, rowIndex, columnName
    Set testData = CreateObject("Scripting.Dictionary")

    DataTable.ImportSheet filePath, 1, worksheetName
    rowCount = DataTable.GetSheet(worksheetName).GetRowCount()

    Dim allColumns, columnValues()
    ReDim allColumns(DataTable.GetSheet(worksheetName).GetParameterCount() - 1)

    For columnIndex = 1 To DataTable.GetSheet(worksheetName).GetParameterCount()
        allColumns(columnIndex - 1) = DataTable.GetSheet(worksheetName).GetParameter(columnIndex).Name
    Next

    For columnIndex = 0 To UBound(requiredColumns)
        columnName = requiredColumns(columnIndex)
        If ColumnExistsInArray(columnName, allColumns) Then
            ReDim columnValues(rowCount - 1)
            For rowIndex = 1 To rowCount
                DataTable.GetSheet(worksheetName).SetCurrentRow(rowIndex)
                columnValues(rowIndex - 1) = DataTable.Value(columnName, worksheetName)
            Next
            testData.Add columnName, columnValues
        Else
            MsgBox "Required column '" & columnName & "' missing in test data", vbExclamation
        End If
    Next

    Set LoadTestDataFromExcel = testData
End Function

Private Function ColumnExistsInArray(columnName, columnsArray)
    Dim i
    For i = LBound(columnsArray) To UBound(columnsArray)
        If StrComp(columnsArray(i), columnName, vbTextCompare) = 0 Then
            ColumnExistsInArray = True
            Exit Function
        End If
    Next
    ColumnExistsInArray = False
End Function

' ==============================================
' Test Cases
' ==============================================
Sub Test_Salesforce_Login_And_Contact_Creation()
    On Error Resume Next

    Dim salesforceAutomation
    Set salesforceAutomation = New SalesforceAutomation

    salesforceAutomation.TestSetup
    salesforceAutomation.LoginToSalesforce "negos11044@dlbazi.com", "Testing@123"

    Dim testDataColumns, testData
    testDataColumns = Array("salutation", "lastName", "phone")
    Set testData = LoadTestDataFromExcel(Environment.Value("TestDir") & "\Default.xlsx", testDataColumns, "Global")

    Dim contactName, successMessage
    contactName = "Dr.  Bugs Buny"
    successMessage = "Contact """ & contactName & """ was created"

    If Not testData Is Nothing Then
        If testData.Exists("salutation") And testData.Exists("lastName") And testData.Exists("phone") Then
            If UBound(testData("salutation")) >= 0 Then
                salesforceAutomation.NavigateToAppSection "Contacts"
                salesforceAutomation.CreateNewContactRecord testData("phone")(0), testData("salutation")(0), testData("lastName")(0)
                If Browser("Salesforce").Page("Salesforce Home").WebElement("text:=" & successMessage).Exist(10) Then
                    Reporter.ReportEvent micPass, "Contact Creation", "The contact was created successfully."
                Else
                    Reporter.ReportEvent micFail, "Contact Creation", "The success message was not found."
                End If
            End If
        End If
    End If

    salesforceAutomation.TestCleanup
    On Error GoTo 0
End Sub

Sub Test_View_Contracts_Associated_With_Assets()
    On Error Resume Next

    Dim salesforceAutomation
    Set salesforceAutomation = New SalesforceAutomation

    salesforceAutomation.TestSetup
    salesforceAutomation.LoginToSalesforce "negos11044@dlbazi.com", "Testing@123"
    salesforceAutomation.NavigateToAppSection "Contracts"

    ' Check if contracts and their related assets are visible
    If Browser("Salesforce").Page("Salesforce Contracts").WebElement("text:=Asset").Exist(10) Then
        Reporter.ReportEvent micPass, "Contract-Asset View", "Contracts associated with assets are visible."
    Else
        Reporter.ReportEvent micFail, "Contract-Asset View", "No asset-related records were found under contracts."
    End If

    salesforceAutomation.TestCleanup
    On Error GoTo 0
End Sub

Sub Test_RickAndMorty_API_Endpoint()
    On Error Resume Next
    Dim salesforceAutomation
    Set salesforceAutomation = New SalesforceAutomation
    Dim apiResponse
    Set apiResponse = salesforceAutomation.ExecuteApiRequest(RICKANDMORTY_API_ENDPOINT, "GET", "")
    On Error GoTo 0
End Sub

Sub Test_Call_Sf_Sanctuary_Cli()
	On Error Resume Next
	' Dim shell, command, exec, output
	
	' Set shell = CreateObject("WScript.Shell")
	' Print "Ruta del test actual: " & Environment.Value("TestDir")
	' shell.CurrentDirectory = Environment.Value("TestDir")

	' Set exec = shell.Exec("cmd /c ")
	' output = exec.StdOut.ReadAll
	' Print "Directorio actual: " & output
	' command = "cmd /c cd sf-sanctuary-cli/ && npm run start -- users deactivate -i 005Pu00000FW5ggIAD"
	
	' Set shell = CreateObject("WScript.Shell")
	' Set exec = shell.Exec(command)
	
	' Do While Not exec.StdOut.AtEndOfStream
  	'	line = exec.StdOut.ReadLine()
  	'	Print "Salida: " & line
	' Loop
	
	'Do While Not exec.StdErr.AtEndOfStream
  	'	line = exec.StdErr.ReadLine()
  	'	Print "Error: " & line
	'Loop
	
	Set exec = Nothing
	Set shell = Nothing
	On Error GoTo 0
End Sub


' ==============================================
' Test Execution
' ==============================================
Call Test_Salesforce_Login_And_Contact_Creation()
' Call Test_View_Contracts_Associated_With_Assets()
' Call Test_RickAndMorty_API_Endpoint()
' Call Test_Call_Sf_Sanctuary_Cli()




