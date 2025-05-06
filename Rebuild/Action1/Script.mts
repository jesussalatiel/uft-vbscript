' Constants for better maintainability
Const SALESFORCE_URL = "https://globanttesting-dev-ed.develop.lightning.force.com/lightning"
Const PAGE_LOAD_WAIT_TIME = 5
Const DEFAULT_WAIT_TIME = 5

' Main Salesforce automation class
Class Salesforce
    ' Object repository - using descriptive names
    Private oUsernameField, oPasswordField, oLoginButton, oPhoneField
    Private oBrowserDescriptor, oPageDescriptor
    Private browserTitlePattern

    ' Initialize all object descriptions
    Public Sub Initialize()
        Set oUsernameField = CreateWebElement("WebEdit", "html id", "username")
        Set oPasswordField = CreateWebElement("WebEdit", "html id", "password")
        Set oLoginButton = CreateWebElement("WebButton", "html id", "Login")
        Set oPhoneField = CreateWebElement("WebEdit", "name", "Phone")

        Set oBrowserDescriptor = CreateBrowserDescriptor(".*")
        Set oPageDescriptor = CreatePageDescriptor(".*Salesforce.*")
    End Sub

    ' Helper method to create web elements
    Private Function CreateWebElement(elementClass, attributeName, attributeValue)
        Dim elementDesc
        Set elementDesc = Description.Create()
        elementDesc("micclass").Value = elementClass
        elementDesc(attributeName).Value = attributeValue
        Set CreateWebElement = elementDesc
    End Function

    ' Helper method to create browser descriptor
    Private Function CreateBrowserDescriptor(namePattern)
        Dim browserDesc
        Set browserDesc = Description.Create()
        browserDesc("micclass").Value = "Browser"
        browserDesc("name").Value = namePattern
        Set CreateBrowserDescriptor = browserDesc
    End Function

    ' Helper method to create page descriptor
    Private Function CreatePageDescriptor(titlePattern)
        Dim pageDesc
        Set pageDesc = Description.Create()
        pageDesc("micclass").Value = "Page"
        pageDesc("title").Value = titlePattern
        Set CreatePageDescriptor = pageDesc
    End Function

    ' Test setup
    Public Sub BeforeTest()
        Initialize
        CloseBrowser
        LaunchBrowser SALESFORCE_URL
        browserTitlePattern = ".*Salesforce.*"

        If Not IsElementVisible(oUsernameField, PAGE_LOAD_WAIT_TIME) Then
            ReportFailure "Page Load", "Username field not found within " & PAGE_LOAD_WAIT_TIME & " seconds"
        End If
    End Sub

    ' Test teardown
    Public Sub AfterTest()
        CloseBrowser
    End Sub

    ' Browser operations
    Private Sub CloseBrowser()
        SystemUtil.CloseProcessByName "msedge.exe"
    End Sub

    Private Sub LaunchBrowser(url)
        SystemUtil.Run "msedge.exe", url
    End Sub

    ' Element interaction methods
    Private Function IsElementVisible(elementDesc, timeout)
        IsElementVisible = Browser(oBrowserDescriptor).Page(oPageDescriptor).WebEdit(elementDesc).Exist(timeout)
    End Function

    Private Sub SetFieldIfExists(elementDesc, value, fieldName)
        If IsElementVisible(elementDesc, DEFAULT_WAIT_TIME) Then
            Browser(oBrowserDescriptor).Page(oPageDescriptor).WebEdit(elementDesc).Set value
        Else
            ReportFailure fieldName & " field", "The " & fieldName & " field is not visible after " & DEFAULT_WAIT_TIME & " seconds"
        End If
    End Sub

    ' Reporting helper
    Private Sub ReportFailure(stepName, message)
        Reporter.ReportEvent micFail, stepName, "âŒ " & message
    End Sub

    ' Business logic methods
    Public Sub Login(username, password)
        SetFieldIfExists oUsernameField, username, "Username"
        SetFieldIfExists oPasswordField, password, "Password"
        ClickLoginButton
    End Sub

    Private Sub ClickLoginButton()
        If Browser(oBrowserDescriptor).Page(oPageDescriptor).WebButton(oLoginButton).Exist(DEFAULT_WAIT_TIME) Then
            Browser(oBrowserDescriptor).Page(oPageDescriptor).WebButton(oLoginButton).Click
        Else
            ReportFailure "Login Button", "Login button not found"
        End If
    End Sub

    Public Sub NavigateToSection(sectionName)
        Select Case sectionName
            Case "Contacts"
                NavigateToContacts
            Case "Home"
                ReportNavigation "Home selected"
            Case Else
                ReportWarning "Navigation", "Invalid option: " & sectionName
        End Select
    End Sub

    Private Sub NavigateToContacts()
        AIUtil.SetContext Browser(oBrowserDescriptor)
        AIUtil.FindTextBlock("Contacts", micFromTop, 1).Click
        AIUtil.FindTextBlock("New", micFromTop, 1).Click
    End Sub

    Private Sub ReportNavigation(message)
        Reporter.ReportEvent micDone, "Navigation", "í ½í³Œ " & message
    End Sub

    Private Sub ReportWarning(stepName, message)
        Reporter.ReportEvent micWarning, stepName, "âš ï¸ " & message
    End Sub

    Public Sub CreateNewContact(phone, salutation, lastName)
        SetFieldIfExists oPhoneField, phone, "Phone"
        SelectSalutation salutation
        AIUtil("text_box", "Last Name").SetText lastName
        AIUtil("button", "Save").Click
    End Sub

    Private Sub SelectSalutation(salutation)
        AIUtil("combobox", "Salutation").Click
        AIUtil.FindTextBlock(salutation).Click
    End Sub

End Class

' Test script
Sub Test_Login_And_Create_Contact()
    On Error Resume Next

    Dim salesforce
    Set salesforce = New Salesforce

    salesforce.BeforeTest
    salesforce.Login "jesus.bustamante@globant.com", "Testing@123"

    Dim columns, data
    columns = Array("salutation", "lastName", "phone")
    Set data = ReadExcel(Environment.Value("TestDir") & "\Default.xlsx", columns)

    If Not data Is Nothing Then
        If data.Exists("salutation") And data.Exists("lastName") And data.Exists("phone") Then
            If UBound(data("salutation")) >= 0 Then
                salesforce.NavigateToSection "Contacts"
                salesforce.CreateNewContact data("phone")(0), data("salutation")(0), data("lastName")(0)
            End If
        End If
    End If

    salesforce.AfterTest
    On Error GoTo 0
End Sub

Function ReadExcel(fileName, desiredColumns)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(fileName) Then
        MsgBox "File does not exist: " & fileName, vbCritical
        Exit Function
    End If

    Dim dataDict, rowCount, i, j, colName
    Set dataDict = CreateObject("Scripting.Dictionary")

    DataTable.ImportSheet fileName, 1, "Global"
    rowCount = DataTable.GetSheet("Global").GetRowCount()

    Dim colCount, columnsInFile()
    colCount = DataTable.GetSheet("Global").GetParameterCount()
    ReDim columnsInFile(colCount - 1)
    For i = 1 To colCount
        columnsInFile(i - 1) = DataTable.GetSheet("Global").GetParameter(i).Name
    Next

    Dim values()
    For i = 0 To UBound(desiredColumns)
        colName = desiredColumns(i)
        If IsInArray(colName, columnsInFile) Then
            ReDim values(rowCount - 1)
            For j = 1 To rowCount
                DataTable.GetSheet("Global").SetCurrentRow(j)
                values(j - 1) = DataTable.Value(colName, "Global")
            Next
            dataDict.Add colName, values
        Else
            MsgBox "Column '" & colName & "' does not exist in the file.", vbExclamation
        End If
    Next

    Set ReadExcel = dataDict
End Function

Private Function IsInArray(valueToFind, arr)
    Dim i
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), valueToFind, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Call Test_Login_And_Create_Contact()
