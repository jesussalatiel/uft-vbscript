# Unified Functional Testing (UFT)

## What is UFT?

**Unified Functional Testing (UFT)** is an automated software testing tool used to test a wide range of applications and environments, including web, desktop, mobile, and API-based applications.

## Supported Technologies

UFT can be used to automate testing of the following types of applications:

- Web
- Java
- .NET
- Flex
- Oracle
- SAP
- Salesforce
- PeopleSoft
- Siebel
- Delphi
- Terminal Emulators
- PowerBuilder
- Stingray
- VisualAge
- QT

## Scripting Language

- **VBScript**

## Supported Browsers

- Chrome
- Firefox
- Safari
- Internet Explorer

## Execution Environments

- Windows
- Linux
- macOS

## Common Object State Properties

Below are commonly used properties in UFT along with examples of how they can be accessed or verified using VBScript.

| Property           | Object Type(s)         | Description                                             | Example (VBScript)                                                                            |
| ------------------ | ---------------------- | ------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `Exist`            | Any                    | Indicates if the object exists on the page/UI           | `If Browser("MyApp").Page("Home").WebButton("Login").Exist(5) Then`                           |
| `Visible`          | WebElement, WebEdit    | Indicates if the object is visible to the user          | `isVisible = Browser("MyApp").Page("Home").WebEdit("Search").GetROProperty("visible")`        |
| `Disabled`         | WebButton, WebEdit     | Indicates whether the object is disabled (`true/false`) | `isDisabled = Browser("MyApp").Page("Home").WebButton("Submit").GetROProperty("disabled")`    |
| `Value`            | Inputs, dropdowns      | The current value of the control                        | `inputValue = Browser("MyApp").Page("Home").WebEdit("Email").GetROProperty("value")`          |
| `Checked`          | Checkbox, radio button | Whether the control is checked (`true/false`)           | `isChecked = Browser("MyApp").Page("Home").WebCheckBox("Subscribe").GetROProperty("checked")` |
| `Selected`         | Lists, combo boxes     | Whether an option is selected                           | `isSelected = Browser("MyApp").Page("Home").WebList("Country").GetROProperty("selected")`     |
| `Class`            | WebElement             | CSS class (useful for detecting visual states)          | `cssClass = Browser("MyApp").Page("Home").WebElement("Banner").GetROProperty("class")`        |
| `Status`           | Browser, Page          | Indicates if the page has fully loaded                  | `status = Browser("MyApp").GetROProperty("status")`                                           |
| `Enabled`          | Windows apps           | Indicates if the control is enabled                     | `isEnabled = Window("MyApp").WinEdit("Username").GetROProperty("enabled")`                    |
| `Text / innerText` | Labels, buttons        | The text displayed to the user                          | `textValue = Browser("MyApp").Page("Home").WebElement("Message").GetROProperty("innertext")`  |
| `Focused`          | WebElement             | Whether the control currently has focus                 | `hasFocus = Browser("MyApp").Page("Home").WebEdit("Password").GetROProperty("focused")`       |

## Advanced Examples

### <a name="mobile-example"></a>

<details>
<summary>Click to expand</summary>

### Mobile App Testing (UFT Mobile / Mobile Center)

```vb
' Test login functionality on a mobile banking app
Set device = MobileDevice("MyAndroid")
device.LaunchApp "com.bank.app"

' Wait for the login screen
If device.App("com.bank.app").MobileEdit("username").Exist(10) Then
    ' Fill in credentials
    device.App("com.bank.app").MobileEdit("username").Set "john.doe"
    device.App("com.bank.app").MobileEdit("password").SetSecure "MyEncryptedPassword"

    ' Click login
    device.App("com.bank.app").MobileButton("login_button").Tap

    ' Validate if home screen is displayed
    If device.App("com.bank.app").MobileText("welcome_message").Exist(15) Then
        Reporter.ReportEvent micPass, "Login Test", "Login successful, home screen displayed"
    Else
        Reporter.ReportEvent micFail, "Login Test", "Home screen not loaded"
    End If
Else
    Reporter.ReportEvent micFail, "Login Screen", "Username field not found"
End If
```

### 🔌 API Testing (via UFT API / Service Test)

```vb
' Create a user via POST and validate using GET
Set createUser = CreateObject("HP.ST.Fwk.RunTimeObjects.RESTActivity")
createUser.Endpoint = "https://api.example.com/users"
createUser.Method = "POST"
createUser.Body = "{""name"": ""John"", ""email"": ""john@example.com""}"
createUser.ContentType = "application/json"
createUser.Send

If createUser.StatusCode = 201 Then
    Reporter.ReportEvent micPass, "User Creation", "User created successfully"

    ' Extract user ID from response
    Dim userId
    userId = ExtractUserId(createUser.ResponseBody)

    ' Fetch the user
    Set getUser = CreateObject("HP.ST.Fwk.RunTimeObjects.RESTActivity")
    getUser.Endpoint = "https://api.example.com/users/" & userId
    getUser.Method = "GET"
    getUser.Send

    If InStr(getUser.ResponseBody, """email"":""john@example.com""") > 0 Then
        Reporter.ReportEvent micPass, "User Verification", "User data matches"
    Else
        Reporter.ReportEvent micFail, "User Verification", "User data mismatch"
    End If
Else
    Reporter.ReportEvent micFail, "User Creation", "Failed with status " & createUser.StatusCode
End If

Function ExtractUserId(response)
    ' Dummy example: simple extraction using RegExp
    Dim regEx, matches
    Set regEx = New RegExp
    regEx.Pattern = """id"":\s*(\d+)"
    regEx.Global = False
    regEx.IgnoreCase = True

    Set matches = regEx.Execute(response)
    If matches.Count > 0 Then
        ExtractUserId = matches(0).SubMatches(0)
    Else
        ExtractUserId = ""
    End If
End Function
```

### Desktop App Testing (Windows-based apps)

```vb
' End-to-end test of login + report generation in a desktop app
If Window("AccountingApp").Exist(10) Then
    Window("AccountingApp").WinEdit("Username").Set "admin"
    Window("AccountingApp").WinEdit("Password").SetSecure "EncryptedPwd"
    Window("AccountingApp").WinButton("Login").Click

    ' Wait for main screen
    If Window("AccountingApp").WinMenu("MainMenu").Exist(15) Then
        Window("AccountingApp").WinMenu("MainMenu").Select "Reports;Generate Monthly Report"

        ' Set parameters
        Window("AccountingApp").WinComboBox("Month").Select "March"
        Window("AccountingApp").WinButton("Generate").Click

        ' Validate report generation
        If Window("AccountingApp").WinStatic("ReportStatus").GetROProperty("text") = "Report Generated Successfully" Then
            Reporter.ReportEvent micPass, "Report Generation", "Report created correctly"
        Else
            Reporter.ReportEvent micFail, "Report Generation", "Report creation failed"
        End If
    Else
        Reporter.ReportEvent micFail, "Main Menu", "Main screen not displayed after login"
    End If
Else
    Reporter.ReportEvent micFail, "App Launch", "Application not found"
End If
```

</details>

### AOM (Automation Object Model)

It is an interface COM expose to UFT allows control UFT with external scripts, examples: VBScript, Poweshell or Python using pywin32.

```py
import win32com.client
import time

# Crear el objeto de UFT
qtApp = win32com.client.Dispatch("QuickTest.Application")

# Iniciar UFT si no está abierto
if not qtApp.Launched:
    qtApp.Launch()
    time.sleep(5)  # espera que termine de cargar

# Mostrar UFT (opcional)
qtApp.Visible = True

# Abrir una prueba existente
test_path = r"C:\Tests\LoginTest"
qtApp.Open(test_path, True)

# Ejecutar la prueba
qtApp.Test.Run()

# Esperar que termine
while qtApp.Test.IsRunning:
    time.sleep(1)

# Obtener resultados
results = qtApp.Test.LastRunResults
print("Test Result: ", results.Status)

# Cerrar la prueba y UFT
qtApp.Test.Close()
qtApp.Quit()

# Liberar objeto
del qtApp
```

### Export multiple files

```vb
ExecuteFile "lib\Logger.vbs"
ExecuteFile "utils\Validator.vbs"
ExecuteFile "factories\BrowserFactory.vbs"
```

### 📁 Organize Your Project by Responsibility

```bash
/tests           → Test cases
/utils           → Shared validation and helper functions
/factories       → BrowserFactory, DeviceFactory, etc.
/data            → External files (CSV, Excel)
```

### Use SetSecure for Sensitive Data

Never hardcode passwords in plain text:

```vb
Browser("App").Page("Login").WebEdit("password").SetSecure "MyEncryptedPassword"
```

### Check for Object Existence Before Interacting

Prevent failures by validating objects before using them:

```vb
If Browser("App").Page("Login").WebEdit("Username").Exist(5) Then
    Browser("App").Page("Login").WebEdit("Username").Set "admin"
End If
```

### 🧪 Use Reporter.ReportEvent for Custom Reporting

Make your test results readable:

```vb
Reporter.ReportEvent micPass, "Login Test", "Login successful"
Reporter.ReportEvent micFail, "Login Test", "Login failed"
```

### Use Dictionaries to Simulate Objects

VBScript doesn’t support real objects, but you can use dictionaries:

```vb
Set user = CreateObject("Scripting.Dictionary")
user.Add "username", "john"
user.Add "password", "12345678"
```

## References

- [Guru99 UFT/QTP Tutorial](https://www.guru99.com/quick-test-professional-qtp-tutorial.html)
