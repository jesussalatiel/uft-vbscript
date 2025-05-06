# Salesforce Automation Framework with UFT

This framework provides automated testing capabilities for Salesforce web applications using UFT (Unified Functional Testing). It includes both UI automation and API testing functionalities.

## Features

- **Salesforce UI Automation**:

  - Login functionality
  - Navigation through different sections
  - Contact creation workflow
  - Robust element identification and interaction

- **API Testing**:

  - REST API call capabilities
  - Response validation
  - Error handling

- **Test Data Management**:

  - Excel-based test data loading
  - Data validation
  - Parameterized testing

- **Reporting**:

  - Detailed test execution reports
  - Failure/warning/info classification
  - Visual indicators in reports

## Prerequisites

### Software Requirements

1. **Micro Focus UFT** (version 14.x or higher)
2. **Microsoft Edge** browser (or configure for your preferred browser)
3. **Microsoft Excel** (for test data management)

### Environment Setup

1. Create an `Environment` file with these variables:

   1. SALESFORCE_URL = [Your Salesforce instance URL]
   2. RICKANDMORTY_API = [Rick and Morty API endpoint]
   3. TestDir = [Path to your test data directory]

2. Prepare test data Excel file (`Default.xlsx`) with these sheets/columns:

- **Global** sheet with columns:
  - `salutation`
  - `lastName`
  - `phone`

## Framework Structure

📦 framework
├── 📜 Main.vbs - Core automation classes
├── 📜 Tests.vbs - Test cases
├── 📂 test-data
│ └── 📜 Default.xlsx - Test data
└── 📂 environment
└── 📜 config.ini - Environment variables

## How to Use

### 1. Running Tests

Execute the test scripts from UFT:

```vbs
Call Test_Salesforce_Login_And_Contact_Creation()
Call Test_RickAndMorty_API_Endpoint()
```

### 2. Main Components

#### SalesforceAutomation Class

- `TestSetup()` - Initializes browser and test environment
- `LoginToSalesforce(username, password)` - Performs login
- `NavigateToAppSection(sectionName)` - Navigates to specified section
- `CreateNewContactRecord(phone, salutation, lastName)` - Creates new contact
- `ExecuteApiRequest(url, method, body)` - Makes API calls

#### Test Data Management

```vbs
Set testData = LoadTestDataFromExcel(filePath, columnsArray, sheetName)
```

### 3. Customizing Tests

**To add new UI flows** :

1. Add new element descriptors in `InitializeUIElements()`
2. Create new interaction methods following existing patterns

**To add new API tests** :

```vbs
Dim apiResponse
Set apiResponse = salesforceAutomation.ExecuteApiRequest(url, "GET|POST|PUT|DELETE", body)
```

## Best Practices

1. **For UI Tests** :

- Always use the element interaction methods (`SetFieldValueWithValidation`, `ClickElementWithValidation`)
- Maintain element descriptors in the repository section
- Keep page transitions validated

1. **For API Tests** :

- Validate status codes
- Process responses consistently
- Handle errors gracefully

1. **For Test Data** :

- Keep test data separate from scripts
- Validate required columns before test execution
- Use descriptive names for parameters

## Troubleshooting

| Issue             | Solution                                    |
| ----------------- | ------------------------------------------- |
| Element not found | Check element descriptors and wait times    |
| Login failures    | Verify credentials and Salesforce URL       |
| API call errors   | Check endpoint URL and network connectivity |
| Test data issues  | Validate Excel file structure and content   |

```vbs
' Test: Create contact with data from Excel
Sub Test_Create_Contact_With_Test_Data()
Dim salesforce, testData

    Set salesforce = New SalesforceAutomation
    salesforce.TestSetup

    ' Login
    salesforce.LoginToSalesforce "user@company.com", "securePassword"

    ' Load test data
    Set testData = LoadTestDataFromExcel(Environment.Value("TestDir") & "\Contacts.xlsx", _
                    Array("salutation", "lastName", "phone"), "TestCases")

    ' Execute test
    salesforce.NavigateToAppSection "Contacts"
    salesforce.CreateNewContactRecord testData("phone")(0), testData("salutation")(0), testData("lastName")(0)

    salesforce.TestCleanup

End Sub
```

## Maintenance

1. **For UI Changes** :

- Update element descriptors when UI changes
- Review and adjust wait times as needed

1. **For New Features** :

- Extend the framework with new methods
- Maintain consistent patterns

1. **Regular Checks** :

- Verify environment variables
- Validate test data files
- Review test reports for flaky tests

## Contributors

- [Jesus Bustamante]

This README provides comprehensive documentation covering:

1. Framework capabilities
2. Setup requirements
3. Usage instructions
4. Best practices
5. Troubleshooting
6. Maintenance guidelines

The document is structured to help new users get started quickly while providing enough detail for advanced customization. You may want to:

- Add specific version requirements
- Include screenshots of sample reports
- Add your team's specific coding standards
- Include any CI/CD integration details if applicable
