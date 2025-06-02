# Salesforce Automation Framework with UFT

This framework enables automated testing for Salesforce web applications using UFT (Unified Functional Testing). It supports both UI automation and API testing.

## Features

- **Salesforce UI Automation**:

  - Login functionality
  - Navigation across various sections
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
  - Failure, warning, and informational classification
  - Visual indicators in reports

## Prerequisites

### Software Requirements

1. **Micro Focus UFT** (version 14.x or higher)
2. **Microsoft Edge** browser (or configure your preferred browser)
3. **Microsoft Excel** (for test data management)

### Environment Setup

1. Create an `Environment` file with the following variables:

   - SALESFORCE_URL = [Your Salesforce instance URL]
   - RICKANDMORTY_API = [Rick and Morty API endpoint]
   - TestDir = [Path to your test data directory]

2. Prepare a test data Excel file (`Default.xlsx`) with the following sheet/columns:

- **Global** sheet with columns:
  - `salutation`
  - `lastName`
  - `phone`

## Framework Structure

```
📦 framework
├── 📜 Main.vbs                # Core automation classes
├── 📜 Tests.vbs               # Test cases
├── 📂 test-data
│   └── 📜 Default.xlsx        # Test data
└── 📂 environment
    └── 📜 config.ini          # Environment variables
```

## How to Use

### 1. Running Tests

Execute the test scripts from UFT:

```vbs
Call Test_Salesforce_Login_And_Contact_Creation()
Call Test_RickAndMorty_API_Endpoint()
```

### 2. Main Components

#### SalesforceAutomation Class

- `TestSetup()` — Initializes the browser and test environment
- `LoginToSalesforce(username, password)` — Performs login
- `NavigateToAppSection(sectionName)` — Navigates to the specified section
- `CreateNewContactRecord(phone, salutation, lastName)` — Creates a new contact
- `ExecuteApiRequest(url, method, body)` — Performs API calls

#### Test Data Management

```vbs
Set testData = LoadTestDataFromExcel(filePath, columnsArray, sheetName)
```

### 3. Customizing Tests

**To add new UI flows**:

1. Add new element descriptors in `InitializeUIElements()`.
2. Create new interaction methods following the existing pattern.

**To add new API tests**:

```vbs
Dim apiResponse
Set apiResponse = salesforceAutomation.ExecuteApiRequest(url, "GET|POST|PUT|DELETE", body)
```

## Best Practices

- **UI Tests**:

  - Always use helper methods like `SetFieldValueWithValidation`, `ClickElementWithValidation`
  - Maintain element descriptors in a central repository
  - Ensure transitions between pages are validated

- **API Tests**:

  - Validate status codes and handle edge cases
  - Process responses consistently
  - Use error handling to ensure test stability

- **Test Data**:
  - Separate data from logic/scripts
  - Validate required columns before running tests
  - Use clear and descriptive parameter names

## Troubleshooting

| Issue             | Solution                                    |
| ----------------- | ------------------------------------------- |
| Element not found | Check element descriptors and wait times    |
| Login failures    | Verify credentials and Salesforce URL       |
| API call errors   | Check endpoint URL and network connectivity |
| Test data issues  | Validate Excel file structure and contents  |

### Sample Test Case

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

- **UI Changes**:

  - Update object descriptors promptly
  - Adjust synchronization/wait logic as needed

- **New Features**:

  - Add new helper methods for UI/API logic
  - Keep implementation consistent with existing architecture

- **Routine Checks**:
  - Verify environment variables
  - Validate test data files
  - Review reports for flaky test patterns

## File Naming Conventions

| File Type          | Naming Convention                           | Example                        |
| ------------------ | ------------------------------------------- | ------------------------------ |
| Test Cases         | `TC_[Module]_[Description].vbs`             | `TC_Login_InvalidPassword.vbs` |
| Reusable Functions | `Functions_[Module].vbs`                    | `Functions_Login.vbs`          |
| Constants          | `Constants.vbs` or `Constants_[Module].vbs` | `Constants_Login.vbs`          |
| Object Repository  | `[Module].tsr`                              | `Login.tsr`                    |
| Data Files         | `[Module]_TestData.xlsx`                    | `Orders_TestData.xlsx`         |
| Recovery Scenarios | `GlobalRecovery.qrs`                        | `AppCrash.qrs`                 |

✅ **Best Practices**:

- **Modular structure**: Organize scripts by functional area (Login, Orders, Payments, etc.)
- **Clear prefixes**: Use `TC_`, `Functions_`, `Constants_` for clarity
- **Avoid generic names**: Use descriptive filenames (not `Test1.vbs`)
- **One logic per test case**: Keep each test case focused on a single scenario
- **Externalized data**: Use Excel/XML files in `/DataTables/`
- **Centralize reusable logic**: E.g., `ClickButton`, `WaitForObject`, `LoginAs`
- **Version control**: Use Git or ALM for tracking changes

## Contributors

- Jesus Bustamante

---

This README provides a complete guide to:

1. Framework features
2. Installation and setup
3. How to run and extend tests
4. Troubleshooting issues
5. Long-term maintenance

Suggestions for further improvements:

- Include specific UFT version notes
- Add screenshots of reports
- Define team coding standards
- Document CI/CD integration steps (if applicable)
