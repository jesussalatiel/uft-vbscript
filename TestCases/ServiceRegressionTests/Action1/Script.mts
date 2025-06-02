'Call LaunchBrowserWithURL(G_SALESFORCE_BASE_URL)

Dim salesforce
Set salesforce = CreateSalesforceTestGenerator()
salesforce.TestSetup()
salesforce.Login "jesus.busta", "myPassword123"
salesforce.TestCleanup()
'Call CloseAllBrowserInstances()
