' WebDriver Init
Dim browser, app 
Set browser = CreateWebDriver()

' Start Browser
browser.LaunchSalesforcePlatform

' Salesforce Management
Set app = Salesforce()
app.Init(browser)
app.Login G_USERNAME, G_PASSWORD
app.Settings("Users")

' Finish Browser 
browser.CloseAllBrowserInstances()







