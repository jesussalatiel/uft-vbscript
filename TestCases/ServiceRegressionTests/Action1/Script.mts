Dim browserMgr, oBrowser, oPage 

' WebDriver Init
 Set oWebDriver = CreateWebDriver()
 oWebDriver.LaunchSalesforcePlatform

' Salesforce Management
Dim object
Set object = CreateSalesforceTestGenerator()
object.Init(oWebDriver)
object.Login "Julieta", "Jesus"
oWebDriver.CloseAllBrowserInstances()





