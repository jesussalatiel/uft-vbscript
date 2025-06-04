Dim browserMgr, oBrowser, oPage 

' WebDriver Init
 Set oWebDriver = CreateWebDriver()
 oWebDriver.LaunchSalesforcePlatform

' Salesforce Management
Dim object
Set object = CreateSalesforceTestGenerator()
object.Init(oWebDriver)
'object.Login "Julieta", "Jesus"
object.EnterUsernameUsingAI "Julieta"
oWebDriver.CloseAllBrowserInstances()


' AIUtil.SetContext Browser("creationtime:=0")
'AIUtil.SetContext Browser("creationtime:=0")
'AIUtil("text_box", "Username").Type "asdsd"





