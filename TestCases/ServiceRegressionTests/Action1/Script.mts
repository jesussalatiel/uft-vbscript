Dim browserMgr 


Set browserMgr = CreateBrowserGenerator()
browserMgr.LaunchBrowserWithURL G_SALESFORCE_BASE_URL, "edge"


' Usa Descriptive Programming sin depender del repositorio
Set oBrowser = browserMgr.GetSalesforceBrowser()
Set oPage = oBrowser.Page("title:=Login \| Salesforce")
Set oUsername = oPage.WebEdit("html id:=username")

oUsername.Set "Hello"

browserMgr.CloseAllBrowserInstances()
