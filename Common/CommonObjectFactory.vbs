' Workaround to instantiate the BrowserManager class defined in a resource
' https://community.opentext.com/devops-cloud/funct-testing/f/discussions/528993/unable-to-define-a-vbscript-class-in-a-function-library
Public Function CreateBrowserGenerator()
    Dim manager
    Set manager = New BrowserManager
    Set CreateBrowserGenerator = manager
End Function
