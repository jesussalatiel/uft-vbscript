' Workaround to instantiate the SalesforceTestManager class defined in a resource
' https://community.opentext.com/devops-cloud/funct-testing/f/discussions/528993/unable-to-define-a-vbscript-class-in-a-function-library
Public Function Salesforce() 
    Dim manager
    Set manager = New SalesforceTestManager
    Set Salesforce = manager
End Function
