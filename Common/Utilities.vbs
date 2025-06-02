' Description: Loads the environment variables script (Environment.vbs).
' This subroutine should be called once at the beginning of your test execution.
' It ensures that variables defined in Environment.vbs (such as SALESFORCE_BASE_URL)
' are globally available to all your scripts and classes.

Sub LoadEnvironmentScript()

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Retrieve the base directory of test data from the UFT Environment object.
    ' This assumes 'TestDir' is correctly set in the UFT Environment Variables or config.ini.
    Dim folderPath
    folderPath = Environment.Value("TestDir")

    ' Check if TestDir is empty or not configured
    If folderPath = "" Then
        Reporter.ReportEvent micFail, "Environment Configuration Failed", _
                            "The environment variable 'TestDir' is not defined or is empty. Cannot load the environment script."
        Set fso = Nothing
        Exit Sub
    End If

    ' Get the parent folder of TestDir.
    ' This is essential if Environment.vbs is located one level above TestDir.
    Dim parentFolderPath
    parentFolderPath = fso.GetParentFolderName(folderPath)

    ' Construct the full path to the Environment.vbs file.
    ' Assumes Environment.vbs is placed directly under the parent folder of TestDir.
    Dim environmentFilePath
    environmentFilePath = parentFolderPath & "\Environment.vbs"

    ' Check if the environment script file exists before attempting to execute it.
    If fso.FileExists(environmentFilePath) Then
        ' Execute the VBScript file. This makes any global variables or functions
        ' defined within Environment.vbs available to the current script context.
        ExecuteFile environmentFilePath
        Reporter.ReportEvent micInfo, "Environment Configuration", _
                            "Environment script successfully loaded from: " & environmentFilePath
    Else
        ' Log a warning if the file is not found, as it is a critical component.
        Reporter.ReportEvent micWarning, "Environment Configuration - Warning", _
                            "Environment script not found at: " & environmentFilePath & ". Ensure the file exists and the 'TestDir' path is correct."
    End If

    Set fso = Nothing ' Release the FileSystemObject
End Sub
