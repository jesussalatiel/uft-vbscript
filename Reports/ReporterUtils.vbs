'==============================================
' Reporting Utilities
' ==============================================

Sub ReportTestFailure(step, message)
    Reporter.ReportEvent micFail, step, "❌ FAILURE: " & message
End Sub

Sub ReportTestWarning(step, message)
    Reporter.ReportEvent micWarning, step, "⚠️ WARNING: " & message
End Sub

Sub ReportTestInfo(step, message)
    Reporter.ReportEvent micDone, step, "ℹ️ INFO: " & message
End Sub
