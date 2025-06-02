' ==============================================
' API Test Utilities
' ==============================================
Function ExecuteApiRequest(endpointUrl, httpMethod, requestBody)
    Dim httpRequest: Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    httpRequest.Open httpMethod, endpointUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.Send requestBody

    If Err.Number <> 0 Then
        ReportTestFailure "API Request", "Error executing API call: " & Err.Description
        Set ExecuteApiRequest = Nothing
        Exit Function
    End If

    ProcessApiResponse httpRequest
    Set ExecuteApiRequest = httpRequest
    On Error GoTo 0
End Function

Sub ProcessApiResponse(httpRequest)
    If httpRequest.Status >= 200 And httpRequest.Status < 300 Then
        Reporter.ReportEvent micDone, "API Response", "✅ SUCCESS: " & httpRequest.Status
        MsgBox httpRequest.responseText
    Else
        Reporter.ReportEvent micWarning, "API Response", "⚠️ UNEXPECTED STATUS: " & httpRequest.Status
        MsgBox httpRequest.responseText
    End If
End Sub
