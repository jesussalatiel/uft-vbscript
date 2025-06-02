' ==============================================
' Test Data Utilities
' ==============================================

Function LoadTestDataFromExcel(filePath, requiredColumns, worksheetName)
    Dim fileSystem
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    If Not fileSystem.FileExists(filePath) Then
        MsgBox "Test data file not found: " & filePath, vbCritical
        Exit Function
    End If

    Dim testData, rowCount, columnIndex, rowIndex, columnName
    Set testData = CreateObject("Scripting.Dictionary")

    DataTable.ImportSheet filePath, 1, worksheetName
    rowCount = DataTable.GetSheet(worksheetName).GetRowCount()

    Dim allColumns, columnValues()
    ReDim allColumns(DataTable.GetSheet(worksheetName).GetParameterCount() - 1)

    For columnIndex = 1 To DataTable.GetSheet(worksheetName).GetParameterCount()
        allColumns(columnIndex - 1) = DataTable.GetSheet(worksheetName).GetParameter(columnIndex).Name
    Next

    For columnIndex = 0 To UBound(requiredColumns)
        columnName = requiredColumns(columnIndex)
        If ColumnExistsInArray(columnName, allColumns) Then
            ReDim columnValues(rowCount - 1)
            For rowIndex = 1 To rowCount
                DataTable.GetSheet(worksheetName).SetCurrentRow(rowIndex)
                columnValues(rowIndex - 1) = DataTable.Value(columnName, worksheetName)
            Next
            testData.Add columnName, columnValues
        Else
            MsgBox "Required column '" & columnName & "' missing in test data", vbExclamation
        End If
    Next

    Set LoadTestDataFromExcel = testData
End Function

Private Function ColumnExistsInArray(columnName, columnsArray)
    Dim i
    For i = LBound(columnsArray) To UBound(columnsArray)
        If StrComp(columnsArray(i), columnName, vbTextCompare) = 0 Then
            ColumnExistsInArray = True
            Exit Function
        End If
    Next
    ColumnExistsInArray = False
End Function