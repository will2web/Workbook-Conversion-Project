Attribute VB_Name = "RemoveEmployeeModule"
Sub DeleteRowWithButton()
    Dim btnTop As Double
    Dim btnLeft As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim shp As Shape
    Dim threeInitial As String
    Dim fullName As String
        
    ' Unprotect Entire workbook so sheet and code can be deleted...
    ' ...and distribution sheet so row can be removed
    ThisWorkbook.Unprotect
    ThisWorkbook.Worksheets("Distribution").Unprotect
        
    'Get the button location and dimensions
    Set shp = ActiveSheet.Shapes(Application.Caller)
    btnTop = shp.Top
    btnLeft = shp.Left
    btnWidth = shp.Width
    btnHeight = shp.Height
    fullName = Mid(shp.Name, InStr(shp.Name, " ") + 1)
    threeInitial = Mid(shp.AlternativeText, InStr(shp.AlternativeText, " ") + 1)
    
    
' @@@@@@@@@@ START TEMP CODE

   ' variable used to iterate through distribution table
    Dim distributionTable As ListObject
    ' The new row that will be added containing the added employees info
    Dim addedEmployeeRow As ListRow
    ' the row containing the added employee's info
    Dim addedEmployeeDataRange As Range
    ' the variable to iterate over the added employee's info to add
    Dim iterateaddedEmployeeDataRange As Long
       

    'Specify the range containing added employee's info
    Set addedEmployeeDataRange = ThisWorkbook.Sheets("Distribution").Range("B2:I2")
   
    
    ' Set list object title to the distribution table
    Set distributionTable = ThisWorkbook.Sheets("Distribution").ListObjects("Distribution")
    ' get how many added employee's info cells to iterate through (should be 7)
    
    Dim headerRow As Range
    Dim headerRowNum As Long
    Set headerRow = distributionTable.HeaderRowRange
    headerRowNum = headerRow.Row
    
    
    Dim rowNumber As Long
    rowNumber = shp.TopLeftCell.Row 'rowNumber = 16,headerNumber = 7 using Sam as test case
    
    Set addedEmployeeRow = distributionTable.ListRows(rowNumber - headerRowNum)
    
    ' For each cell of the added employee's info ...
    For iterateaddedEmployeeDataRange = 1 To addedEmployeeDataRange.Columns.Count - 1
        '... Copy the data from the added employee's info range to the added employees info table row
        
        addedEmployeeDataRange.Cells(1, iterateaddedEmployeeDataRange).Value = _
        addedEmployeeRow.Range.Cells(1, iterateaddedEmployeeDataRange).Value
        
    ' go to next cell in added employee's info data range
    Next iterateaddedEmployeeDataRange

' @@@@@@@@@@ END TEMP CODE
    
    
    'Delete the row
    shp.TopLeftCell.EntireRow.Delete
    
    For Each shp In ActiveSheet.Shapes
        If shp.Name = "Remove " & fullName Then
            ' ... suppress warning about deleting a sheet ...
            
            'Delete the button
            shp.Delete
        End If
    Next shp
    
    
    RemoveEmployeeModuleAndSheet threeInitial, fullName
    
    ' Protect distribution sheet again
    ThisWorkbook.Worksheets("Distribution").Protect
    ThisWorkbook.Protect
    
End Sub





' Subroutine to remove employee's code module and sheet if applicable
Sub RemoveEmployeeModuleAndSheet(threeInitial As String, fullName As String)

'@@@@@@@@@@Sub RemoveEmployeeModuleAndSheet()

    ' employee's ...
    Dim firstName As String ' ... 1st name
    Dim ownSheet As String ' ... ownSheet name
    Dim connectionName As String '... connection name
    Dim sheetName As String
    
    ' set role-removal full name to value from Cell C2
    '@@@@@@@@@@ fullName = Range("C2").Value
    ' obtain first name from full name field by selecting everything before the space
    ' firstName = Left(fullName, InStr(fullName, " ") - 1)
    ' set role-removal ownsheet variable to value from Cell F2
        
    sheetName = threeInitial & " PROJECTS"
    
    connectionName = firstName & " Query"
    ' Call query and connection delete function
'NOT NEEDED    DeleteConnection connectionName
        
    For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = sheetName Then
        ' ... suppress warning about deleting a sheet ...
        Application.DisplayAlerts = False
        ws.Delete
        ' ... delete employees sheet ...
        ' ... and re-enable sheet deletion warnings
        Application.DisplayAlerts = True
    End If
        ownSheet = Range("F2").Value
    Next ws
    
    
    ' Call subroutine to remove employee module ...
    ' ... using employees full name and appending the word "module"
    RemoveEmployeeModule (Replace(fullName, " ", "") & "Module")
        
        
    ' if own sheet string is not empty ...
    If Not (ownSheet = "") Then
        ' ... suppress warning about deleting a sheet ...
        Application.DisplayAlerts = False
        ' ... delete employees sheet ...
        ' ThisWorkbook.Sheets(firstName).Delete
        ' ... and re-enable sheet deletion warnings
        Application.DisplayAlerts = True
    End If
    
    
End Sub





' SubRoutine to remove employees code module ...
' ...taking the condensed full employee name as an argument
Sub RemoveEmployeeModule(fullName As String)

    ' loop variable that will be assigned to each module in the workbook via the loop
    Dim vbComponent As Object
    ' boolean variable to check if module exists
    Dim moduleExists As Boolean
    Dim moduleName As String
    ' initialize module exists boolean variable to false
    moduleExists = False
    moduleName = fullName
    
    
    ' iterate through all modules in the workbook...
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        ' ...If a VBA component is of a type module,  and its name ...
        ' ... matches the module name we are searching for ...
        
        If vbComponent.Type = 1 And vbComponent.Name = moduleName Then
            '... set module exists Boolean variable to true
            moduleExists = True
            ' exit the for loop if the module we are looking for is found
    Exit For
            
        End If
    
    
    Next
    ' if the module that we are looking for exists ...
    If moduleExists Then
        ' ... remove it
        ThisWorkbook.VBProject.VBComponents.Remove vbComponent
    ' if it doesn't exist ...
    Else
    ' ... display a message to the be user saying the module was not found
        MsgBox "Module not found"
    End If
    
End Sub



Sub DeleteConnection(connectionName As String)
    Dim connection As WorkbookConnection
    
    ' Find the connection by name
    For Each connection In ThisWorkbook.Connections
        If connection.Name = connectionName Then
            ' Delete the connection
            connection.Delete
            Exit Sub
        End If
    Next connection
    
    ' Connection not found
    MsgBox "Connection not found: " & connectionName, vbCritical
End Sub





Sub PrintQueriesAndConnections()
    Dim query As WorkbookQuery
    Dim connection As WorkbookConnection
    
    ' Print all queries
    Debug.Print "Queries:"
    For Each query In ThisWorkbook.Queries
        Debug.Print "  " & query.Name & " (" & query.Type & ")"
    Next query
    
    ' Print all connections
    Debug.Print "Connections:"
    For Each connection In ThisWorkbook.Connections
        Debug.Print "  " & connection.Name & " (" & connection.Type & ")"
    Next connection
End Sub

