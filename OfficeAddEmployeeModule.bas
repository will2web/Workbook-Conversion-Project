Attribute VB_Name = "AddEmployeeModule"
' Main Subroutine for creating the new sheet and populating it accordingly
Sub EmployeeAddRoles()
    ' turn off screen updating to prevent flicker
    Application.ScreenUpdating = False
    
    ' Unprotect Entire workbook so sheet and code can be deleted...
    ' ...and distribution sheet so row can be removed
    ThisWorkbook.Unprotect
    ThisWorkbook.Worksheets("Distribution").Unprotect
    
    ' Variables for:
    
    ' employee's ...
    Dim department As String ' ... department
    Dim fullName As String ' ... full name
    Dim firstName As String ' ... 1st name
    Dim twoInitial As String ' ... 2-initial
    Dim threeInitial As String ' ... 3-initial
    
    
    
    ' determines if employee gets their own sheet
    Dim ownSheet As String
    ' determines if employee gets copy of active list
    Dim activeList As String
    ' determines if employee gets copy of finished list
    Dim finishedList As String
    ' the code string that will be save to the employee module for publishing their PDFS
    Dim publishPDFCode As String
    
    
    
    ' set employee's ...
    ' ... department from Cell B2
    department = Range("B2").Value
    ' ... full name value from Cell C2
    fullName = Range("C2").Value
    ' ... 2-inital value from Cell D2
    twoInitial = Range("D2").Value
    ' ... 3-inital value from Cell E2
    threeInitial = Range("E2").Value
    
    
    ' START Exception Handling Code
    ' Exception handling functions on confirming whether a module for the employee already...
    ' ...exists, as all employees receive modules because all employees get at least one sheet
    
    ' Module name variable initialization
    Dim moduleName As String
    ' Module name variable assignment
    'moduleName = threeInitial & "Module"  'Replace(fullName, " ", "") & "Module"
    moduleName = Replace(fullName, " ", "") & "Module"
   
    ' loop variable that will be assigned to each module in the workbook via the loop
    Dim vbComponent As Object
    ' boolean variable to check if module exists
    Dim moduleExists As Boolean
    
    ' initialize module exists boolean variable to false
    moduleExists = False
    
    
    ' iterate through all modules in the workbook...
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        ' ...If a VBA component is of a type module,  and its name ...
        ' ... matches the module name we are searching for ...
        
        If vbComponent.Type = 1 And vbComponent.Name = moduleName Then
            '... set module exists Boolean variable to true
            moduleExists = True
            
            ' Module already exists, show a message box and exit the sub
            MsgBox "Employee " & fullName & " already exists.", vbExclamation
            Exit Sub
        End If
    Next
    ' END Exception Handling Code
    
    
    ' obtain first name from full name field by selecting everything before the space
    firstName = Left(fullName, InStr(fullName, " ") - 1)
    ' if anything is in cell F2, employee gets a copy of their own sheet
    ownSheet = Range("F2").Value
    ' if anything is in cell G2, employee gets a copy of the office sheet
    activeList = Range("G2").Value
    ' if anything is in cell H2, employee gets a copy of the office sheet
    finishedList = Range("H2").Value
    ' initialize publish code string to an empty string because text will be added to it
    publishPDFCode = ""
    
    
    
    ' if own sheet string is not empty ...
    If Not (ownSheet = "") Then
    
    ' ... their publish code string gets the worksheet variable is set to their sheet ...
    ' ... the save location is determ   ined by their department and name  ...
    ' ... their sheet is named after them accordingly
    
    publishPDFCode = "Set ws = Sheets(""" & threeInitial & " PROJECTS" & """)" & vbCrLf & _
                     "saveLocation = ""F:\ACTIVE PROJECTS\BACKUP\USERS\" & department & "\" & _
                     fullName & "\_" & firstName & _
                     " - PROJECT LIST.pdf""" & vbCrLf & _
                     "ws.ExportAsFixedFormat Type:=xlTypePDF, _" & vbCrLf & _
                     "Filename:=saveLocation" & vbCrLf & vbCrLf
                      
    ' employee's full name, first name, two initial version, and ...
    ' ... three initial version  are all used to create their own their own sheet
    CreateEmployeeSheet fullName, firstName, twoInitial, threeInitial, department
    End If
    

    ' if active sheet string is not empty ...
    If Not (activeList = "") Then
    
    ' ... their publish code string gets the worksheet variable is set to the active list sheet ...
    ' ... the save location is determined by their department and name  ...
    ' ... the active list sheet is named accordingly
    publishPDFCode = publishPDFCode & "Set ws = Sheets(""" & "ACTIVE PROJECTS" & """)" & vbCrLf & _
                     "saveLocation = ""F:\ACTIVE PROJECTS\BACKUP\USERS\" & department & "\" & _
                     fullName & "\" & _
                     "ACTIVE PROJECT LIST.pdf""" & vbCrLf & _
                     "ws.ExportAsFixedFormat Type:=xlTypePDF, _" & vbCrLf & _
                     "Filename:=saveLocation" & vbCrLf & vbCrLf
    End If

    ' if active sheet string is not empty ...
    If Not (finishedList = "") Then
    
    ' ... their publish code string gets the worksheet variable is set to the finished list sheet ...
    ' ... the save location is determined by their department and name  ...
    ' ... the finished list sheet is named accordingly
    publishPDFCode = publishPDFCode & "Set ws = Sheets(""" & "FINISHED PROJECTS" & """)" & vbCrLf & _
                     "saveLocation = ""F:\ACTIVE PROJECTS\BACKUP\USERS\" & department & "\" & _
                     fullName & "\" & _
                     "FINISHED PROJECT LIST.pdf""" & vbCrLf & _
                     "ws.ExportAsFixedFormat Type:=xlTypePDF, _" & vbCrLf & _
                     "Filename:=saveLocation" & vbCrLf & vbCrLf
    End If
    
    
    ' Add employee to distribution table
    AddEmployeeToDistributionTable fullName, threeInitial
    
    ' create employees code module: their full name is used to name the module ...
    '  ... and it is populated with their publishing code
    CreateEmpolyeeModule fullName, publishPDFCode
    
    
    ' after everything else is done, refresh the data in the entire workbook to ...
    ' ... have the employees table populate accordingly
    ThisWorkbook.RefreshAll
    
    ' re-enable screen update
    Application.ScreenUpdating = True
    ' return focus to distribution sheet
    Worksheets("Distribution").Select
    
    ' Protect distribution sheet again
    ThisWorkbook.Worksheets("Distribution").Protect
    ThisWorkbook.Protect
End Sub





' SubRoutine that uses employee's full name, first name, two initial version, ...
' ... and three initial version to create their own sheet
Sub CreateEmployeeSheet(fullName As String, firstName As String, _
                        twoInitial As String, threeInitial As String, department As String)

    ' the employee sheet Workwsheet object
    Dim employeeSheet As Worksheet
    ' query table name that is populated with their work orders
    Dim queryTableName As ListObject


    ' Unhide template sheet
    ThisWorkbook.Sheets("____").Visible = True
    ' copy template sheet to the last position in the workbook
    ThisWorkbook.Sheets("____").Copy After:=Sheets(Sheets.Count)
    ' hide template sheet again
    ThisWorkbook.Sheets("____").Visible = False
    
    
    ' set employee sheet object to the active sheet just copied
    Set employeeSheet = ActiveSheet

    ' Name employee sheet object to employee's three-initial
    employeeSheet.Name = threeInitial & " PROJECTS"
    

    ' Set title to "CHAPARRAL - OFFICE WORKING LIST - " &  employee's 1st name
    ' Set title to "CHAPARRAL - PROJECT SUMMARY - ACTIVE PROJECTS                                                         "& employee's fullName - Department
    employeeSheet.Range("A1").Value = "CHAPARRAL - PROJECT SUMMARY - ACTIVE PROJECTS                                                         " & fullName & " - " & department
    
    
    ' Set query table name to table on newly created sheet
    Set queryTableName = employeeSheet.ListObjects(1)
    ' Change the table name to the employee's first name
    queryTableName.Name = firstName & "_Query_Table"
       
       
    ' call the subroutine to populate the template's query connection information ...
    ' ... which will be named after the employee, and template query will have ...
    ' .. information replaced with employee's to initial and three initial version
    QueryPopulator firstName, twoInitial, threeInitial

End Sub





' SubRoutine to populate  query
' ... which will be named after the employee's three initial version, and template query will have ...
' .. information replaced with employee's two initial and three initial version
Sub QueryPopulator(firstNameVersion As String, _
                   twoInitialVersion As String, threeInitialVersion As String)


    ' employee's just-created sheet Data Connection
    Dim employeeSheetDataConnection As WorkbookConnection
    ' query->connection properties->definition->command text
    Dim queryCommandText As String
   

    ' set the data connection to the template data connection just created
    Set employeeSheetDataConnection = ThisWorkbook.Connections("____ Query1")
    ' Rename the data connection to the employee's first name
    employeeSheetDataConnection.Name = threeInitialVersion & " Query"


    ' initialize the query command text string ...
    ' ... to the employee sheet's data connection command text
    queryCommandText = employeeSheetDataConnection.ODBCConnection.CommandText
    ' Replace "_@@" place holder with employee's 2-initial version
    queryCommandText = Replace(queryCommandText, "_@@", twoInitialVersion)
    ' Replace "_@_@" place holder with employee's 3-initial version
    queryCommandText = Replace(queryCommandText, "_@_@", threeInitialVersion)
   
    
    ' Update the connection's query command text
    employeeSheetDataConnection.ODBCConnection.CommandText = queryCommandText
    
End Sub





' SubRoutine to create employees code module: their full name is used to ...
'  ...  name the module and it is populated with their publishing code
Sub CreateEmpolyeeModule(fullName As String, publishPDFCode As String)

    'Create a new module
    Dim newModule As Object
    Dim moduleName As String
    
    
    ' set the new module to a new VBA code module
    Set newModule = ThisWorkbook.VBProject.VBComponents.Add(1)
    ' remove the space in the full name because modules names can't have spaces
    fullName = Replace(fullName, " ", "")
    'name the new module after employees full name and append word "Module"
    'moduleName = threeInitial & "Module"
    moduleName = fullName & "Module"
    newModule.Name = moduleName
    
    'Add a subroutine to the module ...
    ' ... who's code is specified in the published code string argument
    'newModule.CodeModule.AddFromString "Sub " & threeInitial & "PublishAllPDFs()" & _
                                        vbCrLf & publishPDFCode & vbCrLf & "End Sub"
    newModule.CodeModule.AddFromString "Sub " & fullName & "PublishAllPDFs()" & _
                                        vbCrLf & publishPDFCode & vbCrLf & "End Sub"
    
    'Save the workbook
    ThisWorkbook.Save
End Sub





' SubRoutine to add employee to the distribution list when "Add Employee" Button is pressed
Sub AddEmployeeToDistributionTable(fullName As String, threeInitial As String)

    ' variable used to iterate through distribution table
    Dim distributionTable As ListObject
    ' The new row that will be added containing the added employees info
    Dim addedEmployeeRow As ListRow
    ' the row containing the added employee's info
    Dim addedEmployeeDataRange As Range
    ' the variable to iterate over the added employee's info to add
    Dim iterateaddedEmployeeDataRange As Long
    ' Remove Employee Button object
    Dim removeEmployeeButton As button
    ' Cell for which to add remove Employee Button
    Dim removeButtonColumn As Range


    'Specify the range containing added employee's info
    Set addedEmployeeDataRange = ThisWorkbook.Sheets("Distribution").Range("B2:H2")
    ' Set list object title to the distribution table
    Set distributionTable = ThisWorkbook.Sheets("Distribution").ListObjects("Distribution")
    ' get how many added employee's info cells to iterate through (should be 7)
    Set addedEmployeeRow = distributionTable.ListRows.Add(distributionTable.ListRows.Count + 1)
    
    
    
   ' For each cell of the added employee's info ...
    For iterateaddedEmployeeDataRange = 1 To addedEmployeeDataRange.Columns.Count
        '... Copy the data from the added employee's info range to the added employees info table row
        addedEmployeeRow.Range.Cells(1, iterateaddedEmployeeDataRange).Value = _
        addedEmployeeDataRange.Cells(1, iterateaddedEmployeeDataRange).Value
        
    ' go to next cell in added employee's info data range
    Next iterateaddedEmployeeDataRange
    
    
    ' Add a button to the last cell in the new row
    Set removeButtonColumn = addedEmployeeRow.Range.Cells(1, addedEmployeeDataRange.Columns.Count + 1)



    Set removeEmployeeButton = ThisWorkbook.Sheets("Distribution").Buttons.Add _
                               (Left:=removeButtonColumn.Left, Top:=removeButtonColumn.Top, _
                               Width:=removeButtonColumn.Width, Height:=12)
    

    ' Add RemoveRowWithButton macro to the button
    With removeEmployeeButton
        .OnAction = "DeleteRowWithButton"
        .Name = "Remove " & fullName ' named after employee's 3-initial to distinguish it from other buttons
        .Text = "Remove " & threeInitial '" Remove " & Left(fullName, InStr(fullName, " ") - 1)
    End With

    ' Move the button with the row
    removeEmployeeButton.Placement = xlMove

    
   
    'Clear the values from the non-table range
    addedEmployeeDataRange.ClearContents
End Sub


