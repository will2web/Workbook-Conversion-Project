Attribute VB_Name = "PublishingModule"
' Subroutine Macro to publish applicable all sheets


Sub PublishAll()

    Application.ScreenUpdating = False

    ' employee's full name
    Dim fullName As String
    ' variable used to iterate through distribution table
    Dim distributionTable As ListObject
    ' variable to hold the distribution table range
    Dim allEmployeeRange As Range
    ' Variable to iterate through for loop
    Dim eachEmployee As Integer



    ' Set list object title to the distribution table
    Set distributionTable = ThisWorkbook.Sheets("Distribution").ListObjects("Distribution")
    ' set range variable to distribution tables range
    Set allEmployeeRange = distributionTable.ListColumns(1).DataBodyRange






    ' for loop: initialize loop iteration variable to 1 distribution table range ...
    ' ... to count of all records in table and iterates through distribution table
    For eachEmployee = 1 To allEmployeeRange.Rows.Count

        ' Set full name string get full name from table, remove space ...
        ' ... between first and last name, and append "PublishAllPDFs"
        fullName = Replace(allEmployeeRange.Cells(eachEmployee, 2).Value _
                           , " ", "") & "PublishAllPDFs"

        ' run each employee's module
        Application.Run fullName



    Next eachEmployee




    Application.ScreenUpdating = True

End Sub


