Attribute VB_Name = "JOSEPHDOEModule"
Sub JOSEPHDOEPublishAllPDFs()
Set ws = Sheets("JAD PROJECTS")
saveLocation = "F:\ACTIVE PROJECTS\BACKUP\USERS\ACCOUNTING\JOSEPH DOE\_JOSEPH - PROJECT LIST.pdf"
ws.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation

Set ws = Sheets("ACTIVE PROJECTS")
saveLocation = "F:\ACTIVE PROJECTS\BACKUP\USERS\ACCOUNTING\JOSEPH DOE\ACTIVE PROJECT LIST.pdf"
ws.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation

Set ws = Sheets("FINISHED PROJECTS")
saveLocation = "F:\ACTIVE PROJECTS\BACKUP\USERS\ACCOUNTING\JOSEPH DOE\FINISHED PROJECT LIST.pdf"
ws.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation


End Sub

