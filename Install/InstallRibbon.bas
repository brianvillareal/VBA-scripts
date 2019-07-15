Attribute VB_Name = "InstallRibbon"
Option Explicit

Sub GetLabelInstall(control As IRibbonControl, ByRef returnedVal)
    
    Dim strWorkbookVersion As String, strLastSaveTime As String

    strWorkbookVersion = Split(ThisWorkbook.Name, "v")(1)
    strWorkbookVersion = Left(strWorkbookVersion, 4)
    strLastSaveTime = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "m/d/yy h:mm AM/PM")
    
    returnedVal = "Version " & vbCrLf & strWorkbookVersion & vbCrLf & "Updated " & vbCrLf & strLastSaveTime

End Sub
