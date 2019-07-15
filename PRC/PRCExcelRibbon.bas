Attribute VB_Name = "PRCExcelRibbon"
Option Explicit

Sub GetLabelPRC(control As IRibbonControl, ByRef returnedVal)

    Dim strWorkbookVersion  As String, _
        strLastSaveTime     As String

    strWorkbookVersion = Left(Right(ThisWorkbook.Name, 9), 4)
    strLastSaveTime = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "m/d/yy h:mm AM/PM")
    
    returnedVal = "Version " & vbNewLine & _
                    strWorkbookVersion & vbNewLine & _
                    "Updated " & vbNewLine & _
                    strLastSaveTime
    
End Sub

Sub Auto_Add()

	'* Auto_Add prevents errors when end-users use installation script to update add-in.

End Sub

