Attribute VB_Name = "RT999ExcelRibbon"
Option Explicit

Sub GetLabelRT999(control As IRibbonControl, ByRef returnedVal)

    Dim strWorkbookVersion  As String, _
        strLastSaveTime     As String

    strWorkbookVersion = Left(Right(ThisWorkbook.Name, 9), 4)
    strLastSaveTime = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "m/d/yy h:mm AM/PM")
    
    returnedVal = "Version" _
        & vbNewLine & strWorkbookVersion _
        & vbNewLine & "Updated" _
        & vbNewLine & strLastSaveTime

End Sub

Sub Auto_Add()

	'* Auto_Add prevents errors when end-users use installation script to update add-in.

End Sub