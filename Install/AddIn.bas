Attribute VB_Name = "AddIn"
Option Explicit

Sub UpdateDictionary(control As IRibbonControl)

    Const strCurrentDic As String = "\\Lorem\ipsum$\dolor sit amet\Brian\Macros\Archive\Aging Dictionary.DIC"
    Dim strOldDic As String
    strOldDic = Environ("APPDATA") & "\Microsoft\UProof\Aging Dictionary.DIC"
    FileCopy strCurrentDic, strOldDic
    MsgBox "The aging dictionary has successfully updated."

End Sub

Sub InstallAddIn(control As IRibbonControl)
    
    Const strFolder As String = "\\Lorem\ipsum$\dolor sit amet\Brian\Macros\Current Add-ins"
    Dim AddIn As Excel.AddIn
    Dim objFile As Object, objFSO As Object
    Dim strFile As String, strAddIn As String, strInstall As String, strOld As String, strCRM As String
    Dim strDaily As String, strOther As String, strPRC As String, strRT10 As String, strRT10GA As String
    Dim wbkButton As Workbook
    
    Set wbkButton = ActiveWorkbook
    
    Select Case control.ID
    
        Case "OIR"
        
            strAddIn = "OIR v"

        Case "PRC"
        
            strAddIn = "PRC v"
            
        Case "Morning"
        
            strAddIn = "Daily Download v"
            
        Case "Other"
        
            strAddIn = "Other Loads v"
            
        Case "RT10"
        
            strAddIn = "RT10 Loads v"

        Case "buttonRT999"
        
            strAddIn = "Athena RT 999 v"
            
    End Select
    
    strFile = Dir(strFolder & "\*")
    
    Do While Len(strFile) > 0
    
        If InStr(1, strFile, strAddIn) Then

            strOld = Dir(Application.UserLibraryPath & "*")
            
            Do While Len(strOld) > 0

                If InStr(1, strOld, strAddIn) Then

                    If strOld = strFile Then
                    
                        MsgBox "It looks like you already have the latest" & vbCrLf & "version of the " & strAddIn & "Add-in.", vbInformation, "Hmm..."
                        wbkButton.Close
                        DoEvents
                        Exit Sub
                    
                    End If

                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    Set AddIn = Application.AddIns.Add(Application.UserLibraryPath & strOld)
                    AddIn.Installed = False
                    DoEvents
                    Kill Application.UserLibraryPath & strOld
                
                End If
                
                strOld = Dir

            Loop

            strInstall = strFolder & "\" & strFile
            FileCopy strInstall, Application.UserLibraryPath & strFile
            Set AddIn = Application.AddIns.Add(Application.UserLibraryPath & strFile)
            AddIn.Installed = True
            
            Exit Do
    
        End If
        
        strFile = Dir

    Loop

    wbkButton.Close
    DoEvents
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

' Brian Villareal.
' Changelog.
' 1.01 - 02/05/2019:
'   Daily Download added as an option for GA Web Client users (InstallAddIn).
'   strAddIn variables extended to differentiate GA Web Client add-ins from their respective base (InstallAddIn).
