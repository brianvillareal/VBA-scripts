Attribute VB_Name = "PriorDay"
Option Explicit
    
    Const strRT999     As String = "\\Etiam\tempor\orci\eu\lobortis\", _
          strHold      As String = "Athena Holding"
          
    Dim objFSO         As Object, _
        strAcronymCD   As String, _
        strAcronymPD   As String, _
        strFile        As String, _
        strFolderCD    As String, _
        strFolderPD    As String, _
        strSource      As String, _
        strDestination As String

Private Sub InitPriorDay(control As IRibbonControl)

    Const strPD As String = "Prior Day", _
          strUA As String = "Unapplied", _
          strUI As String = "Unidentified", _
          strUP As String = "Unpostable"
    
    strFolderCD = strRT999 & strHold
    
    Select Case control.ID
    
        Case "OriginalMove"
            
            strAcronymCD = "acd"
            strAcronymPD = "apd"
            strFolderPD = strRT999 & strPD & " File for Next Day"
                
        Case strUA & "Move"
        
            strAcronymCD = "cdua"
            strAcronymPD = "pdua"
            strFolderCD = strFolderCD & " " & strUA
            strFolderPD = strRT999 & strUA & " " & strPD
                
        Case strUI & "Move"
            
            strAcronymCD = "cdui"
            strAcronymPD = "pdui"
            strFolderCD = strFolderCD & " " & strUI
            strFolderPD = strRT999 & strUI & " (Revenue) " & strPD
                
        Case strUP & "Move"
            
            strAcronymCD = "cdup"
            strAcronymPD = "pdup"
            strFolderCD = strFolderCD & " " & strUP
            strFolderPD = strRT999 & strUP & " " & strPD
            
    End Select
    
    Call PrepareFiles
    Call DisplayMessagePD
    
End Sub

Private Sub PrepareFiles()
        
    strFile = Dir(strFolderPD & "\" & strAcronymPD & "*")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Call MovePD
    strFile = Dir(strFolderCD & "\" & strAcronymCD & "*")
    Call MoveCD
    Set objFSO = Nothing

End Sub

Private Sub MovePD()

    Do While Len(strFile) > 0

        strSource = strFolderPD & "\" & strFile
        strDestination = strRT999 & strFile
        
        objFSO.MoveFile Source:=strSource, _
                        Destination:=strDestination
                        
        strFile = Dir
    
    Loop

End Sub

Private Sub MoveCD()
    
    Do While Len(strFile) > 0
            
        strSource = strFolderCD & "\" & strFile
        strDestination = strRT999 & strFile
            
        objFSO.MoveFile Source:=strSource, _
                        Destination:=strDestination
            
        strSource = strDestination
        strFile = Replace(strFile, strAcronymCD, strAcronymPD)
        strDestination = strFolderPD & "\" & strFile
    
        objFSO.CopyFile Source:=strSource, _
                        Destination:=strDestination
                            
        strFile = Dir
    
    Loop

End Sub

Private Sub DisplayMessagePD()

    MsgBox "The move operation is complete.", vbOKOnly

End Sub
