Attribute VB_Name = "LoadedRT1"
Option Explicit
          
    Dim strArchive As String, _
        strButton  As String, _
        strMessage As String, _
        strRT1     As String, _
        wbkArchive As Workbook, _
        wbkRecon   As Workbook, _
        wksRecon   As Worksheet

Private Sub InitLoadedRT1(control As IRibbonControl)

    strButton = Replace(control.ID, "ReconRT1", vbNullString)
    Call SetAthenaType
    Application.ScreenUpdating = False
    Set wbkArchive = OpenFileWait(strArchive)
    strArchive = vbNullString
    Set wbkRecon = OpenFileWait(strRT1)
    strRT1 = vbNullString
    Set wksRecon = wbkRecon.Worksheets(1)
    Call ImportBalance("C", "G", "I")
    Set wksRecon = Nothing
    wbkRecon.Close SaveChanges:=False
    Set wbkRecon = Nothing
    wbkArchive.Close SaveChanges:=True
    Set wbkArchive = Nothing
    Application.ScreenUpdating = True
    Call ShowMessage

End Sub

Private Function SetAthenaType()

    'Const strBalancing As String = "\\Lorem\ipsum$\dolor sit amet\Brian\Macros\In Progress\Test\Athena Balancing ",
    Const strBalancing As String = "\\Lorem\ipsum$\dolor sit amet\IS\Athena Balancing Files\Athena Balancing ", _
          strReports   As String = "\\Etiam\tempor\orci\eu\lobortis\FTP Holding\"

    strArchive = strBalancing & strButton & ".xlsx"
    strRT1 = strReports & "RT1 " & strButton & ".xls"

End Function

Private Function OpenFileWait(strFile As String) As Workbook

    Set OpenFileWait = Application.Workbooks.Open(strFile)
        
    Do Until OpenFileWait.ReadOnly = False
        
        OpenFileWait.Close
        Application.Wait Now + TimeValue("00:00:01")
        Set OpenFileWait = Application.Workbooks.Open(strFile)
        
    Loop

End Function

Private Sub ImportBalance(AccountNumber As String, LoadColumn As String, DifferenceColumn As String)

    Dim dblBalance As Double, _
        i          As Long, _
        lngAccount As Long, _
        lngLastRow As Long, _
        strTab     As String, _
        wksArchive As Worksheet
    
    lngLastRow = wksRecon.Cells(Rows.Count, "D").End(xlUp).Row - 1
    
    For i = 2 To lngLastRow
    
        lngAccount = wksRecon.Range(AccountNumber & i)
        dblBalance = wksRecon.Range("F" & i)
        strTab = ReturnAccountName(lngAccount)
        Set wksArchive = wbkArchive.Worksheets("Daily Process-" & strTab)
        lngLastRow = wksArchive.Cells(Rows.Count, "C").End(xlUp).Row
        
        If strButton = "Unpostable" Then
        
            dblBalance = dblBalance * -1
            
        End If
        
        If strButton = "Unidentified" Then
            
            dblBalance = Abs(dblBalance)
            
            If Range("D" & lngLastRow) < Range("E" & lngLastRow) Then

                dblBalance = dblBalance * -1
            
            End If
        
        End If
        
        wksArchive.Range(LoadColumn & lngLastRow) = dblBalance
            
        If wksArchive.Range(DifferenceColumn & lngLastRow) <> 0 Then
            
            strMessage = strMessage & vbNewLine & _
                         strTab
                
        End If
        
    Next i
    
    Set wksArchive = Nothing

End Sub

Private Function ReturnAccountName(Account As Long) As String
    
    Select Case Account
    
        Case 88, 404, 405, 442
        
            ReturnAccountName = "FHMG"

        Case 91, 438, 439, 459
            ReturnAccountName = "FHNS"

        Case 132, 270, 436, 437, 458
        
            ReturnAccountName = "Gordon"

        Case 144, 408, 409, 444
        
            ReturnAccountName = "Memorial"

        Case 164, 418, 419, 449
        
            ReturnAccountName = "Park Ridge"

        Case 242, 402, 403, 441
        
            ReturnAccountName = "AHP"

        Case 244, 432, 433, 456
        
            ReturnAccountName = "LOHP"

        Case 246, 416, 417, 448
        
            ReturnAccountName = "Chippewa"

        Case 248, 414, 415, 447
        
            ReturnAccountName = "Deland"

        Case 250, 434, 435, 457
        
            ReturnAccountName = "Fish"

        Case 252, 424, 425, 452
        
            ReturnAccountName = "Flagler"

        Case 254, 400, 401, 440
        
            ReturnAccountName = "Heart"

        Case 256, 430, 431, 455
        
            ReturnAccountName = "FHM"

        Case 258, 412, 413, 446
        
            ReturnAccountName = "FHNP"

        Case 260, 428, 429, 454
        
            ReturnAccountName = "Tampa"

        Case 262, 420, 421, 450
        
            ReturnAccountName = "FHZ"

        Case 264, 426, 427, 453
        
            ReturnAccountName = "Metroplex"

        Case 266, 410, 411, 445
        
            ReturnAccountName = "SMPG"

        Case 268, 422, 423, 451
        
            ReturnAccountName = "Huguley"

        Case 269, 480, 481, 482
        
            ReturnAccountName = "HCP"

        End Select
    
End Function

Private Sub ShowMessage()

    Dim strStatus As String

    If strMessage = "" Then
        
        strStatus = "No upload differences found."
        
    Else
        
        strStatus = "Facilities with upload differences to review:"
    
    End If
    
    MsgBox "Load balances imported from " & strButton & " RT1 reports." & vbNewLine & _
           vbNewLine & _
           strStatus & strMessage, vbOKOnly
    
    strButton = vbNullString
    strMessage = vbNullString

End Sub
