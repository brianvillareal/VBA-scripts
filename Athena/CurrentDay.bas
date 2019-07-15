Attribute VB_Name = "CurrentDay"
Option Explicit

    Dim arrFacilities(19) As String, _
        arrTabs(19)       As String, _
        dteBalanceDate    As Date, _
        dblSum            As Double, _
        lngLastDate       As Long, _
        strAcronym        As String, _
        strArchive        As String, _
        strColumn         As String, _
        strCSV            As String, _
        strHeader         As String, _
        strFolderPD       As String, _
        wbkCurrent        As Workbook, _
        wksCurrent        As Worksheet, _
        wksArchive        As Worksheet

Private Sub InitCurrentDay(control As IRibbonControl)
        
    Call SetArrays
    Call SelectAthenaType(control.ID)
    Application.ScreenUpdating = False
    Call PrepArchive
    Application.ScreenUpdating = True
    MsgBox "Current/Prior File balances updated.", vbOKOnly
    
End Sub

Private Sub SetArrays()

    arrFacilities(0) = "prh"
    arrFacilities(1) = "memorial"
    arrFacilities(2) = "gordon"
    arrFacilities(3) = "deland"
    arrFacilities(4) = "tampa"
    arrFacilities(5) = "fhmg"
    arrFacilities(6) = "fhns"
    arrFacilities(7) = "fhmem"
    arrFacilities(8) = "heart"
    arrFacilities(9) = "flag"
    arrFacilities(10) = "fish"
    arrFacilities(11) = "fhnp"
    arrFacilities(12) = "fhz"
    arrFacilities(13) = "smpg"
    arrFacilities(14) = "hug"
    arrFacilities(15) = "lohp"
    arrFacilities(16) = "metro"
    arrFacilities(17) = "chip"
    arrFacilities(18) = "ahp"
    arrFacilities(19) = "hcp"
    arrTabs(0) = "Park Ridge"
    arrTabs(1) = "Memorial"
    arrTabs(2) = "Gordon"
    arrTabs(3) = "Deland"
    arrTabs(4) = "Tampa"
    arrTabs(5) = "FHMG"
    arrTabs(6) = "FHNS"
    arrTabs(7) = "FHM"
    arrTabs(8) = "Heart"
    arrTabs(9) = "Flagler"
    arrTabs(10) = "Fish"
    arrTabs(11) = "FHNP"
    arrTabs(12) = "FHZ"
    arrTabs(13) = "SMPG"
    arrTabs(14) = "Huguley"
    arrTabs(15) = "LOHP"
    arrTabs(16) = "Metroplex"
    arrTabs(17) = "Chippewa"
    arrTabs(18) = "AHP"
    arrTabs(19) = "HCP"

End Sub

Private Sub SelectAthenaType(controlID As String)

    Const strBalancing As String = "\\Lorem\ipsum$\dolor sit amet\IS\Athena Balancing Files\Athena Balancing ", _
          strRT999     As String = "\\Etiam\tempor\orci\eu\lobortis\", _
          strPD        As String = "Prior Day", _
          strOG        As String = "Original", _
          strUA        As String = "Unapplied", _
          strUI        As String = "Unidentified", _
          strUP        As String = "Unpostable"

    Select Case controlID
    
        Case strOG & "Balance"
            
            strAcronym = "apd"
            strArchive = strBalancing & strOG & " (version 1).xlsb.xlsx"
            strColumn = "G"
            strHeader = "Amount"
            strFolderPD = strRT999 & strPD & " File for Next Day"
                
        Case strUA & "Balance"
            
            strAcronym = "pdua"
            strArchive = strBalancing & strUA & ".xlsx"
            strColumn = "D"
            strFolderPD = strRT999 & strUA & " " & strPD
            strHeader = "unappliedamt"
                
        Case strUI & "Balance"
            
            strAcronym = "pdui"
            strArchive = strBalancing & strUI & ".xlsx"
            strColumn = "K"
            strFolderPD = strRT999 & strUI & " (Revenue) " & strPD
            strHeader = "Amount"
                
        Case strUP & "Balance"
            
            strAcronym = "pdup"
            strArchive = strBalancing & strUP & ".xlsx"
            strColumn = "K"
            strFolderPD = strRT999 & strUP & " " & strPD
            strHeader = "Amount"
            
    End Select

End Sub

Private Sub PrepArchive()

    Dim i              As Long, _
        strDate        As String, _
        wbkArchive     As Workbook

    Set wbkArchive = Application.Workbooks.Open(strArchive)
    DoEvents
    
    For i = 0 To 19
    
        Set wksArchive = wbkArchive.Worksheets("Daily Process-" & arrTabs(i))
        strCSV = Dir(strFolderPD & "\*" & arrFacilities(i) & "*")
    
        Do While Len(strCSV) > 0
    
            Call OpenFileWait
            Set wksCurrent = wbkCurrent.Worksheets(1)
            strDate = Mid(wbkCurrent.Name, Len(strAcronym) + 1, 6)
            dteBalanceDate = CDate(Left(strDate, 2) & "/" & Mid(strDate, 3, 2) & "/" & Right(strDate, 2))
            lngLastDate = wksArchive.Cells(Rows.Count, "C").End(xlUp).Row
            dblSum = SumAmountColumn
            Call UpdateArchive
            wbkCurrent.Close SaveChanges:=False
            strCSV = Dir

        Loop
    
    Next i
    
    Set wksCurrent = Nothing
    Set wbkCurrent = Nothing
    Set wksArchive = Nothing
    wbkArchive.Close SaveChanges:=True
    Set wbkArchive = Nothing

End Sub

Private Sub OpenFileWait()

    Set wbkCurrent = Application.Workbooks.Open(strFolderPD & "\" & strCSV)
        
    Do Until wbkCurrent.ReadOnly = False
        
        wbkCurrent.Close
        Application.Wait Now + TimeValue("00:00:01")
        Set wbkCurrent = Application.Workbooks.Open(strFolderPD & "\" & strCSV)
        
    Loop

End Sub

Private Function SumAmountColumn() As Double
    
    Dim lngLastAmount As Long
    
    '* Column A and D can't be used to set lngLastAmount due to metadata in those columns at the end of each Unapplied file.
    lngLastAmount = wksCurrent.Cells(Rows.Count, "B").End(xlUp).Row
    
    Select Case strHeader
    
        Case wksCurrent.Range(strColumn & "2")
        
            SumAmountColumn = Application.Sum(wksCurrent.Range(strColumn & "3", strColumn & lngLastAmount))
        
        Case wksCurrent.Range(strColumn & "1")
        
            SumAmountColumn = Application.Sum(wksCurrent.Range(strColumn & "2", strColumn & lngLastAmount))
    
    End Select

End Function

Private Sub UpdateArchive()

    If dteBalanceDate = wksArchive.Range("C" & lngLastDate) Then
            
        wksArchive.Range("D" & lngLastDate) = wksArchive.Range("D" & lngLastDate) + dblSum
            
    Else
                
        wksArchive.Range("C" & lngLastDate + 1) = dteBalanceDate
        wksArchive.Range("E" & lngLastDate + 1) = wksArchive.Range("D" & lngLastDate)
        wksArchive.Range("D" & lngLastDate + 1) = dblSum
            
    End If

End Sub
