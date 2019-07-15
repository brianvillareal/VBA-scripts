Attribute VB_Name = "FilesFTP"
Option Explicit
    
    Dim arrCentral(2)           As String, _
        arrEastern(4)           As String, _
        arrFlorida()            As String, _
        blnCurrentQuarter       As Boolean, _
        blnPriorQuarter         As Boolean, _
        blnFile                 As Boolean, _
        lngCurrentQuarterStart  As Long, _
        lngCurrentQuarterEnd    As Long, _
        lngPriorQuarterStart    As Long, _
        lngPriorQuarterEnd      As Long, _
        strAcronym              As String, _
        strDate                 As String, _
        strFile                 As String, _
        strFileCategory         As String, _
        strFolder               As String, _
        strHolding              As String, _
        strIssueMessage         As String, _
        strMissingMessage       As String, _
        strRegion               As String, _
        strSuffix               As String, _
        varFacility             As Variant, _
        wbkAthena               As Workbook

Private Sub InitFilesFTP(control As IRibbonControl)

    Call UnzipFiles
    Application.ScreenUpdating = False

    Select Case control.ID
    
        Case "buttonOriginal"
        
            strFileCategory = "Original"
            strHolding = "\"
            strAcronym = "acd "
            Call AssignArraysACD
            Call PrepareOther
                
        Case "buttonUnapplied"
        
            strFileCategory = "Unapplied"
            strHolding = " Unapplied\"
            strAcronym = "cdua "
            strSuffix = ""
            Call AssignArraysCDUA
            Call PrepareCDUA
                
        Case "buttonUnidentified"
        
            strFileCategory = "Unidentified"
            strHolding = " Unidentified\"
            strAcronym = "cdui "
            Call AssignArraysCDUI
            Call PrepareOther
                
        Case "buttonUnpostable"
        
            strFileCategory = "Unpostable"
            strHolding = " Unpostable\"
            strAcronym = "cdup "
            Call AssignArraysCDUP
            Call PrepareOther
            
    End Select
    
    Application.ScreenUpdating = True
    Call DisplayMessage
    
End Sub

Private Sub UnzipFiles()

    Dim objApp  As Object, _
        strZip  As String
    
    strFolder = "\\Etiam\tempor\orci\eu\lobortis\FTP Holding\"
    strZip = Dir(strFolder & "documents*")
    strZip = strFolder & strZip
    strFolder = strFolder & "Files\"
    Set objApp = CreateObject("Shell.Application")
    objApp.Namespace((strFolder)).CopyHere objApp.Namespace((strZip)).Items
    DoEvents
    Set objApp = Nothing
    Kill strZip
    DoEvents

End Sub

Private Sub AssignArraysACD()

    arrCentral(0) = "Gordon"
    arrCentral(1) = "Memorial"
    arrCentral(2) = "PRMA"
    
    arrEastern(0) = "Chippewa"
    arrEastern(1) = "Huguley"
    arrEastern(2) = "LOHP"
    arrEastern(3) = "Metroplex"
    arrEastern(4) = "SMPG"
    
    ReDim arrFlorida(9)
    arrFlorida(0) = "Deland"
    arrFlorida(1) = "FHMG"
    arrFlorida(2) = "FHNP"
    arrFlorida(3) = "FHNS"
    arrFlorida(4) = "FHZ"
    arrFlorida(5) = "FH Memorial"
    arrFlorida(6) = "Fish"
    arrFlorida(7) = "Flagler"
    arrFlorida(8) = "heartland"
    arrFlorida(9) = "Tampa"

End Sub

Private Sub PrepareOther()

    '* AHP and HCP are always missing from these categories.
    strMissingMessage = vbNewLine & "AHP (Amita) Prior Quarter" & _
                vbNewLine & "AHP (Amita) Current Quarter" & _
                vbNewLine & "HCP (Florida)"

    Call SetQuarterDates
    
    strRegion = "Central"
    
    For Each varFacility In arrCentral
    
        Call ProcessQuarterFiles
    
    Next varFacility
    
    strRegion = "Eastern"
    
    For Each varFacility In arrEastern
    
        Call ProcessQuarterFiles
    
    Next varFacility
    
    strRegion = "Florida"
    
    For Each varFacility In arrFlorida
    
        Call ProcessSingleFile
    
    Next varFacility

End Sub

Private Sub AssignArraysCDUA()

    arrCentral(0) = "Gordon"
    arrCentral(1) = "Manchester_Memorial_PSMH"
    arrCentral(2) = "Park_Ridge"
    
    arrEastern(0) = "CVH_Durand"
    arrEastern(1) = "HMA_Huguley"
    arrEastern(2) = "LOHP_Live_OAK"
    arrEastern(3) = "MCP_Metroplex"
    arrEastern(4) = "Shawnee"
    
    ReDim arrFlorida(11)
    arrFlorida(0) = "FHHMC"
    arrFlorida(1) = "FHMG1"
    arrFlorida(2) = "FHMG2"
    arrFlorida(3) = "FHNP"
    arrFlorida(4) = "FHPG"
    arrFlorida(5) = "FHTHM"
    arrFlorida(6) = "FHZ"
    arrFlorida(7) = "HCPD"
    arrFlorida(8) = "HCPF"
    arrFlorida(9) = "HCPM"
    arrFlorida(10) = "HCPNS"
    arrFlorida(11) = "HFPFM"

End Sub

Private Sub PrepareCDUA()

    '* AHP is unique to Amita.
    varFacility = "AHP"
    strRegion = "Amita"
    Call ProcessSingleFile
    
    strRegion = "Central"

    For Each varFacility In arrCentral
        
        Call ProcessSingleFile
        
    Next varFacility

    strRegion = "Eastern"

    For Each varFacility In arrEastern
        
        Call ProcessSingleFile
    
    Next varFacility

    strRegion = "Florida"

    For Each varFacility In arrFlorida
            
        Call ProcessSingleFile
            
    Next varFacility

End Sub

Private Sub AssignArraysCDUI()

    arrCentral(0) = "Gordon"
    arrCentral(1) = "Manchester Memorial"
    arrCentral(2) = "Park Ridge"
    
    arrEastern(0) = "Chippewa"
    arrEastern(1) = "Huguley"
    arrEastern(2) = "LOHP"
    arrEastern(3) = "Metroplex"
    arrEastern(4) = "Shawnee"
    
    ReDim arrFlorida(9)
    arrFlorida(0) = "Deland"
    arrFlorida(1) = "FHMG"
    arrFlorida(2) = "FHNP"
    arrFlorida(3) = "FHNS"
    arrFlorida(4) = "FHZ"
    arrFlorida(5) = "Fish"
    arrFlorida(6) = "Flagler"
    arrFlorida(7) = "Heartland"
    arrFlorida(8) = "Memorial"
    arrFlorida(9) = "Tampa"

End Sub

Private Sub AssignArraysCDUP()

    arrCentral(0) = "Gordon"
    arrCentral(1) = "Manchester Memorial"
    arrCentral(2) = "Park Ridge"
    
    arrEastern(0) = "Chippewa"
    arrEastern(1) = "Huguley"
    arrEastern(2) = "LOHP"
    arrEastern(3) = "Metroplex"
    arrEastern(4) = "SMPG"
    
    '* FHMG needs to be processed before FHM or there will be an overwrite issue.
    ReDim arrFlorida(9)
    arrFlorida(0) = "Deland"
    arrFlorida(1) = "FHMG"
    arrFlorida(2) = "FHM"
    arrFlorida(3) = "FHNP"
    arrFlorida(4) = "FHNS"
    arrFlorida(5) = "FHZ"
    arrFlorida(6) = "Fish"
    arrFlorida(7) = "Flagler"
    arrFlorida(8) = "Heartland"
    arrFlorida(9) = "Tampa"

End Sub

Private Sub SetQuarterDates()

    lngPriorQuarterStart = DateAdd("q", DatePart("q", Date) - 2, DateSerial(Year(Date), 1, 1))
    lngCurrentQuarterStart = DateAdd("q", 1, lngPriorQuarterStart)
    lngPriorQuarterEnd = DateAdd("d", -1, lngCurrentQuarterStart)
    lngCurrentQuarterEnd = DateAdd("q", 1, lngCurrentQuarterStart) - 1

End Sub

Private Sub ProcessQuarterFiles()
    
    strFile = Dir(strFolder & "\*" & varFacility & "*")

    Do While Len(strFile) > 0

        Call OpenFileWait
        Call FilterQuarters
        strFile = Dir
        
    Loop
    
    Call EditQuarterMessage(varFacility, strRegion)

End Sub

Private Sub ProcessSingleFile()

    strFile = Dir(strFolder & "\*" & varFacility & "*")

    Do While Len(strFile) > 0
        
        Call OpenFileWait
        strSuffix = ""
        blnFile = True
        Call SaveRT999Athena
        Kill strFolder & "\" & strFile
        DoEvents
        strFile = Dir
        
    Loop
    
    Call EditOtherMessage(varFacility, strRegion)

End Sub

Private Sub OpenFileWait()

    Set wbkAthena = Application.Workbooks.Open(strFolder & "\" & strFile)
        
    Do Until wbkAthena.ReadOnly = False
        
        wbkAthena.Close
        Application.Wait Now + TimeValue("00:00:01")
        Set wbkAthena = Application.Workbooks.Open(strFolder & "\" & strFile)
        
    Loop

End Sub

Private Sub FilterQuarters()

    wbkAthena.Sheets(1).Range("A1").AutoFilter Field:=3, _
                                        Criteria1:=">=" & lngPriorQuarterStart, _
                                        Operator:=xlAnd, _
                                        Criteria2:="<=" & lngPriorQuarterEnd

    If wbkAthena.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
    
        Call CheckForMultiples(blnPriorQuarter)
        strSuffix = " A"
        Call SaveRT999Athena
        blnPriorQuarter = True
        Kill strFolder & "\" & strFile
        DoEvents
            
    Else
        
        wbkAthena.Sheets(1).Range("A1").AutoFilter Field:=3, _
                                            Criteria1:=">=" & lngCurrentQuarterStart, _
                                            Operator:=xlAnd, _
                                            Criteria2:="<=" & lngCurrentQuarterEnd
                                                
        If wbkAthena.Sheets(1).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            
            Call CheckForMultiples
            strSuffix = " B"
            Call SaveRT999Athena
            blnCurrentQuarter = True
            Kill strFolder & "\" & strFile
            DoEvents
            
        Else
            
            wbkAthena.Close SaveChanges:=False
            DoEvents
            Set wbkAthena = Nothing
            strIssueMessage = strIssueMessage & vbNewLine & "Incorrect data in " & varFacility & " (" & strRegion & ") file"
            Kill strFolder & "\" & strFile
            DoEvents

        End If

    End If

End Sub

Private Sub CheckForMultiples(Optional blnPriorQuarter)

    If IsMissing(blnPriorQuarter) = True Then
    
        If blnCurrentQuarter = True Then

            strIssueMessage = strIssueMessage & vbNewLine & "Multiple " & varFacility & " (" & strRegion & ") Current Quarter files"
        
        End If
        
    Else
    
        If blnPriorQuarter = True Then
    
            strIssueMessage = strIssueMessage & vbNewLine & "Multiple " & varFacility & " (" & strRegion & ") Prior Quarter files"
        
        End If

    End If

End Sub

Private Sub SaveRT999Athena()

    Const strRT999  As String = "\\Etiam\tempor\orci\eu\lobortis\Athena Holding"

    Dim strSave     As String

    Call PullFileDate(varFacility)
    strSave = strRT999 & strHolding & strAcronym & strDate & varFacility & strSuffix
    
    Application.DisplayAlerts = False
    wbkAthena.SaveAs Filename:=strSave, _
                    FileFormat:=xlCSV
    DoEvents
    wbkAthena.Close SaveChanges:=False
    DoEvents
    Set wbkAthena = Nothing
    Application.DisplayAlerts = True

End Sub

Private Sub PullFileDate(varFacility)

    Select Case strFileCategory
    
        Case "Original", "Unidentified", "Unpostable"
        
            strDate = Split(wbkAthena.Name, "_")(4) & " "
                
        Case "Unapplied"
        
            strDate = Split(wbkAthena.Name, varFacility)(1) & " "
            strDate = Split(strDate, ".")(0) & " "
            
    End Select

End Sub

Private Sub EditQuarterMessage(varFacility, strRegion)

    If blnPriorQuarter = False Then
            
        strMissingMessage = strMissingMessage & vbNewLine & varFacility & " (" & strRegion & ") Prior Quarter"
        
    Else
        
        blnPriorQuarter = False
        
    End If

    If blnCurrentQuarter = False Then
    
        strMissingMessage = strMissingMessage & vbNewLine & " (" & strRegion & ") Current Quarter"
        
    Else
    
        blnCurrentQuarter = False
    
    End If
    
End Sub

Private Sub EditOtherMessage(varFacility, strRegion)

    If blnFile = False Then
            
        strMissingMessage = strMissingMessage & vbNewLine & varFacility & " (" & strRegion & ")"
        
    Else
        
        blnFile = False
        
    End If

End Sub

Private Sub DisplayMessage()

    If strMissingMessage = "" Then
    
        strMissingMessage = vbNewLine & "None"
    
    End If
    
    If strIssueMessage = "" Then
    
        strIssueMessage = vbNewLine & "None"
    
    End If

    MsgBox "The renaming operation is complete." & _
        vbNewLine & _
        vbNewLine & "Missing:" & _
        strMissingMessage & _
        vbNewLine & _
        vbNewLine & "Issues:" & _
        strIssueMessage, _
        vbOKOnly
        
    strMissingMessage = ""
    strIssueMessage = ""

End Sub
