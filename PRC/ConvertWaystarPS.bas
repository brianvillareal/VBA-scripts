Attribute VB_Name = "ConvertWaystarPS"
Option Explicit

'Private Const strPRC        As String = "\\Lorem\ipsum$\dolor sit amet\Brian\Macros\PRC\"
Private Const strPRC        As String = "\\Lorem\ipsum$\dolor sit amet\Brian\Macros\In Progress\"
Private blnInput            As Boolean
Private adtePostRange()     As Date
Private dteStart            As Date, _
        dteEnd              As Date, _
        dtePost             As Date
Private strInput            As String, _
        strUnmapped         As String, _
        strFacility         As String
Private lngVisible          As Long, _
        lngLastDelete       As Long, _
        lngUnmappedColumn   As Long, _
        lngLastRow          As Long, _
        lngBuUnit           As Long, _
        lngOpUnit           As Long, _
        lngFirstGL          As Long, _
        lngSecondGL         As Long, _
        j                   As Long, _
        k                   As Long, _
        lngFirstRow         As Long, _
        lngLastWorked       As Long, _
        lngLastDRCR         As Long
Private wbkJE               As Workbook, _
        wbkSource           As Workbook, _
        wbkWaystar          As Workbook

Sub InitConvert(control As IRibbonControl)
    
    Call GetPostDate

    Application.ScreenUpdating = False
    
    Call CopySource
    
    '* Delete redundant columns.
    Union(wbkJE.Sheets(1).Columns("A:D"), _
            wbkJE.Sheets(1).Columns("F:H"), _
            wbkJE.Sheets(1).Columns("J"), _
            wbkJE.Sheets(1).Columns("L"), _
            wbkJE.Sheets(1).Columns("N:Q"), _
            wbkJE.Sheets(1).Columns("T:V")).EntireColumn.Delete

    Call DeleteBlankDates
    Call DeleteUserDates
    
    With wbkJE.Sheets(1)
    
        .Range("J1") = "PeopleSoft Bank Account"
        .Range("K1") = "PeopleSoft GL Account"
        .Range("L1") = "Description"
        .Range("M1") = "Flipped Payment Amount"
    
    End With
    
    Call DeleteRebatch
    Call DeleteMatchStatus
    Call DeletePaymentAccounts
    Call DeletePaymentAmounts
    Call DeleteBankOfAmerica
    Call DeleteUSBank
    Call DeleteWellsFargo

    wbkJE.Sheets(1).Range("A1:M" & lngLastRow).Sort Key1:=wbkJE.Sheets(1).Range("I1:I" & lngLastRow), _
                                                    Order1:=xlAscending, _
                                                    Header:=xlYes, _
                                                    Key2:=wbkJE.Sheets(1).Range("E1:I" & lngLastRow), _
                                                    Order1:=xlAscending, _
                                                    Header:=xlYes
    
    wbkJE.Sheets(1).Range("C2:C" & lngLastRow).NumberFormat = "0.00_);[Red](0.00)"
    Union(wbkJE.Sheets(1).Range("B2:B" & lngLastRow), wbkJE.Sheets(1).Range("G2:G" & lngLastRow)).NumberFormat = "0"
 
    Call MapPeopleSoft
    
    Application.DisplayAlerts = False
    wbkSource.Close savechanges:=False
    DoEvents
    wbkJE.Close savechanges:=False
    DoEvents
    
    With Application
    
        .DisplayAlerts = True
        .ScreenUpdating = True
        
    End With

End Sub

Private Sub GetPostDate()

    Dim lngResponse As Long

    lngResponse = MsgBox("Are you posting a range of dates?", vbYesNo, "Question")
        
        If lngResponse = vbYes Then
        
            Call GetStartDate

        Else

            Do Until blnInput = True
        
                strInput = Application.InputBox("Enter a date in MMDDYY format" & vbNewLine & _
                                                "(e.g. 020619 for February 6, 2019)", "What are you posting?")
    
                If strInput = "False" Then

                    End
        
                End If
    
                strInput = Left(strInput, 2) & "/" & Mid(strInput, 3, 2) & "/" & Right(strInput, 2)
    
                If IsDate(strInput) Then
                    
                    ReDim adtePostRange(0)
                    adtePostRange(0) = strInput
                    MsgBox "Match date: " & adtePostRange(0), vbInformation, "Confirmation"
                    blnInput = True
        
                Else
        
                    MsgBox "Invalid date: " & strInput, vbCritical, "Hmm.."

                End If
            
            Loop
        
        End If
    
    blnInput = False

End Sub

Private Sub GetStartDate()

    Do Until blnInput = True
        
    strInput = Application.InputBox("Enter the start date in MMDDYY format" & vbNewLine & _
                                    "(e.g. 020619 for February 6, 2019)", "What are you posting?")
    
        If strInput = "False" Then

            End
        
        End If
    
        strInput = Left(strInput, 2) & "/" & Mid(strInput, 3, 2) & "/" & Right(strInput, 2)
    
        If IsDate(strInput) Then
                
            dteStart = strInput
            MsgBox "Start date: " & dteStart, vbInformation, "Confirmation"
            Call GetEndDate
        
        Else
        
            MsgBox "Invalid date: " & strInput, vbCritical, "Hmm.."

        End If
            
    Loop

End Sub

Private Sub GetEndDate()

    '* Ask for end date after getting valid start date.
    Do Until blnInput = True
                
        strInput = Application.InputBox("Enter the end date in MMDDYY format" & vbNewLine & _
                                        "(e.g. 020619 for February 6, 2019)", "What are you posting?")
    
        If strInput = "False" Then
                            
            End
        
        End If
    
        strInput = Left(strInput, 2) & "/" & Mid(strInput, 3, 2) & "/" & Right(strInput, 2)
    
        If IsDate(strInput) Then

            dteEnd = strInput
                            
            If dteEnd < dteStart Then
                            
                MsgBox "The end date you entered (" & dteEnd & ") occurs before the start date (" & dteStart & ").", vbCritical, "Hmm.."
                Exit Do
                            
            ElseIf dteEnd = dteStart Then
                                
                MsgBox "The end date you entered (" & dteEnd & ") is the same as the start date (" & dteStart & ").", vbCritical, "Hmm.."
                Exit Do
                            
            Else
                                
                MsgBox "End date: " & dteEnd, vbInformation, "Confirmation"
                Call ParseDateRange
                blnInput = True
                            
            End If
        
        Else
        
            MsgBox "Invalid date: " & strInput, vbCritical, "Hmm.."

        End If
                    
    Loop

End Sub

Private Sub ParseDateRange()

    Dim m As Long, _
        n As Long
    
    k = DateDiff("d", dteStart, dteEnd)
    
    If k = 1 Then
    
        ReDim adtePostRange(k)
        adtePostRange(0) = dteStart
        adtePostRange(k) = dteEnd
    
    Else
        
        ReDim adtePostRange(k)
        adtePostRange(0) = dteStart
        adtePostRange(k) = dteEnd
        n = k - 1

        For m = 1 To n Step 1
    
            adtePostRange(m) = dteStart + m
        
        Next
        
    End If

End Sub

Private Sub CopySource()

    '* Create copy of source report.
    Set wbkJE = ActiveWorkbook
    Set wbkSource = Workbooks.Add
    DoEvents
    wbkJE.Sheets(1).Cells.Copy wbkSource.Sheets(1).Range("A1")
    Application.CutCopyMode = False

End Sub

Private Sub DeleteBlankDates()

    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=9, _
                                    Criteria1:=""

    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(9).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub FinishDelete()

    If lngVisible > 1 Then
    
        lngLastDelete = wbkJE.Sheets(1).Cells.Find(What:="*", _
                                                    After:=Range("A1"), _
                                                    LookAt:=xlPart, _
                                                    LookIn:=xlFormulas, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlPrevious, _
                                                    MatchCase:=False).Row
        
        wbkJE.Sheets(1).Range("A1:M" & lngLastDelete).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
    End If

End Sub

Private Sub DeleteUserDates()

    '* Delete rows with no match to user input.
    If UBound(adtePostRange) - LBound(adtePostRange) + 1 = 1 Then
        
        With wbkJE.Sheets(1)

            .AutoFilterMode = False
            .Range("A1:M1").AutoFilter Field:=9, _
                                        Criteria1:="<>" & adtePostRange(0)

        End With
        
        lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(9).SpecialCells(xlCellTypeVisible).Cells.Count

        Call FinishDelete
    
    Else
        
        Call DeleteUserRange
        
    End If

End Sub

Private Sub DeleteUserRange()

    Dim p As Long

    wbkJE.Sheets(1).AutoFilterMode = False
        
    For p = 0 To k Step 1
            
        With wbkJE.Sheets(1)
            
            .Cells(1, 10 + p) = " Matched Date"
            .Cells(2, 10 + p) = "<>" & adtePostRange(p)
                
            
        End With
        
    Next
        
    lngLastRow = wbkJE.Sheets(1).Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookAt:=xlPart, _
                                            LookIn:=xlFormulas, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row
        
    With wbkJE.Sheets(1)
        
        .Activate
        .Range("A1:I" & lngLastRow).AdvancedFilter Action:=xlFilterInPlace, _
                                                    CriteriaRange:=wbkJE.Sheets(1).Range(Cells(1, 10), Cells(2, 9 + p)), _
                                                    Unique:=False

    End With
        
    lngVisible = wbkJE.Sheets(1).Columns(9).SpecialCells(xlCellTypeVisible).Cells.Count

    Call FinishDelete
    
    With wbkJE.Sheets(1)
    
        If .FilterMode Then
    
            .ShowAllData
    
        End If
    
        .Range(Cells(1, 10), Cells(2, 9 + p)).Delete
    
    End With

End Sub

Private Sub DeleteRebatch()
    
    '* Delete rows with anything other than in Rebatch Indicator.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=6, _
                                    Criteria1:="<>Child"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(6).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeleteMatchStatus()

    '* Delete rows with Workable Unmatched in Match Status.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=4, _
                                    Criteria1:="Workable Unmatched"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeletePaymentAccounts()

    With wbkJE.Sheets(1)
    
        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=5, _
                                    Criteria1:="<>TIN 123456789*", _
                                    Criteria2:="<>TIN 123456789*"
                                                    
    End With
    
    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(5).SpecialCells(xlCellTypeVisible).Cells.Count
    
    If lngVisible > 1 Then
    
        MsgBox "The Waystar file includes at least one facility that hasn't been mapped to a PeopleSoft GL." & vbNewLine & _
                "Any such activity will be ignored.", vbInformation, "Hmm.."
    
    End If
    
    Call FinishDelete

    '* Delete rows with Lorem Claim Remits or Exception/Hospital Remits in Payment Account Name.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=5, _
                                    Criteria1:=Array("Exception/Hospital Remits", "Lorem Claim Remits"), _
                                    Operator:=xlFilterValues
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(5).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeletePaymentAmounts()
    
    '* Delete rows with $0.00 in Payment Amount.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=3, _
                                    Criteria1:="$0.00"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(3).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeleteBankOfAmerica()
    
    '* Delete rows with *Home* in Payment Account Name when Bank of America is the Bank Name.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=8, _
                                    Criteria1:="Bank of America"
        .Range("A1:M1").AutoFilter Field:=5, _
                                    Criteria1:="*Home*"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(8).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeleteUSBank()
    
    '* Delete rows with *Athena* in Payment Account Name when US Bank is the Bank Name.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=8, _
                                    Criteria1:="US Bank"
        .Range("A1:M1").AutoFilter Field:=5, _
                                    Criteria1:="*Athena*"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(8).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete

End Sub

Private Sub DeleteWellsFargo()
    
    '* Delete rows with *Cerner or *Series* in Payment Account Name when Wells Fargo is the Bank Name.
    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=8, _
                                    Criteria1:="Wells Fargo"
        .Range("A1:M1").AutoFilter Field:=5, _
                                    Criteria1:=Array("*Cerner", "*Series*"), _
                                    Operator:=xlFilterValues
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(8).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete
    
    wbkJE.Sheets(1).AutoFilterMode = False
    
    lngLastRow = wbkJE.Sheets(1).Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookAt:=xlPart, _
                                            LookIn:=xlFormulas, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row
                                            
    Call CheckRedundancy

End Sub

Private Sub CheckRedundancy()
    
    If wbkJE.Sheets(1).Range("A2") = "" Then

        If UBound(adtePostRange) - LBound(adtePostRange) + 1 = 1 Then

            MsgBox "There are no transactions from " & adtePostRange(0) & " to post.", vbInformation, "Empty"
        
        Else
        
            MsgBox "There are no transactions from " & adtePostRange(0) & " to " & adtePostRange(k) & " to post.", vbInformation, "Empty"
        
        End If
        
        Application.DisplayAlerts = False
        wbkJE.Close savechanges:=False
        DoEvents
        wbkSource.Close savechanges:=False
        DoEvents
        
        With Application
        
            .DisplayAlerts = True
            .ScreenUpdating = True
        
        End With
        
        End
    
    End If

End Sub

Private Sub MapPeopleSoft()
    
    j = 2
    lngFirstRow = 2

    Do While j <= lngLastRow
        
        Call MapBanks
        Call MapPayments

        wbkJE.Sheets(1).Range("J" & j) = lngFirstGL
        wbkJE.Sheets(1).Range("K" & j) = lngSecondGL
        wbkJE.Sheets(1).Range("M" & j) = -wbkJE.Sheets(1).Range("C" & j)
        
        '* Remove scientific notation from Payment Number when concatenating.
        wbkJE.Sheets(1).Range("L" & j) = Format(wbkJE.Sheets(1).Range("B" & j), "0") & " " & wbkJE.Sheets(1).Range("A" & j)
        
        '* Limit this field to 30 characters.
        If Len(wbkJE.Sheets(1).Range("L" & j)) > 30 Then
        
            wbkJE.Sheets(1).Range("L" & j) = Left(wbkJE.Sheets(1).Range("L" & j), 30)
        
        End If
        
        '* Create a separate Waystar JE file for each Facility and Journal Date.
        If wbkJE.Sheets(1).Range("I" & j + 1) = "" Then
            
            lngLastWorked = j
            dtePost = wbkJE.Sheets(1).Range("I" & j)
            Call MapFacility
            Call CreateWaystarJE
            lngFirstRow = j + 1
            
        ElseIf wbkJE.Sheets(1).Range("I" & j) <> wbkJE.Sheets(1).Range("I" & j + 1) Or _
                Split(wbkJE.Sheets(1).Range("E" & j), "-")(0) <> Split(wbkJE.Sheets(1).Range("E" & j + 1), "-")(0) Then
                
            lngLastWorked = j
            dtePost = wbkJE.Sheets(1).Range("I" & j)
            Call MapFacility
            Call CreateWaystarJE
            lngFirstRow = j + 1
        
        End If
        
        j = j + 1
    
    Loop

End Sub

Private Sub MapBanks()

    Dim strBank As String

    '* Identify PeopleSoft GL account number for DR/CR.
    strBank = wbkJE.Sheets(1).Range("H" & j)

    Select Case strBank
        
        Case "Wells Fargo"
            
            lngFirstGL = 123456
            
        Case "Bank of America"
            
            lngFirstGL = 234567
            
        Case Else
            
            lngUnmappedColumn = 8
            strUnmapped = strBank
            MsgBox Chr(34) & strUnmapped & Chr(34) & " hasn't been mapped to a PeopleSoft GL." & vbNewLine & _
                                                        "Activity for this Bank Name will be ignored.", vbInformation, "Hmm.."
            Call DeleteUnmapped
            
            If j > 2 Then
            
                j = j - 1
            
            End If
            
            Call MapBanks
            
    End Select

End Sub

Private Sub DeleteUnmapped()

    With wbkJE.Sheets(1)

        .AutoFilterMode = False
        .Range("A1:M1").AutoFilter Field:=lngUnmappedColumn, _
                                    Criteria1:=strUnmapped & "*"
    
    End With

    lngVisible = wbkJE.Sheets(1).AutoFilter.Range.Columns(lngUnmappedColumn).SpecialCells(xlCellTypeVisible).Cells.Count
    
    Call FinishDelete
    
    wbkJE.Sheets(1).AutoFilterMode = False
    
    lngLastRow = wbkJE.Sheets(1).Cells.Find(What:="*", _
                                            After:=Range("A1"), _
                                            LookAt:=xlPart, _
                                            LookIn:=xlFormulas, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row
                                            
    Call CheckRedundancy

End Sub

Private Sub MapPayments()

    Dim strPayment As String
    
    strPayment = Split(wbkJE.Sheets(1).Range("E" & j), " ")(2)
        
    Select Case strPayment
        
        Case "Cerner", "Series"
            
            lngSecondGL = 987654
                
        Case "Athena/NexGen", "Athena/NextGen"
            
            lngSecondGL = 876543
                
        Case "Home Health", "HomeCare"
            
            lngSecondGL = 765432
                
        Case "Exceptions", "PLB"
            
            lngSecondGL = 654321
                
        Case Else
            
            lngUnmappedColumn = 5
            strUnmapped = wbkJE.Sheets(1).Range("E" & j)
            MsgBox Chr(34) & strUnmapped & Chr(34) & " hasn't been mapped to a PeopleSoft GL." & vbNewLine & _
                    "Activity for this Payment Account Name will be ignored.", vbInformation, "Hmm.."
            Call DeleteUnmapped
            
            If j > 2 Then
            
                j = j - 1
            
            End If
            
            Call MapBanks
            Call MapPayments
    
    End Select

End Sub

Private Sub MapFacility()

    Dim strTIN As String

    '* Identify account name.
    strTIN = Split(wbkJE.Sheets(1).Range("E" & j), "-")(0)
    
    Select Case strTIN
        
        Case "TIN 123456789"
        
            strFacility = "AdventHealth Sebring"
            lngBuUnit = 12345
            lngOpUnit = 1234
            
    End Select

End Sub

Private Sub CreateWaystarJE()

    Const strTemplate As String = "Waystar JE Template.xlsx"
    
    Application.Workbooks.Open Filename:=strPRC & strTemplate, _
                                UpdateLinks:=False
    DoEvents
    Set wbkWaystar = ActiveWorkbook
    
    '* Copy/paste source report.
    wbkSource.Sheets(1).Cells.Copy wbkWaystar.Sheets(3).Range("A1")
    Application.CutCopyMode = False
    
    Call PopulateWaystarJE
    Call SaveWaystarJE

End Sub

Private Sub PopulateWaystarJE()

    With wbkWaystar.Sheets(2)
    
        .Range("E1") = strFacility
        .Range("E2") = lngBuUnit
        .Range("M1") = dtePost
        
    End With

    '* Amounts & Description.
    Union(wbkJE.Sheets(1).Range("C" & lngFirstRow & ":C" & lngLastWorked), _
            wbkJE.Sheets(1).Range("L" & lngFirstRow & ":L" & lngLastWorked), _
            wbkJE.Sheets(1).Range("M" & lngFirstRow & ":M" & lngLastWorked)).Copy
            
    wbkWaystar.Sheets(2).Range("K6").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    lngLastDRCR = wbkWaystar.Sheets(2).Cells.Find(What:="*", _
                                                    After:=Range("A1"), _
                                                    LookAt:=xlPart, _
                                                    LookIn:=xlFormulas, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlPrevious, _
                                                    MatchCase:=False).Row

    wbkWaystar.Sheets(2).Range("L6:L" & lngLastDRCR).Copy wbkWaystar.Sheets(2).Range("L" & lngLastDRCR + 1)
    Application.CutCopyMode = False

    wbkWaystar.Sheets(2).Range("M6:M" & lngLastDRCR).Cut wbkWaystar.Sheets(2).Range("K" & lngLastDRCR + 1)
    Application.CutCopyMode = False
    
    wbkJE.Sheets(1).Range("J" & lngFirstRow & ":K" & lngLastWorked).Copy
    wbkWaystar.Sheets(2).Range("C6").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    wbkWaystar.Sheets(2).Range("D6:D" & lngLastDRCR).Cut wbkWaystar.Sheets(2).Range("C" & lngLastDRCR + 1)
    
    Application.CutCopyMode = False
    
    lngLastDRCR = wbkWaystar.Sheets(2).Cells.Find(What:="*", _
                                                    After:=Range("A1"), _
                                                    LookAt:=xlPart, _
                                                    LookIn:=xlFormulas, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlPrevious, _
                                                    MatchCase:=False).Row
                        
    Union(wbkWaystar.Sheets(2).Range("A6:A" & lngLastDRCR), wbkWaystar.Sheets(2).Range("P6:P" & lngLastDRCR)) = lngBuUnit

    wbkWaystar.Sheets(2).Range("B6:B" & lngLastDRCR) = "ACTUAL"
    wbkWaystar.Sheets(2).Range("F6:F" & lngLastDRCR) = lngOpUnit
    wbkWaystar.Sheets(2).Range("J6:J" & lngLastDRCR) = "USD"
    
    wbkWaystar.Sheets(2).Range("K6:K" & lngLastDRCR).NumberFormat = "0.00_);[Red](0.00)"
    wbkWaystar.Sheets(2).Range("D1:R" & lngLastDRCR).Columns.AutoFit
    wbkWaystar.Sheets(2).Activate
    wbkWaystar.Sheets(2).Range("A1").Activate
    
    wbkWaystar.Sheets(1).Range("B7") = dtePost

End Sub

Private Sub SaveWaystarJE()
    
    Const strOutput As String = "Test\"
    Dim strWaystar  As String, _
        strSave     As String

    '* Save out, close, move on.
    Application.DisplayAlerts = False
    strWaystar = "Waystar " & Split(strFacility, " ")(1) & " " & Format(dtePost, "yyyy-mm-dd")
    strSave = strPRC & strOutput & strWaystar
    wbkWaystar.SaveAs Filename:=strSave
    DoEvents
    wbkWaystar.Close
    DoEvents
    Application.DisplayAlerts = True

End Sub

