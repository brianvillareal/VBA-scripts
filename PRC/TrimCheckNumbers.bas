Attribute VB_Name = "TrimCheckNumbers"
Option Explicit

Private Sub InitTrim(control As IRibbonControl)

    Dim lngLastRow   As Long, _
        i            As Long, _
        j            As Long, _
        strCheck     As String, _
        strCharacter As String

    lngLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    
    For i = 1 To lngLastRow
            
        strCheck = Range("C" & i)
            
        For j = 1 To Len(strCheck)
                
            strCharacter = Mid(strCheck, j, 1)
                
            If IsNumeric(strCharacter) Then
                
                Select Case j
                
                    Case Is > 4
                    
                        If Mid(strCheck, j - 3, 3) <> "EFT" Then
                    
                            
                            Range("C" & i) = Mid(strCheck, j)
                    
                        Else
                    
                            Range("C" & i) = Mid(strCheck, j - 3)
                    
                        End If
                    
                    Case 4
                    
                        If Left(strCheck, 3) <> "EFT" Then
                    
                            Range("C" & i) = Mid(strCheck, 4)
                    
                        End If
                
                    Case Is > 1
                    
                        Range("C" & i) = Mid(strCheck, j)
                
                End Select
                
                Exit For

            End If
            
        Next j
        
    Next i
    
    Columns(3).NumberFormat = "0"

End Sub
