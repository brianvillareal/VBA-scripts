Attribute VB_Name = "15Characters"
Option Explicit

Private Sub InitTruncate(control As IRibbonControl)

    Dim rngCell As Range
    
    For Each rngCell In Selection
    
        If Len(rngCell) > 15 Then
            
            If IsNumeric(rngCell) Then
                
                rngCell.NumberFormat = "0"
                
                If Mid(rngCell, 16, 1) = "E" Then
                
                    rngCell = Left(rngCell, 15)
                
                Else
                
                    rngCell = Left(rngCell, 16)
                
                End If
                
                rngCell = rngCell * 100000000000000#
            
            Else

                rngCell = Left(rngCell, 15)
                rngCell.NumberFormat = "0"
                
            End If
            
        Else
        
            rngCell.NumberFormat = "0"
        
        End If
    
    Next rngCell

End Sub
