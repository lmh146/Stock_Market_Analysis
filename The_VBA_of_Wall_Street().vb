Sub The_VBA_of_Wall_Street()

    'Declare Variables
        Dim Ticker As String
        Dim Yr_Chng As Double
        Dim Pct_Chng As Double
        Dim Ttl_Stock As Double
        Dim open_price As Double
        Dim EndRow As Long
        Dim EndRow2 As Long
        Dim I As Long
        Dim j As Long
        Dim WS As Worksheet
        
        j = 0
        
    'Assign Title Values
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yr_Chng"
        Range("K1").Value = "Pct_Chng"
        Range("L1").Value = "Ttl_Stock"
        
   'Assign ForLoop Needed Values
        EndRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        EndRow2 = Cells(Rows.Count, 11).End(xlUp).Row
        
        open_price = Cells(2, 3).Value
        
        Ws_Count = ActiveWorkbook.Worksheets.Count
        
        
    'Worksheet Loop
    For Each WS In Worksheets
        WS.Activate
        
    'Establish Loop
        For I = 2 To EndRow
        
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            
                If open_price = 0 Then
                    Pct_Chng = 0
                Else
                           
                'Ticker
                Range("I" & 2 + j).Value = Cells(I, 1).Value
                
                'close_price
                close_price = Cells(I, 6).Value
                
                
                'Yr_Chng Calculation
                Yr_Chng = (close_price - open_price)
                Range("J" & 2 + j).Value = Yr_Chng
                
                               
                'Pct_Chng Calculation
                Pct_Chng = Yr_Chng / open_price
                Range("K" & 2 + j).Value = Pct_Chng
                Range("K" & 2 + j).NumberFormat = "0.00%"
                
                
                'Ttl_Stock Calculation
                Ttl_Stock = Ttl_Stock + Cells(I, 7)
                Range("L" & 2 + j).Value = Ttl_Stock
                
                j = j + 1
              
                Ttl_Stock = 0
                
                
                'Following open prices
                    If Cells(I + 1, 3).Value = 0 Then
                        open_price = Cells(I + 2, 3).Value
                        
                            ElseIf Cells(I + 2, 3).Value = 0 Then
                                open_price = Cells(I + 3, 3).Value
                                
                                    Else
                                      open_price = Cells(I + 1, 3).Value
                        End If
                
                End If
                
            Else
                
                Ttl_Stock = Ttl_Stock + Cells(I, 7)
                
                     
            
            End If
            
                        
        Next I
        
        'Conditional Format for Yr_Change
        For I = 1 To EndRow2
            If Range("J" & 1 + I).Value < 0 Then
                Range("J" & 1 + I).Interior.Color = RGB(255, 0, 0)
                
            Else
                Range("J" & 1 + I).Interior.Color = RGB(0, 255, 0)
                
                
            End If
                    
        Next I
        
    
    Next WS

            
    
                
                
        
End Sub


