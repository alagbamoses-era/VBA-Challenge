# VBA-Challenge


Sub MultipleStock()

    Dim TotalVolume As Double
    Dim RowCount As Long
    Dim ws As Worksheet
    Dim CurrentTicker As String
    Dim NextTicker As String
    Dim OutputPosition As Long
    Dim OpenStock As Double
    Dim CloseStock As Double
    Dim TF_FirstRecord As Boolean
    Dim GreatTicker As String
    Dim GreaterTicker As String
    Dim TotalStock As Long
    
    
    
    
    
    
    For Each ws In Worksheets
        'setting the titles
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Value"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
       
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        TotalVolume = 0
        OutputPosition = 2
        OpenStock = 0
        CloseStock = 0
        TF_FirstRecord = True
        
        
        For i = 2 To RowCount
            CurrentTicker = ws.Cells(i, 1).Value
            NextTicker = ws.Cells(i + 1, 1).Value
            
            If TF_FirstRecord = True Then
                OpenStock = ws.Cells(i, 3)
                TF_FirstRecord = False
            End If
                        
            
            If CurrentTicker <> NextTicker Then
                'since it is no longer the same, write out the Total
                
                TotalVolume = TotalVolume + ws.Cells(i, 7) 'calculate last total volume
                CloseStock = ws.Cells(i, 6).Value
                
                'Output session
                               
                ws.Cells(OutputPosition, 9).Value = CurrentTicker
                ws.Cells(OutputPosition, 10).Value = CloseStock - OpenStock
                If OpenStock = 0 Then
                    ws.Cells(OutputPosition, 11).Value = 0
                ElseIf OpenStock > 0 Then
                    ws.Cells(OutputPosition, 11).Value = (CloseStock - OpenStock) / OpenStock * 100
                End If
                
                ws.Cells(OutputPosition, 12).Value = TotalVolume
                
                TotalVolume = 0
                OutputPosition = OutputPosition + 1
                TF_FirstRecord = True
                
                
                
                
                
                
            Else
            
                TotalVolume = TotalVolume + ws.Cells(i, 7)
                
            End If
            
        
        Next i
        
        'create a new rowcount (on I-column), from row 2,
        RowCount = Cells(Rows.Count, "I").End(xlUp).Row
        
        
        ' initialise Totalstock = 0
        'TotalStock = 0
        'initialise ticker = 0
        'Ticker = 0
        'retrieve the 1st stock, replace( totalstock = cellcurrent)
        
       ' ws.Cells(OutputPosition, 16).Value = TotalStock
        
        'create a new loop from 2 to end of row count
        
       ' For j = 2 To RowCount
            GreatTicker = ws.Cells(i, 9).Value
            GreaterTicker = ws.Cells(i + 1, 9).Value
        
        
       '      If GreatTicker <> GreaterTicker Then
                'since it is no longer the same, write out the Total
                
               
       ' Next j
       
       
       
        
        
        
        
        
    Next ws
    
    




End Sub



'Array


Public Sub Greatest()


    Dim mycell As Range
    Dim myrange As Range
    Dim Counter As Long
    Dim ws As Worksheet
    
    
    For Each ws In Worksheets
    
            Counter = 1
            
            Set myrange = ws.Range("K2:K3019")
            Greatest_Increase = WorksheetFunction.Max(myrange)
            Greatest_Decrease = WorksheetFunction.Min(myrange)
            
            
            
            For Each mycell In myrange
                Counter = Counter + 1
                If mycell.Value = Greatest_Increase Then
                    ws.Range("Q2") = mycell.Value
                    ws.Range("P2") = Cells(Counter, 9)
                    
                    
                    
                End If
                
                If mycell.Value = Greatest_Decrease Then
                    ws.Range("Q3") = mycell.Value
                    ws.Range("P3") = Cells(Counter, 9)
                    
                End If
                
            
            Next mycell
            
            Set myrangeVol = ws.Range("L2:L3019")
            Greatest_Increase = WorksheetFunction.Max(myrangeVol)
            
            Counter = 1
            For Each mycell In myrangeVol
                Counter = Counter + 1
                If mycell.Value = Greatest_Increase Then
                    ws.Range("Q4") = mycell.Value
                    ws.Range("P4") = Cells(Counter, 9)
                    
                End If
            Next mycell
    Next ws
    

End Sub

Sub Final()

    
    Call MultipleStock
    Call Greatest

End Sub


