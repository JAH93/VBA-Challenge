Attribute VB_Name = "Module1"
Sub StockData():
    

        For Each ws In Worksheets
        
            Dim worksheetname As String
    
             Dim i As Long
            
             Dim j As Long
        
            Dim tickcount As Long

            Dim lastrowA As Long
        
            Dim lastrowI As Long
        
            Dim PerChange As Double
            
            
            worksheetname = ws.Name
            
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            tickcount = 2
            

            j = 2
            
        
            lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            
            
                For i = 2 To lastrowA
                
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
                    
                    ws.Cells(tickcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
            
                      If ws.Cells(tickcount, 10).Value < 0 Then
                    
        
                     ws.Cells(tickcount, 10).Interior.ColorIndex = 3
                    
                      Else
                    
                
                     ws.Cells(tickcount, 10).Interior.ColorIndex = 4
                    
                     End If
                        
        
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                        
            
                     ws.Cells(tickcount, 11).Value = Format(PerChange, "percent")
                        
                      Else
                        
                     ws.Cells(tickcount, 11).Value = Format(0, "percent")
                        
                     End If
                        
               ws.Cells(tickcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                    
            
              tickcount = tickcount + 1
                    
            
             j = i + 1
                    
            End If
                
            Next i
            
        lastrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
               
            
        For i = 2 To lastrowI
                        
                
        Next i
                

         Worksheets(worksheetname).Columns("A:Z").AutoFit
                
        Next ws
            
    End Sub



