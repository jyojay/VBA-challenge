Sub stock_summary()

    'Setting variables required in project
        
    Dim TotalVol As Double
    Dim OpenVal As Double
    Dim CloseVal As Double
    Dim numofrows As Double
    Dim ResultRow As Double
    Dim Increase As Double
    Dim Decrease As Double
    Dim GreatTotVol As Double
    Dim RowInc As Double
    Dim RowDec As Double
    Dim RowTotVol As Double

    'Looping for all sheets in a workbook
    For Each ws In Worksheets
    
       'Initializing Variables
        
        TotalVol = 0
        ResultRow = 2
        OpenVal = ws.Cells(2, 3).Value
        CloseVal = ws.Cells(2, 6).Value
        
        'Populating headers for result table in the sheet
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Determining last row with data in the sheet
        
        numofrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop to calculate data for unique Ticker values from the available data rows
        
        For i = 2 To numofrows
           
            'When current and next ticker value in data is different
         
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
             
               'Saving Closing value in a variable and sum of Volume for the current ticker
                        
                CloseVal = ws.Cells(i, 6).Value
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                
               'Populate results in the required location
                
                ws.Cells(ResultRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(ResultRow, 10).Value = CloseVal - OpenVal 'Yearly Change
                ws.Cells(ResultRow, 10).NumberFormat = "0.00"  'Number formatting for Yearly change
                ws.Cells(ResultRow, 11).Value = (CloseVal - OpenVal) / OpenVal ' Percent Change
                ws.Cells(ResultRow, 11).NumberFormat = "0.00%" ' % formatting for Percent change
                ws.Cells(ResultRow, 12).Value = TotalVol
                
               'Conditional Colour formatting for Yearly change positive as green and negative as red
               'If no change then there will be no colour formatting
               
                If ws.Cells(ResultRow, 10).Value > 0 Then
                    ws.Cells(ResultRow, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(ResultRow, 10).Value < 0 Then
                    ws.Cells(ResultRow, 10).Interior.ColorIndex = 3
                End If
             
               'Reset total variable and increment result row by 1 and save opening value of next row of data
               
                TotalVol = 0
                ResultRow = ResultRow + 1
                OpenVal = ws.Cells(i + 1, 3).Value
                           
            Else
            ' When current and next ticker value is same
        
               ' Cumulate Total Volume for the ticker
                TotalVol = TotalVol + ws.Cells(i, 7).Value
         
            End If
            
        Next i
        
        '---------------------------------------------------------------------------------------
        ' Bonus: To determine Greatest % increase, Greatest % decrease and Greatest Total Volume
        '---------------------------------------------------------------------------------------
        
        'Initializing variables
        

        Increase = ws.Cells(2, 10).Value
        Decrease = ws.Cells(2, 10).Value
        GreatTotVol = ws.Cells(2, 12).Value
      

        'Looping through all rows of data to determine Greatest % increase, Greatest % decrease and Greatest Total Volume
        
        For i = 2 To ResultRow -1

            ' Finding Greatest % increase
            If ws.Cells(i, 10).Value > Increase Then
                Increase = ws.Cells(i, 10).Value
                RowInc = i ' saving corresponding row number
            End If

            ' Finding Greatest % Decrease
            If ws.Cells(i, 10).Value < Decrease Then
                Decrease = ws.Cells(i, 10).Value
                RowDec = i ' saving corresponding row number
            End If

            ' Finding Greatest % Total Volume
            If ws.Cells(i, 12).Value > GreatTotVol Then
            GreatTotVol = ws.Cells(i, 12).Value
            RowTotVol = i ' saving corresponding row number
            End If

        Next i

        
        'Populating result table with headers
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Populating result table with values
        
        ws.Range("P2").Value = ws.Cells(RowInc, 9).Value
        ws.Range("P3").Value = ws.Cells(RowDec, 9).Value
        ws.Range("P4").Value = ws.Cells(RowTotVol, 9).Value
        ws.Range("Q2").Value = Str(ws.Cells(RowInc, 10).Value) + "%"
        ws.Range("Q3").Value = Str(ws.Cells(RowDec, 10).Value) + "%"
        ws.Range("Q4").Value = ws.Cells(RowTotVol, 12).Value
        
        
        ' Autofitting to display data
        ws.Columns("A:Q").AutoFit
    
    Next ws

End Sub

