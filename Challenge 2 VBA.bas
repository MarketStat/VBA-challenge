Attribute VB_Name = "Module1"
Sub MultipleYearStockData()

'Temporarily turn off Excel Screen Activity
Application.ScreenUpdating = False

    For Each ws In Worksheets
        
        'Each Worksheet
        Dim WsNm As String
        
        'Get the WorksheetName
        WsNm = ws.Name
        
        'Integer Variables - starting row is 2 - ticker integer in For loop
        Dim i As Long, j As Long, StockRow As Long
        
        StockRow = 2
        
        j = 2
        
        'Last Row Variables
        Dim lrtick As Long, LRSumTick As Long
        
        'Variables for percent change calculation
        Dim StockChange As Double, ValG As Double, ValD As Double, MostStock As Double
        
        'All Stock Output Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Max Values Headers and Labels
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Loop through each row/ticker
        lrtick = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ' Complete calculations on each ticker in the column
            For i = 2 To lrtick
            
                ' Determine if the Ticker equals the Ticker in the next cell
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ' Fill the Summed Ticker Row
                    ws.Cells(StockRow, 9).Value = ws.Cells(i, 1).Value
                
                    ' Yearly Change equals the Stock Value at Year End minus the Stock Value at Year Open
                    ws.Cells(StockRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                        ' Conditionally format the Yearly Change, if greater than 0 format as green, else red
                        If ws.Cells(StockRow, 10).Value < 0 Then
                
                            ' Red
                            ws.Cells(StockRow, 10).Interior.ColorIndex = 3
                
                        Else
                        
                            ' Green
                            ws.Cells(StockRow, 10).Interior.ColorIndex = 4
                
                        End If
                    
                    ' Change the value format of the stock change in column K to percent
                    ' Else fill 0 and format as percent
                    If ws.Cells(j, 3).Value <> 0 Then
                    
                        StockChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                        ws.Cells(StockRow, 11).Value = Format(StockChange, "Percent")
                    
                    Else
                    
                        ws.Cells(StockRow, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ' Put the total volume in Column L
                ws.Cells(StockRow, 12).Value = Application.WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                ' Go to next row
                StockRow = StockRow + 1
                
                ' Go to next row
                j = i + 1
                
                End If
            
            Next i
            
        ' Determine last used cell in Column I
        LRSumTick = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Set starting rows to search for each variable
        MostStock = ws.Range("L2").Value
        
        ValG = ws.Range("K2").Value
        
        ValD = ws.Range("K2").Value
        
            'Loop for summary
            For i = 2 To LRSumTick
            
                ' Compare Stock Values - loop through column L and assign the largest to a variable
                ' Else keep the value stored in MostStock
                If ws.Cells(i, 12).Value > MostStock Then
                    
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                    
                    MostStock = ws.Cells(i, 12).Value
                
                Else
                
                    MostStock = MostStock
                
                End If
                
                ' Compare Stock Value Increases - loop through Percent Change and assign the largest increase to a variable
                ' Else keep the value stored in ValG
                If ws.Cells(i, 11).Value > ValG Then
                
                    ValG = ws.Cells(i, 11).Value
                    
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                    ValG = ValG
                
                End If
                
                ' Compare Stock Value Increases - loop through Percent Change and assign the largest decrease to a variable
                ' Else keep the value stored in ValG
                If ws.Cells(i, 11).Value < ValD Then
                
                    ValD = ws.Cells(i, 11).Value
                    
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                    ValD = ValD
                
                End If
                
            ' Print values to their corresponding cells
            
            ' Print the Stock/Ticker with the most volume with Scientific format
            ws.Range("Q4").Value = Format(MostStock, "Scientific")
            
            ' Print the Stock/Ticker with the largest value increase with Percent format
            ws.Range("Q2").Value = Format(ValG, "Percent")
            
            ' Print the Stock/Ticker with the largest value decrease with Percent format
            ws.Range("Q3").Value = Format(ValD, "Percent")
            
            Next i
            
        ' Autofit columns on each worksheet
        Worksheets(WsNm).Columns("A:Z").AutoFit
            
    Next ws
    
' Turn Screen Updating back on
Application.ScreenUpdating = True
        
End Sub
