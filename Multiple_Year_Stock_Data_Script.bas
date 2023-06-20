Attribute VB_Name = "Module1"
Sub stockdata()

    'Create a for Loop for 'ws' to return values on all sheets of the excel file
    For Each ws In Worksheets

        'Declare all variables that are required for the entire script
        Dim Ticker As String
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Dim Summary_Table_Row As Integer
        Dim Result_1 As Double
        Dim Result_2 As Double
        Dim Result_3 As Double
            
        'Assign variables to a value
        Open_Price = Range("C2").Value
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
    
        'Create new columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Create Table for Calculated Values
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
                
        'Determine Last Row of the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        'Create a For Loop to determine each of the values (total stock volume, yearly change, percent change)
        For i = 2 To LastRow
                    
            'Set an If argument for when the value of the stock name changes in row 1
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                    
                'Add to the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                        
                'Print the Ticker in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                        
                'Print the Total Stock Volume in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
                'Calculate the Yearly Change
                Close_Price = ws.Cells(i, 6).Value
                Yearly_Change = (Close_Price - Open_Price)
            
                'Input the Yearly Change velues onto Column J
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                    'Create a Nested If argument to Calculate the Percent Change. This will return the Percent Change within the For Loop
                    If (Open_Price) = 0 Then
                
                        Percent_Change = 0
                        
                    Else
                    
                        Percent_Change = Yearly_Change / Open_Price
    
                    End If
                
                'Input the Percent Change values onto the Excel Sheet in column K
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Add 1 to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Total Stock Volume
                Total_Stock_Volume = 0
                
                'Reset Open Price
                Open_Price = ws.Cells(i + 1, 3)
                    
            Else
                
                'Add to the Total Stock Volume. The Else argument will add up all values for the same stock name
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'Create a For Loop to set colour code of each cell in the Yearly Change Column
        For j = 2 To LastRow2
        
            'Determine which cells will be filled with green and red by writing an If/Else argument
            If ws.Cells(j, 10).Value > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 10
            
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
                
            End If
            
        Next j
        
        'Assign Variables to Integers for the Calculated Values
        Result_1 = 0
        Result_2 = 0
        Result_3 = 0
        
        'Create a For Loop to return the values for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        For k = 2 To LastRow
            
            If ws.Cells(k, 11).Value > Result_1 Then
            
                Result_1 = ws.Cells(k, 11).Value
                ws.Range("P2").Value = Result_1
                ws.Range("P2").NumberFormat = "0.00%"
                ws.Range("O2").Value = ws.Cells(k, 9).Value
                
            End If
    
        Next k
        
        For l = 2 To LastRow
        
            If ws.Cells(l, 11).Value < Result_2 Then
            
                Result_2 = ws.Cells(l, 11).Value
                ws.Range("P3").Value = Result_2
                ws.Range("P3").NumberFormat = "0.00%"
                ws.Range("O3").Value = ws.Cells(l, 9).Value
            
            End If
        
        Next l
        
        For m = 2 To LastRow
        
            If ws.Cells(m, 12).Value > Result_3 Then
            
                Result_3 = ws.Cells(m, 12).Value
                ws.Range("P4").Value = Result_3
                ws.Range("O4").Value = ws.Cells(m, 9).Value
                
            End If
            
        Next m
        
    Next ws
    
End Sub
