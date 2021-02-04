Attribute VB_Name = "Module1"
Sub Stock_Market():
    'Go through Each Worsheet
    For Each ws In Worksheets
        'Assign Variables to each of the Variables and assign Starting values
        Dim Ticker As String
        Dim Yearly_Change As Double
            Yearly_Change = 0
        Dim Percent_Change As Double
            Percent_Change = 0
        Dim Total_Stock_Volume As Double
            Total_Stock_Volume = 0
        Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
        Dim LastRow As Long
        Dim Sopen As Double
        Dim SClose As Double
    
    
        ws.Columns("A:Q").AutoFit
        
        'Assign Ranges For Summary Table Row
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'Determine where the Last Row Is
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Start loop 1
        For i = 2 To LastRow
        
            'Calculate Total_Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            
            'Validate that we are remianing within the same stock name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then  ' true
                
                'Create Definitions for calculations of Summary Data and Bonus
                Ticker = ws.Cells(i, 1).Value
                Sopen = ws.Cells(i, 3)
                SClose = ws.Cells(i, 6)
                Yearly_Change = SClose - Sopen
                
                'End Loop 1
                
                'Set Up the Summary_Table_Row
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                
                Sopen = ws.Cells(i + 1, 3)
                'Start Loop 2
                If Sopen = 0 Then
                    Percent_Change = 0
                    Else
                    Percent_Change = (Yearly_Change / Sopen)
                
                End If
                'End Loop 2
                
                
                'Start Loop 3
                'Develop Conditional Formatting
                If Range("J" & Summary_Table_Row).Value >= 0 Then
                   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        'End Loop 3
                    End If
                    'Declare Bonuses
                    ws.Range("O1").Value = "Ticker"
                    ws.Range("P1").Value = "Value"
                    Cells(2, 14).Value = "Greatest % Increase"
                    Cells(3, 14).Value = "Greatest % Decrease"
                    Cells(4, 14).Value = "Greatest Total Volume"
                End If
            'Else   ' condition is false
                       
            
            End If
        
        Next i
    
    Next ws

End Sub

