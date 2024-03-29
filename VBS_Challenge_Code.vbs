Attribute VB_Name = "Module1"
Sub Year_Stock_Data()

For Each ws In Worksheets

    Dim Worksheet As String
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Volume As Variant
    
    Dim i, j As Integer
    
    Dim open_value, close_value As Double
    Dim Summary_Table_Row As Integer

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Summary_Table_Row = 2
    open_value = ws.Range("C2").Value
    Volume = 0
    Percentage_Change = 0
    
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("O2:O4").EntireColumn.AutoFit
    ws.Range("I2:I" & LastRow).EntireColumn.AutoFit
    ws.Range("J2:J" & LastRow).EntireColumn.AutoFit
    ws.Range("K2:K" & LastRow).EntireColumn.AutoFit
    ws.Range("L2:L" & LastRow).EntireColumn.AutoFit
    
    
    
    For i = 2 To LastRow
    
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
        close_value = ws.Cells(i, 6).Value
        Volume = Volume + ws.Cells(i, 7).Value
        Yearly_Change = close_value - open_value
        
        'MsgBox (Yearly_Change)
        
        'MsgBox (open_value)
        
        Percentage_Change = Yearly_Change / open_value
        
        'MsgBox (Percentage_Change)
        
        Ticker = ws.Cells(i, 1).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("L" & Summary_Table_Row).Value = Volume
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    
            If (Yearly_Change < 0) Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

            Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
            End If
    
        ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
            If (Yearly_Change < 0) Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

            Else
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    
            End If

        Summary_Table_Row = Summary_Table_Row + 1
        
        open_value = ws.Cells(i + 1, 3).Value
        close_value = 0
        Volume = 0
        
        Else
        
        Volume = Volume + ws.Cells(i, 7).Value
        
        
        End If

    
    Next i


    Dim Low As Double
    Dim High As Double
    Dim Top As Double
    Dim Low_Ticker As String
    Dim High_Ticker As String
    Dim Top_Ticker As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Low = ws.Range("K2").Value
    
    For i = 2 To LastRow
    
    If ws.Cells(i + 1, 11).Value < Low Then
    
    Low = ws.Cells(i + 1, 11).Value
    Low_Ticker = ws.Cells(i + 1, 9).Value
    'MsgBox (Low)
    'MsgBox (Low_Ticker)
    
    End If
    
    Next i
    
    High = ws.Range("K2").Value
    
    For i = 2 To LastRow
    
    If ws.Cells(i + 1, 11).Value > High Then
    
    High = ws.Cells(i + 1, 11).Value
    High_Ticker = ws.Cells(i + 1, 9).Value
    'MsgBox (High)
    'MsgBox (High_Ticker)


    
    End If
    
    
    Next i
    
    Top = ws.Range("L2").Value
    
    For i = 2 To LastRow
    
    If ws.Cells(i + 1, 12).Value > Top Then
    
    Top = ws.Cells(i + 1, 12).Value
    Top_Ticker = ws.Cells(i + 1, 9).Value
    'MsgBox (Top)
    'MsgBox (Top_Ticker)
    
    End If
    
    
    Next i
    
    
    ws.Range("Q2").Value = High
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = Low
    ws.Range("P2").Value = High_Ticker
    ws.Range("P3").Value = Low_Ticker
    ws.Range("Q4").Value = Top
    ws.Range("P4").Value = Top_Ticker
    ws.Range("Q2:Q4").EntireColumn.AutoFit
    

'A more optimized code for the above loops, I have learnt below. However, i had help from the Tutor to get it right, hence included as a comment.
'    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
'    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
'    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
'
'    Dim Max_Change_Index As Double
'    Dim Min_Change_Index As Double
'    Dim Max_Volume_Index As Double
'
'    Min_Change_Index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & LastRow), 0)
'    Max_Change_Index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & LastRow), 0)
'    Max_Volume_Index = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & LastRow), 0)
'
'    ws.Range("P3").Value = ws.Cells(Min_Change_Index + 1, 9).Value
'    ws.Range("P2").Value = ws.Cells(Max_Change_Index + 1, 9).Value
'    ws.Range("P4").Value = ws.Cells(Max_Volume_Index + 1, 9).Value
'    ws.Range("Q3").NumberFormat = "0.00%"


 Next ws

End Sub




