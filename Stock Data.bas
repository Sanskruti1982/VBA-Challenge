Attribute VB_Name = "Module1"
Sub stock()
'Loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'Set an intial variable for holding the ticker name
Dim Ticker As String

'Set an intial variable for holding the total volume per ticker, ticker open and close price, difference in prices and percent change
Dim Ticker_Vol As Double
Dim Ticker_Start As Double
Dim Ticker_End As Double
Dim Ticker_Diff As Double
Dim Ticker_Per As Double
Ticker_Vol = 0

'Set variables for max, min and max vol
Dim maxperinc As Double
Dim minperdec As Double
Dim maxstockvol As Double
Dim Tickermax As Double
Dim Tickermin As Double
Dim Tickermaxvol As Double


'Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Summary_Table_Row1 As Integer
Summary_Table_Row1 = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stock values
For i = 2 To lastrow
    
    'Check if we are still within the same ticker, if it is not then
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the ticker
        Ticker = ws.Cells(i, 1).Value
      
        
        'Add to the Ticker Volume
        Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
        
        'Print the ticker in the summary table
        ws.Range("M" & Summary_Table_Row).Value = Ticker
        
        'Print the ticker volume in the summary table
        ws.Range("N" & Summary_Table_Row).Value = Ticker_Vol
    
        
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the ticker vol
        Ticker_Vol = 0
        
        'If the cellimmediately following a row is the same ticker
        Else
        
        'hold the ticker volume
        Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
        
        
        
    End If
        
Next i
        

'Loop through all high and low values
For i = 1 To lastrow
    
    'Check if we are still within the same ticker, if it is not then
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'set the open price
        Ticker_Start = ws.Cells(i + 1, 3).Value
                
        'Print the ticker in the summary table
        ws.Range("O" & Summary_Table_Row1).Value = Ticker_Start
        
        'Add one to the summary table row
        Summary_Table_Row1 = Summary_Table_Row1 + 1
       
        
    'If the cellimmediately following a row is the same ticker
    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    
        'Set the ticker start
       Ticker_End = ws.Cells(i + 1, 6).Value
      
        'Print the ticker start in the summary table
        ws.Range("P" & Summary_Table_Row1 - 1).Value = Ticker_End
        
              
        
    End If
    
    'If Ticker Start price is greater than zero then calculate the difference and percentage difference
    If Ticker_Start > 0 Then
       
       Ticker_Diff = Ticker_End - Ticker_Start
       ws.Range("Q" & Summary_Table_Row1 - 1).Value = Ticker_Diff
       Ticker_Per = (Ticker_Diff / Ticker_Start)
       ws.Range("R" & Summary_Table_Row1 - 1).Value = Ticker_Per
      
       
    Else
       
    End If
     
     
    'If the ticker percentage increase is greater than zero then colour green or else red
    If Ticker_Per > 0 Then
       ws.Range("R" & Summary_Table_Row1 - 1).Interior.ColorIndex = 4
       
    Else
       
       ws.Range("R" & Summary_Table_Row1 - 1).Interior.ColorIndex = 3
       
 
    End If
       
       
       
       
Next i

'-------------------------------------'
'CHALLENGE'
'-------------------------------------'


    'Getting the max and matching ticker value and formatting it to %
    maxperinc = WorksheetFunction.Max(ws.Range("R2:R" & lastrow))
    Tickermax = WorksheetFunction.Match(maxperinc, ws.Range("R2:R" & lastrow), 0)
    ws.Cells(2, 22).Value = maxperinc
    ws.Cells(2, 22).NumberFormat = "0.00%"
    ws.Cells(2, 21).Value = ws.Cells(Tickermax + 1, 13)
    ws.Range("R2:R" & lastrow).NumberFormat = "0.00%"

    'Getting the min and matching ticker value and formatting it to %
    minperdec = WorksheetFunction.Min(ws.Range("R2:R" & lastrow))
    Tickermin = WorksheetFunction.Match(minperdec, ws.Range("R2:R" & lastrow), 0)
    ws.Cells(3, 22).Value = minperdec
    ws.Cells(3, 22).NumberFormat = "0.00%"
    ws.Cells(3, 21).Value = ws.Cells(Tickermin + 1, 13)

    'Getting the max stock value and matching ticker value
    maxstockvol = WorksheetFunction.Max(ws.Range("N2:N" & lastrow))
    Tickermaxvol = WorksheetFunction.Match(maxstockvol, ws.Range("N2:N" & lastrow), 0)
    ws.Cells(4, 22).Value = maxstockvol
    ws.Cells(4, 21).Value = ws.Cells(Tickermaxvol + 1, 13)

    'Print all the headings
    ws.Range("M1").Value = "Ticker"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker Start"
    ws.Range("p1").Value = "Ticker End"
    ws.Range("Q1").Value = "Yearly Change"
    ws.Range("R1").Value = "Percent Change"
    ws.Range("T2").Value = "Greatest % Increase"
    ws.Range("T2:T5").Columns.AutoFit
    ws.Range("T3").Value = "Greatest % Decrease"
    ws.Range("T4").Value = "Greatest Total Volume"
    ws.Range("U1").Value = "Ticker"
    ws.Range("V1").Value = "Value"

Next ws



End Sub

