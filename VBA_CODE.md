
#Create a VBA Script to loop through 800k rows of stock value and return Performance data. 


Sub StockData()

'Declare Variables

Dim TickerName As String
Dim WorksheetName As String
Dim counter As Integer
Dim tickerstart As Double
Dim tickerend As Double

'---------------------------------------
'Loop through all worksheets
'---------------------------------------

For Each ws In Worksheets


WorksheetName = ws.Name

'Set initial Ticker Value to 0

Dim TickerValue As Double
TickerValue = 0
  
Dim Summary_TableRow As Integer

Summary_TableRow = 2

'determine the last row

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set the first ticker position (for Year end Changes)

counter = 1

'Set the Loop

For i = 2 To lastrow

'write conditionals

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Find the Last Ticker Position
        
        tickerend = ws.Cells(i, 6).Value
        
        'set the ticker name
        
        TickerName = ws.Cells(i, 1).Value
        
        'Add to the ticker value
        
        TickerValue = TickerValue + ws.Cells(i, 7).Value
        
        'Print tickername into summary table
        
        ws.Range("I" & Summary_TableRow).Value = TickerName
        
         'Print the Ticker Amount to the Summary Table
        
        ws.Range("L" & Summary_TableRow).Value = TickerValue
        
        'Print year end change to the summary table
        
        ws.Range("J" & Summary_TableRow).Value = tickerend - tickerstart
        
        
        'get rid of division by 0 error

        If tickerstart = 0 Then
        ws.Cells(i, 11).Value = "N/A"
        
        Else
        
        'find percent change and print to summary table
        
        ws.Range("K" & Summary_TableRow).Value = ((tickerend - tickerstart) / tickerstart)
        
        End If
        
        'conditional formatting
    
        If ws.Range("J" & Summary_TableRow).Value < 0 Then
        ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 3
        
        Else
        ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 4
        ws.Range("J" & Summary_TableRow).Style = "currency"
        
        End If
        
        'change column k to percent
        
        ws.Range("K" & Summary_TableRow).Style = "percent"
                
           ' Add one to the summary table row
           
           Summary_TableRow = Summary_TableRow + 1
              
           'change column J to currency
           
              ws.Range("L2", "L" & Summary_TableRow).Style = "currency"
           
              
              ' Reset the Ticker Total and counter
              
              TickerValue = 0
                
              counter = 1
    
        ' If the cell immediately following a row is the same ticker...
        Else
            
            'find the first ticker using a counter
            
            If counter = 1 Then
                tickerstart = ws.Cells(i, 3).Value
                counter = 2
            End If
                      
          ' Add to the ticker Total
          
          TickerValue = TickerValue + ws.Cells(i, 7).Value
    
        End If
    
        
Next i


'populate the summary table

    lastrow2 = ws.Cells(Rows.Count, 1).End(xlUp).Row
       ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K2", "K" & lastrow2))
       ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K2", "K" & lastrow2))
       ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L2", "L" & lastrow2))

    'find the ticker associated with Max and Min
    
    For j = 2 To lastrow2
    
    If ws.Cells(j, 11).Value = ws.Range("Q2").Value Then
    ws.Range("P2").Value = ws.Cells(j, 9).Value
    ws.Range("Q2").Style = "percent"
    
    End If
    
    If ws.Cells(j, 11).Value = ws.Range("Q3").Value Then
    ws.Range("P3").Value = ws.Cells(j, 9).Value
    ws.Range("Q3").Style = "percent"
    End If
    
    If ws.Cells(j, 12).Value = ws.Range("Q4").Value Then
    ws.Range("P4").Value = ws.Cells(j, 9).Value
    ws.Range("Q4").Style = "currency"
    End If
    
    Next j
    
    'create summary table headers
    
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
Next ws

End Sub