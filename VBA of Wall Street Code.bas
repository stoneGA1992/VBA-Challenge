Attribute VB_Name = "Module1"
Sub tticker()

  Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
  
 
For Each ws In ActiveWorkbook.Worksheets

    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"

  ' Set an initial variable for holding the ticker
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per ticker
  Dim Ticker_Total As Double
    Ticker_Total = 0
    
    ' Set up Dims for Open And Close Prices and yearly change
    
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim yearly_change As Double
    Dim yearly_percent_change As Double
    
    
    
    ' Set up last row
   Lastrow = Sheet2.Range("A" & Rows.Count).End(xlUp).Row


  ' Keep track of the location for each Ticker and opening price in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  year_open_price = ws.Cells(2, 3).Value
  
    
    'Begin LOOP
  For i = 2 To Lastrow

    ' Check if we are still within the same ticker, if it is not...
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the variables
      Ticker_Name = ws.Cells(i, 1).Value
      
      
      year_close_price = ws.Cells(i, 6).Value
      
      ' Calculations
      'Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
      
      yearly_change = year_close_price - year_open_price
      
        If year_open_price > 0 Then
             yearly_percent_change = yearly_change / year_open_price
            
                  ws.Range("k" & Summary_Table_Row).Value = yearly_percent_change
                  ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
                  
            End If
            
                If ws.Range("j" & Summary_Table_Row).Value > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                    ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                    ElseIf ws.Range("J" & Summary_Table_Row).Value = "" Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = ""
                    
                End If
            
      

      ' Print the Ticker in the Summary Table
      ws.Range("i" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Amount to the Summary Table
      ws.Range("l" & Summary_Table_Row).Value = Ticker_Total
      
      
      ' print yearly change
      ws.Range("j" & Summary_Table_Row).Value = yearly_change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0
     
     ' Reset Open price
     year_open_price = Cells(i + 1, 3).Value
      
       
      ' Add to the Ticker Total
      'Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
  'MsgBox ActiveWorkbook.Worksheets(j).Name'
  
  MsgBox "Loop check"
Next ws



  
End Sub



