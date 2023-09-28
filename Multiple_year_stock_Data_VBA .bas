VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_analysis()


' To loop through eash worksheet

  Dim ws As Worksheet
  For Each ws In Worksheets
  
' Set initial variables
  
  Dim Ticker_name As String
  Dim Open_price As Double
  Dim Close_price As Double
  
  Dim Yearly_change As Double
  Yearly_change = 0
  
  Dim Percent_change As Double
  
  Dim Total_volume As Double
  Total_volume = 0
  
  Dim Report_row As Integer
  Report_row = 2
  
' Name the reporting columns
  
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  ws.Columns("I:Q").AutoFit
  
'Set the open price
 
  Open_price = ws.Cells(2, 3).Value
   
'Identify the Last Raw
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  
  For i = 2 To LastRow
  
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
        Ticker_name = ws.Cells(i, 1).Value
        ws.Range("I" & Report_row).Value = Ticker_name
        
        Close_price = ws.Cells(i, 6).Value
        Yearly_change = Close_price - Open_price
        ws.Range("J" & Report_row).Value = Yearly_change
        
        
        Percent_change = Yearly_change / Open_price
        ws.Range("K" & Report_row).Value = Percent_change
        ws.Range("K" & Report_row).NumberFormat = "0.00%"
        
        Total_volume = Total_volume + ws.Cells(i, 7).Value
        ws.Range("L" & Report_row).Value = Total_volume
         
         Report_row = Report_row + 1
         
         Total_volume = 0
         
         Open_price = ws.Cells(i + 1, 3).Value
         
     Else
        
         Total_volume = Total_volume + ws.Cells(i, 7).Value
         
     End If
     
   Next i
   
   
   'Conditional Formatting The Yearly change column
   
   Last_Report_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row
   
   For j = 2 To Last_Report_Row
      If ws.Cells(j, 10).Value > 0 Then
         ws.Cells(j, 10).Interior.ColorIndex = 4
         
      ElseIf ws.Cells(j, 10).Value < 0 Then
           ws.Cells(j, 10).Interior.ColorIndex = 3
      End If
      
    Next j
    
    'Finding the Max and Min
    
    'Define the Variables
    
    
    Dim Max_Percent As Double
    Dim Min_percent As Double
    Dim Max_volume As Double
    Dim Max_ticker As String
    Dim Min_ticker As String
    Dim Vol_ticker As String
    
    
    ' Name the Rows and columns
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' set initial values
    Max_Percent = ws.Cells(2, 11).Value
    Min_percent = ws.Cells(2, 11).Value
    Max_volume = ws.Cells(2, 12).Value
    
    
    For k = 2 To Last_Report_Row
    
       If ws.Cells(k, 11).Value > Max_Percent Then
          Max_Percent = ws.Cells(k, 11).Value
          Max_ticker = ws.Cells(k, 9).Value
          
          ws.Cells(2, 16).Value = Max_ticker
          ws.Cells(2, 17).Value = Max_Percent
          ws.Cells(2, 17).NumberFormat = "0.00%"
       End If
          
       If ws.Cells(k, 11).Value < Min_percent Then
          Min_percent = ws.Cells(k, 11).Value
          Min_ticker = ws.Cells(k, 9).Value
          
          ws.Cells(3, 16).Value = Min_ticker
          ws.Cells(3, 17).Value = Min_percent
          ws.Cells(3, 17).NumberFormat = "0.00%"
       End If
       
       If ws.Cells(k, 12).Value > Max_volume Then
          Max_volume = ws.Cells(k, 12).Value
          Vol_ticker = ws.Cells(k, 9).Value
          
          ws.Cells(4, 16).Value = Vol_ticker
          ws.Cells(4, 17).Value = Max_volume
          
       End If
       
    Next k
   
   Next ws
        
        
        
 End Sub

