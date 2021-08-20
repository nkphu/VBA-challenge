

Sub stock_analysis()
 
' Set worksheet variable.
Dim CurrentWs As Worksheet
    
' Loop through all of the worksheets
For Each CurrentWs In Worksheets
    
' Set variable for holding the ticker index
Dim ticker_index As String
ticker_index = " "
        
' Set variable for holding total stock volume
Dim total_stock_volume As Double
total_stock_volume = 0
        
' Set variable for yearly change price

Dim yearly_open_price As Double
yearly_open_price = 0
Dim yearly_close_price As Double
yearly_close_price = 0
Dim yearly_change_price As Double
yearly_change_price = 0
Dim yearly_change_percentage As Double
yearly_change_percentage = 0

' Set variable for greatest increase and decrease + total volume
Dim MAX_TICKER_INDEX As String
MAX_TICKER_INDEX = " "
Dim MIN_TICKER_INDEX As String
MIN_TICKER_INDEX = " "
Dim MAX_PERCENTAGE As Double
MAX_PERCENTAGE = 0
Dim MIN_PERCENTAGE As Double
MIN_PERCENTAGE = 0
Dim MAX_VOLUME_TICKER As String
MAX_VOLUME_TICKER = " "
Dim MAX_VOLUME As Double
MAX_VOLUME = 0
   
'summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
        
Dim Lastrow As Long
Dim i As Long
        
Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

' Set headings for tables
CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"

CurrentWs.Range("O2").Value = "Greatest % Increase"
CurrentWs.Range("O3").Value = "Greatest % Decrease"
CurrentWs.Range("O4").Value = "Greatest Total Volume"
CurrentWs.Range("P1").Value = "Ticker"
CurrentWs.Range("Q1").Value = "Value"

yearly_open_price = CurrentWs.Cells(2, 3).Value
        
' Loop
For i = 2 To Lastrow
        
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
        
ticker_index = CurrentWs.Cells(i, 1).Value
                
' yearly change price formula
yearly_close_price = CurrentWs.Cells(i, 6).Value
yearly_change_price = yearly_close_price - yearly_open_price

If yearly_open_price <> 0 Then

yearly_change_percentage = (yearly_change_price / yearly_open_price) * 100
               
End If
                
'set total stock volume formula
total_stock_volume = total_stock_volume + CurrentWs.Cells(i, 7).Value
              
                
'set values in summary table
CurrentWs.Range("I" & Summary_Table_Row).Value = ticker_index
CurrentWs.Range("J" & Summary_Table_Row).Value = yearly_change_price

'condition format colours
If (yearly_change_price > 0) Then

'set green if over 0
CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

ElseIf (yearly_change_price <= 0) Then
'set red if =<0
CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If
                
'set values in summary table
CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(yearly_change_percentage) & "%")
              
CurrentWs.Range("L" & Summary_Table_Row).Value = total_stock_volume
                
'summary table
Summary_Table_Row = Summary_Table_Row + 1
        
yearly_change_price = 0
                
yearly_close_price = 0
                
yearly_open_price = CurrentWs.Cells(i + 1, 3).Value
              
                
'calulate percentage
If (yearly_change_percentage > MAX_PERCENTAGE) Then
MAX_PERCENTAGE = yearly_change_percentage
MAX_TICKER_INDEX = ticker_index

ElseIf (yearly_change_percentage < MIN_PERCENTAGE) Then

MIN_PERCENTAGE = yearly_change_percentage

MIN_TICKER_INDEX = ticker_index

End If
                       
If (total_stock_volume > MAX_VOLUME) Then

MAX_VOLUME = total_stock_volume
MAX_VOLUME_TICKER = ticker_index

End If
                
yearly_change_percentage = 0
total_stock_volume = 0
                               
Else

'set total stock formula
total_stock_volume = total_stock_volume + CurrentWs.Cells(i, 7).Value


End If
      
Next i

'percentage values allocation
CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENTAGE) & "%")
CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENTAGE) & "%")
CurrentWs.Range("P2").Value = MAX_TICKER_INDEX
CurrentWs.Range("P3").Value = MIN_TICKER_INDEX
CurrentWs.Range("Q4").Value = MAX_VOLUME
CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER

        
Next CurrentWs

End Sub

