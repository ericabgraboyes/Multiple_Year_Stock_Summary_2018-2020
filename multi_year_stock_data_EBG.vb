Option Explicit

Sub Performance()
 With Application
   .StatusBar = "The macro is running"
   .ScreenUpdating = False
   .DisplayAlerts = False
 End With
 
 Dim Sh As Worksheet
 
 For Each Sh In ThisWorkbook.Worksheets

  Const SRow As Byte = 2                    'specify constant for starting row of data range
  Dim R As Long                             'row counter
  Dim LastRowData As Long                   'identify last row for data range
  Dim LastRowSum As Long                    'identify last row of summary table (for biggest inc/dec table)
  Dim Ticker As String                      'identify each stock ticker
  Dim OpenPrice As Double                   'identify opening price (1st day of year)
  Dim ClosePrice As Double                  'identify closing price (last day of year)
  Dim YoYChange As Double                   'identify YoY Price Change ($)
  Dim PrctChange As Double                  'identify % Price Change
  Dim Volume As Double                      'holds volume associated with each ticker (additive measure)
  Dim SumRow As Integer                     'Keep track of location for ticker in summary table
  Dim PriceRow As Long                      'Keep track of row where open price stored (per ticker)
    
'Determine Last Row of Dataset
  LastRowData = Sh.Range("A" & Rows.Count).End(xlUp).Row
  
'Assign initial values for Price Row, Volume and Summary Table
  PriceRow = 2
  SumRow = 2
  Volume = 0

'Label columns for Summary Table
   Sh.Range("I1").Value = "Ticker"
   Sh.Range("J1").Value = "Yearly Change"
   Sh.Range("K1").Value = "Percent Change"
   Sh.Range("L1").Value = "Total Stock Volume"
   Sh.Range("O2").Value = "Greatest % Increase"
   Sh.Range("O3").Value = "Greatest % Decrease"
   Sh.Range("O4").Value = "Greatest Total Volume"
   Sh.Range("P1").Value = "Ticker"
   Sh.Range("Q1").Value = "Value"
   
'Loop through dataset in columns A - G, identify different tickers, and stock performance metrics by ticker
 For R = SRow To LastRowData
   If Sh.Range("A" & R + 1).Value <> Sh.Range("A" & R).Value Then      'If Ticker Row Below <> current row....
    Ticker = Sh.Range("A" & R).Value                                     'Set Ticker
    Volume = Volume + Sh.Range("G" & R).Value                            'Add Total Stock Volume
    OpenPrice = Sh.Range("C" & PriceRow).Value                           'Set Open Price
    ClosePrice = Sh.Range("F" & R).Value                                 'Set Close Price
    YoYChange = ClosePrice - OpenPrice                                   'Calculate YoY Change
      
    If OpenPrice = 0 Then
     PrctChange = 0                                                      'If Open Price is 0, no % change
    Else
     PrctChange = YoYChange / OpenPrice                                  '% Change from Open Price
    End If
      
    Sh.Range("I" & SumRow).Value = Ticker                                'Write Ticker to Summary
    Sh.Range("J" & SumRow).Value = YoYChange                             'Write YoY Change to Summary
    Sh.Range("K" & SumRow).Value = PrctChange                            'Write Percent Change to Summary
    Sh.Range("L" & SumRow).Value = Volume                                'Write Total Volume to Summary
    
    Sh.Range("J" & SumRow).NumberFormat = "#,##0.00_);(#,##0.00)"
    Sh.Range("K" & SumRow).NumberFormat = "0.00%"
      
    If Sh.Range("J" & SumRow).Value > 0 Then
     Sh.Range("J" & SumRow).Interior.Color = vbGreen
    Else
     Sh.Range("J" & SumRow).Interior.Color = vbRed
    End If
    
    If Sh.Range("K" & SumRow).Value > 0 Then
     Sh.Range("K" & SumRow).Interior.Color = vbGreen
    Else
     Sh.Range("K" & SumRow).Interior.Color = vbRed
    End If
      
    'Add new row to summary table
    SumRow = SumRow + 1
    
    'update row reference for opening price
    PriceRow = R + 1
    
    'reset volume counter to 0
    Volume = 0
    
 'if the ticker in the row below is the same as the current row, then add volume and go to next row
   Else
    Volume = Volume + Sh.Range("G" & R).Value
   End If
  Next R

'Define variables for summary table loop logic
  Dim BiggestInc As Double                                          'identify largest price increase
  Dim BiggestDec As Double                                          'identify largest price decrease
  Dim MostVol As Double                                             'identify greatest volume traded
  Dim BiggestInc_Ticker As String                                   'ticker associated with largest price increase
  Dim BiggestDec_Ticker As String                                   'ticker associated with largest price decrease
  Dim MostVol_Ticker As String                                      'ticker associated with most volume

'Set 1st Ticker as Biggest Increase, Decrease, Volume
  BiggestInc = Sh.Range("K2").Value
  BiggestDec = Sh.Range("K2").Value
  MostVol = Sh.Range("L2").Value

'Determine Last Row of Summary Table
 LastRowSum = Sh.Range("I" & Rows.Count).End(xlUp).Row
 
 For R = 2 To LastRowSum
   If Sh.Range("K" & R + 1).Value > BiggestInc Then                     'If Ticker in row below has greater Inc....
    BiggestInc = Sh.Range("K" & R + 1).Value                              'Update Greatest Inc %
    BiggestInc_Ticker = Sh.Range("I" & R + 1).Value                       'Update Ticker for Inc %
    
   ElseIf Sh.Range("K" & R + 1).Value < BiggestDec Then
    BiggestDec = Sh.Range("K" & R + 1).Value                              'Update Greatest Dec %
    BiggestDec_Ticker = Sh.Range("I" & R + 1).Value                       'Update Ticker for Dec %
    
   ElseIf Sh.Range("L" & R + 1).Value > MostVol Then
    MostVol = Sh.Range("L" & R + 1).Value                                 'Update Most Volume
    MostVol_Ticker = Sh.Range("I" & R + 1).Value                          'Update Ticker for Most volume
   End If
  Next R
  
  Sh.Range("P2").Value = BiggestInc_Ticker
  Sh.Range("P3").Value = BiggestDec_Ticker
  Sh.Range("P4").Value = MostVol_Ticker
  Sh.Range("Q2").Value = BiggestInc
  Sh.Range("Q3").Value = BiggestDec
  Sh.Range("Q4").Value = MostVol
  Sh.Range("Q2:Q3").NumberFormat = "0.00%"
   
 Sh.Range("I:L").EntireColumn.AutoFit
 Sh.Range("O:Q").EntireColumn.AutoFit
 
 Next Sh

 With Application
   .StatusBar = ""
   .ScreenUpdating = True
   .DisplayAlerts = True
End With

End Sub


