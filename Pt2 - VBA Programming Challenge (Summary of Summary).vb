Option Explicit

Sub Performance()
 With Application
   .StatusBar = "The macro is running"
   .ScreenUpdating = False
   .DisplayAlerts = False
 End With
 
 Dim Sh As Worksheet
 
 For Each Sh In ThisWorkbook.Worksheets
'Label columns for Summary Table
   Sh.Range("O2").Value = "Greatest % Increase"
   Sh.Range("O3").Value = "Greatest % Decrease"
   Sh.Range("O4").Value = "Greatest Total Volume"
   Sh.Range("P1").Value = "Ticker"
   Sh.Range("Q1").Value = "Value"
   
'Define variables for summary table loop logic
  Dim R As Long                                                     'row counter
  Dim LastRowSum As Long                                            'identify last row of summary table (for biggest inc/dec table)
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



