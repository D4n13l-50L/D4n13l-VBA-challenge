Attribute VB_Name = "VBA_ofWAll_street"
Sub TickerIII()




'Create header

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


'set initial variable for sholding the brand name
Dim Ticker_Name As String
'Set Closing tickerDAteMax and open_tickerDAteMin
Dim closing_tickerDAteMax As Double
Dim open_tickerDAteMin As Double

'Set an initial variable for holding the total$ per credit card brand
Dim Ticker_Total As Double
'Set an initial variable for holding the open value
Dim Ticker_open As Double
'Set an initial variable for holding the close value
Dim Ticker_close As Double

'As always when ever you compare or sum, need to stablis initial value
Ticker_Total = 0
Ticker_open = 0
Ticker_close = 0
closing_tickerDAteMax = 0
open_tickerDAteMin = Range("C2").Value


'creamos otras variable Auxiliar para llevar el control de la fila de cada credit card in the summary table
Dim Summary_Table_Row As Integer
Dim Summary_Table_RowII As Integer

'Se pone un valor incial de "2" porque la lista inicia en la fila 2
Summary_Table_Row = 2

'loop throug the list
Dim i As Double

'create lastrow variable
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    'check if we are still within the same Ticker name,
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set Ticker name
        Ticker_Name = Cells(i, 1).Value
        'Add to the ticker total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        'Print the Ticker in the sumary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Name
        'Print the Ticker volumen to the summary Table
        Range("L" & Summary_Table_Row).Value = Ticker_Total
        'Print the yearly change
        Range("J" & Summary_Table_Row).Value = closing_tickerDAteMax - open_tickerDAteMin
        If Range("J" & Summary_Table_Row).Value > 0 Then
           Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If

        'if divided "0" print 0
        If open_tickerDAteMin = 0 Then
            Range("K" & Summary_Table_Row).Value = 0
        'Print the percentaje
        Else
            Range("K" & Summary_Table_Row).Value = (closing_tickerDAteMax - open_tickerDAteMin) / open_tickerDAteMin
        'End If
        'If Range("K" & Summary_Table_Row).Value > 0 Then
            'Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        'Else
            'Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        'this is our ticker first day of the year opening value
        open_tickerDAteMin = Cells(i + 1, 3).Value

         'Add one to the summary table Row
         Summary_Table_Row = Summary_Table_Row + 1
     
        'Reset the total vol to "0"
        Ticker_Total = 0
     
        'since, we tink as aconditional, IF the cell inmediately following a row is the same brand but we use the Else since is the contrary to if in this scenario.
    Else
    
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        'this is our ticker last day of the year clossing value
        closing_tickerDAteMax = Cells(i + 1, 6).Value
    
    End If
    Next i
'convert to %
  
       Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"

MsgBox ("Done")
End Sub
