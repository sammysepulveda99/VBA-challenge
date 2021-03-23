VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stockdata()

'define each variable
Dim Ticker As String
Dim VolumeStock As Long
Dim YearOpenPrice As Double
Dim YearClosePrice As Double
Dim YearChange As Double
Dim PercentChange As Double
Dim SummaryTable As Integer

'Putting the names on the columns'
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
VolumeStock = 0
SummaryTable = 2
YearOpenPrice = Cells(2, 3).Value

'Last row function to find the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
'Start loop
For i = 2 To LastRow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
       'Assigning value to the variables
        Ticker = Cells(i, 1).Value
        VolumeStock = VolumeStock + Cells(i, 7).Value
        YearClosePrice = Cells(i, 6).Value
        YearChange = YearClosePrice - YearOpenPrice
        'If  the YearOpen has a value of 0
        If YearOpenPrice = 0 Then
             PercentChange = 0
        Else
             PercentChange = YearChange / YearOpenPrice
        End If
          'insert values into summary
            Cells(SummaryTable, 9).Value = Ticker
            Cells(SummaryTable, 10).Value = YearChange
            Cells(SummaryTable, 11).Value = PercentChange
            Cells(SummaryTable, 12).Value = VolumeStock
            If Cells(SummaryTable, 10) < 0 Then
                Cells(SummaryTable, 10).Interior.Color = vbRed
            Else
                Cells(SummaryTable, 10).Interior.Color = vbGreen
            End If
            
            SummaryTable = SummaryTable + 1
            VolumeStock = 0
            YearOpenPrice = Cells(i + 1, 3).Value
    'Assigning a percentage format to all K columns'
    Columns("K").NumberFormat = "0.00%"
    End If

'close loop
Next i

End Sub

