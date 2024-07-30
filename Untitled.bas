Attribute VB_Name = "Module1"
Sub alphabetical_testing()

' Define Varibles
Dim I As Long
Dim LastRow As Long
Dim ws As Worksheet
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Quarterly As Double
Dim TotalVolume As Double
Dim Volume As Double
Dim Percentage As Double
Dim GIncrease As Double
Dim GDecrease As Double
Dim GVolume As Double
Dim ITicker As String
Dim DTicker As String
Dim VTicker As String

For Each ws In Worksheets

'Count rows in first column
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Headers for calculated
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Volume"

Dim summary_table As String
        summary_table = 2
Dim TickCount As Long
        TickCount = 2
        
'Set the start Value
TotalVolume = 0

'Loop

OpenPrice = 0
ClosePrice = 0
Volume = 0
GIncrease = 0
GDecrease = 0
GVolume = 0
ITicker = ""
DTicker = ""
VTicker = ""

For I = 2 To LastRow 'main for loop, loop through each row on ws
    
    'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
        OpenPrice = OpenPrice + ws.Cells(I, 3).Value
        
        ClosePrice = ClosePrice + ws.Cells(I, 6).Value
        
      'Sum up Total Stock Volume

      Volume = Volume + ws.Cells(I, 7).Value
    
    'When current cell value is differant then the next then save
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            'set ticker name
            Dim Ticker As String
            Ticker = ws.Cells(I, 1).Value
            ws.Cells(summary_table, 9).Value = Ticker
        
             Quarterly = OpenPrice - ClosePrice
             ws.Cells(summary_table, 10).Value = Quarterly
              
             'Set color pos as green and neg as red
            If Quarterly > 0 Then
                ws.Cells(summary_table, 10).Interior.ColorIndex = 10 ' Green
            ElseIf Quarterly < 0 Then
                ws.Cells(summary_table, 10).Interior.ColorIndex = 3 ' Red
            End If
            
             ws.Cells(summary_table, 11).Interior.ColorIndex = 0 'no fill color for column 11
            
            Percentage = (Quarterly / OpenPrice) * 100
             ws.Cells(summary_table, 11).Value = Percentage
             
            'set total stock volume
            ws.Cells(summary_table, 12).Value = Volume
            
            'increment summary table
                 summary_table = summary_table + 1
             'greastest increase
    
              If GIncrease = 0 Or Percentage > GIncrease Then
                    GIncrease = Percentage
                    ITicker = Ticker
            End If
            'greates decrease
            If GDecrease = 0 Or Percentage < GDecrease Then
                GDecrease = Percentage
                DTicker = Ticker
                End If
                
            'greatest Total volume
            If GVolume = 0 Or Volume > GVolume Then
                    GVolume = Volume
                    VTicker = Ticker
            End If
    
           
            OpenPrice = 0
            ClosePrice = 0
            Volume = 0
        
    
    End If 'main end if
        
 
Next I
 '   "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
  
ws.Cells(2, 16).Value = FormatPercent(GIncrease)
ws.Cells(3, 16).Value = FormatPercent(GDecrease)
ws.Cells(4, 16).Value = GVolume

ws.Cells(2, 15).Value = ITicker
ws.Cells(3, 15).Value = DTicker
ws.Cells(4, 15).Value = VTicker


Next ws

 
End Sub

