Sub Main()
        
        Dim LastCell As Long
        Dim difference As Double 'closing - open
        Dim ftc As Long 'first ticker cell
        Dim openAmount As Double '
        Dim closeAmount As Double
        Dim percentageChange As Double
        Dim j  As Long 'yearchangetrcker
        Dim totalVol As Double
        
For Each ws_tabs In Worksheets
         
        'activate worksheets and set autofit
        ws_tabs.Activate
        'set column headers
          
        ws_tabs.Cells(1, 9).Value = "Ticker"
        ws_tabs.Cells(1, 10).Value = "Yearly Change"
        ws_tabs.Cells(1, 11).Value = "Percentage Change"
        ws_tabs.Cells(1, 12).Value = "Total Stock Volume"

ftc = 2
totalVol = 0
j = 2 ' track which row to insert found data
LastCell = Cells(Rows.Count, "A").End(xlUp).Row
'newTickercol = Cells(Rows.Count, "I").End(xlUp).Row

For i = 2 To LastCell
    totalVol = totalVol + Cells(i, 7).Value
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i + 1, 1).Value <> "" Then
          
           'set close amount and open ammount
           openAmount = Cells(ftc, 3).Value
           closeAmount = Cells(i, 6).Value
           
           'calculate the change
           
           difference = closeAmount - openAmount
           
           'set the ticker value and change value in new column
           
           Range("I" & j).Value = Cells(i, 1).Value
           Range("J" & j).Value = difference
           
           'set formating for the price change
           
           If Range("J" & j).Value > 0 Then
               Range("J" & j).Interior.ColorIndex = 4
           ElseIf Range("J" & j).Value < 0 Then
               Range("J" & j).Interior.ColorIndex = 3
           End If
               
          'set the total volume
           Range("L" & j).Value = totalVol
           
          'calculate the percent change
          
           If openAmount > 0 Then
                percentageChange = (difference / openAmount) * 100
                Range("K" & j).Value = percentageChange
                
            Else
                Range("K" & j).Value = 0
            End If
           
            j = j + 1
            totalVol = 0
            ftc = i + 1

    End If
Next i
ws_tabs.Columns("A:L").AutoFit
ws_tabs.Columns.Range("K:K").NumberFormat = "0.00\%"

Next ws_tabs
MsgBox ("All Worksheets have been processed")
End Sub

'+++++++++++++++++++++++Bonus challenge function +++++++++++++++++++++++++++++++++++

Sub bonus_challenge()

Dim k As Integer
k = Cells(Rows.Count, "J").End(xlUp).Row



For Each ws_tabs In Worksheets

    'activate worksheets
    ws_tabs.Activate
    
    'enter column names
    ws_tabs.Cells(1, 15).Value = "Ticker"
    ws_tabs.Cells(1, 16).Value = "Value"
    
    'name ranges for max and min values
    
    Set cellrange = ws_tabs.Range("J:J")
                                    
    Set volcellrange = ws_tabs.Range("L:L")

        
 
        Dim MaxCell As Double 'cell value with max percentage increase
        Dim MinCell As Double  'Min cell value for min percentage decrease
        Dim MaxTotalvol As Double 'max value for max total stock volume
   'set named ranges to use in max and min functions
         
            
                     
            'set summary table field lables
            
            Range("N2").Value = "Greatest % Increase"
            Range("N3").Value = "Greatest % decrease"
            Range("N4").Value = "Greatest Total Volume"
            
            'Find Max and min values
            
            MaxCell = Application.WorksheetFunction.Max(cellrange)
            Range("P2").Value = MaxCell
            MinCell = Application.WorksheetFunction.Min(cellrange)
            Range("P3").Value = MinCell
            MaxTotalvol = Application.WorksheetFunction.Max(volcellrange)
            Range("P4").Value = MaxTotalvol
        
            For i = 2 To k
                If Cells(i, 10).Value = MaxCell Then
                    Range("O2").Value = Cells(i, 9).Value
                ElseIf Cells(i, 10).Value = MinCell Then
                    Range("O3").Value = Cells(i, 9).Value
                ElseIf Cells(i, 12).Value = MaxTotalvol Then
                    Range("O4").Value = Cells(i, 9).Value
                End If
            Next i
    ws_tabs.Columns("N:P").AutoFit
    
 Next ws_tabs
MsgBox ("All Worksheets have been processed")
End Sub



