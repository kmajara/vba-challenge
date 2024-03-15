Attribute VB_Name = "Module1"


 Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim CurrentWorksheet As Integer
         Dim ws As Worksheet

         ' Set WS_Count equal to the number of worksheets in the active workbook.
         
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the outside (worksheets) loop.
         For CurrentWorksheet = 1 To WS_Count
             Set ws = ActiveWorkbook.Worksheets(CurrentWorksheet)

      ' Define source and destination columns
        
            Dim openingPriceColumn As Integer
            Dim closingPriceColumn As Integer
            Dim YearlyChangeColumn As Integer
            
            openingPriceColumn = 3
            closingPriceColumn = 6
            YearlyChangeColumn = 10
            
            Dim lastRow As Long
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    ' initialize variables
            Dim I As Long, j As Long
            Dim YearlyChange As Double
            Dim PercentageChange As Double
            Dim TotalVolume As LongLong
            Dim column As Integer
            Dim TickerName As String
            Dim TickerCount As Integer
            Dim WorksheetName As String
            Dim SummaryTableRow As Integer
            SummaryTableRow = 2
            

' Assign Column Headers
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
         End With

'--------------------------------------------------------------------------------------------------------------
 'Loop that inputs information in the Summary table for Ticker, Yearly Change, Percentage Change, and Total stock Volume Columns
 '--------------------------------------------------------------------------------------------------------------
   

    For I = 2 To lastRow
    TickerCount = I - 252
  
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
     
                'Set the Ticker Name
                  TickerName = ws.Cells(I, 1).Value
                      
                'Update Total Stock Volume
                  TotalVolume = TotalVolume + ws.Cells(I, 7).Value
                      
                'Calculate yearly change andpercentage change
                  Dim BeginningPrice As Double
                  Dim EndingPrice As Double
                  BeginningPrice = ws.Cells(TickerCount, openingPriceColumn).Value
                  EndingPrice = ws.Cells(I, closingPriceColumn).Value
                  YearlyChange = EndingPrice - BeginningPrice
                  PercentageChange = (YearlyChange / BeginningPrice) * 100
                  
            
                  'Print Summary Table
                  With ws
                    .Range("I" & SummaryTableRow).Value = TickerName
                    .Range("J" & SummaryTableRow).Value = YearlyChange
                    .Range("K" & SummaryTableRow).Value = Round(PercentageChange, 2)
                    .Range("L" & SummaryTableRow).Value = TotalVolume
                  End With
                  
                  'Advance to the next SummaryTableRow
                  SummaryTableRow = SummaryTableRow + 1
                  
                   'Reset Total Volume
                   TotalVolume = 0
        Else
                ' Add Cumulative Volume if current ticker is the same as next ticker
                TotalVolume = TotalVolume + Cells(I, 7).Value
                
    End If
    
  Next I

    
 '------------------------------------------------------------------------------
 'Apply conditional formating to make negative cells red and positive cells green
 '-------------------------------------------------------------------------------
 For Each cell In ws.Range("J1:K500")
    
     ' Check if the cell value is negative
    If cell.Value < 0 Then
      cell.Interior.ColorIndex = 3
    ElseIf cell.Value > 0 Then
     cell.Interior.ColorIndex = 10
    End If

  Next cell

 '------------------------------------------------------------------------------
 'Calculate  and update the worksheet with the 3 Metrics: Greates % Increase, Greatest % Decrease, and Greatest Total Volume
 '-------------------------------------------------------------------------------

  ' Assign Column Headers for the Metrics
    With ws
        .Range("P2").Value = "Greatest % Increase"
        .Range("P3").Value = "Greatest % decrease"
        .Range("P4").Value = "Greatest total volume"
        .Range("R1").Value = "Ticker"
        .Range("S1").Value = "Value"
        .Range("T1").Value = "Reference Row"
    End With
    
    'Create variables to store Metrics values
        Dim maxPCDecrease As Double
        Dim maxPCIncrease As Double
        Dim maxVol As LongLong
        
        Dim pcRange As Range
        Dim volRange As Range
        
        ' Set ranges for the Percent change and volume
        Set pcRange = ws.Range("K2:K500")
        Set volRange = ws.Range("L2:L500")
        
        'Define range containing ticker - offset 2 and 3 columns to the left of the percentage and total volume, respectively.
        Dim tickerRange As Range
        Set tickerRange = pcRange.Offset(0, -2)
        Dim volTickerRange As Range
        Set volTickerRange = volRange.Offset(0, -3)
        
        'Calculate values in range
        maxPCDecrease = WorksheetFunction.Min(ws.Range("K2:K500"))
        maxPCIncrease = WorksheetFunction.Max(ws.Range("K2:K500"))
        maxVol = WorksheetFunction.Max(ws.Range("L2:L500"))
        
        'Find Row of Max Values
        Dim maxRowPCIncrease As Double
        Dim maxRowPCDecrease As Double
        Dim MaxRowVol As Long
        
        maxRowPCIncrease = Application.Match(maxPCIncrease, pcRange, 0)
        maxRowPCDecrease = Application.Match(maxPCDecrease, pcRange, 0)
        MaxRowVol = Application.Match(maxVol, volRange, 0)
        
        'Assign values to respective cells the result
        
        'Greatest Percentage Increase
        With ws
            .Range("S2").Value = maxPCIncrease
            .Range("R2").Value = tickerRange.Cells(maxRowPCIncrease, 1).Value
            .Range("T2").Value = maxRowPCIncrease + 1
        End With
        'Greatest Percentage Decrease
        With ws
            .Range("S3").Value = maxPCDecrease
            .Range("R3").Value = tickerRange.Cells(maxRowPCDecrease, 1).Value
            .Range("T3").Value = maxRowPCDecrease + 1
        End With
        'Greatest Total Volume
        With ws
            .Range("S4").Value = maxVol
            .Range("R4").Value = volTickerRange.Cells(MaxRowVol, 1).Value
            .Range("T4").Value = MaxRowVol + 1
        End With
    Next CurrentWorksheet
End Sub
    
        
        
    
    
    
    









