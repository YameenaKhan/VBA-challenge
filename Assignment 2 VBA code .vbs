Attribute VB_Name = "Module1"
Sub StocksData_Summary()
    
    'Looping all Sheets
    Ws_count = ActiveWorkbook.Worksheets.Count
    For w = 1 To Ws_count
    
    Dim Worksheet As String
    Dim Ticker As String
    Dim Open_price As Long
    Dim close_price As Long
    Dim high_price As Long
    Dim low_price As Long
    Dim total_volume As Double
    total_volume = 0
    Dim StartValue As Long
    StartValue = 2
    
    Dim Yearly_change As Double
    Dim Percentage_change As Double
    Dim summaryTablerow As Long
    summaryTablerow = 2
    
    Dim ConditFormatRange As Range
    
    
    Worksheet = Worksheets(w).Name
    Worksheets(Worksheet).Activate
    
    'Adding Column Labels
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percentage Change"
        Cells(1, 13).Value = "Total Volume"
        
        'Bonus
        Cells(2, 16).Value = "Greatest% Increase"
        Cells(3, 16).Value = "Greatest% Decrease"
        Cells(4, 16).Value = "Greatest Total Volume"
        Cells(1, 17).Value = "Ticker"
        Cells(1, 18).Value = "Value"
    
    'Determining the lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Setting the For loop to read data and running condition to get calculations
     For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                Ticker = Cells(i, 1).Value
                total_volume = total_volume + Cells(i, 7).Value
                Yearly_change = Cells(i, 6).Value - Cells(StartValue, 3).Value

                Percentage_change = (Cells(i, 6).Value - Cells(StartValue, 3).Value) / Cells(StartValue, 3).Value
                
               'Printing Values in correct columns
               
                Range("J" & summaryTablerow).Value = Ticker
                Range("M" & summaryTablerow).Value = total_volume
                Range("K" & summaryTablerow).Value = Yearly_change
                Range("L" & summaryTablerow).Value = Percentage_change
                
                'Resetting for next stock
                summaryTablerow = summaryTablerow + 1
                total_volume = 0
                StartValue = i + 1
                
        Else
                total_volume = total_volume + Cells(i, 7).Value
        End If
       
      
      'Formatting text and cells
        Range("L" & summaryTablerow).NumberFormat = "0.00%"
       
       
    
    Next i
    
    'Conditional Formatting Cells
    STlastrow = Cells(Rows.Count, 11).End(xlUp).Row
    
    For c = 2 To STlastrow
        If Cells(c, 11).Value < 0 Then
                Cells(c, 11).Interior.ColorIndex = 3
        Else
                Cells(c, 11).Interior.ColorIndex = 4
        End If
        
    Next c
    
     'Bonus Part
    Dim PercentageChangeRange As Range
    Dim TotalVolumeRange As Range
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim tickerGI As String
    Dim tickerGd As String
    Dim tickerGTV As String

    

    Set PercentageChangeRange = Worksheets(w).Range("L2", Range("L2").End(xlDown))

    Set TotalVolumeRange = Worksheets(w).Range("M2", Range("M2").End(xlDown))
    
    GreatestIncrease = Application.WorksheetFunction.Max(PercentageChangeRange)
    GreatestDecrease = Application.WorksheetFunction.Min(PercentageChangeRange)
    GreatestTotalVolume = Application.WorksheetFunction.Max(TotalVolumeRange)
    
    
    Cells(2, 18).Value = GreatestIncrease
    Cells(3, 18).Value = GreatestDecrease
    Cells(4, 18).Value = GreatestTotalVolume
    
    
    BonusLastrow = Cells(Rows.Count, 12).End(xlUp).Row
    For j = 2 To BonusLastrow
    
    
        If Cells(j, 12).Value = GreatestIncrease Then
        
        tickerGI = Cells(j, 10).Value
        
        End If
        
        
        If Cells(j, 12).Value = GreatestDecrease Then
        
        tickerGd = Cells(j, 10).Value
        
        End If
        
        If Cells(j, 13).Value = GreatestTotalVolume Then
        
        tickerGTV = Cells(j, 10).Value
        
        
        
        End If
        
        Cells(2, 17).Value = tickerGI
        Cells(3, 17).Value = tickerGd
        Cells(4, 17).Value = tickerGTV
        
        Cells(2, 18).NumberFormat = "0.00%"
        Cells(3, 18).NumberFormat = "0.00%"
        
        Next j
        
        
       
       
       
       Next w
End Sub


