Sub Stonks():

'Create the Loop through each Worksheet
    Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets
    
'Give the new columns their headers
        WS.Range("I1").Value = "Ticker"
        WS.Range("J1").Value = "Yearly Change"
        WS.Range("K1").Value = "Percent Change"
        WS.Range("L1").Value = "Total Stock Volume"
        
'Call the Ticker from each sheet
        Dim Ticker As String
    
        i = 2
        j = 2
        Do Until IsEmpty(WS.Cells(i, 1).Value)
        
            If (WS.Cells(i, 1).Value = WS.Cells((i + 1), 1).Value) Then
                i = i + 1
            
            Else
                WS.Cells(j, 9).Value = WS.Cells(i, 1).Value
                i = i + 1
                j = j + 1
            
            End If
        
        Loop
 
 'Look for the yearly change for each Ticker
        Dim Yearly As Double
        Yearly = 0
 
        i = 2
        j = 2
        Yearly = WS.Cells(i, 3).Value
    
        Do Until IsEmpty(WS.Cells(i, 1).Value)
    
            If (WS.Cells(i, 1).Value = WS.Cells((i + 1), 1).Value) Then
                i = i + 1
            
            Else
                Yearly = WS.Cells(i, 6).Value - Yearly
                WS.Cells(j, 10).Value = Yearly
                i = i + 1
                j = j + 1
                Yearly = WS.Cells(i, 3).Value
            
            End If
    
        Loop
    
'Colour the Yearly Change values by positive or negative changes
        i = 2
        Do Until IsEmpty(WS.Cells(i, 10).Value)
    
            If (WS.Cells(i, 10).Value >= 0) Then
                WS.Cells(i, 10).Interior.ColorIndex = 4
                i = i + 1
            
            Else
                WS.Cells(i, 10).Interior.ColorIndex = 3
                i = i + 1
        
            End If
    
        Loop
    
'Calculate the Percent Change and paste the value in the Percent Change column.
        Dim Percent As Double
        Percent = 0
 
        i = 2
        j = 2
        Percent = WS.Cells(i, 3).Value
    
        Do Until IsEmpty(WS.Cells(i, 1).Value)
    
            If (WS.Cells(i, 1).Value = WS.Cells((i + 1), 1).Value) Then
                i = i + 1
            
            Else
                Percent = (WS.Cells(i, 6).Value - Percent) / Percent
                WS.Cells(j, 11).Value = Percent
                i = i + 1
                j = j + 1
                Percent = WS.Cells(i, 3).Value
            
            End If
    
        Loop
    
'Format percentages in Percent Change column
        i = 2
    
        Do Until IsEmpty(WS.Cells(i, 11).Value)
            WS.Cells(i, 11).Value = FormatPercent(WS.Cells(i, 11).Value)
            i = i + 1
        Loop
    
'Compile the total volume per stock and past in the Total Stock Volume column
        Dim Volume As Double
 
        i = 2
        j = 2
        Volume = 0
    
        Do Until IsEmpty(WS.Cells(i, 1).Value)
    
            If (WS.Cells(i, 1).Value = WS.Cells((i + 1), 1).Value) Then
                Volume = Volume + WS.Cells(i, 7).Value
                i = i + 1
            
            Else
                Volume = Volume + WS.Cells(i, 7).Value
                WS.Cells(j, 12).Value = Volume
                i = i + 1
                j = j + 1
                Volume = 0
            
            End If
    
        Loop
        
'Create Columns for Ticker & Value
        WS.Range("P1").Value = "Ticker"
        WS.Range("Q1").Value = "Value"
        
'Create the Rows for Greatest % Increase, Decrease and Total Volume
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        
'Find the value for Greatest Increase
        i = 2
        GreatestPercent = WS.Cells(i, 10).Value
        
        Do Until IsEmpty(WS.Cells(i, 10).Value)
    
            If (WS.Cells(i, 10).Value <= GreatestPercent) Then
                i = i + 1
            
            Else
                GreatestPercent = WS.Cells(i, 10).Value
                Ticker = WS.Cells(i, 9).Value
                i = i + 1
            
            End If
    
        Loop
        
        WS.Range("P2").Value = Ticker
        WS.Range("Q2").Value = GreatestPercent
        
'Find the value for Greatest Decrease
        i = 2
        GreatestDecrease = WS.Cells(i, 10).Value
        
        Do Until IsEmpty(WS.Cells(i, 10).Value)
    
            If (WS.Cells(i, 10).Value >= GreatestDecrease) Then
                i = i + 1
            
            Else
                GreatestDecrease = WS.Cells(i, 10).Value
                Ticker = WS.Cells(i, 9).Value
                i = i + 1
            
            End If
    
        Loop
        
        WS.Range("P3").Value = Ticker
        WS.Range("Q3").Value = GreatestDecrease
        
'Find the value for Greatest Total Volume
        i = 2
        GreatestVolume = WS.Cells(i, 12).Value
        
        Do Until IsEmpty(WS.Cells(i, 12).Value)
    
            If (WS.Cells(i, 12).Value <= GreatestVolume) Then
                i = i + 1
            
            Else
                GreatestVolume = WS.Cells(i, 12).Value
                Ticker = WS.Cells(i, 9).Value
                i = i + 1
            
            End If
    
        Loop
        
        WS.Range("P4").Value = Ticker
        WS.Range("Q4").Value = GreatestVolume
       
    Next WS
       
End Sub