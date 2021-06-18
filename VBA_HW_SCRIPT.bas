Attribute VB_Name = "Module1"

Sub totalCalc():
    Dim yearChange(2) As Double
    Dim volume As Double
    Dim maxPer As Double
    Dim minPer As Double
    Dim maxVol As Double
    Dim maxTickVol As String
    Dim minTick As String
    Dim maxTick As String
    volume = 0
    
    'Add increment variables
    Dim counter As Integer
    Dim counter2 As Integer
    Dim counter3 As Integer
    'counter for ticker
    counter = 0
    'counter for price change and percentage
    counter2 = 0
    'counter for volume
    counter3 = 0
    
    'Get the first tricker symbol
    Cells(2, 8).Value = Cells(2, 1).Value
    counter = counter + 1
    
    'Get the first start price of the first stock
    yearChange(0) = Cells(2, 3).Value
    
    'This loop gets all the info needed in excel sheet
    For i = 3 To Rows.Count
        'Get the volume of the day
        volume = volume + Cells(i - 1, 7).Value
        If Not (StrComp(Cells(i, 1).Value, Cells(i - 1, 1).Value) = 0) Then
            'Put total volume into the excel sheet
            Cells(counter3 + 2, 11).Value = volume
            'Find the max volume
            If Cells(counter3 + 2, 11).Value > maxVol Then
                maxVol = Cells(counter3 + 2, 11).Value
                maxTickVol = Cells(counter3 + 2, 8).Value
            End If
            'Reset volume variable
            volume = 0
            'increment counter
            counter3 = counter3 + 1
            'Get the end price
            yearChange(1) = Cells(i - 1, 6).Value
            'Add price change to the excel sheet
            Cells(counter2 + 2, 9).Value = yearChange(1) - yearChange(0)
            'Color the cells
            If Cells(counter2 + 2, 9).Value >= 0 Then
                Cells(counter2 + 2, 9).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(counter2 + 2, 9).Interior.Color = RGB(255, 0, 0)
            End If
            'Make sure not divided by a zero for calculating percentage
            'Add percentage onto the excel sheet
            If yearChange(0) = 0 Then
                Cells(counter2 + 2, 10).Value = 0
                Cells(counter2 + 2, 10).NumberFormat = "0.00%"
            Else
                Cells(counter2 + 2, 10).Value = Cells(counter2 + 2, 9).Value / yearChange(0)
                Cells(counter2 + 2, 10).NumberFormat = "0.00%"
            End If
            
            'Check to store max, min
            If Cells(counter2 + 2, 10).Value > maxPer Then
                maxPer = Cells(counter2 + 2, 10).Value
                maxTick = Cells(counter2 + 2, 8).Value
            ElseIf Cells(counter2 + 2, 10).Value < minPer Then
                minPer = Cells(counter2 + 2, 10).Value
                minTick = Cells(counter2 + 2, 8).Value
            End If
            
            
            'Get the new start price
            yearChange(0) = Cells(i, 3).Value
            'increment counter
            counter2 = counter2 + 1
            
            'Add ticker to the excel sheet
            Cells(counter + 2, 8).Value = Cells(i, 1).Value
            'increment counter
            counter = counter + 1
        End If
    Next i
    
    'Add in titles
    Cells(1, 8).Value = "Ticker"
    Cells(1, 9).Value = "Yearly Change"
    Cells(1, 10).Value = "Percentage Change"
    Cells(1, 11).Value = "Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % decrease"
    Cells(4, 14).Value = "Greatest Volume"
    
    Cells(2, 15).Value = maxTick
    Cells(3, 15).Value = minTick
    Cells(4, 15).Value = maxTickVol
    Cells(2, 16).Value = maxPer
    Cells(3, 16).Value = minPer
    Cells(4, 16).Value = maxVol

    
    


End Sub

