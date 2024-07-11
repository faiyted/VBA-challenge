Attribute VB_Name = "Module1"
'Function for each quarter and combine to one sheet
Sub Combine()


    
    Sheets.Add.name = "Combined_Data"
    
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
    Set combined_sheet = Worksheets("Combined_Data")

    
    For Each ws In Worksheets

        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        
        lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowYear - 1) + lastRow)).Value = ws.Range("A2:K" & (lastRowYear + 1)).Value

    Next ws

   
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
   
    combined_sheet.Columns("A:G").AutoFit

End Sub

'Function run though Ticker/ Quarterly change/Percentage change/Total stock volume. Can use this fuction for each quarter sheet
Sub ticker()

    Dim tickerName As String
    Dim earliestDate As Date
    Dim earliestValC As Double
    Dim latestDate As Date
    Dim latestValF As Double
    Dim profit As Double
    Dim percentageChange As Double
    Dim totalVal As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    
    earliestDate = DateValue("01/01/9999")
    earliestValC = 0
    latestDate = DateValue("01/01/1900")
    latestValF = 0
    profit = 0
    percentageChange = 0
    totalVal = 0

    
    For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row

        
        If i = 2 Or Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            tickerName = Cells(i, 1).Value
            
            
            earliestDate = Cells(i, 2).Value
            earliestValC = Cells(i, 3).Value
            latestDate = Cells(i, 2).Value
            latestValF = Cells(i, 6).Value
            
            
            totalVal = Cells(i, 7).Value

        Else
            
            totalVal = totalVal + Cells(i, 7).Value
            
            
            If Cells(i, 2).Value < earliestDate Then
                earliestDate = Cells(i, 2).Value
                earliestValC = Cells(i, 3).Value
            End If
            
            
            If Cells(i, 2).Value > latestDate Then
                latestDate = Cells(i, 2).Value
                latestValF = Cells(i, 6).Value
            End If
        End If

       
        If i = 379501 Or Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            profit = latestValF - earliestValC
            
            
            If earliestValC <> 0 Then
                percentageChange = (latestValF - earliestValC) / earliestValC * 100
            Else
                
                percentageChange = 0
            End If
            
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Qtrly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"

            Range("I" & Summary_Table_Row).Value = tickerName
            Range("L" & Summary_Table_Row).NumberFormat = "0"
            Range("L" & Summary_Table_Row).Value = totalVal
            Range("J" & Summary_Table_Row).Value = profit
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            
            Range("K" & Summary_Table_Row).Value = percentageChange / 100
            If percentageChange / 100 > 0 Then
                
                Range("K" & Summary_Table_Row).Interior.Color = RGB(144, 238, 144)
            ElseIf percentageChange / 100 < 0 Then
                
                Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 99, 71)
            Else
                
                Range("K" & Summary_Table_Row).Interior.ColorIndex = xlNone
            End If
            
            If profit > 0 Then
                
                Range("J" & Summary_Table_Row).Interior.Color = RGB(144, 238, 144)
            ElseIf profit < 0 Then
                
                Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 99, 71)
            Else
                
                Range("J" & Summary_Table_Row).Interior.ColorIndex = xlNone
            End If
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            
            earliestDate = DateValue("01/01/9999")
            earliestValC = 0
            latestDate = DateValue("01/01/1900")
            latestValF = 0
            profit = 0
            percentageChange = 0
            totalVal = 0
        End If
        
    Next i

End Sub

Sub highVol()

Dim hightestPercent As Double
hightestPercent = 0


Dim lowPercent As Double
lowPercent = 0


Dim highestVol As Double
highestVol = 0

Dim tickerName As String
tickerName = ""

Dim vol As Double

Dim percent As Double

For i = 2 To 1501

    percent = Cells(i, "K").Value

    If percent > hightestPercent Then

    hightestPercent = percent
    tickerName = Cells(i, "I")
    
    Range("N2").Value = "Greatest % increase"
    Range("O2").Value = tickerName
    Range("P2").NumberFormat = "0.00%"
    Range("P2").Value = percent
    End If

Next i


For i = 2 To 1501

    percent = Cells(i, "K").Value

    If percent < lowPercent Then

    lowPercent = percent
    tickerName = Cells(i, "I")
    
    Range("N3").Value = "Greatest % decrease"
    Range("O3").Value = tickerName
    Range("P3").NumberFormat = "0.00%"
    Range("P3").Value = percent
    End If
Next i

For i = 2 To 1501
    vol = Cells(i, "L").Value

    If vol > highestVol Then

    highestVol = vol
    tickerName = Cells(i, "I")

    Range("N4").Value = "Greatest total volume"
    Range("O4").Value = tickerName
    Range("P4").Value = vol

    End If
Next i

End Sub


