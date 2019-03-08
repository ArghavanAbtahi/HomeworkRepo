Attribute VB_Name = "Module1"
Sub Calculate()
    'repeating same steps for each worksheet
    For Each ws In Worksheets
        ws.Select
            Cells(1, 11).Value = "Ticker"
            Cells(1, 12).Value = "Yearly Change"
            Cells(1, 13).Value = "Percent Change"
            Cells(1, 14).Value = "Total Stock Volume"

            Dim TicketName As String
            Dim TickerVolume As Double
            TickerVolume = 0
            Dim opn As Double
            Dim cls As Double
            Dim difference As Double
            Dim Summary As Integer
            Summary = 2
            Dim percent As Double
                        
            'Define Last Row
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            'Change format in M2 column to percent
            Range("M2").EntireColumn.NumberFormat = "0.000%"
            
            For i = 2 To lastrow
                'getting the name and volume
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    TickerName = Cells(i, 1).Value
                    TickerVolume = Cells(i, 7).Value + TickerVolume
                    Cells(Summary, 11).Value = TickerName
                    Cells(Summary, 14).Value = TickerVolume
                Else: TickerVolume = Cells(i, 7).Value + TickerVolume
                End If
                'getting the difference
                If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                    opn = Cells(i, 3).Value
                ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    cls = Cells(i, 6).Value
                    difference = cls - opn
                    Cells(Summary, 12).Value = difference
                'getting the percent difference
                If opn = 0 Then
                    Cells(Summary, 13).Value = "NA"
                Else: percent = difference / opn
                    Cells(Summary, 13).Value = percent
                End If
                'changing column colors
                If Cells(Summary, 12).Value > 0 Then
                    Cells(Summary, 12).Interior.ColorIndex = 4
                Else: Cells(Summary, 12).Interior.ColorIndex = 3
                End If
                'resetting values
                Summary = Summary + 1
                TickerVolume = 0
                cls = 0
                opn = 0
                End If
            Next i
            
            'finding max increase, decrease and volume
            Dim rng As Range
            Dim rng2 As Range
            Dim max As Double
            Dim min As Double
            Dim maxvolume As Double
            Dim maxticker As String
            Dim minticker As String
            Dim volumeticker As String
                
            Cells(2, 16).Value = "Greatest % Increase"
            Cells(3, 16).Value = "Greatest % Decrease"
            Cells(4, 16).Value = "Greatest total volume"
            Cells(1, 17).Value = "Ticker"
            Cells(1, 18).Value = "Value"
            Cells(2, 18).NumberFormat = "0.000%"
            Cells(3, 18).NumberFormat = "0.000%"
            
            'finding max percent increase and decrease
            Set rng = Range("M2").EntireColumn
                max = Application.WorksheetFunction.max(rng)
                Cells(2, 18).Value = max
                min = Application.WorksheetFunction.min(rng)
                Cells(3, 18).Value = min
            
            'finding max volume
            Set rng2 = Range("N2").EntireColumn
                maxvolume = Application.WorksheetFunction.max(rng2)
                Cells(4, 18).Value = maxvolume
            
            'finding and copying ticker symbol of maxes
            For i = 2 To lastrow
                If Cells(i, 13).Value = max Then
                    maxticker = Cells(i, 11).Value
                    Cells(2, 17).Value = maxticker
                ElseIf Cells(i, 13).Value = min Then
                    minticker = Cells(i, 11).Value
                    Cells(3, 17).Value = minticker
                End If
                If Cells(i, 14).Value = maxvolume Then
                    volumeticker = Cells(i, 11).Value
                    Cells(4, 17).Value = volumeticker
                End If
            Next i
        Next ws
End Sub
