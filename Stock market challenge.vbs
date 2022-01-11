Sub StockMarketData()
	
	For Each ws In Worksheets
		
		Dim WorkSheetName As String
		
		Dim i As Long
		
		Dim j As Long
		
		Dim TickerCount As Long
		
		Dim LastRowA As Long
		
		Dim LastRowI As Long
		
		Dim PerChange As Double
		
		Dim GreatIncr As Double
		
		Dim GreatDecr As Double
		
		Dim GreatVol As Double
		
		WorkSheetName = ws.Name
		
		ws.Cells(1, 9).Value = "Ticker"
		
		ws.Cells(1, 10).Value = "Yearly Change"
		
		ws.Cells(1, 11).Value = "Percent Change"
		
		ws.Cells(1, 12).Value = "Total Stock Volume"
		
		ws.Cells(1, 17).Value = "Value"
		
		ws.Cells(2, 15).Value = "Greatest % Increase"
		
		ws.Cells(3, 15).Value = "Greatest % Decrease"
		
		ws.Cells(4, 15).Value = "Greatest Total Volume"
		
		TickerCount = 2
		
		j = 2
		
		LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
		
		For i = 2 To LastRowA
			
			If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
				
				ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
				
				ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
				
				If ws.Cells(TickerCount, 10).Value < 0 Then
					
					
					ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
					
				Else
					
					
					ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
					
				End If
				
				
				If ws.Cells(j, 3).Value <> 0 Then
					
					PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
					
					
					ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
					
				Else
					
					ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
					
				End If
				
				
				ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
				
				
				TickerCount = TickerCount + 1
				
				
				j = i + 1
				
			End If
			
			Next i
			
			
			LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
			
			
			
			GreatVol = ws.Cells(2, 12).Value
			GreatIncr = ws.Cells(2, 11).Value
			GreatDecr = ws.Cells(2, 11).Value
			
			
			For i = 2 To LastRowI
				
				
				If ws.Cells(i, 12).Value > GreatVol Then
					GreatVol = ws.Cells(i, 12).Value
					ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
					
				Else
					
					GreatVol = GreatVol
					
				End If
				
				
				If ws.Cells(i, 11).Value > GreatIncr Then
					GreatIncr = ws.Cells(i, 11).Value
					ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
					
				Else
					
					GreatIncr = GreatIncr
					
				End If
				
				
				If ws.Cells(i, 11).Value < GreatDecr Then
					GreatDecr = ws.Cells(i, 11).Value
					ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
					
				Else
					
					GreatDecr = GreatDecr
					
				End If
				
				
				ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
				ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
				ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
				
				Next i
				
				
				Worksheets(WorkSheetName).Columns("A:Z").AutoFit
				
				Next ws
				
			End Sub
			
