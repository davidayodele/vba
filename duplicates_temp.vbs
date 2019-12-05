Sub subroutine8()    
	Dim count    
	Dim count2        
	For i = 1 To 55        
		Sheets("Form Responses 1").Cells([i], [7]).Value = Sheets("Form Responses 1").Cells([4], [5]).Value        
		Sheets("Form Responses 1").Cells([i], [7]).ClearContents        Sheets("Form Responses 1").Cells([i], [7]).Value = Sheets("Form Responses 1").Cells([i], [5]).Value                
		Sheets("Form Responses 1").Cells([i], [8]).Value = Sheets("Form Responses 1").Cells([4], [5]).Value        
		Sheets("Form Responses 1").Cells([i], [8]).ClearContents                
		Sheets("Form Responses 1").Cells([i], [10]).Value = Sheets("Form Responses 1").Cells([4], [5]).Value        
		Sheets("Form Responses 1").Cells([i], [10]).ClearContents    
	Next i        

	'============= We need to remove empty cells in order to plot/evaluate the data    
	count = 1    
	For j = 2 To 55        
		If Not IsEmpty(Sheets("Form Responses 1").Cells([j], [7])) Then            
			count = count + 1            
			Sheets("Form Responses 1").Cells([count], [8]).Value = Sheets("Form Responses 1").Cells([j], [7]).Value            
			Sheets("Form Responses 1").Cells([count], [10]).Value = Sheets("Form Responses 1").Cells([j], [7]).Value        
		End If    
	Next j     
	   
	'============ Creating a tally list (removing duplicates) ===========    
	LastRow_in_ColumnJ = Sheets("Form Responses 1").Range("J99").End(xlUp).Row    
	Sheets("Form Responses 1").Range("$J$1:$J$" & LastRow_in_ColumnJ).RemoveDuplicates Columns:=1, Header:=xlYes        

	'Autofit columns    
	Sheets("Form Responses 1").Columns(10).AutoFit    
	Sheets("Form Responses 1").Columns(8).AutoFit        

	count2 = 1    
	For k = 1 To 55        
		If Not IsEmpty(Sheets("Form Responses 1").Cells([k], [10])) Then            
			count2 = count2 + 1        
		End If    
	Next k            

	'We now create our tally    
	For m = 2 To count        
		For n = 2 To count2            
			If Sheets("Form Responses 1").Cells([n], [10]).Value = Sheets("Form Responses 1").Cells([m], [8]).Value Then                
				Sheets("Form Responses 1").Cells([n], [12]).Value = Sheets("Form Responses 1").Cells([n], [12]).Value + 1            
			End If        
		Next n    
	Next m            

	'Now we plot    
	Dim chart1 As Chart    
	Set chart1 = Charts.Add    
	Set src1 = chart1.SeriesCollection.NewSeries        
	For p = 1 To count2        
		Sheets("Sheet1").Cells([p], [2]).Value = Sheets("Form Responses 1").Cells([p], [10]).Value        
		Sheets("Sheet1").Cells([p], [3]).Value = Sheets("Form Responses 1").Cells([p], [12]).Value    
	Next p        
	Sheets("Sheet1").Cells([1], [2]).Value = "Birthday"    
	Sheets("Sheet1").Cells([1], [3]).Value = "Tally"        
	src1.Name = "Birthdays"    
	src1.XValues = Range(Sheets("Sheet1").Range("B2"), Sheets("Sheet1").Range("B50").End(xlUp))    
	src1.Values = Range(Sheets("Sheet1").Range("C2"), Sheets("Sheet1").Range("C50").End(xlUp))        
	chart1.HasTitle = True    chart1.ChartTitle.Text = "Tally vs Birthday"        
	chart1.Axes(xlCategory, xlPrimary).HasTitle = True    
	chart1.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Birthday"        
	chart1.Axes(xlValue, xlPrimary).HasTitle = True    
	chart1.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Tally"        
	chart1.ChartType = xlColumnClustered

End Sub
