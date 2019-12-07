'=================== Script - full - working (updated 3/18/18) ==========================
Sub temp2()     
'Const fileName1 As String = "TSC OR Schedule.docx"    
'excel vba script to export data to word and pdf    
' enable MS Word Object lib: tools -> references -> check Microsoft Word 12.0 Object Library    
'or    'ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files\Common Files\Microsoft 'Shared\OFFICE14\MSO.DLL"    
'To list references (paths) in Excel:    'Dim ref As Reference    
'For Each ref In ActiveWorkbook.VBProject.References    
'    Debug.Print ref.Description; " -- "; ref.FullPath    
'Next
    'Result:    
'Microsoft Excel 14.0 Object Library -- C:\Program Files\Microsoft Office\Office14\EXCEL.EXE    
'OLE Automation -- C:\Windows\system32\stdole2.tlb    'Microsoft Forms 2.0 Object Library -- C:\Windows\system32\FM20.DLL
    'Dim ref As Reference
    'Word objects
    
Dim wdApp As Word.Application    
Dim wdDoc As Word.Document    
Dim wdbmRange As Word.Range
        
'Name Word and PDF docs    
Dim fileName1 As String    
Dim fileName2 As String    
fileName1 = "tsc_sched_" & Replace(Trim(Split(Worksheets("tsc_sched_3-6-18_csv").Cells([1], [3]), "--")(0)), "/", "-") & ".docx"    
fileName2 = "tsc_sched_" & Replace(Trim(Split(Worksheets("tsc_sched_3-6-18_csv").Cells([1], [3]), "--")(0)), "/", "-") & ".pdf"        
'Declare Excel objects    
Dim wbBook As Workbook    
Dim wsSheet As Worksheet   
Dim xlSheet As Range
        
'==========  Begin Data extraction =============    
Dim f, g, h, i, j, k, m, n, r, count, count2, count3, count4 As Integer    
Dim s, s2, s3, s4, s5, t, sCellVal  As String    
Dim rCell, rFilled, rSel As Range

    'sCellVal = LCase$(Range("C1").Value (to make case sensitive)    
'sCellVal = Sheets("Sheet1").Range("A1").Value
    'Set sheet font (to print on 11x8)    
With ActiveSheet        
	.Cells.Font.Size = "7"        
	.Cells.Font.Name = "Arial"    
End With
    'Autofit cols of interest (obscured data will be pulled otherwise)    
Worksheets("tsc_sched_3-6-18_csv").Columns(21).AutoFit    
Worksheets("tsc_sched_3-6-18_csv").Columns(26).AutoFit    
Worksheets("tsc_sched_3-6-18_csv").Columns(27).AutoFit    
Worksheets("tsc_sched_3-6-18_csv").Columns(28).AutoFit    
Worksheets("tsc_sched_3-6-18_csv").Columns(29).AutoFit

count = 1    
count2 = 1    
count3 = 1    
count4 = 1    

For i = 1 To 50    
	j, k, m, n = 1
	s = ""    
	s = Worksheets("tsc_sched_3-6-18_csv").Cells([i], [19]).Text
    ' "Like" Case sensitive at begining of string    
	' Last If executes first (call stack) => cannot see vars set by above Ifs        
	If s Like "Room: TSC OR 01*" Then            
		count = count + 1            
		Sheets("Sheet4").Cells([j], [2]).Value = Trim(Split(s, ":")(1)) 'Or Rtrim/LTrim            
		Sheets("Sheet4").Cells([j], [2]).Font.Bold = True            
		Sheets("Sheet4").Cells([j], [2]).Interior.ColorIndex = 36            
		Sheets("Sheet4").Columns(2).AutoFit                        
		Do                
			f = i + j + 1                
			t = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [19]).Text                
			j = j + 1                                
			Sheets("Sheet4").Cells([j], [1]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [21]).Text                
			Sheets("Sheet4").Cells([j], [1]).WrapText = True                
			Sheets("Sheet4").Columns(1).AutoFit                
			Sheets("Sheet4").Cells([j], [2]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [26]).Text                
			Sheets("Sheet4").Cells([j], [2]).WrapText = True                
			Sheets("Sheet4").Cells([j], [2]).Interior.ColorIndex = 36                
			Sheets("Sheet4").Columns(2).AutoFit                
			Sheets("Sheet4").Cells([j], [3]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [27]).Text                
			Sheets("Sheet4").Cells([j], [3]).WrapText = True                
			Sheets("Sheet4").Columns(3).AutoFit                
			Sheets("Sheet4").Cells([j], [4]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [28]).Text                
			Sheets("Sheet4").Cells([j], [4]).WrapText = True                
			Sheets("Sheet4").Columns(4).AutoFit                
			Sheets("Sheet4").Cells([j], [5]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [29]).Text                
			Sheets("Sheet4").Cells([j], [5]).WrapText = True                
			Sheets("Sheet4").Columns(5).AutoFit                
			Sheets("Sheet4").Cells([j], [6]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([j], [31]).Text                
			Sheets("Sheet4").Cells([j], [6]).WrapText = True                
			Sheets("Sheet4").Columns(6).AutoFit
			            
		Loop Until t Like "Room: TSC OR 02*"            
		s = ""            
		Sheets("Sheet4").Columns(7).ColumnWidth = 0            
		Sheets("Sheet4").Columns(8).ColumnWidth = 0        
	End If
	                    
	If s Like "Room: TSC OR 02*" And Not (s Like "Room: TSC OR 01*") Then            
		count2 = count2 + 1            
		Sheets("Sheet4").Cells([k], [10]).Value = Trim(Split(s, ":")(1))            
		Sheets("Sheet4").Cells([k], [10]).Font.Bold = True            
		Sheets("Sheet4").Cells([k], [10]).Interior.ColorIndex = 37            
		Sheets("Sheet4").Columns(10).AutoFit                        
		Do                
			f = count + k - 1  'i + k                
			k = k + 1                
			t = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [19]).Text                
			Sheets("Sheet4").Cells([k], [9]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [21]).Text                
			Sheets("Sheet4").Cells([k], [9]).WrapText = True                
			Sheets("Sheet4").Columns(9).AutoFit                
			Sheets("Sheet4").Cells([k], [10]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [26]).Text                
			Sheets("Sheet4").Cells([k], [10]).WrapText = True                
			Sheets("Sheet4").Cells([k], [10]).Interior.ColorIndex = 37                
			Sheets("Sheet4").Columns(10).AutoFit                
			Sheets("Sheet4").Cells([k], [11]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [27]).Text                
			Sheets("Sheet4").Cells([k], [11]).WrapText = True                
			Sheets("Sheet4").Columns(11).AutoFit                
			Sheets("Sheet4").Cells([k], [12]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [28]).Text                
			Sheets("Sheet4").Cells([k], [12]).WrapText = True                
			Sheets("Sheet4").Columns(12).AutoFit                
			Sheets("Sheet4").Cells([k], [13]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [29]).Text                
			Sheets("Sheet4").Cells([k], [13]).WrapText = True                
			Sheets("Sheet4").Columns(13).AutoFit                
			Sheets("Sheet4").Cells([k], [14]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [31]).Text                
			Sheets("Sheet4").Cells([k], [14]).WrapText = True                
			Sheets("Sheet4").Columns(14).AutoFit            
		Loop Until t Like "Room: TSC OR 03*"            
		s = ""            
		Sheets("Sheet4").Columns(15).ColumnWidth = 0            
		Sheets("Sheet4").Columns(16).ColumnWidth = 0        
	End If                
	If s Like "Room: TSC OR 03*" Then            
		count3 = count3 + 1            
		Sheets("Sheet4").Cells([m], [18]).Value = Trim(Split(s, ":")(1))            
		Sheets("Sheet4").Cells([m], [18]).Font.Bold = True            
		Sheets("Sheet4").Cells([m], [18]).Interior.ColorIndex = 39            
		Sheets("Sheet4").Columns(18).AutoFit                        
		Do                
			f = count + count2 + m - 2                
			t = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [19]).Text                
			m = m + 1                
			Sheets("Sheet4").Cells([m], [17]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [21]).Text                
			Sheets("Sheet4").Cells([m], [17]).WrapText = True                
			Sheets("Sheet4").Columns(17).AutoFit                
			Sheets("Sheet4").Cells([m], [18]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [26]).Text                
			Sheets("Sheet4").Cells([m], [18]).WrapText = True                
			Sheets("Sheet4").Cells([m], [18]).Interior.ColorIndex = 39                
			Sheets("Sheet4").Columns(18).AutoFit                
			Sheets("Sheet4").Cells([m], [19]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [27]).Text                
			Sheets("Sheet4").Cells([m], [19]).WrapText = True                
			Sheets("Sheet4").Columns(19).AutoFit                
			Sheets("Sheet4").Cells([m], [20]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [28]).Text                
			Sheets("Sheet4").Cells([m], [20]).WrapText = True                
			Sheets("Sheet4").Columns(20).AutoFit                
			Sheets("Sheet4").Cells([m], [21]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [29]).Text                
			Sheets("Sheet4").Cells([m], [21]).WrapText = True                
			Sheets("Sheet4").Columns(21).AutoFit                
			Sheets("Sheet4").Cells([m], [22]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [31]).Text                
			Sheets("Sheet4").Cells([m], [22]).WrapText = True                
			Sheets("Sheet4").Columns(22).AutoFit            
		Loop Until t Like "Room: TSC OR 04*"            
		s = ""            
		Sheets("Sheet4").Columns(23).ColumnWidth = 0            
		Sheets("Sheet4").Columns(24).ColumnWidth = 0        
	End If                
	If s Like "Room: TSC OR 04*" Then            
		count4 = count4 + 1            
		Sheets("Sheet4").Cells([n], [26]).Value = Trim(Split(s, ":")(1))            
		Sheets("Sheet4").Cells([n], [26]).Font.Bold = True            
		Sheets("Sheet4").Cells([n], [26]).Interior.ColorIndex = 40            
		Sheets("Sheet4").Columns(26).AutoFit            
		Do                
			f = count + count2 + count3 + n - 3                
			t = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [19]).Text                
			n = n + 1                
			Sheets("Sheet4").Cells([n], [25]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [21]).Text                
			Sheets("Sheet4").Cells([n], [25]).WrapText = True                
			Sheets("Sheet4").Columns(25).AutoFit                
			Sheets("Sheet4").Cells([n], [26]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [26]).Text                
			Sheets("Sheet4").Cells([n], [26]).WrapText = True                
			Sheets("Sheet4").Cells([n], [26]).Interior.ColorIndex = 40                
			Sheets("Sheet4").Columns(26).AutoFit                
			Sheets("Sheet4").Cells([n], [27]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [27]).Text                
			Sheets("Sheet4").Cells([n], [27]).WrapText = True                
			Sheets("Sheet4").Columns(27).AutoFit                
			Sheets("Sheet4").Cells([n], [28]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [28]).Text                
			Sheets("Sheet4").Cells([n], [28]).WrapText = True                
			Sheets("Sheet4").Columns(28).AutoFit                
			Sheets("Sheet4").Cells([n], [29]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [29]).Text                
			Sheets("Sheet4").Cells([n], [29]).WrapText = True                
			Sheets("Sheet4").Columns(29).AutoFit                
			Sheets("Sheet4").Cells([n], [30]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [31]).Text                
			Sheets("Sheet4").Cells([n], [30]).WrapText = True                
			Sheets("Sheet4").Columns(30).AutoFit            
		Loop Until t Like "Room: TSC OR 05*"            
		s = ""            
		Sheets("Sheet4").Columns(31).ColumnWidth = 0            
		Sheets("Sheet4").Columns(32).ColumnWidth = 0        
	End If                
	r = 1        
	If s Like "Room: TSC OR 05*" Then            
		Sheets("Sheet4").Cells([r], [34]).Value = Trim(Split(s, ":")(1))            
		Sheets("Sheet4").Cells([r], [34]).Font.Bold = True            
		Sheets("Sheet4").Cells([r], [34]).Interior.ColorIndex = 35            
		Sheets("Sheet4").Columns(34).AutoFit            
		While t Like "Room: TSC OR 05*"                
			f = count + count2 + count3 + count4 + r - 4                
			t = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [19]).Text                
			r = r + 1                
			Sheets("Sheet4").Cells([r], [33]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [21]).Text                
			Sheets("Sheet4").Cells([r], [33]).WrapText = True                
			Sheets("Sheet4").Columns(33).AutoFit                
			Sheets("Sheet4").Cells([r], [34]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [26]).Text                
			Sheets("Sheet4").Cells([r], [34]).WrapText = True                
			Sheets("Sheet4").Cells([r], [34]).Interior.ColorIndex = 35                
			Sheets("Sheet4").Columns(34).AutoFit                
			Sheets("Sheet4").Cells([r], [35]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [27]).Text                
			Sheets("Sheet4").Cells([r], [35]).WrapText = True                
			Sheets("Sheet4").Columns(35).AutoFit                
			Sheets("Sheet4").Cells([r], [36]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [28]).Text                
			Sheets("Sheet4").Cells([r], [36]).WrapText = True                
			Sheets("Sheet4").Columns(36).AutoFit                
			Sheets("Sheet4").Cells([r], [37]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [29]).Text                
			Sheets("Sheet4").Cells([r], [37]).WrapText = True                
			Sheets("Sheet4").Columns(37).AutoFit                
			Sheets("Sheet4").Cells([r], [38]).Value = Worksheets("tsc_sched_3-6-18_csv").Cells([f], [31]).Text                
			Sheets("Sheet4").Cells([r], [38]).WrapText = True                
			Sheets("Sheet4").Columns(38).AutoFit            
		Wend        
	End If            
Next i    '==========  End Data extraction =============        
'Initialize Excel objects, use only non-empty cells    
Set wbBook = ThisWorkbook    
Set wsSheet = wbBook.Worksheets("Sheet4")    
'lastRow = ActiveSheet.Cells(200, 39).End(xlUp).Row    
lastRow = ActiveSheet.Range("A" & Rows.count).End(xlUp).Row  'Range will determine size in wdDoc    
'Set rTable = Range("A1:AZ" & lastRow)    
'Set rSel = Nothing    
'For Each rCell In rTable    
'    If rCell.Value <> "" Then    
'        If rSel Is Nothing Then    
'            Set rSel = rCell    
'        Else    
'            Set rSel = Union(rSel, rCell)    
'        End If    
'    End If    
'Next rCell    
'If Not rSel Is Nothing Then rSel.Select        
Set xlSheet = wsSheet.Range(Cells.Address)    
'Set xlSheet = ActiveSheet.UsedRange    
'Set xlSheet = wsSheet.Range("A1:AM" & lastRow)    
'Set xlSheet = wsSheet.Range("A1:AM50")        
'Initialize Word objets    
Set wdApp = New Word.Application    
Set wdApp = CreateObject("word.application")    
wdApp.Visible = True    
'Set wdDoc = wdApp.Documents.Open(wbBook.Path &; "\" &; fileName1) 
'for preexisting docs    Set wdDoc = wdApp.Documents.Add    
'wdDoc.Activate    
wdDoc.Bookmarks.Add Name:="Report"    
Set wdbmRange = wdDoc.Bookmarks("Report").Range
    'If the macro has been run before, clean up any artifacts before trying to paste the table in again    
On Error Resume Next    
With wdDoc.InlineShapes(1)        
.Select        
.Delete    
End With    
On Error GoTo 0
'Turn off screen updating    
Application.ScreenUpdating = False
'Copy sheet to the clipboard    
xlSheet.Copy        
With wdDoc        
'.Save        
.PageSetup.Orientation = wdOrientLandscape  
'Orientation must be set to landscape BEFORE paste    
End With        
'Select the range defined by the "Report" bookmark and paste in the report from clipboard.    
With wdbmRange        
.Select        
.PasteSpecial Link:=False, _                      
DataType:=wdPasteMetafilePicture, _                      
Placement:=wdInLine, _                      
DisplayAsIcon:=False    
End With
    'Reorient doc, save as fileName1, save pdf as fileName2, print, close all    
With wdDoc        
'.Save        
'.PrintOut        
'.PageSetup.Orientation = wdOrientPortrait        
.SaveAs (fileName1)        
.ExportAsFixedFormat OutputFileName:=fileName2, ExportFormat:=wdExportFormatPDF        
.Close    
End With
    'Quit Word    
wdApp.Quit
    'Null out variables.    
Set wdbmRange = Nothing    
Set wdDoc = Nothing    
Set wdApp = Nothing
    'Clear out the clipboard, and turn screen updating back on.    
With Application        
.CutCopyMode = False        
.ScreenUpdating = True    
End With
    MsgBox "The report has successfully been " & vbNewLine & _           
"transferred and saved as " & fileName1, vbInformation
End Sub'======================= End Script ===============================