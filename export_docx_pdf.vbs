Sub temp1()
' Sanitizes data with removal of empty rows

Dim ws_orig As Worksheet
Dim ws_new As Worksheet
Set ws_orig = Sheets("AZ Active 20100818")
Set ws_new = Sheets("cleaned")

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer

Worksheets("cleaned").Cells([1], [1]) = Worksheets("AZ Active 20100818").Cells([1], [3])
Worksheets("cleaned").Cells([1], [2]) = Worksheets("AZ Active 20100818").Cells([1], [4])
Worksheets("cleaned").Cells([1], [3]) = Worksheets("AZ Active 20100818").Cells([1], [14])
Worksheets("cleaned").Cells([1], [4]) = Worksheets("AZ Active 20100818").Cells([1], [17])

j = 1
For i = 1 To 1451
  If WorksheetFunction.CountA(ws_orig.Range("Q" & i)) <> 0 Then
    If ((Worksheets("AZ Active 20100818").Cells([i], [14]) = "AZ000") Or (Worksheets("AZ Active 20100818").Cells([i], [14]) = "AZ0000") Or (Worksheets("AZ Active 20100818").Cells([i], [14]) = "AZ0140") Or (Worksheets("AZ Active 20100818").Cells([i], [14]) = "140")) Then
          
          Worksheets("cleaned").Cells([j], [1]) = Worksheets("AZ Active 20100818").Cells([i], [3])
          Worksheets("cleaned").Cells([j], [2]) = Worksheets("AZ Active 20100818").Cells([i], [4])
          
          Worksheets("cleaned").Cells([j], [3]) = Worksheets("AZ Active 20100818").Cells([i], [14])
          Worksheets("cleaned").Cells([j], [4]) = Worksheets("AZ Active 20100818").Cells([i], [17])
          If (ws_orig.Range("Q" & i).Interior.Color = RGB(255, 199, 206)) Then
              ws_new.Range("D" & j).Interior.Color = RGB(255, 192, 0)
          End If
    j = j + 1
    End If
  End If
Next i

' Worksheets("03-30-2015").Cells([j], [5]) = "0" + Worksheets("03-30-2015").Cells([j], [5])
  'End If
  
' ws_orig.CountA(Range("E1:E1451")) <> 0

End Sub

----------------------------------------------------------

Sub temp2()

'excel vba script to export data to word and pdf
' enable MS Word Object lib: tools -> references -> check Microsoft Word 12.0 Object Library
'or
'ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files\Common Files\Microsoft 'Shared\OFFICE14\MSO.DLL"
'To list references (paths) in Excel: 
'Dim ref As Reference
'For Each ref In ActiveWorkbook.VBProject.References
'    Debug.Print ref.Description; " -- "; ref.FullPath
'Next

'Result:
'Microsoft Excel 14.0 Object Library -- C:\Program Files\Microsoft Office\Office14\EXCEL.EXE
'OLE Automation -- C:\Windows\system32\stdole2.tlb
'Microsoft Forms 2.0 Object Library -- C:\Windows\system32\FM20.DLL

Dim i As Integer
Dim r As Integer
Dim wordApp As Word.Application
Dim wordDoc As Word.Document
Dim fileName1 As String
Dim fileName2 As String

Set wordApp = CreateObject("word.application")
wordApp.Visible = True

Set wordDoc = wordApp.Documents.Add

r = 3   'r = row, change row to change file/candidate

'Worksheets("Sheet1").Cells([1], [1]) = fileName 
For i = 1 To 66
    fileName1 = "C:\Users\Public\Desktop\survey_response_" & Cells(r, 3) & ".docx"
    fileName2 = "C:\Users\Public\Desktop\survey_response_" & Cells(r, 3) & ".pdf"
    wordDoc.Content.InsertAfter Cells([1], [i]) & vbCrLf
    wordDoc.Content.InsertAfter Cells([r], [i]) & vbCrLf & vbCrLf & vbCrLf
Next i

wordDoc.ExportAsFixedFormat OutputFileName:=fileName2, ExportFormat:=wdExportFormatPDF
wordDoc.SaveAs (fileName1)
wordApp.Quit

Set wordDoc = Nothing
Set wordApp = Nothing

End Sub