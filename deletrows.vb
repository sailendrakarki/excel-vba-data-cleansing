Private Sub delbtn_Click()

 Dim lastrow As Long, erow As Long, i As Integer, mainworkbook As Workbook
 Set mainworkbook = ActiveWorkbook

 For j = 6 To mainworkbook.Sheets.Count
    sheetname = mainworkbook.Sheets(j).Name
    'Debug.Print sheetname
    lastrow = Worksheets(sheetname).Cells(Rows.Count, 2).End(xlUp).Row
    i = 1
    While i <= lastrow
     If Worksheets(sheetname).Cells(i, 2).Value Like "Grade*" Then
       If Worksheets(sheetname).Cells(i, 2).Value <> "Grade: 12" Then
           'Debug.Print sheetname
           Worksheets(sheetname).Rows(i).EntireRow.Delete
       End If
     End If
     i = i + 1
    Wend
  Next
 
End Sub