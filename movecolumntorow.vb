Private Sub copyallsheetdata_Click()
    Dim lastrow As Long, erow As Long,mainworkbook As Workbook
    Set mainworkbook = ActiveWorkbook
    scrsheetname = "Carson City"
    destsheetname = "incode"
    lastrow = Worksheets(scrsheetname).Cells(Rows.Count, 2).End(xlUp).Row
    titleheader = Worksheets(scrsheetname).Cells(1, 2).Value
    'check last inserted row number
    erow = Worksheets(destsheetname).Cells(Rows.Count, 1).End(xlUp).Row 
    i = 0
    while i <= lastrow
        
        'copy school name and its constant
        Worksheets(scrsheetname).Cells(1, 2).Copy
        Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 1)
        
        'school code
        If Worksheets(scrsheetname).Cells(i, 1).Value Then
            Worksheets(scrsheetname).Cells(i, 1).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 2)
            'school name
            Worksheets(scrsheetname).Cells(i, 2).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 3)
            'total
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 5)
        End If
        'Female
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Gender:F" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(ierow+1, 6)
        End If
        'Male
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Gender:M" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 7)
        End If
        'Asian
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: A" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 8)
        End If
        'Black
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: B" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 9)
        End If
        'C
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: C" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 10)
        End If
        
        'Hisapnic
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: H" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 11)
        End If
        
        'I american native
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: I" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 12)
        End If
        
        'Two or more M
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: M" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 13)
        End If
        
        'Pacific
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Ethnicity: P" Then
            Worksheets(scrsheetname).Cells(i, 3).Copy
            Worksheets(scrsheetname).Paste Destination:=Worksheets(destsheetname).Cells(erow+1, 14)
        End If
        'Grade
        If Worksheets(scrsheetname).Cells(i, 2).Value = "Grade: 12" Then
            Worksheets(destsheetname).Cells(erow+1, 4) = 12
        End If
         
         'check if it need to go new line or current line 
         If Worksheets(destsheetname).Cells(erow+1, 4).Value Then
             erow = Worksheets(destsheetname).Cells(Rows.Count, 1).End(xlUp).Row
        End If

        Worksheets(destsheetname).Cells(i, 15).Value = "2015-2016"
       i = i+1  
    wend
    
End Sub
