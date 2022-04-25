Sub BOM()
'Optimize Macro Speed
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
'-------------------------------------------------------------------------'

'Declare variables
Dim bomFile As Variant
Dim openBook As Workbook

'Get next row in actual macro file
NextRowMacro = Sheets(1).Cells(Cells.Rows.Count, 2).End(xlUp).Row + 1

' Clear content
If NextRowMacro > 3 Then
    Sheets(1).Range("B3:I" & NextRowMacro).ClearContents
    'Reset to row 3
    NextRowMacro = 3
End If

'Open dialog box to get BOM Excel file
bomFile = Application.GetOpenFilename(Title:="Choose the SystemairCAD BOM file to open", FileFilter:="Excel Files (*.xls*), *xls*", MultiSelect:=False)

If bomFile <> False Then
    'Open BOM file
    Set bomBook = Application.Workbooks.Open(bomFile)
    
    'Get number of ahus
    no_ahus = bomBook.Sheets.Count
    
    'Go sheet by sheet and get the data
    For i = 1 To no_ahus
        'Get name of sheet
        sheet_name = bomBook.Sheets(i).Name
    
        'Get last row of BOM file
        LastRowBOM = bomBook.Sheets(i).Cells(Cells.Rows.Count, 2).End(xlUp).Row
        no_items = LastRowBOM - 5
        
        'Copy data from BOM file
        bomBook.Sheets(i).Range("B6:H" & LastRowBOM).Copy
        
        'Paste data in this workbook
        ThisWorkbook.Sheets(1).Cells(NextRowMacro, 2).PasteSpecial xlPasteValues
        
        'Paste name of the AHU in column I
        ThisWorkbook.Sheets(1).Range("I" & NextRowMacro & ":I" & NextRowMacro + no_items - 1) = sheet_name
        
        'Reset NextRowMacro
        NextRowMacro = NextRowMacro + no_items
        
    Next i
    
    'Close BOM file as usual
    bomBook.Close False
End If

'Done message
MsgBox ("Done bro!")


'-------------------------------------------------------------------------'
'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

End Sub