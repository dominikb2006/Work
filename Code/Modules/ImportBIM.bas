Attribute VB_Name = "ImportBIM"
Sub Test()
    
    DeleteRow
    AddTopo
    CombineSheetsWithDifferentHeaders
    'DeleteSheets

    AddColumn R_REINFORCEMENT
    AddColumn R_FORMWORK
    AddColumn R_NAME
    AddColumn R_NAME_FINAL
    AddColumn R_VOLUME2
    AddColumn R_INTERVAL_HEIGHT
    AddColumn R_FORMWORK2
    
    ReorganizeColumns
    Headers
    Filter
    DeleteRowsFromImportBIM
    DeleteUnnecessaryColumn
End Sub

Sub Main()

    DeleteRow
    AddTopo
    CombineSheetsWithDifferentHeaders
    DeleteSheets

    AddColumn R_REINFORCEMENT
    AddColumn R_FORMWORK
    AddColumn R_NAME
    AddColumn R_NAME_FINAL
    AddColumn R_VOLUME2
    AddColumn R_INTERVAL_HEIGHT
    AddColumn R_FORMWORK2
    
    ReorganizeColumns
    Headers
    Filter
    DeleteRowsFromImportBIM
    DeleteUnnecessaryColumn
End Sub

Private Sub DeleteRow()
'
'Delete second row if Cell(1,2) is empty
'Usuniecie drugiego wiersza, jezeli komorka (1,2) jest pusta
'
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If (Sheets(i).name <> A_PRICE_LIST _
            And Sheets(i).name <> A_ASSUMPTIONS _
            And Sheets(i).name <> A_MAN_HOUR _
            And Sheets(i).name <> A_PROFILES _
            And Sheets(i).name <> A_COMMENTS _
            And Sheets(i).name <> A_CALCULATION2) Then
            Sheets(i).Select
            If IsEmpty(Cells(2, 1)) Then
                Rows("2:2").Delete
            End If
        End If
    Next i
End Sub

Private Sub AddTopo()
'
'Add Names in Topograpy Workshet
'Dodanie nazw w Topografii
'
    Sheets("Topography").Select
    Dim j As Integer
    For j = 2 To Cells(Rows.Count, 1).End(xlUp).row
        Cells(j, 3) = "Topography"
        Cells(j, 4) = "Topography"
        Cells(j, 5) = "TOPO"
        Cells(j, 6) = "Topography"
        Cells(j, 8) = "ZIE"
        Cells(j, 9) = "ZIE"
    Next j
End Sub

Private Sub DeleteSheets()
'
'Delete unnessecary worksheets
'Usuwa niepotrzebe arkusze
'
        Sheets("Walls").Delete
        Sheets("Floors").Delete
        Sheets("Generic Models").Delete
        Sheets("Structural Foundations").Delete
        Sheets("Structural Columns").Delete
        Sheets("Structural Framing").Delete
        Sheets("Topography").Delete
        Sheets("Floors - Slab Edges").Delete
End Sub

Private Sub AddColumn(name As Variant)
'
'Add column
'Dodanie kolumny
'
    Columns(1).Insert
    Cells(1, 1).value = name
End Sub

Private Sub Headers()
'
'Headers in BoQ
'Naglowki
'
    'Czcionka
    Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 10
    End With
    
    'Wysrodkowanie
    Range(Cells(1, 1), Cells(1, UBound(BOQ_COLUMNS) + 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        lib.Borders
        .Interior.ColorIndex = 15
        .Font.Bold = True
    End With

    'Wypelnienie
    Range(Cells(1, BoQIndex(R_PHASE_CREATED)), Cells(1, BoQIndex(R_PROFILE))).Interior.ColorIndex = 24
    Range(Cells(1, BoQIndex(R_CPI_KEY)), Cells(1, BoQIndex(R_5D4D_CODE))).Interior.ColorIndex = 45
    Range(Cells(1, BoQIndex(R_MATERIAL)), Cells(1, BoQIndex(R_WATERPROOF))).Interior.ColorIndex = 40
    Range(Cells(1, BoQIndex(R_COUNT)), Cells(1, BoQIndex(R_PERIMETER))).Interior.ColorIndex = 43
    Range(Cells(1, BoQIndex(R_CUT)), Cells(1, BoQIndex(R_NET_CUT_FILL))).Interior.ColorIndex = 50
     
    'Grupowanie
    Range(Cells(1, BoQIndex(R_PHASE_CREATED) + 1), Cells(1, BoQIndex(R_PROFILE))).Columns.Group
    Range(Cells(1, BoQIndex(R_NAME) + 1), Cells(1, BoQIndex(R_WATERPROOF))).Columns.Group
    Range(Cells(1, BoQIndex(R_FOUND_THICKNESS) + 1), Cells(1, BoQIndex(R_SLOPE))).Columns.Group
    Range(Cells(1, BoQIndex(R_CUT) + 1), Cells(1, BoQIndex(R_NET_CUT_FILL))).Columns.Group
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlBelow
        .SummaryColumn = xlLeft
    End With
End Sub

Private Sub Filter()
'
'Filtr + Autofit
'
    'Filtr
    Cells.AutoFilter
    'Autofit
    Cells.EntireColumn.AutoFit
End Sub

Private Sub DeleteRowsFromImportBIM()
'
'Usuniecie wierszy z zawartoœci¹ nag³ówka
'
    Sheets(A_IMPORT_BIM).Select
    Dim j As Integer
    j = 2
    Do While j <= Cells(Rows.Count, 1).End(xlUp).row
        If (Cells(j, 1).value = "Phase Created" _
            Or Cells(j, 1).value = "") Then
            Rows(j).Delete
            GoTo continueDO
        End If
        j = j + 1
continueDO:
    Loop
'    For j = 2 To Cells(Rows.Count, 1).End(xlUp).row
'        If (Cells(j, 1).VALUE = "Phase Created" _
'            Or Cells(j, 1).VALUE = "") Then
'            Rows(j).Delete
'            continue For
'        End If
'    Next j
End Sub

Private Sub DeleteUnnecessaryColumn()
'
'Delete unnecessary columns
'
    ThisWorkbook.Worksheets(A_IMPORT_BIM).Range(Columns(UBound(BOQ_COLUMNS) + 2), Columns(Columns.Count)).Select
    Selection.Delete Shift:=xlToLeft
End Sub

Private Sub ReorganizeColumns()
'
'Reorganize columns in Excel based on column header
'Reorganizacja kolumn
'
    ' Reorganize Columns Macro
    '
    ' Developer: If you want to know, please contact Winko Erades van den Berg
    ' E-mail : winko at winko-erades.nl
    ' Developed: 11-11-2013
    ' Modified: 11-11-2013
    ' Version: 1.0
    '
    ' Description: Reorganize columns in Excel based on column header
    Dim x As Variant, findfield As Variant
    Dim oCell As Range
    Dim iNum As Long
    For x = LBound(BOQ_COLUMNS) To UBound(BOQ_COLUMNS)
        findfield = BOQ_COLUMNS(x)
        iNum = iNum + 1
        Set oCell = Worksheets(A_IMPORT_BIM).Rows(1).Find( _
            What:=findfield, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
        If Not oCell.column = iNum Then
            Columns(oCell.column).Cut
            Columns(iNum).Insert Shift:=xlToRight
        End If
    Next x
End Sub

Private Sub CombineSheetsWithDifferentHeaders()
'
'Polacz wiele arkuszy z innymi naglowkami
'
    
    Dim wksDst As Worksheet, wksSrc As Worksheet
    Dim lngIdx As Long, lngLastSrcColNum As Long, _
        lngFinalHeadersCounter As Long, lngFinalHeadersSize As Long, _
        lngLastSrcRowNum As Long, lngLastDstRowNum As Long
    Dim strColHeader As String
    Dim varColHeader As Variant
    Dim rngDst As Range, rngSrc As Range
    Dim dicFinalHeaders As Scripting.Dictionary
    Set dicFinalHeaders = New Scripting.Dictionary
    
    'Set references up-front
    dicFinalHeaders.CompareMode = vbTextCompare
    lngFinalHeadersCounter = 1
    lngFinalHeadersSize = dicFinalHeaders.Count
    Set wksDst = ThisWorkbook.Worksheets.Add
    wksDst.name = A_IMPORT_BIM
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Phase 1: Prepare Final Headers and Destination worksheet'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'First, we loop through all of the data worksheets,
    'building our Final Headers dictionary
    For Each wksSrc In ThisWorkbook.Worksheets
   
        'Make sure we skip the Destination worksheet!
        If (((wksSrc.name <> wksDst.name) _
        And wksSrc.name <> A_PRICE_LIST And wksSrc.name <> A_ASSUMPTIONS And wksSrc.name <> A_MAN_HOUR And wksSrc.name <> A_PROFILES And wksSrc.name <> A_COMMENTS And wksSrc.name <> A_CALCULATION2)) Then
            With wksSrc
        
                'Loop through all of the headers on this sheet,
                'adding them to the Final Headers dictionary
                lngLastSrcColNum = LastOccupiedColNum(wksSrc)
                For lngIdx = 1 To lngLastSrcColNum
                
                    'If this column header does NOT already exist in the Final
                    'Headers dictionary, add it and increment the column number
                    strColHeader = Trim(CStr(.Cells(1, lngIdx)))
                    If Not dicFinalHeaders.Exists(strColHeader) Then
                        dicFinalHeaders.Add KEY:=strColHeader, _
                                            item:=lngFinalHeadersCounter
                        lngFinalHeadersCounter = lngFinalHeadersCounter + 1
                    End If
                
                Next lngIdx
                
            End With
            
        End If
        
    Next wksSrc
    
    'Wahoo! The Final Headers dictionary now contains every column
    'header name from the worksheets. Let's write these values into
    'the Destination worksheet and finish Phase 1
    For Each varColHeader In dicFinalHeaders.Keys
        wksDst.Cells(1, dicFinalHeaders(varColHeader)) = CStr(varColHeader)
    Next varColHeader
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'End Phase 1: Final Headers are ready to rock!'
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Phase 2: write the data from each worksheet to the Destination!'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'We begin just like Phase 1 -- by looping through each sheet
    For Each wksSrc In ThisWorkbook.Worksheets
      
        'Once again, make sure we skip the Destination worksheet!
        If (((wksSrc.name <> wksDst.name) _
        And wksSrc.name <> A_PRICE_LIST And wksSrc.name <> A_ASSUMPTIONS And wksSrc.name <> A_MAN_HOUR And wksSrc.name <> A_PROFILES And wksSrc.name <> A_COMMENTS And wksSrc.name <> A_CALCULATION2)) Then
        
            With wksSrc
        
                'Identify the last row and column on this sheet
                'so we know when to stop looping through the data
                lngLastSrcRowNum = LastOccupiedRowNum(wksSrc)
                lngLastSrcColNum = LastOccupiedColNum(wksSrc)
                
                'Identify the last row of the Destination sheet
                'so we know where to (eventually) paste the data
                lngLastDstRowNum = LastOccupiedRowNum(wksDst)
                
                'Loop through the headers on this sheet, looking up
                'the appropriate Destination column from the Final
                'Headers dictionary and creating ranges on the fly
                For lngIdx = 1 To lngLastSrcColNum
                    strColHeader = Trim(CStr(.Cells(1, lngIdx)))
                    
                    'Set the Destination target range using the
                    'looked up value from the Final Headers dictionary
                    Set rngDst = wksDst.Cells(lngLastDstRowNum + 1, _
                                              dicFinalHeaders(strColHeader))
                                              
                    'Set the source target range using the current
                    'column number and the last-occupied row
                    Set rngSrc = .Range(.Cells(2, lngIdx), _
                                        .Cells(lngLastSrcRowNum, lngIdx))
                    
                    'Copy the data from this sheet to the destination!
                    rngSrc.Copy Destination:=rngDst
                    
                Next lngIdx
            
            End With
        
        End If
    
    Next wksSrc
    
    'Yay! Let the user know that the data has been combined
'    MsgBox "Data combined!"
End Sub

Private Function LastOccupiedRowNum(Sheet As Worksheet) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last row
'OUTPUT      : Long, the last occupied row
'SPECIAL CASE: if Sheet is empty, return 1
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
End Function

Private Function LastOccupiedColNum(Sheet As Worksheet) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last column
'OUTPUT      : Long, the last occupied column
'SPECIAL CASE: if Sheet is empty, return 1
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              After:=.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng
End Function
