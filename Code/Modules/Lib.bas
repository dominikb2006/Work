Attribute VB_Name = "lib"
Public Sub Test()

    Renaming A_PRICE_LIST, A_CALCULATION2, A_CALCULATION
    AlignmentFiltr
'    DeleteSheet
    SheetsOrder
    NumericDataFormat
End Sub
Public Sub Main()

    Renaming A_PRICE_LIST, A_CALCULATION2, A_CALCULATION
    AlignmentFiltr
    DeleteSheet
    SheetsOrder
    NumericDataFormat
End Sub

Public Function CalcIndex(ByVal i As Variant) As Variant
'
'Find in Array CALC_COLUMNS
'
    CalcIndex = WhereInArray(CALC_COLUMNS, i) + 1
End Function

Public Function BoQIndex(ByVal i As Variant) As Variant
'
'Find in Array BOQ_COLUMNS
'
    BoQIndex = WhereInArray(BOQ_COLUMNS, i) + 1
End Function

Public Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'
'Find in Array (start from 0)
'
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) = vFind Then
            WhereInArray = i
            Exit Function
        End If
    Next i
    WhereInArray = Null
End Function

Public Function cell(row, column, Optional reference As Boolean = False)
'
'Shorter way to give cells into formulas (previusly: Cells(row, col).Address(0, 0))
'
    If reference = False Then
        cell = Cells(row, column).Address(0, 0)
    Else
        cell = Cells(row, column).Address(1, 1)
    End If
End Function

Public Sub NumericDataFormat()
'
'Zmiana formatu danych na numeryczne - zakres Kalkulacja - K22:O1048576
'
    Sheets(A_CALCULATION).Select
    Range("K" + CStr(HEADLINE_ROW + 1) + ":O1048576").Select
    Selection.NumberFormat = "#,##0.00"
End Sub

Public Sub AlignmentFiltr()
'
'Autofit w Kalkulacji plus Filtr
'
'Wyrownanie
    Worksheets(A_CALCULATION).Activate
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
'Filtr
    Rows(CStr(HEADLINE_ROW) + ":1048576").Select
    Selection.AutoFilter
End Sub

Public Sub DeleteSheet()
'
'Usuniecie Kalkulacji2
'
    Sheets(A_CALCULATION2).Select
    ActiveWindow.SelectedSheets.Delete
End Sub

Public Sub SheetsOrder()
'
'Sheets order
'Porz¹dek arkuszy
'
    Worksheets(A_TABLE).Move After:=Worksheets(A_IMPORT_BIM)
    Worksheets(A_CALCULATION).Move After:=Worksheets(A_TABLE)
    Worksheets(A_PRICE_LIST).Move After:=Worksheets(A_CALCULATION)
    Worksheets(A_MAN_HOUR).Move After:=Worksheets(A_PRICE_LIST)
    Worksheets(A_PROFILES).Move After:=Worksheets(A_MAN_HOUR)
End Sub


Public Sub Renaming(wrksh As String, change_what As String, for_what As String)
'
'Zmiana w Cenniku w formulach Kalkulacja2 na Kalkulacja
'
    Sheets(wrksh).Select
    Cells.Select
    Selection.Replace What:=change_what, Replacement:=for_what, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Public Sub Renaming2()
'
'Zmiana nazwy w Cenniku z Kalkulacja na Kalkulacja2
'
    Sheets(A_PRICE_LIST).Select
    Cells.Select
    Selection.Replace What:=A_CALCULATION, Replacement:=A_CALCULATION2, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Public Sub NewSheet()
'
'Nowy Arkusz o nazwie Kalkulacja2
'
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = A_CALCULATION2
End Sub

Public Sub UsuniecieKalkulacja()
'
'Usuniecie arkusza Kalkulacja
'
    Sheets(A_CALCULATION).Select
    ActiveWindow.SelectedSheets.Delete
End Sub

Public Sub UsuniecieNowych()
'
'Usuniecie arkuszy ImportBIM, Tabela zbiorcza oraz Kalkulacja
'
    Sheets(A_IMPORT_BIM).Select
    ActiveWindow.SelectedSheets.Delete
    Sheets(A_TABLE).Select
    ActiveWindow.SelectedSheets.Delete
    Sheets(A_CALCULATION).Select
    ActiveWindow.SelectedSheets.Delete
End Sub

Public Sub Borders()
'
'Obramowanie wokol komórek(procz przekatnych)
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub
