Attribute VB_Name = "Table"
Option Explicit
Const Table = "Tabela zbiorcza"             'nazwa tabeli przestawnej

Public Function TABLE_COLUMNS(calc_item)
'
'Array columns in Table/Calculation
'
    Dim it As Variant
    Dim dict As New Scripting.Dictionary

    With dict
        .Add R_COUNT, xlSum
        .Add R_THICKNESS, xlAverage
        .Add R_WIDTH, xlAverage
        .Add R_LENGTH, xlSum
        .Add R_LENGTH_CUT, xlSum

        .Add R_UNCONNECT_HEIGHT, xlAverage
        .Add R_AREA, xlSum
        .Add R_VOLUME, xlSum
        .Add R_REINFORCEMENT_INDICATOR, xlAverage
        .Add R_REINFORCEMENT, xlSum

        .Add R_FORMWORK, xlSum
        .Add R_FORMWORK2, xlSum
        .Add R_FORMWORK_HEIGHT, xlAverage
        .Add R_PERIMETER, xlSum
        .Add R_ANGLE, xlAverage

        .Add R_CUT, xlSum
        .Add R_FILL, xlSum

        For Each it In .Keys
            If it = calc_item Then TABLE_COLUMNS = .item(it)
        Next
    End With
End Function

Sub Test()
    'Dodanie listy
    ADD_LVL_LIST
    CALC_COLUMNS
    
    CreateTable
    OrganizeTable
    NoSubtotals
    RefreshTable
End Sub

Sub Main()
    'Dodanie listy
    ADD_LVL_LIST
    CALC_COLUMNS
    
    CreateTable
    OrganizeTable
    NoSubtotals
    RefreshTable
End Sub

Private Function table_columns_name(name)
'
'Take string from Dictonary based on Item (xlSum or xlAverage)
'
'INPUT: name of
'OUTPUT: Level Name
'
    If TABLE_COLUMNS(name) = xlSum Then
        table_columns_name = "Suma z " + name
    ElseIf TABLE_COLUMNS(name) = xlAverage Then
        table_columns_name = "�rednia z " + name
    Else
        table_columns_name = "-1"
    End If
End Function

Sub CreateTable()
'
'Tworzenie tabeli przestawnej
'
    Dim Table_source_data  As Variant
    Table_source_data = A_IMPORT_BIM + "!R1C1:R1048576C" + CStr(UBound(BOQ_COLUMNS) + 1)

    Sheets(A_IMPORT_BIM).Select
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Table_source_data). _
    CreatePivotTable _
        TableDestination:="", _
        TableName:=Table
    
    ActiveSheet.name = A_TABLE
    
    With Sheets(A_TABLE).Tab
        .ColorIndex = 40
        .TintAndShade = 0.599993896298105
    End With

    With ActiveSheet.PivotTables(Table).PivotFields(R_5D4D_REGION)
        .Orientation = xlRowField
        .Position = 1
        
    End With
    With ActiveSheet.PivotTables(Table).PivotFields(R_5D4D_LVL)
        .Orientation = xlRowField
        .Position = 2
    End With
    
    With ActiveSheet.PivotTables(Table).PivotFields(R_5D4D_CODE)
        .Orientation = xlRowField
        .Position = 3
    End With

    With ActiveSheet.PivotTables(Table).PivotFields(R_NAME_FINAL)
        .Orientation = xlRowField
        .Position = 4
    End With
    
    Dim it As Variant
    Dim i As Integer
    i = 0

    'Add columns in Table depends of Dictoniary
    For Each it In CALC_COLUMNS_2
        ActiveSheet.PivotTables(Table).AddDataField ActiveSheet.PivotTables(Table).PivotFields(CALC_COLUMNS_2(i)), table_columns_name(CALC_COLUMNS_2(i)), TABLE_COLUMNS(CALC_COLUMNS_2(i))
        i = i + 1
    Next
    
    With ActiveSheet.PivotTables(Table)
        .InGridDropZones = True
        .ShowValuesRow = False
        .RowAxisLayout xlTabularRow
    End With
       
    'Autofit
    Cells.EntireColumn.AutoFit
End Sub

Sub OrganizeTable()
'
'Sorting by R_5D4D_LVL using list number 6 (file->option->..->lists)
'Sortowanie tabeli po R_5D4D_LVL i liscie nr 6 (plik->opcje->zaawanowane->ogolne->listy niestandardowe)
'
    Dim PvtTbl As PivotTable
    Set PvtTbl = Worksheets(Table).PivotTables(Table)
    
    PvtTbl.SortUsingCustomLists = True
    PvtTbl.PivotFields(R_5D4D_LVL).DataRange.Sort Order1:=xlAscending, Type:=xlSortLabels, OrderCustom:=6
End Sub

Sub RefreshTable()
'
'Odswiezanie tabeli
'
    Dim Sheet As Worksheet, Pivot As PivotTable

    For Each Sheet In ThisWorkbook.Worksheets
        For Each Pivot In Sheet.PivotTables
            Pivot.RefreshTable
            Pivot.Update
        Next
    Next
End Sub

Sub NoSubtotals()
'
'Do not show subtotals in PivotTable
'

    Dim pt As PivotTable
    Dim pf As PivotField
    On Error Resume Next
    For Each pt In ActiveSheet.PivotTables
        pt.ManualUpdate = True
        For Each pf In pt.PivotFields
            'First, set index 1 (Automatic) to True,
            'so all other values are set to False
            pf.Subtotals(1) = True
            pf.Subtotals(1) = False
        Next pf
        pt.ManualUpdate = False
    Next pt
End Sub
