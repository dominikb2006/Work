Attribute VB_Name = "Calculation"
Option Explicit

Public Const C_POZ = "POZ"
Public Const C_OPT = "OPT"
Public Const C_LP = "Lp."
Public Const C_AREA = "Dzia�ka"
Public Const C_LVL = "Poziom"
Public Const C_ELEMENT = "Element"
Public Const C_MR = "M/R"
Public Const C_MATERIAL = "Materia�"
Public Const C_PRICE = "Cena"
Public Const C_DESCRIPTION = "Opis"
Public Const C_UNIT = "Jedn."
Public Const C_AMOUNT = "Ilo��"
Public Const C_AMOUNT2 = "Ilo�� do Kosztorysu"
Public Const C_UNITPRICE = "CJ"
Public Const C_SCOPE = "Zakres"
Public Const C_VALUE = "Warto��"
Public Const C_NETVALUE = "Warto�� Netto"
Public Const C_COMMENTS = "Uwagi"
Public Const C_TAKEOFF = "Przedmiary"
Public Const C_WORKVALUE = "Warto�� prac"
Public Const C_ELEMENT2 = "Nazwa Elementu"


Public CalcRow As Long 'tymczasowy wiersz
Public CalcRow_Table As Long 'tymczasowy wiersz z nazwa

Public CALC_COLUMNS_1 As Variant
Public CALC_COLUMNS_2 As Variant
Public CALC_COLUMNS_3 As Variant

Sub Main()

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).name = A_CALCULATION
    End With

    Optimizations
    Scopes
    Headline
    Positioning
    ExtraPosition
    Grouping
End Sub

Sub Test()

    Application.Volatile False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
        
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).name = A_CALCULATION
    End With

    Optimizations
    Scopes
    Headline
    Positioning
    ExtraPosition
    Grouping
    
    Application.Calculate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Public Function CALC_COLUMNS() As Variant
'
'Add array columns in Calculation
'
    CALC_COLUMNS_1 = Array(C_POZ, C_OPT, C_LP, C_AREA, C_LVL, C_ELEMENT, C_MR, C_MATERIAL, C_PRICE, C_DESCRIPTION, _
        C_UNIT, C_AMOUNT, C_AMOUNT2, C_SCOPE, C_UNITPRICE, C_VALUE, C_NETVALUE, C_COMMENTS, C_TAKEOFF)

    CALC_COLUMNS_2 = Array(R_COUNT, R_THICKNESS, R_WIDTH, R_LENGTH, R_LENGTH_CUT, _
        R_UNCONNECT_HEIGHT, R_AREA, R_VOLUME, R_REINFORCEMENT_INDICATOR, R_REINFORCEMENT, _
        R_FORMWORK, R_FORMWORK2, R_FORMWORK_HEIGHT, R_PERIMETER, R_ANGLE, _
        R_CUT, R_FILL)

    CALC_COLUMNS_3 = Array(C_WORKVALUE, C_ELEMENT2)
    
    CALC_COLUMNS = Split(Join(CALC_COLUMNS_1, ",") & "," & Join(CALC_COLUMNS_2, ",") & "," & Join(CALC_COLUMNS_3, ","), ",")
End Function

Function HEADLINE_ROW()
    HEADLINE_ROW = WorksheetFunction.Max(Range(OPT_START).Cells.row, Range(SCOPE_START).Cells.row) + WorksheetFunction.Max(OPT_COUNT, SCOPE_COUNT) + 3
End Function

Sub ExtraPosition()
'
'Additional position at end
'Dodanie pozycji dodatkowych, jak dodatki do betonu
'
    Dim cTable As New C_Table
     
    Worksheets(A_CALCULATION).Activate
    
    'P0
    Cells(CalcRow, CalcIndex(C_AREA)).Formula = ""
    cTable.P0
    
    'P1
    Cells(CalcRow, CalcIndex(C_LVL)).Formula = "DOD"
    cTable.P1
    
    'P2
    Cells(CalcRow, CalcIndex(C_ELEMENT)).Formula = "DOD"
    cTable.P2
    
    'P3
    Cells(CalcRow, CalcIndex(C_DESCRIPTION)).Formula = "Dodatki do stanu surowego"
    CalcRow_Table = CalcRow
    cTable.P3
    
    'P4
    cTable.P4
End Sub

Sub Grouping()
'
'Grouping rows
'Grupowanie wierszy
'
    Dim i As Long
    Dim j As Long
    Worksheets(A_CALCULATION).Activate

    For i = HEADLINE_ROW + 1 To ActiveCell.SpecialCells(xlLastCell).row
        Select Case Cells(i, CalcIndex(C_POZ)).Text
            Case "P0"
                For j = i + 1 To ActiveCell.SpecialCells(xlLastCell).row
                    If Cells(j, CalcIndex(C_POZ)).Text = "P0" Then Exit For
                Next j
                If (i < j And Cells(i, CalcIndex(C_POZ)).Text <> Cells(i + 1, CalcIndex(C_POZ)).Text) Then
                    Range(Cells(i + 1, CalcIndex(C_POZ)), Cells(j - 1, CalcIndex(C_POZ))).Rows.Group
                End If
            Case "P1"
                For j = i + 1 To ActiveCell.SpecialCells(xlLastCell).row
                    If (Cells(j, CalcIndex(C_POZ)).Text = "P1" Or Cells(j, CalcIndex(C_POZ)).Text = "P0") Then Exit For
                Next j
                If (i < j And Cells(i, CalcIndex(C_POZ)).Text <> Cells(i + 1, CalcIndex(C_POZ)).Text) Then
                    Range(Cells(i + 1, CalcIndex(C_POZ)), Cells(j - 1, CalcIndex(C_POZ))).Rows.Group
                End If
            Case "P2"
                For j = i + 1 To ActiveCell.SpecialCells(xlLastCell).row
                    If (Cells(j, CalcIndex(C_POZ)).Text = "P2" Or Cells(j, CalcIndex(C_POZ)).Text = "P1" Or Cells(j, CalcIndex(C_POZ)).Text = "P0") Then Exit For
                Next j
                If (i < j And Cells(i, CalcIndex(C_POZ)).Text <> Cells(i + 1, CalcIndex(C_POZ)).Text) Then
                    Range(Cells(i + 1, CalcIndex(C_POZ)), Cells(j - 1, CalcIndex(C_POZ))).Rows.Group
                End If
            Case "P3"
                For j = i + 1 To ActiveCell.SpecialCells(xlLastCell).row
                    If (Cells(j, CalcIndex(C_POZ)).Text = "P3" Or Cells(j, CalcIndex(C_POZ)).Text = "P2" Or Cells(j, CalcIndex(C_POZ)).Text = "P1" Or Cells(j, CalcIndex(C_POZ)).Text = "P0") Then Exit For
                Next j
                If (i < j And Cells(i, CalcIndex(C_POZ)).Text <> Cells(i + 1, CalcIndex(C_POZ)).Text) Then
                    Range(Cells(i + 1, CalcIndex(C_POZ)), Cells(j - 1, CalcIndex(C_POZ))).Rows.Group
                End If
            Case "P4"
                For j = i + 1 To ActiveCell.SpecialCells(xlLastCell).row
                    If (Cells(j, CalcIndex(C_POZ)).Text = "P4" Or Cells(j, CalcIndex(C_POZ)).Text = "P3" Or Cells(j, CalcIndex(C_POZ)).Text = "P2" Or Cells(j, CalcIndex(C_POZ)).Text = "P1" Or Cells(j, CalcIndex(C_POZ)).Text = "P0") Then Exit For
                Next j
                If (i < j And Cells(i, CalcIndex(C_POZ)).Text <> Cells(i + 1, CalcIndex(C_POZ)).Text) Then
                    Range(Cells(i + 1, CalcIndex(C_POZ)), Cells(j - 1, CalcIndex(C_POZ))).Rows.Group
                End If
        End Select
    Next i
    
    'delete errors of grouping
    Dim lrow As Integer
    lrow = Cells(Rows.Count, CalcIndex(C_LP)).End(xlUp).row
    
    Sheets(A_CALCULATION).Select
    For i = 1 To 10
        Rows(lrow + 1).EntireRow.Delete
    Next i
End Sub

Sub Positioning()
'
'Add Main Table
'Pozycjonowanie
'
    Dim cTable As New C_Table
    
    Dim Tab_S As Range
    'Oznaczenia kolumn
    Const T_AREA = 1
    Const T_LVL = 2
    Const T_CODE = 3
    Const T_NAME = 4
    Dim TabRow  As Long
    
    Set Tab_S = Worksheets(A_TABLE).Range("A1") 'poczatek tabeli
    Worksheets(A_TABLE).Activate
    Tab_S.Select
    Tab_S.Offset(2, 0).Select
    
    For TabRow = Tab_S.Offset(2, 0).row To ActiveCell.SpecialCells(xlLastCell).row
        Worksheets(A_TABLE).Activate
        
        If (Cells(TabRow, T_AREA) <> "" And Cells(TabRow, T_AREA) <> "(puste)" And Cells(TabRow, T_AREA) <> "Suma ko�cowa") Then
            Worksheets(A_CALCULATION).Activate
            CalcRow = Cells(ActiveCell.SpecialCells(xlLastCell).row + 1, CalcIndex(C_POZ)).row
            
            'P0
            Cells(CalcRow, CalcIndex(C_AREA)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_AREA).Formula
            cTable.P0
            
            'P1
            Cells(CalcRow, CalcIndex(C_LVL)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_LVL).Formula
            cTable.P1
            
            'P2
            Cells(CalcRow, CalcIndex(C_ELEMENT)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_CODE).Formula
            cTable.P2
            
            'P3
            Cells(CalcRow, CalcIndex(C_DESCRIPTION)).Formula = Worksheets(A_TABLE).Cells(TabRow, T_NAME).Formula
            CalcRow_Table = CalcRow
            cTable.P3
            
            'P4
            cTable.P4
        
        ElseIf (Cells(TabRow, T_LVL) <> "" And Cells(TabRow, T_LVL) <> "(puste)") Then
            Worksheets(A_CALCULATION).Activate
            
            'P1
            Cells(CalcRow, CalcIndex(C_LVL)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_LVL).Formula
            cTable.P1
            
            'P2
            Cells(CalcRow, CalcIndex(C_ELEMENT)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_CODE).Formula
            cTable.P2
            
            'P3
            Cells(CalcRow, CalcIndex(C_DESCRIPTION)).Formula = Worksheets(A_TABLE).Cells(TabRow, T_NAME).Formula
            CalcRow_Table = CalcRow
            cTable.P3
            
            'P4
            cTable.P4

        ElseIf (Cells(TabRow, T_CODE) <> "" And Cells(TabRow, T_CODE) <> "(puste)" And Cells(TabRow, T_CODE) <> "") Then
            Worksheets(A_CALCULATION).Activate

            'P2
            Cells(CalcRow, CalcIndex(C_ELEMENT)).Formula = "'" & Worksheets(A_TABLE).Cells(TabRow, T_CODE).Formula
            cTable.P2
            
            'P3
            Cells(CalcRow, CalcIndex(C_DESCRIPTION)).Formula = Worksheets(A_TABLE).Cells(TabRow, T_NAME).Formula
            CalcRow_Table = CalcRow
            cTable.P3
            
            'P4
            cTable.P4
            
         ElseIf (Cells(TabRow, T_NAME) <> "" And Cells(TabRow, T_NAME) <> "(puste)") Then
            Worksheets(A_CALCULATION).Activate

            'P3
            Cells(CalcRow, CalcIndex(C_DESCRIPTION)).Formula = Worksheets(A_TABLE).Cells(TabRow, T_NAME).Text
            CalcRow_Table = CalcRow
            cTable.P3
            
            'P4
            cTable.P4
            
        End If
    Next TabRow
End Sub

Sub Headline()
'
'Add headline
'Dodanie naglowkow
'

    Dim i As Integer
    Dim Tabela  As Range

'TEXT
    For i = 0 To UBound(CALC_COLUMNS)
        Cells(HEADLINE_ROW, CalcIndex(CALC_COLUMNS(i))).Formula = CALC_COLUMNS(i)
    Next i
    
'FORMAT
'COLUMNS A:B
    Range(Columns(CalcIndex(C_POZ)), Columns(CalcIndex(C_OPT))).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Font.ColorIndex = 15
    End With
'COLUMNS C:R
    Range(Cells(HEADLINE_ROW, CalcIndex(C_LP)), Cells(HEADLINE_ROW, CalcIndex(C_COMMENTS))).Select
    lib.Borders
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.ThemeColor = xlThemeColorAccent4
        .Interior.TintAndShade = 0.399975585192419
    End With
'COLUMNS S:AJ
    Range(Cells(HEADLINE_ROW, CalcIndex(C_TAKEOFF)), Cells(HEADLINE_ROW, CalcIndex(R_FILL))).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0.599993896298105
    End With
'COLUMNS AK:AL
    Cells(HEADLINE_ROW, CalcIndex(C_WORKVALUE)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.ThemeColor = xlThemeColorLight2
        .Interior.TintAndShade = 0.399975585192419
    End With 'Kolumna AM
    Cells(HEADLINE_ROW, CalcIndex(C_ELEMENT2)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.ThemeColor = xlThemeColorAccent3
        .Interior.TintAndShade = 0.399975585192419
    End With 'Kolumny AN:AP

'GROUPING
    Range(Cells(HEADLINE_ROW, CalcIndex(C_AREA)), Cells(HEADLINE_ROW, CalcIndex(C_PRICE))).Columns.Group
    Range(Cells(HEADLINE_ROW, CalcIndex(R_COUNT)), Cells(HEADLINE_ROW, CalcIndex(R_FILL))).Columns.Group
    
    With ActiveSheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlLeft
    End With
    
'HIDE RIGHT UNNECESARY COLUMNS
    Columns(CalcIndex(C_ELEMENT2) + 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    
'SECOND ROW
'NAME
    Cells(HEADLINE_ROW + 1, CalcIndex(C_DESCRIPTION)).Select
    With Selection
        .Formula = INVESTMENT_NAME
        .Font.Bold = True
        .Font.ColorIndex = 3
        .HorizontalAlignment = xlCenter
    End With
'SUM
    Cells(HEADLINE_ROW + 1, CalcIndex(C_NETVALUE)).Select
    With Selection
        .Formula = "=SUM(" & Columns(CalcIndex(C_VALUE)).Address(0, 0) & ")"
        .Font.Bold = True
        .Font.ColorIndex = 3
        .NumberFormat = "#,##0.00"
    End With
    
    Range(Cells(HEADLINE_ROW, CalcIndex(C_LP)), Cells(HEADLINE_ROW, CalcIndex(C_COMMENTS))).Offset(1, 0).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
    End With
    lib.Borders
End Sub

Private Sub Scopes()
'
'Add Scope Table
'Dodanie tabelki zakres�w
'
    Dim i As Integer
    
    Sheets(A_CALCULATION).Select
    Range(SCOPE_START).Select
    
    Range(SCOPE_START).Formula = "Kalk. w�asna"
    
    'Kolory branz
    ActiveCell.Offset(0, 1).Formula = "STRABAG"
    ActiveCell.Offset(0, 1).Interior.COLOR = S_COLOR_P4(z_str)
    ActiveCell.Offset(z_rob, 1).Formula = "Roboczogodziny"
    ActiveCell.Offset(z_rob, 1).Interior.COLOR = S_COLOR_P4(z_rob)
    ActiveCell.Offset(z_bet, 1).Formula = "Beton 1"
    ActiveCell.Offset(z_bet, 1).Interior.COLOR = S_COLOR_P4(z_bet)
    ActiveCell.Offset(z_sza, 1).Formula = "Szalunki 1"
    ActiveCell.Offset(z_sza, 1).Interior.COLOR = S_COLOR_P4(z_sza)
    ActiveCell.Offset(z_zbr, 1).Formula = "Zbrojenie 1"
    ActiveCell.Offset(z_zbr, 1).Interior.COLOR = S_COLOR_P4(z_zbr)
    ActiveCell.Offset(z_zie, 1).Formula = "Roboty ziemne 1"
    ActiveCell.Offset(z_zie, 1).Interior.COLOR = S_COLOR_P4(z_zie)
    ActiveCell.Offset(z_ber, 1).Formula = "Berlinka 1"
    ActiveCell.Offset(z_ber, 1).Interior.COLOR = S_COLOR_P4(z_ber)
    ActiveCell.Offset(z_roz, 1).Formula = "Rozbiorki 1"
    ActiveCell.Offset(z_roz, 1).Interior.COLOR = S_COLOR_P4(z_roz)
    ActiveCell.Offset(z_sta, 1).Formula = "Konstrukcja stalowa 1"
    ActiveCell.Offset(z_sta, 1).Interior.COLOR = S_COLOR_P4(z_sta)
    ActiveCell.Offset(z_mur, 1).Formula = "Mury 1"
    ActiveCell.Offset(z_mur, 1).Interior.COLOR = S_COLOR_P4(z_mur)
    ActiveCell.Offset(z_inny, 1).Formula = "Inny zakres 1"
    ActiveCell.Offset(z_inny, 1).Interior.COLOR = S_COLOR_P4(z_inny)
    ActiveCell.Offset(z_11, 1).Formula = "Zakres11a"
    ActiveCell.Offset(z_11, 1).Interior.COLOR = S_COLOR_P4(z_11)
    ActiveCell.Offset(z_12, 1).Formula = "Zakres12a"
    ActiveCell.Offset(z_12, 1).Interior.COLOR = S_COLOR_P4(z_12)
    
    'JEZELI POJAWIA SI� TUTAJ BLAD 1004, TO NALEZY ZMIENIC W MENAGERZE NAZW ZAKRES NAZWY Z CENNIKA NA SKOROSZYT (ABY NAZWY ZAKRESOW BYLY WIDOCZNE W INNYCH ARKUSZACH NIZ CENNIKU)
    For i = 1 To SCOPE_COUNT
        If i < 13 Then ' Kodowanie z zerem w nazwach
            ActiveCell.Offset(i, 0).Formula = "Zakres " & i
            ActiveCell.Offset(i, 1).Validation. _
                Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=zak" & i & "_nazwa"
                
        '        ElseIf i < 13 Then 'kodowanie dziesietnych
        '                ActiveCell.Offset(i, 0).Formula = "Zakres " & i
        '                ActiveCell.Offset(i, 1).Validation. _
        '                    Add Type:=xlValidateList, _
        '                    AlertStyle:=xlValidAlertStop, _
        '                    Operator:=xlBetween, _
        '                    Formula1:="=zak" & i & "_nazwa"
        '        Else 'Randomowe kolory zeby byly
        '            ActiveCell.Offset(i, 0).Formula = "Zakres " & i
        '            ActiveCell.Offset(i, 1).Validation. _
        '                Add Type:=xlValidateList, _
        '                AlertStyle:=xlValidAlertStop, _
        '                Operator:=xlBetween, _
        '                Formula1:="=zak" & i & "_nazwa"
        '                ActiveCell.Offset(i, 1).Interior.ColorIndex = Rnd
        
        End If
    Next i
    ActiveCell.Range(Cells(1, 1), Cells(SCOPE_COUNT + 1, 2)).Select
    lib.Borders
End Sub

Sub Optimizations()
'
'Add optimizations table
'Dodanie tabelki z optymalizacjami
'
    Dim j As Integer
    
    'Tworzenie tabeli
    Sheets(A_CALCULATION).Select
    Range(OPT_START).Select
    Range(OPT_START).Formula = "Dokumentacja"
    
    ActiveCell.Offset(0, 1).Formula = "Wersja"
    
    ActiveCell.Offset(0, 2).Formula = "Oferta bazowa"
    ActiveCell.Offset(0, 2).Font.ColorIndex = 3
    
    ActiveCell.Offset(0, 3).Formula = "Oferta optymalizacyjna"
    ActiveCell.Offset(0, 3).Font.ColorIndex = 3
     
    ActiveCell.Range("A1:B1").Interior.ColorIndex = 15
    
    For j = 1 To OPT_COUNT
        ActiveCell.Offset(j, 0).Formula = "Optymalizacja " & j
        ActiveCell.Offset(j, 1).Validation. _
            Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=j & "N," & j & "T"
            
        ActiveCell.Offset(j, 1).Formula = j & "N"
        ActiveCell.Offset(j, 2).Formula = j & "N"
        ActiveCell.Offset(j, 2).Font.ColorIndex = 3
        
        ActiveCell.Offset(j, 2).HorizontalAlignment = xlCenter
        ActiveCell.Offset(j, 3).HorizontalAlignment = xlCenter
    Next j
        
    ActiveCell.Range(Cells(1, 1), Cells(OPT_COUNT + 1, 2)).Select
    lib.Borders
    
    ActiveCell.Range(Cells(2, 2), Cells(OPT_COUNT + 1, 2)).Select
    lib.Borders
    ActiveWorkbook.Names.Add name:="optymalizacje", RefersTo:=Selection
End Sub
