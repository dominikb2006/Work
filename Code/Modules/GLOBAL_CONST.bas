Attribute VB_Name = "GLOBAL_CONST"
Option Explicit
'Worksheets names
Public Const A_IMPORT_BIM = "ImportBIM"
Public Const A_TABLE = "Tabela zbiorcza"
Public Const A_CALCULATION = "Kalkulacja"
Public Const A_CALCULATION2 = "Kalkulacja2"
Public Const A_PRICE_LIST = "Cennik"
Public Const A_MAN_HOUR = "Roboczogodziny"
Public Const A_PROFILES = "STA Profile"
Public Const A_ASSUMPTIONS = "Za�o�enia"
Public Const A_COMMENTS = "Uwagi"

'INFO IN CALCULATION

'Column A
Public Const POZ = 1
'Investment Name
Public Const INVESTMENT_NAME = "Nazwa inwestycji"  'Nazwa inwestycji

'Scope - count of scopes and initial location
Public Const SCOPE_COUNT = 12 'MAX 12
Public Const SCOPE_START = "I2"

'Optimalization - count of variants and initial location
Public Const OPT_COUNT = 12
Public Const OPT_START = "L2"


'Colors in BoQ
Public Const COLOR_BOQ_EDIT = 36
Public Const COLOR_BOQ_MUST_HAVE = 37
 
'Constant names from Revit
Public Const R_PHASE_CREATED = "Phase Created":
Public Const R_PHASE_DEMOLISHED = "Phase Demolished":
Public Const R_FAMILY = "Family":
Public Const R_TYPE = "Type":
Public Const R_CPI_KEY = "cpiFitMatchKey":
Public Const R_CPI_TYPE = "cpiComponentType":
Public Const R_COMMENT = "Comments":
Public Const R_5D4D_LVL = "5d4d-Geschoss":
Public Const R_5D4D_CODE = "5d4d-Bauelement":
Public Const R_5D4D_REGION = "5d4d-Bauabschnitt":

Public Const R_MATERIAL = "Material":
Public Const R_MATERIAL_STRUCT = "Structural Material":
Public Const R_MATERIAL_TYPE = "5d Materialg�te":
Public Const R_WATERPROOF = "5dki WU":
Public Const R_REINFORCEMENT_INDICATOR = "5d Bewehrung Stab kg pro m3":
Public Const R_REINFORCEMENT = "Reinforcement":
Public Const R_FORMWORK = "Formwork":
Public Const R_FORMWORK2 = "Formwork2":
Public Const R_NAME = "Name":
Public Const R_NAME_FINAL = "Final Name":

Public Const R_COUNT = "Count":
Public Const R_LENGTH = "Length":
Public Const R_LENGTH_CUT = "Cut Length":
Public Const R_PERIMETER = "Perimeter":
Public Const R_AREA = "Area":

Public Const R_SLOPE = "Slope":
Public Const R_VOLUME = "Volume":
Public Const R_VOLUME2 = "Volume2":
Public Const R_THICKNESS = "Default Thickness":
Public Const R_FOUND_THICKNESS = "Foundation Thickness":
Public Const R_WIDTH = "Width":
Public Const R_UNCONNECT_HEIGHT = "Unconnected Height":
Public Const R_DICKE_DE = "Dicke":
Public Const R_TIEFE_DE = "Tiefe":
Public Const R_BREITE_DE = "Breite":
Public Const R_LAUFBREITE_DE = "5d Laufbreite":

Public Const R_HOHE_DE = "H�he":
Public Const R_FORMWORK_HEIGHT = "5dki Schalh�he":
Public Const R_CUT = "Cut":
Public Const R_FILL = "Fill":
Public Const R_NET_CUT_FILL = "Net cut/fill":
Public Const R_INTERVAL_HEIGHT = "Interval High":

Public Const R_DIAMETER = "Durchmesser":

Public Const R_PRESTRESSED = "5d vorgespannt":
Public Const R_ANGLE = "5d Winkel":
Public Const R_PROFILE = "Profile":

'Kolejnosc i nr zakresow
Public Const z_nazwa = "zak" ' na razie nie zmieniac -> wpierw ujednolicic formule do wyszukiwania CJ
Public Const z_str = 0
Public Const z_rob = 1
Public Const z_bet = z_rob + 1 '2
Public Const z_sza = z_bet + 1 '3
Public Const z_zbr = z_sza + 1 '4
Public Const z_zie = z_zbr + 1 '5
Public Const z_ber = z_zie + 1 '6
Public Const z_roz = z_ber + 1 '7
Public Const z_sta = z_roz + 1 '8
Public Const z_mur = z_sta + 1 '9
Public Const z_inny = z_mur + 1 '10
Public Const z_11 = z_inny + 1 '11
Public Const z_12 = z_11 + 1 '12

Public Const COLOR_P0 = xlThemeColorAccent2
Public Const COLOR_P0v = 0
Public Const COLOR_P1 = xlThemeColorAccent6
Public Const COLOR_P1v = 0.4
Public Const COLOR_P2 = xlThemeColorAccent1
Public Const COLOR_P2v = 0.4
Public Const COLOR_P3 = xlThemeColorAccent2
Public Const COLOR_P3v = 1
Public Const COLOR_P4 = xlThemeColorAccent2
Public Const COLOR_P4v = 0.6
Public Const COLOR_P5 = xlThemeColorAccent2
Public Const COLOR_P5v = 0.8

Public Function S_COLOR_P4(S_NUMBER) As Variant
'
'Scope colors
'get color of scope based on number of scope
'
    Dim it As Variant
    Dim dict As New Scripting.Dictionary

    With dict
        .Add z_str, RGB(255, 204, 153)
        .Add z_rob, RGB(112, 173, 71)
        .Add z_bet, RGB(68, 114, 196)
        .Add z_sza, RGB(91, 167, 213)
        .Add z_zbr, RGB(255, 153, 0)
        .Add z_zie, RGB(255, 97, 97)

        .Add z_ber, RGB(166, 121, 255)
        .Add z_roz, RGB(51, 204, 204)
        .Add z_sta, RGB(63, 191, 127)
        .Add z_mur, RGB(255, 255, 0)
        .Add z_inny, RGB(153, 204, 0)

        .Add z_11, RGB(218, 176, 0)
        .Add z_12, RGB(145, 142, 0)

        For Each it In .Keys
            If it = S_NUMBER Then S_COLOR_P4 = .item(it)
        Next
    End With
End Function

Public Function S_COLOR_P5(S_NUMBER) As Variant
'
'get color of scope based on number of scope
'
    Dim it As Variant
    Dim dict As New Scripting.Dictionary

    With dict
        .Add z_str, RGB(255, 226, 197)
        .Add z_rob, RGB(198, 224, 180)
        .Add z_bet, RGB(180, 198, 231)
        .Add z_sza, RGB(139, 184, 225)
        .Add z_zbr, RGB(255, 184, 79)
        .Add z_zie, RGB(255, 139, 139)

        .Add z_ber, RGB(208, 185, 255)
        .Add z_roz, RGB(126, 226, 224)
        .Add z_sta, RGB(132, 214, 173)
        .Add z_mur, RGB(255, 255, 155)
        .Add z_inny, RGB(188, 255, 1)

        .Add z_11, RGB(255, 217, 55)
        .Add z_12, RGB(196, 191, 0)

        For Each it In .Keys
            If it = S_NUMBER Then S_COLOR_P5 = .item(it)
        Next
    End With
End Function

Public Function BOQ_COLUMNS() As Variant
'
'Array columns in ImportBIM
'
    BOQ_COLUMNS = Array( _
        R_PHASE_CREATED, R_PHASE_DEMOLISHED, R_FAMILY, R_TYPE, R_CPI_TYPE, R_PROFILE, R_CPI_KEY, _
        R_5D4D_REGION, R_5D4D_LVL, R_5D4D_CODE, _
        R_NAME, R_MATERIAL, R_MATERIAL_STRUCT, R_MATERIAL_TYPE, R_WATERPROOF, _
        R_NAME_FINAL, R_COUNT, R_THICKNESS, R_WIDTH, R_LENGTH, R_LENGTH_CUT, R_UNCONNECT_HEIGHT, R_AREA, R_VOLUME, R_VOLUME2, _
        R_REINFORCEMENT_INDICATOR, R_REINFORCEMENT, _
        R_FORMWORK, R_FORMWORK2, _
        R_FOUND_THICKNESS, R_DICKE_DE, R_TIEFE_DE, R_BREITE_DE, R_LAUFBREITE_DE, R_HOHE_DE, R_FORMWORK_HEIGHT, _
        R_PERIMETER, R_INTERVAL_HEIGHT, R_DIAMETER, R_PRESTRESSED, R_ANGLE, R_SLOPE, R_COMMENT, _
        R_CUT, R_FILL, R_NET_CUT_FILL)
End Function

Public Function P1_Description(P1_Code)
'
'Take LEVEL NAME from Dictonary based on CODE
'
'INPUT: Level Code
'OUTPUT: Level Name
'
    Dim it As Variant
    With CreateObject("scripting.dictionary")
        .Add "ZIE", "Roboty ziemne i rozbi�rkowe"
        .Add "FD", "Poziom FD - Fundamenty"
        .Add "EG", "Poziom EG - Parter"
        .Add "DG", "Poziom DG - Dach"
        .Add "DOD", "Pozycje dodatkowe"
        
        Dim i As Integer
        For i = 1 To 50
            .Add "OG" & i, "Poziom " & i & ".OG - Poziom +" & i
        Next
        For i = 1 To 10
            .Add "UG" & i, "Poziom " & i & ".UG - Poziom -" & i
        Next
        
        For Each it In .Keys
            If it = P1_Code Then P1_Description = .item(it)
        Next
    End With
End Function

Public Sub ADD_LVL_LIST()
'
'Add Custom List in Range UG10 - OG50
'
    Application.AddCustomList Array("ZIE", "FD", "UG10", "UG9", "UG8", "UG7", "UG6", "UG5", "UG4", "UG3", "UG2", "UG1", "EG", "OG1", "OG2", "OG3", "OG4", "OG5", "OG6", "OG7", "OG8", "OG9", "OG10", "OG11", "OG12", "OG13", "OG14", "OG15", "OG16", "OG17", "OG18", "OG19", "OG20", "OG21", "OG22", "OG23", "OG24", "OG25", "OG26", "OG27", "OG28", "OG29", "OG30", "OG31", "OG32", "OG33", "OG34", "OG35", "OG36", "OG37", "OG38", "OG39", "OG40", "OG41", "OG42", "OG43", "OG44", "OG45", "OG46", "OG47", "OG48", "OG49", "OG50", "DG", "DOD")
End Sub
