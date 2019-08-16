Attribute VB_Name = "BoQ"
Option Explicit
Public BoQRow As Long

Sub Main()

    CalculationQuantity
    Application.Calculate
End Sub

Sub Test()

    Application.Volatile False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    CalculationQuantity
    Application.Calculate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Public Sub CalculationQuantity()
'
'Quantity calculation
'
    Sheets(A_IMPORT_BIM).Select
    Dim cBoQ As New C_BoQ
        
    'Dim BoQRow As Long
    For BoQRow = 2 To Cells(Rows.Count, 3).End(xlUp).row
        cBoQ.BoQRow = BoQRow

        If Cells(BoQRow, BoQIndex(R_COUNT)).Formula = "" Then
            Cells(BoQRow, BoQIndex(R_COUNT)).Formula = "1"
            Cells(BoQRow, BoQIndex(R_COUNT)).Interior.ColorIndex = COLOR_BOQ_EDIT
        End If
            
        Cells(BoQRow, BoQIndex(R_5D4D_REGION)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
        Cells(BoQRow, BoQIndex(R_5D4D_LVL)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE

        'Case depends of CPI_KEY code
        Select Case Cells(BoQRow, BoQIndex(R_CPI_KEY)).Text

            Case "5D-KI-STB-OB-FU-EF" 'Stopa fundamentowa
                cBoQ.SpreadFoundation

            Case "5D-KI-STB-OB-FU-SF" '�awa fundamentowa
                cBoQ.StripFoundation

            Case "5D-KI-STB-OB-BoPla" 'P�yta/Rampa fundamentowa
                cBoQ.FoundationSlab

            Case "5D-KI-STB-OB-FU-AUB" 'P�yta fundamentowa podszybia
                cBoQ.FoundationSlab_PitElevator

            Case "5D-KI-STB-OB-FU-MF" 'Fundament pod maszyny
                cBoQ.FoundationSlab_Machine

            Case "5D-KI-SFB-OB-BoPla" 'P�yta fundamentowa ze zbrojeniem rozproszonym
                cBoQ.FoundationSlab_FiberReinforced

            Case "5D-KI-WB-OB-BoPla" 'P�yta fundamentowa z betonu wa�owanego
                cBoQ.FoundationSlab_RollerConcrete

            Case "5D-KI-STB-OB-W" '�ciana �elbetowa
                cBoQ.Wall_Concrete

            Case "5D-KI-STB-OB-W-1S" '�ciana �elbetowa jednostronnie szalowana
                cBoQ.Wall_Concrete_1S

            Case "5D-KI-STB-OB-WT" '�ciana �elbetowa szachtu
                cBoQ.Wall_Duct

            Case "5D-KI-STB-OB-W-K" 'Tarcza �elbetowa
                cBoQ.Wall_Disc

            Case "5D-KI-STB-OB-TR-UEZ" 'Attyka/nadci�g �elbetowy
                cBoQ.Wall_Attic

            Case "5D-KI-STB-OB-FU-AUW" '�ciana �elbetowa podszybia
                cBoQ.Wall_ElevatorPit

            Case "5D-KI-STB-OB-W-E" '�ciana prefabrykowana typu Filigran
                cBoQ.Wall_Filigran

            Case "5D-KI-STB-OB-TR-UZ" 'Podci�g �elbetowy
                cBoQ.Beam

            Case "5D-KI-MW-W-AP" '�ciana z Porothermu
                cBoQ.BrickWall_Porotherm

            Case "5D-KI-MW-W-PB" '�ciana z bloczk�w Ytong
                cBoQ.BrickWall_Ytong

            Case "5D-KI-MW-W-KSKF" '�ciana z silikat�w drobnowymiarowych
                cBoQ.BrickWall_SilkaSmall

            Case "5D-KI-MW-W-KSGF" '�ciana z silikat�w wielkowymiarowych
                cBoQ.BrickWall_SilkaBig

            Case "5D-KI-MW-W-SA" 'Obudowa szachtu z silikat�w
                cBoQ.BrickWall_SilkaSzacht

            Case "5D-KI-MW-W-TAB" '�ciana z TeknoAmerBlok
                cBoQ.BrickWall_TeknoAmber

            Case "5D-KI-MW-STB-AS" 'Trzpie� �ciany murowanej
                cBoQ.BrickWall_Tang

            Case "5D-KI-MW-STB-AR" 'Wieniec �ciany murowanej
                cBoQ.BrickWall_Grommet

            Case "5D-KI-UB-OB-SkS" 'Chudy beton
                cBoQ.LeanConcrete

            Case "5D-KI-UB-OB-FU-SF", _
                "5D-KI-UB-OB-FU-EF", _
                "5D-KI-UB-OB-BoPla" 'Chudy beton wype�niaj�cy
                cBoQ.LeanConcrete_Fulfil

            Case "5D-KI-STB-OB-V" 'Sko�na kraw�d� p�yty
                cBoQ.SlantingEdges

            Case "5D-KI-STB-OB-ST-REC" 'S�up �elbetowy prostok�tny
                cBoQ.Column_Rec

            Case "5D-KI-STB-OB-ST-RUN" 'S�up �elbetowy okr�g�y
                cBoQ.Column_Round

            Case "5D-KI-STB-OB-ST-LIS" 'Pilaster �elbetowy
                cBoQ.Column_Rec_Pilaser

            Case "5D-KI-STB-OB-ST-STK" 'G�owica �elbetowa
                cBoQ.FlaringHeads

            Case "5D-KI-STB-OB-TRE-POD" 'Spocznik �elbetowy
                cBoQ.Landing

            Case "5D-KI-STB-OB-TRE" 'Bieg schodowy
                cBoQ.Stairs

            Case "5D-KI-STB-OB-DE" 'Strop �elbetowy
                cBoQ.Floor

            Case "5D-KI-STB-OB-DE-EA" 'Strop prefabrykowany
                cBoQ.Floor_Filigran

            Case "5D-KI-STB-OB-M-KB" 'Wspornik zelbetowy
                cBoQ.Bracket

            Case "5D-KI-STB-OB-M-KON" 'Konsola zelbetowa
                cBoQ.Console

            Case "5D-KI-FU-MFP" 'Mata antywibracyjna pod fundamenty pod maszyny
                cBoQ.Foundation_VibrationInsulation

            Case "5D-KI-EBT-W-TPB" 'Dylatacja ze styropianu
                cBoQ.PartitionPanel

            Case "5D-KI-WDI-W", _
                "5D-KI-WDI-DE", _
                "5D-KI-WDI-AM" 'Izolacja wewn�trzna �ciany
                cBoQ.InteriorInsulation

            Case "5D-KI-WD-W", _
                "5D-KI-WD-DE", _
                "5D-KI-WD-AM" 'Izolacja �ciany
                cBoQ.WallInsulation

            Case "5D-KI-WD-BoPla-W", _
                "5D-KI-WD-BoPla-DE", _
                "5D-KI-WD-BoPla-V" 'Izolacja fundamentu
                cBoQ.Foundation_Insulation

            Case "5D-KI-WD-SGS-DE", _
                "5D-KI-WD-SGS-V" 'Podsypka keramzytowa
                cBoQ.FoamGlass

            Case "5D-KI-EBT-W-SGL" 'Podk�adka elastomerowa pasmowa
                cBoQ.ElastomericSlidingWasher_Strip

            Case "5D-KI-EBT-M-FGL" 'Podk�adka elastomerowa punktowa
                cBoQ.ElastomericSlidingWasher_Point
                
            Case "5D-KI-EBT-FT-TRE-PTR", _
            "5D-KI-EBT-FT-TRE-LTR", _
            "5D-KI-EBT-FT-TRE-FTP" 'Podk�adka elastomerowa pod schody
                cBoQ.Elastomeric_Pref_Stairs

            Case "5D-KI-EBT-M-HBT-T1" 'Zbrojenie Comax jednorz�dowe
                cBoQ.Comax_T1

            Case "5D-KI-EBT-M-HBT-T5" 'Zbrojenie Comax dwurz�dowe
                cBoQ.Comax_T5

            Case "5D-KI-STB-FT-TR-UZ" 'Prefabrykowana belka
                cBoQ.Prefabric_Beam

            Case "5D-KI-STB-FT-DE-MP" 'Prefabrykowany strop
                cBoQ.Prefabric_Slab

            Case "5D-KI-STB-FT-DE-SHP" 'Prefabrykowany strop kana�owy
                cBoQ.Prefabric_Slab_Hollow

            Case "5D-KI-STB-FT-DE-TTP" 'Prefarykowana p�yta TT z nadbetonem
                cBoQ.Prefabric_SlabTT

            Case "5D-KI-SB-HP-REC", _
                "5D-KI-SB-HP-RUN", _
                "5D-KI-SB-WP-I", _
                "5D-KI-SB-WP-LU", _
                "5D-KI-SB-VP", _
                "5D-KI-SB-SP" 'Konstrukcja stalowa
                cBoQ.Steel

            Case "TOPO" 'Roboty ziemne
                cBoQ.EarthMoving

            Case "5D-KI-STB-FT-ST-REC" 'Prefabrykowana kolumna
                cBoQ.Prefabric_Column
                
            Case "5D-KI-STB-FT-TRE" 'Prefabrykowany bieg schodowy
                cBoQ.Prefabric_Stairs
                
            Case "5D-KI-STB-FT-TRE-POD" 'Prefabrykowany spocznik
                cBoQ.Prefabric_Landing
                
'            Case "5D-TB-SW" '�cianka szczelinowa
'                SlurryWall BoQRow
'
'            Case "BER" '�cianka berli�ska
'                berlinka BoQRow
                           
        End Select
    Next BoQRow
End Sub

'-------------------------------------------------------------
'-------------------------------------------------------------
'STARE ---> DO ZMIANY ALBO USUNIECIA
'-------------------------------------------------------------
'-------------------------------------------------------------
'Private Sub SlurryWall(BoQRow)
''
''Scianka szczelinowa
''
'    Dim BoQ As New clsBoQ
'    'Dim MyCol As New clsMyColumns
'
'    Cells(BoQRow, BoQIndex(R_NAME)).Formula = "�ciana szczelinowa"
'    Cells(BoQRow, BoQIndex(R_NAME_FINAL)).Formula = BoQ.NameIntervalHigh(BoQRow)
'
'End Sub
