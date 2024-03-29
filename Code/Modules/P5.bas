Attribute VB_Name = "P5"
Option Explicit
Public Enum ZBR_Type
    FD
    STR_
    SC
    POD
    SL
End Enum

'-----------------------------------------------------------------------------
'----------------------------------ZBROJENIE----------------------------------
'-----------------------------------------------------------------------------
Public Sub Zbrojenie_M(Optional element As ZBR_Type)
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "ZBR", " - zbrojenie M", "t", "=" & colT(R_REINFORCEMENT), z_zbr

    Select Case element
        Case FD
            cElement.element "M", "ZBR-10", " - zbrojenie M - dodatek do " & ChrW(216) & "10-12", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_FD_10", z_zbr
        Case STR_
            cElement.element "M", "ZBR-10", " - zbrojenie M - dodatek do " & ChrW(216) & "10-12", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_STR_10", z_zbr
        Case SC
            cElement.element "M", "ZBR-6", " - zbrojenie M - dodatek do " & ChrW(216) & "6-8", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_SC_6", z_zbr
            cElement.element "M", "ZBR-10", " - zbrojenie M - dodatek do " & ChrW(216) & "10-12", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_SC_10", z_zbr
        Case POD
            cElement.element "M", "ZBR-6", " - zbrojenie M - dodatek do " & ChrW(216) & "6-8", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_POD_6", z_zbr
            cElement.element "M", "ZBR-10", " - zbrojenie M - dodatek do " & ChrW(216) & "10-12", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_POD_10", z_zbr
        Case SL
            cElement.element "M", "ZBR-10", " - zbrojenie M - dodatek do " & ChrW(216) & "10-12", "t", "=" & colT(R_REINFORCEMENT) & "*ZBR_SL_10", z_zbr
    End Select
End Sub

Public Sub Zbrojenie_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "ZBR", " - zbrojenie R", "t", "=" & colT(R_REINFORCEMENT), z_zbr
End Sub

'------------------------------------------------------------------------
'----------------------------------STAL----------------------------------
'------------------------------------------------------------------------
Public Sub Stal_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "STA", " - stal konstrukcyjna M", "t", "=" & colT(R_REINFORCEMENT), z_sta
End Sub

Public Sub Stal_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "STA", " - stal konstrukcyjna R", "t", "=" & colT(R_REINFORCEMENT), z_sta
End Sub

Public Sub Stal_ppoz_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "ZAB", " - zabezpieczenie stali p.po�. M", "m2", "=" & colT(R_FORMWORK), z_sta
End Sub

Public Sub Stal_ppoz_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "ZAB", " - zabezpieczenie stali p.po�. R", "m2", "=" & colT(R_FORMWORK), z_sta
End Sub
'-------------------------------------------------------------------------------------
'----------------------------------ROBOTY ZIEMNE--------------------------------------
'-------------------------------------------------------------------------------------
Public Sub Wykop_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "WYK", " - wykop pod fundamenty (bez wywozu) R", "m3", "=" & colT(R_CUT), z_zie
End Sub

Public Sub Odwoz_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "ODW", " - wyw�z urobku R", "m3", "=" & colT(R_CUT) & "-" & colT(R_FILL), z_zie
End Sub
    
Public Sub Zasyp_R()
    Dim cElement As New C_ElementP5
    cElement.element "R", "ZAS", " - zasyp fundament�w R", "m3", "=" & colT(R_FILL), z_zie
End Sub

Public Sub Zasyp_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "ZAS", " - zasyp fundament�w M", "m3", "=" & colT(R_FILL), z_zie
End Sub

Public Sub Podsypka_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "POD", " - podsypka keramzytowa M", "m3", "=" & colT(R_VOLUME), z_zie
End Sub

Public Sub Podsypka_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "POD", " - podsypka keramzytowa R", "m3", "=" & colT(R_VOLUME), z_zie
End Sub
'-------------------------------------------------------------------------
'----------------------------------BETON----------------------------------
'-------------------------------------------------------------------------
Public Sub Beton_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", _
    "=""BET""&IFERROR(""-""&MID(" & colT(C_DESCRIPTION) & ",SEARCH("" C*?/??""," & colT(C_DESCRIPTION) & ")+1,SEARCH(""?/??""," & colT(C_DESCRIPTION) & ")+4-(SEARCH("" C""," & colT(C_DESCRIPTION) & ")+1))&IFERROR(MID(" & colT(C_DESCRIPTION) & ",SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1,SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+3-(SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1)),""""),"""")", _
    "="" - beton "" &IFERROR(MID(" & colT(C_DESCRIPTION) & ",SEARCH("" C*?/??""," & colT(C_DESCRIPTION) & ")+1,SEARCH(""?/??""," & colT(C_DESCRIPTION) & ")+4-(SEARCH("" C""," & colT(C_DESCRIPTION) & ")+1))&IFERROR("" ""&MID(" & colT(C_DESCRIPTION) & ",SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1,SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+3-(SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1)),"""")&"" "","""")&""M""", _
    "m3", "=" & colT(R_VOLUME), z_bet
End Sub

Public Sub Pompa()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "POM", " - pompa R", "m3", "=" & colT(R_VOLUME), z_bet
End Sub

Public Sub Beton_Badania()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "BAD", " - badania betonu", "szt.", "=ROUNDUP(" & colT(R_VOLUME) & ",0)", z_bet
End Sub

Public Sub Beton_Dodatki()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "DOD", " - dodatki do betonu", "szt.", "=ROUNDUP(BET_DOD*" & colT(R_VOLUME) & ",0)", z_bet, "=""za�o�ono ""&TEXT(BET_DOD*100,""0"")&""% betonu z dodatkami"""
End Sub

Public Sub Beton_Rozkurz()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "=""BET""&IFERROR(""-""&MID(" & colT(C_DESCRIPTION) & ",SEARCH("" C*?/??""," & colT(C_DESCRIPTION) & ")+1,SEARCH(""?/??""," & colT(C_DESCRIPTION) & ")+4-(SEARCH("" C""," & colT(C_DESCRIPTION) & ")+1))&IFERROR(MID(" & colT(C_DESCRIPTION) & ",SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1,SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+3-(SEARCH("" W8""," & colT(C_DESCRIPTION) & ")+1)),""""),"""")", _
    "="" - rozkurz betonu ""&TEXT(rozkurz,""0%"")", "m3", "=" & colT(R_VOLUME) & "*rozkurz", z_bet
End Sub

'-------------------------------------------------------------------------------------
'----------------------------------ROBOCIZNA - BETON----------------------------------
'-------------------------------------------------------------------------------------
Public Sub BET_CHB()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "CHB", " - chudy beton R", "m3", "=" & colT(R_VOLUME), z_rob
End Sub

Public Sub BET_CHB_Wyrownanie()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "CHB-W", " - chudy beton R - wyr�wnanie", "m2", "=" & colT(R_AREA), z_rob
End Sub

Public Sub BET_Plyta()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "=""BET""&IFS(" & colT(R_THICKNESS) & "<=0.1,""<0,1""," & colT(R_THICKNESS) & "<=0.2,""<0,2""," & colT(R_THICKNESS) & ">0.2,"">0,2"")", _
    " - beton R", "m3", "=" & colT(R_VOLUME), z_rob
End Sub

Public Sub BET_Wyrownanie_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "BET-W-R", " - beton R - wyr�wnanie z r�ki", "m2", "=" & colT(R_AREA), z_rob
End Sub

Public Sub BET_Zatarcie_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "BET-Z-R", " - beton R - zatarcie z r�ki", "m2", "=" & colT(R_AREA), z_rob
End Sub

Public Sub BET_Wygladzenie_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "BET-G-R", " - beton R - wyg�adzenie z r�ki", "m2", "=" & colT(R_AREA), z_rob
End Sub

Public Sub BET_Wygladzenie_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "BET-G-M", " - beton R - wyg�adzenie maszynowe", "m2", "=" & colT(R_AREA), z_rob
End Sub

Public Sub BET_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "BET", " - beton R", "m3", "=" & colT(R_VOLUME), z_rob
End Sub

Public Sub BET_Slup()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "=""BET""&IFS(" & colT(R_AREA) & "<=0.04,""<0,04""," & colT(R_AREA) & "<=0.1,""<0,10""," & colT(R_AREA) & "<=0.25,""<0,25""," & colT(R_AREA) & ">0.25,"">0,25"")", _
    " - beton R", "m3", "=" & colT(R_VOLUME), z_rob
End Sub

Public Sub BET_Sciana()
    Dim cElement As New C_ElementP5

    cElement.element "R", _
    "=""BET""&IFS(" & colT(R_WIDTH) & "<=0.1,""<10""," & colT(R_WIDTH) & "<=0.15,""<15""," & colT(R_WIDTH) & "<=0.2,""<20""," & colT(R_WIDTH) & "<=0.3,""<30""," & colT(R_WIDTH) & "<=0.5,""<50""," & colT(R_WIDTH) & ">0.5,"">50"")", _
    " - beton R", "m3", "=" & colT(R_VOLUME), z_rob
End Sub

Public Sub BET_D_nach()
    Dim cElement As New C_ElementP5

    cElement.element "R", _
    "=""BET-NA""&IFS(" & colT(R_ANGLE) & "<=5,""<5""," & colT(R_ANGLE) & "<=10,""<10""," & colT(R_ANGLE) & "<=15,""<15""," & colT(R_ANGLE) & ">15,"">15"")", _
    " - dodatek - nachylenie", "m2", "=" & colT(R_AREA), z_rob
End Sub

'----------------------------------------------------------------------------------------
'----------------------------------ROBOCIZNA - SZALUNEK----------------------------------
'----------------------------------------------------------------------------------------

Public Sub SZA_Plyta()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", _
    "=""SZA""&IFS(" & colT(R_THICKNESS) & "<=0.12,""<12""," & colT(R_THICKNESS) & "<=0.25,""<25""," & colT(R_THICKNESS) & "<=0.50,""<50""," & colT(R_THICKNESS) & ">0.50,"">50"")", _
    " - szalunek R - szalunek systemowy", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_System()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA", " - szalunek R - szalunek systemowy", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_D_Deski()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA-D", " - dodatek - zamiana na deski", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA", " - szalunek R", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_R_Schody()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", _
    "=""SZA""&IFS(" & colT(R_FORMWORK_HEIGHT) & "<=3.5,""<3,5""," & colT(R_FORMWORK_HEIGHT) & "<=4.5,""<4,5""," & colT(R_FORMWORK_HEIGHT) & "<=6,""<6,0""," & colT(R_FORMWORK_HEIGHT) & ">6,"">6,0"")", _
     " - szalunek R", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_R_Podciag()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "=""SZA""&IFS(" & colT(R_PERIMETER) & "<=0.9,""<0,9""," & colT(R_PERIMETER) & "<=1.5,""<1,5""," & colT(R_PERIMETER) & "<=2.1,""<2,1""," & colT(R_PERIMETER) & ">2.1,"">2,1"")", _
    " - szalunek R", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Sciana_DC()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA-DC", " - dodatek - demonta� i czyszczenie", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Strop()
    Dim cElement As New C_ElementP5

    cElement.element "R", "=""SZA""&IF(" & colT(R_AREA) & "<=100,""<100"","">100"")", _
    " - szalunek R", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_D_DC()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA-DC", " - dodatek - demonta� i czyszczenie", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_D_C()
    Dim cElement As New C_ElementP5

    cElement.element "R", "SZA-C", " - dodatek - czyszczenie", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Strop_D_gr()
    Dim cElement As New C_ElementP5

    cElement.element "R", "SZA-GR", " - dodatek - d > 30 cm", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Strop_D_A()
    Dim cElement As New C_ElementP5

    cElement.element "R", "SZA-A", " - dodatek - A <= 5 m2", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Strop_D_nach()
    Dim cElement As New C_ElementP5

    cElement.element "R", _
    "=""SZA-NA""&IFS(" & colT(R_ANGLE) & "<=15,""<15""," & colT(R_ANGLE) & "<=30,""<30""," & colT(R_ANGLE) & ">30,"">30"")", _
    " - dodatek - nachylenie", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Strop_K()
    Dim cElement As New C_ElementP5

    cElement.element "R", "=""SZA-K""&IFS(" & colT(R_THICKNESS) & "<=0.2,""<20""," & colT(R_THICKNESS) & "<=0.5,""<50""," & colT(R_THICKNESS) & ">0.5,"">50"")", _
    " - szalunek R - kraw�dzie", "m2", "=" & colT(R_FORMWORK2), z_rob
End Sub

Public Sub SZA_Strop_K_D_C()
    Dim cElement As New C_ElementP5

    cElement.element "R", "SZA-K-C", " - kraw�dzie - dodatek - czyszczenie", "m2", "=" & colT(R_FORMWORK2), z_rob
End Sub
Public Sub SZA_Konsola()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "=""SZA""&IF(" & colT(R_AREA) & "<=0.5,""<0,5"","">0,5"")", " - szalunek R", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_Spocznik_Polaczenie()
    Dim cElement As New C_ElementP5

    cElement.element "R", "SZA-P", " - po��czenie z deskowaniem schod�w", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

Public Sub SZA_D_H()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "SZA-H", " - dodatek - wysoko�� powy�ej 3,5 m", "kpl.", _
    "=IF(" & colT(R_FORMWORK_HEIGHT) & "<=3.5,0,CEILING((" & colT(R_FORMWORK_HEIGHT) & "-3.5)/0.25,1))", _
    z_rob
End Sub

Public Sub SZA_Adhezja_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "ADH", " - malowanie warstwy adhezyjnej", "m2", "=" & colT(R_FORMWORK), z_rob
End Sub

'--------------------------------------------------------------------------------------------
'----------------------------------ROBOCIZNA - PREFABRYKATY----------------------------------
'--------------------------------------------------------------------------------------------

Public Sub PREF_R(kod As String, colT_name As String)
    Dim cElement As New C_ElementP5
    If colT_name = R_AREA Then
        cElement.element "R", kod, " - monta� R", "m2", "=" & colT(R_AREA), z_rob
    ElseIf colT_name = R_VOLUME Then
        cElement.element "R", kod, " - monta� R", "m3", "=" & colT(R_VOLUME), z_rob
    ElseIf colT_name = R_COUNT Then
        cElement.element "R", kod, " - monta� R", "szt.", "=" & colT(R_COUNT), z_rob
    Else
        MsgBox "Enter R_AREA, R_VOLUME or R_COUNT"
    End If
End Sub
'-----------------------------------------------------------------------------
'----------------------------------SZALUNEK-----------------------------------
'-----------------------------------------------------------------------------
Public Sub Szalunek_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "SZA", " - szalunek M", "m2", "=" & colT(R_FORMWORK), z_sza
End Sub

Public Sub Szalunek_M_H()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "=""SZA-""&TEXT(CEILING(" & colT(R_FORMWORK_HEIGHT) & ",2),""0"")", _
    " - szalunek M", "m2", "=" & colT(R_FORMWORK), z_sza
End Sub

Public Sub Szalunek_K()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "SZA-K", " - szalunek kraw�dzi M", "m2", "=" & colT(R_FORMWORK_HEIGHT), z_sza
End Sub

'------------------------------------------------------------------------
'----------------------------------INNE----------------------------------
'------------------------------------------------------------------------

Public Sub Izolacja_MR(mr As String, kod As String, Opis As String)
    Dim cElement As New C_ElementP5
    
    cElement.element mr, kod, Opis, "m3", "=" & colT(R_VOLUME), z_str
End Sub

Public Sub Biala_wanna_MR()
    Dim cElement As New C_ElementP5
    
    cElement.element "MR", "BW", " - uszczelnienie betonu MR", "m2", "=" & colT(R_AREA), z_str
End Sub

Public Sub Zab_ppoz_MR()
    Dim cElement As New C_ElementP5
    
    cElement.element "MR", "EI120", " - zabezpieczenie p.po�. EI120 MR", "m.b.", "=" & colT(R_LENGTH), z_str
End Sub

Public Sub Elastomer_MR()
    Dim cElement As New C_ElementP5
    
    cElement.element "MR", "ELA", " - elastomer MR", "m.b.", "=2*" & colT(R_WIDTH), z_str
End Sub

Public Sub Elektronagrzew_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", "ELEK", " - elektronagrzew zbrojenia R", "m3", "=" & colT(R_VOLUME), z_str
End Sub

Public Sub Comaxy_MR()
    Dim cElement As New C_ElementP5
    
    cElement.element "MR", "CMX", " - zbrojenie typu Comax MR", "szt.", "=" & colT(R_COUNT), z_str
End Sub

Public Sub Niezdefiniowane_x()
    Dim cElement As New C_ElementP5
    
    cElement.element "x", "x", " - x x", "x", "=x", z_str
End Sub

Public Sub DOD_Akcesoria_szal()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "SZAL", " - akcesoria szalunkowe", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

Public Sub DOD_Drewno()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "DRE", " - drewno", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

Public Sub DOD_Elastomery()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "ELA", " - elastomery", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

Public Sub DOD_El_zbr()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "ELZBR", " - elementy zbrojarskie", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

Public Sub DOD_Fun_masz()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "FUN", " - fundamenty pod maszyny", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

Public Sub DOD_Wiercenie()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", "WIER", " - wiercenia i ci�cia", "m3", _
    "=SUMIFS(" & Columns(CalcIndex(C_AMOUNT2)).Address(0, 0) & "," & Columns(CalcIndex(C_MR)).Address(0, 0) & ",""M""," & Columns(CalcIndex(C_MATERIAL)).Address(0, 0) & ",""BET*"")", z_str, _
    "m3 betonu konstrukcyjnego"
End Sub

'------------------------------------------------------------------------
'----------------------------------MURY----------------------------------
'------------------------------------------------------------------------
Public Sub Mur_M()
    Dim cElement As New C_ElementP5
    
    cElement.element "M", _
    "=""MUR-""&TEXT(CEILING(" & colT(R_WIDTH) & "*100,1),""00"")", _
    " - mur M", "m2", "=" & colT(R_AREA), z_mur
    
End Sub
Public Sub Mur_R()
    Dim cElement As New C_ElementP5
    
    cElement.element "R", _
    "=""MUR-""&TEXT(CEILING(" & colT(R_UNCONNECT_HEIGHT) & ",1),""0,00"")", _
    " - mur R", "m2", "=" & colT(R_AREA), z_mur
End Sub

'--------------------------------------------------------------------------------
'----------------------------------PREFABRYKATY----------------------------------
'--------------------------------------------------------------------------------
Public Sub Prefabrykat_M_V(kod As String)
    Dim cElement As New C_ElementP5
    
    cElement.element "M", kod, " - prefabrykat M", "m3", "=" & colT(R_VOLUME), z_inny
    
End Sub
Public Sub Prefabrykat_M_A(kod As String)
    Dim cElement As New C_ElementP5
    
    cElement.element "M", kod, " - prefabrykat M", "m2", "=" & colT(R_AREA), z_inny
    
End Sub

