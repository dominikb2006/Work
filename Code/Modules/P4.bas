Attribute VB_Name = "P4"
Private Function chck(column)
'
'Shorter way to give cells for check Text
'
        chck = Cells(CalcRow_Table, CalcIndex(column)).Text
End Function

'-------------------------------------------------------------------------------------
'----------------------------------ROBOTY ZIEMNE--------------------------------------
'-------------------------------------------------------------------------------------
Public Sub Wykop_i_odwoz()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Wykop i odw�z", "m3", "=" & colT(R_CUT), z_zie
    Wykop_R
    Odwoz_R
End Sub
Public Sub Zasyp()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Zasyp", "m3", "=" & colT(R_FILL), z_zie
    Zasyp_R
    Zasyp_M
End Sub

Public Sub Podsypka()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Podsypka keramzytowa", "m3", "=" & colT(R_VOLUME), z_zie
    Podsypka_M
    Podsypka_R
End Sub
'-------------------------------------------------------------------------
'----------------------------------BETON----------------------------------
'-------------------------------------------------------------------------
Public Sub Beton()
    Dim cEleP4 As New C_elementP4
    
    Select Case colT(C_ELEMENT)
        Case "CHB"
            cEleP4.ElementP4 "Beton", "m3", "=" & colT(R_VOLUME), z_bet, _
            "=IFERROR(""gr. ""&TEXT(IF(" & colT(R_THICKNESS) & "<>0," & colT(R_THICKNESS) & "," & colT(R_VOLUME) & "/" & colT(R_AREA) & ")*100,""0"")&"" cm"","""")"
        Case Else
            cEleP4.ElementP4 "Beton", "m3", "=" & colT(R_VOLUME), z_bet
    End Select
    
    Beton_M
    Pompa
    Beton_Badania
    Beton_Dodatki
    Beton_Rozkurz
End Sub

'-----------------------------------------------------------------------------
'----------------------------------SZALUNEK-----------------------------------
'-----------------------------------------------------------------------------
Public Sub Szalunek()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Szalunek", "m2", "=" & colT(R_FORMWORK), z_sza
    
    Szalunek_M
End Sub

Public Sub Szalunek_H()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Szalunek", "m2", "=" & colT(R_FORMWORK), z_sza
    
    Szalunek_M_H
End Sub

Public Sub Szalunek_H_K()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Szalunek", "m2", "=" & colT(R_FORMWORK), z_sza
    
    Szalunek_M_H
    Szalunek_K
End Sub

'-----------------------------------------------------------------------------
'----------------------------------ZBROJENIE----------------------------------
'-----------------------------------------------------------------------------
Public Sub Zbrojenie(element As ZBR_Type)
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Zbrojenie", "t", "=" & colT(R_REINFORCEMENT), z_zbr, _
    "=TEXT(" & colT(R_REINFORCEMENT_INDICATOR) & ",""0"")&"" kg/m3"""
    
    Zbrojenie_M element
    Zbrojenie_R
End Sub

'------------------------------------------------------------------------
'----------------------------------STAL----------------------------------
'------------------------------------------------------------------------
Public Sub Stal()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Stal konstrukcyjna", "t", "=" & colT(R_REINFORCEMENT), z_sta
    
    Stal_M
    Stal_R
End Sub
Public Sub Stal_ppoz()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Zabezpieczenie p.po�. stali", "m2", "=" & colT(R_FORMWORK), z_sta
    
    Stal_ppoz_M
    Stal_ppoz_R
End Sub

'------------------------------------------------------------------------
'----------------------------------INNE----------------------------------
'------------------------------------------------------------------------
Public Sub Biala_wanna()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Bia�a wanna", "kpl.", "1", z_str
    
    Biala_wanna_MR
End Sub

Public Sub Zab_ppoz()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Zabezpieczenie p.po�.", "kpl.", "1", z_str
    
    Zab_ppoz_MR
End Sub

Public Sub Elektronagrzew()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Elektronagrzew", "kpl.", "1", z_str
    
    Elektronagrzew_R
End Sub


Public Sub Elastomer()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Podk�adka elastomerowa", "m.b.", "=2*" & colT(R_WIDTH), z_str
    
    Elastomer_MR
End Sub
Public Sub Dodatki_do_betonu()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Dodatki do betonu", "kpl.", "1", z_str
    
    DOD_Akcesoria_szal
    DOD_Drewno
    DOD_Elastomery
    DOD_El_zbr
    DOD_Fun_masz
    DOD_Wiercenie
End Sub

Public Sub Izolacja()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Izolacja", "kpl.", "1", z_str
    
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "styropian") <> 0 Then
        Izolacja_MR "M", "STY", " - styropian M"
        Izolacja_MR "R", "STY", " - styropian R"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "antywibracyjna") <> 0 Then
        Izolacja_MR "M", "AWIB", " - mata antywibracyjna M"
        Izolacja_MR "R", "AWIB", " - mata antywibracyjna R"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "elastomerowa") <> 0 Then
        Elastomer_MR
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "Izolacja") <> 0 Then
        Izolacja_MR "M", "IZO", " - izolacja M"
        Izolacja_MR "R", "IZO", " - izolacja R"
    Else
        Izolacja_MR "M", "x", " - x M"
        Izolacja_MR "R", "x", " - x R"
    End If
End Sub

Public Sub Comaxy()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Zbrojenie typu Comax", "szt.", "1", z_str
    
    Comaxy_MR
End Sub

Public Sub Niezdefiniowane()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Niezdefiniowane", "x", "x", z_str
    
    Niezdefiniowane_x
End Sub

'-------------------------------------------------------------------------------
'----------------------------------PREFABRKATY----------------------------------
'-------------------------------------------------------------------------------
Public Sub Prefabrykat()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Prefabrykat", "m3", "=" & colT(R_VOLUME), z_inny
    
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "s�up") <> 0 Then
        Prefabrykat_M_V "SLP"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "�ciana typu Filigran") <> 0 Then
        Prefabrykat_M_A "FIL-SC"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "spocznik") <> 0 Then
        Prefabrykat_M_A "SPO"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "bieg schodowy") <> 0 Then
        Prefabrykat_M_V "SCH"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "belka") <> 0 Then
        Prefabrykat_M_V "BEL"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop kana�owy") <> 0 Then
        Prefabrykat_M_A "KAN"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop") <> 0 Then
        Prefabrykat_M_A "STR"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "p�yta TT") <> 0 Then
        Prefabrykat_M_A "TT"
    ElseIf InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop typu Filigran") <> 0 Then
        Prefabrykat_M_A "FIL-STR"
    End If
End Sub

'------------------------------------------------------------------------
'----------------------------------MURY----------------------------------
'------------------------------------------------------------------------
Public Sub Mur()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "�ciana murowana", "m2", "=" & colT(R_AREA), z_mur
    
    Mur_M
    Mur_R
End Sub

'-----------------------------------------------------------------------------
'----------------------------------ROBOCIZNA----------------------------------
'-----------------------------------------------------------------------------
Public Sub Rob_Chudy_Beton()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_CHB
    BET_CHB_Wyrownanie
End Sub

Public Sub Rob_Plyta_fund()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_Plyta
    BET_Wyrownanie_R
    BET_Zatarcie_R
    BET_Wygladzenie_R
    SZA_Adhezja_M
    SZA_Plyta
    SZA_D_Deski
End Sub

Public Sub Rob_Lawa_fund()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    BET_Wyrownanie_R
    BET_Zatarcie_R
    SZA_Adhezja_M
    SZA_System
    SZA_D_Deski
End Sub

Public Sub Rob_Stopa_fund()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    BET_Wyrownanie_R
    BET_Zatarcie_R
    SZA_Adhezja_M
    SZA_System
End Sub

Public Sub Rob_Krawedz()
    Dim cEleP4 As New C_elementP4
    
    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
End Sub

Public Sub Rob_Sciana()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m2", "=" & colT(R_AREA), z_rob
    
    BET_Sciana
    SZA_Adhezja_M
    SZA_R
    SZA_Sciana_DC
    If chck(R_FORMWORK_HEIGHT) > 3.5 Then SZA_D_H
End Sub

Public Sub Rob_Attyka()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m2", "=" & colT(R_AREA), z_rob
    
    BET_Sciana
    SZA_Adhezja_M
    SZA_R
    SZA_Sciana_DC
End Sub

Public Sub Rob_Podciag()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    SZA_Adhezja_M
    SZA_R_Podciag
    SZA_D_DC
    If chck(R_FORMWORK_HEIGHT) > 3.5 Then SZA_D_H
End Sub

Public Sub Rob_Slup_prost()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_Slup
    SZA_Adhezja_M
    SZA_R
    SZA_D_DC
    If chck(R_FORMWORK_HEIGHT) > 3.5 Then SZA_D_H
End Sub

Public Sub Rob_Slup_okrag()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_Slup
    SZA_Adhezja_M
    SZA_R
    If chck(R_FORMWORK_HEIGHT) > 3.5 Then SZA_D_H
End Sub

Public Sub Rob_Konsola()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    SZA_Adhezja_M
    SZA_Konsola
End Sub

Public Sub Rob_Wspornik()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m2", "=" & colT(R_AREA), z_rob
    
    BET_Plyta
    BET_Wyrownanie_R
    BET_Zatarcie_R
    BET_Wygladzenie_R
    SZA_Adhezja_M
    SZA_R
End Sub

Public Sub Rob_Strop()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m2", "=" & colT(R_AREA), z_rob
    
    BET_Plyta
    BET_Wyrownanie_R
    BET_Zatarcie_R
    BET_Wygladzenie_R
    If IsNumeric(chck(R_ANGLE)) And chck(R_ANGLE) > 5 Then BET_D_nach
    SZA_Adhezja_M
    SZA_Strop
    SZA_D_DC
    If chck(R_THICKNESS) > 0.3 Then SZA_Strop_D_gr
    If chck(R_AREA) <= 5 Then SZA_Strop_D_A
    If IsNumeric(chck(R_ANGLE)) And chck(R_ANGLE) > 15 Then SZA_Strop_D_nach
    If chck(R_FORMWORK_HEIGHT) > 3.5 Then SZA_D_H
    SZA_Strop_K
    SZA_Strop_K_D_C
End Sub

Public Sub Rob_Spocznik()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    BET_Wyrownanie_R
    BET_Zatarcie_R
    BET_Wygladzenie_R
    SZA_Adhezja_M
    SZA_R
    SZA_D_C
    SZA_Spocznik_Polaczenie
End Sub

Public Sub Rob_Wieniec()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    SZA_Adhezja_M
    SZA_R
    SZA_D_Deski
End Sub

Public Sub Rob_Schody()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "m3", "=" & colT(R_VOLUME), z_rob
    
    BET_R
    SZA_Adhezja_M
    SZA_R_Schody
End Sub

Public Sub Rob_Pref()
    Dim cEleP4 As New C_elementP4

    cEleP4.ElementP4 "Robocizna", "kpl.", "=1", z_rob
    
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "s�up") <> 0 Then PREF_R "SLP", R_COUNT
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "spocznik") <> 0 Then PREF_R "SPO", R_COUNT
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "bieg schodowy") <> 0 Then PREF_R "SCH", R_COUNT
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "belka") <> 0 Then PREF_R "BEL", R_VOLUME
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop") <> 0 Then PREF_R "STR", R_AREA
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop kana�owy") <> 0 Then PREF_R "KAN", R_AREA
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "p�yta TT") <> 0 Then PREF_R "TT", R_AREA
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "�ciana typu Filigran") <> 0 Then PREF_R "FIL-SC", R_AREA
    If InStr(Cells(CalcRow_Table, CalcIndex(C_DESCRIPTION)).Text, "strop typu Filigran") <> 0 Then PREF_R "FIL-STR", R_AREA
End Sub
