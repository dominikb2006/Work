VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dNAME As Object
Private dUNIT As Object
Private dAMOUNT As Object

Private Sub Class_Initialize()
'
'Description and units
'Opis i jednostki - DO ZMIANY W PRZYPADKU ZMIAN KODU
'
    Set dNAME = CreateObject("Scripting.Dictionary")
    Set dUNIT = CreateObject("Scripting.Dictionary")
    Set dAMOUNT = CreateObject("Scripting.Dictionary")
    
             'KEY    NAME             UNIT  AMOUNT
             
    Add_value "ZIE", "Roboty ziemne", "m3", R_CUT
    Add_value "SZCZ", "�cianka szczelinowa", "m3", R_VOLUME
    Add_value "CHB", "Chudy beton", "m3", R_VOLUME
    Add_value "KRA", "Kraw�dzie p�yt", "m3", R_VOLUME
    Add_value "FLW", "�awy fundamentowe", "m3", R_VOLUME
    Add_value "FPL", "P�yty fundamentowe", "m3", R_VOLUME
    Add_value "FST", "Stopy fundamentowe", "m3", R_VOLUME
    Add_value "POD", "Podci�gi �elbetowe", "m3", R_VOLUME
    Add_value "KON", "Konsole �elbetowe", "m3", R_VOLUME
    Add_value "SCH", "Schody �elbetowe", "szt.", R_COUNT
    Add_value "SPO", "Spoczniki �elbetowe", "m2", R_AREA
    Add_value "RMP", "Rampy �elbetowe", "m2", R_AREA
    Add_value "SCZ", "�ciany �elbetowe", "m2", R_AREA
    Add_value "SCS", "�ciany �elbetowe szachtu", "m2", R_AREA
    Add_value "SLP", "S�upy �elbetowe prostok�tne", "m3", R_VOLUME
    Add_value "SLO", "S�upy �elbetowe okr�g�e", "m3", R_VOLUME
    Add_value "STR", "Stropy �elbetowe", "m2", R_AREA
    Add_value "SCM", "�ciany murowane", "m2", R_AREA
    Add_value "DOD", "Elementy dodatkowe", "kpl.", "1"
    Add_value "USZ", "Uszczelnienia", "kpl.", "1"
    Add_value "PREF", "Prefabrykaty", "kpl.", "1"
    Add_value "PRZE", "Przes�ony poziome", "kpl.", "1"
    Add_value "STA", "Konstrukcje stalowe", "t", R_REINFORCEMENT
    Add_value "GRZ", "Grzybki �elbetowe", "m3", R_VOLUME
    Add_value "TRZ", "Trzpienie i wie�ce �cian murowanych", "m.b.", R_LENGTH
    Add_value "CMX", "Zbrojenie Comax", "szt.", R_COUNT
    Add_value "IZO", "Izolacje", "kpl.", "1"
    Add_value "WSP", "Wsporniki �elbetowe", "m3", R_VOLUME
    Add_value "ATT", "Attyki �elbetowe", "m2", R_AREA
    Add_value "KER", "Podsypka keramzytowa", "m3", R_VOLUME
    
'    Add_value "ROZ", "Rozbi�rki", "kpl.", "1"
'    Add_value "BER", "�cianka berli�ska", "m2", R_AREA
End Sub

Private Sub Add_value(KEY As Variant, name As Variant, unit As Variant, Optional amount As Variant)
    dNAME.Add item:=name, KEY:=KEY
    dUNIT.Add item:=unit, KEY:=KEY
    dAMOUNT.Add item:=amount, KEY:=KEY
End Sub

Function GetNAME(KEY As String)
    If dNAME.Exists(KEY) Then
        GetNAME = dNAME(KEY)
    Else
        GetNAME = -1 'or raise an error...
    End If
End Function

Function GetUNIT(KEY As String)
    If dUNIT.Exists(KEY) Then
        GetUNIT = dUNIT(KEY)
    Else
        GetUNIT = -1 'or raise an error...
    End If
End Function

Function GetAMOUNT(KEY As String)
    If dAMOUNT.Exists(KEY) Then
        GetAMOUNT = dAMOUNT(KEY)
    Else
        GetAMOUNT = C_TAKEOFF 'or raise an error...
    End If
End Function
