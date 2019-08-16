VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dNAME As Object
Private dCOLOR As Object
Private dCOLOR2 As Object

Const c_P0 = 0
Const c_P1 = 1
Const c_P2 = 2
Const c_P3 = 3
Const c_P4 = 4
Const c_P5 = 5

Private Sub Class_Initialize()
'
'Description and units
'Opis i jednostki - DO ZMIANY W PRZYPADKU ZMIAN KODU
'
    Set dNAME = CreateObject("Scripting.Dictionary")
    Set dCOLOR = CreateObject("Scripting.Dictionary")
    Set dCOLOR2 = CreateObject("Scripting.Dictionary")
    
             'KEY    NAME    COLOR   COLOR2
             
    Add_value z_str, "STRABAG", 40, 1
    
    Add_value z_rob, "Roboczogodziny", 35, 1
    Add_value z_bet, "Beton 1", 17, 1
    Add_value z_sza, "Szalunki 1", 24, 1
    Add_value z_zbr, "Zbrojenie 1", 45, 1
    Add_value z_zie, "Roboty ziemne 1", 22, 1
    
    Add_value z_ber, "Berlinka 1", 26, 1
    Add_value z_roz, "Rozbiorki 1", 42, 1
    Add_value z_sta, "Konstrukcja stalowa 1", 50, 1
    Add_value z_mur, "Mury 1", 27, 1
    Add_value z_inny, "Prefabrykaty 1", 12, 1
    
    Add_value z_11, "Zakres11a", 43, 1
    Add_value z_12, "Zakres12a", 44, 1
    
    Add_value c_P0, "P0", xlThemeColorAccent2, 0
    Add_value c_P1, "P1", xlThemeColorAccent6, 0.4
    Add_value c_P2, "P2", xlThemeColorAccent1, 0.4
    Add_value c_P3, "P3", xlThemeColorAccent2, 1
    Add_value c_P4, "P4", xlThemeColorAccent2, 0.6
    Add_value c_P5, "P5", xlThemeColorAccent2, 0.8
End Sub

Sub Add_value(KEY As Variant, name As String, COLOR As Variant, color2 As Variant)
    dNAME.Add item:=name, KEY:=KEY
    dCOLOR.Add item:=COLOR, KEY:=KEY
    dCOLOR2.Add item:=color2, KEY:=KEY
End Sub

Function GetNAME(KEY As String)
    If dNAME.Exists(KEY) Then
        GetNAME = dNAME(KEY)
    Else
        GetNAME = -1 'or raise an error...
    End If
End Function

Function GetCOLOR(KEY As String)
    If dCOLOR.Exists(KEY) Then
        GetCOLOR = dCOLOR(KEY)
    Else
        GetCOLOR = -1 'or raise an error...
    End If
End Function

Function GetCOLOR2(KEY As String)
    If dCOLOR2.Exists(KEY) Then
        GetCOLOR2 = dCOLOR2(KEY)
    Else
        GetCOLOR2 = -1 'or raise an error...
    End If
End Function

