VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_BoQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public BoQRow As Long

Private Sub Add_val_B(colBumn, value)
'
'Shorter way to add value int ocell
'
        Cells(BoQRow, BoQIndex(colBumn)).Interior.ColorIndex = COLOR_BOQ_EDIT
        Cells(BoQRow, BoQIndex(colBumn)).Formula = value
End Sub

Public Sub SpreadFoundation()
'
'Stopa fundamentowa - 5D-KI-STB-OB-FU-EF
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Foundation Slab"
            'NOTHING TO DO
         
        Case "5dki Einzelfundament Fundament"
            Add_val_B R_PERIMETER, "=2*(" & colB(R_WIDTH) & "+" & colB(R_LENGTH) & ")"
            
    End Select
    
    Add_val_B R_5D4D_CODE, "FST"
    Add_val_B R_THICKNESS, foundThickness
    Add_val_B R_FORMWORK, thickness_M_Perimeter
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_NAME, "Stopa fundamentowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m, ""&TEXT(MIN(" & colB(R_WIDTH) & "," & colB(R_LENGTH) & "),""0,00"")&""x""&TEXT(MAX(" & colB(R_WIDTH) & "," & colB(R_LENGTH) & "),""0,00"")&""m"""
'FORMAT
    Cells(BoQRow, BoQIndex(R_REINFORCEMENT_INDICATOR)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Public Sub StripFoundation()
'
'£awa fundamentowa - 5D-KI-STB-OB-FU-SF
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Foundation Slab"
            Add_val_B R_THICKNESS, foundThickness
            
        Case "5dki Streifenfundament Skelettbau"
            Add_val_B R_THICKNESS, hohe_DE
            Add_val_B R_WIDTH, breite_DE
            Add_val_B R_AREA, "=" & colB(R_WIDTH) & "*" & colB(R_LENGTH_CUT)
            Add_val_B R_PERIMETER, "=2*(" & colB(R_WIDTH) & "+" & colB(R_LENGTH_CUT) & ")"
    End Select

    Add_val_B R_5D4D_CODE, "FLW"
    Add_val_B R_FORMWORK, thickness_M_Perimeter
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_NAME, "£awa fundamentowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", B=""&TEXT(" & colB(R_WIDTH) & ",""0,00"")&""m, H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
'FORMAT
    Cells(BoQRow, BoQIndex(R_REINFORCEMENT_INDICATOR)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Public Sub FoundationSlab()
'
'P³yta/Rampa fundamentowa
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Foundation Slab"
            Add_val_B R_THICKNESS, foundThickness
    End Select
    
    Add_val_B R_5D4D_CODE, "FPL"
    Add_val_B R_FORMWORK, thickness_M_Perimeter
    Add_val_B R_REINFORCEMENT, reinforcement
    'NAME DEPENDS OF TYPE
    If InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "5dki Bodenplattenrampe") <> 0 Then
        Add_val_B R_NAME, "Rampa ¿elbetowa"
    ElseIf InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "5dki Bodenplatte") <> 0 Then
        Add_val_B R_NAME, "P³yta fundamentowa"
    Else
        Add_val_B R_NAME, "-1"
    End If

    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&IF(" & colB(R_MATERIAL_TYPE) & "<>"""","" "",)&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
End Sub

Public Sub FoundationSlab_PitElevator()
'
'P³yta fundamentowa podszybia
'
    FoundationSlab
    Add_val_B R_NAME, "P³yta fundamentowa podszybia"
End Sub

Public Sub FoundationSlab_Machine()
'
'Fundament pod maszyny
'
    FoundationSlab
    Add_val_B R_NAME, "Fundament pod maszyny"
End Sub

Public Sub FoundationSlab_FiberReinforced()
'
'P³yta fundamentowa ze zbrojeniem rozproszonym
'
    FoundationSlab
    Add_val_B R_NAME, "P³yta fundamentowa ze zbrojeniem rozproszonym"
End Sub

Public Sub FoundationSlab_RollerConcrete()
'
'P³yta fundamentowa z betonu wa³owanego
'
    FoundationSlab
    Add_val_B R_NAME, "P³yta fundamentowa z betonu wa³owanego"
End Sub

Public Sub Wall_Concrete()
'
'Œciana ¿elbetowa
'
    Add_val_B R_5D4D_CODE, "SCZ"
    Add_val_B R_FORMWORK, "=2*" & colB(R_LENGTH) & "*" & colB(R_UNCONNECT_HEIGHT)
    Add_val_B R_FORMWORK_HEIGHT, "=" & colB(R_UNCONNECT_HEIGHT)
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_INTERVAL_HEIGHT, intervalHeight
    Add_val_B R_NAME, "Œciana ¿elbetowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&IF(" & colB(R_MATERIAL_TYPE) & "="""","""","" ""&" & colB(R_MATERIAL_TYPE) & ")&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, B=""&TEXT(" & colB(R_WIDTH) & ",""0,00"")&""m"""
End Sub

Public Sub Wall_Concrete_1S()
'
'Œciana ¿elbetowa jednostronnie szalowana
'
    Wall_Concrete
    Add_val_B R_NAME, "Œciana ¿elbetowa jednostronnie szalowana"
    Add_val_B R_FORMWORK, "=" & colB(R_LENGTH) & "*" & colB(R_UNCONNECT_HEIGHT)
End Sub

Public Sub Wall_Duct()
'
'Œciana ¿elbetowa szachtu
'
    Wall_Concrete
    Add_val_B R_5D4D_CODE, "SCS"
    Add_val_B R_NAME, "Œciana ¿elbetowa szachtu"
End Sub

Public Sub Wall_Disc()
'
'Tarcza ¿elbetowa
'
    Wall_Concrete
    Add_val_B R_NAME, "Tarcza ¿elbetowa"
End Sub

Public Sub Wall_Attic()
'
'Attyka/nadci¹g ¿elbetowy
'
    Wall_Concrete
    Add_val_B R_NAME, "Attyka ¿elbetowa"
    
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "5dki Überzug Skelettbau"
            Add_val_B R_WIDTH, breite_DE
            Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
    End Select
    
    Add_val_B R_5D4D_CODE, "ATT"
End Sub

Public Sub Wall_ElevatorPit()
'
'Œciana ¿elbetowa podszybia
'
    Wall_Concrete
    Add_val_B R_NAME, "Œciana ¿elbetowa podszybia"
End Sub

Public Sub Wall_Filigran()
'
'Œciana prefabrykowana typu Filigran
'
    Wall_Concrete
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefabrykowana œciana typu Filigran"
    Add_val_B R_FORMWORK, ""
    Cells(BoQRow, BoQIndex(R_FORMWORK)).Interior.ColorIndex = 0
End Sub

Public Sub Beam()
'
'Podci¹g ¿elbetowy - 5D-KI-STB-OB-TR-UZ
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Basic Wall"
            Add_val_B R_LENGTH_CUT, "=" & colB(R_LENGTH)
            
        Case "5dki Unterzug Skelettbau"
            Add_val_B R_WIDTH, breite_DE
            Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
    End Select
    
    Add_val_B R_5D4D_CODE, "POD"
    Add_val_B R_FORMWORK, "=" & colB(R_LENGTH_CUT) & "*(" & colB(R_WIDTH) & "+2*" & colB(R_UNCONNECT_HEIGHT) & ")"
    Add_val_B R_REINFORCEMENT, "=" & colB(R_VOLUME2) & "*0.001*" & colB(R_REINFORCEMENT_INDICATOR)
    Add_val_B R_INTERVAL_HEIGHT, intervalHeight
    Add_val_B R_NAME, "Podci¹g ¿elbetowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, B=""&TEXT(" & colB(R_WIDTH) & ",""0,00"")&""m, H=""&TEXT(" & colB(R_UNCONNECT_HEIGHT) & ",""0,00"")&""m (H.cal=""&TEXT(" & colB(R_UNCONNECT_HEIGHT) & "+" & colB(R_COMMENT) & ",""0,00"")&""m)"""
    Add_val_B R_VOLUME2, "=" & colB(R_WIDTH) & "*(" & colB(R_UNCONNECT_HEIGHT) & "+" & colB(R_COMMENT) & ")*" & colB(R_LENGTH_CUT)
    Add_val_B R_PERIMETER, "=2*(" & colB(R_WIDTH) & "+" & colB(R_UNCONNECT_HEIGHT) & ")"
    
    Cells(BoQRow, BoQIndex(R_VOLUME2)).ClearComments
    Cells(BoQRow, BoQIndex(R_VOLUME2)).AddComment "Objetoœæ podci¹gu z pó³k¹"
    Cells(BoQRow, BoQIndex(R_COMMENT)).ClearComments
    Cells(BoQRow, BoQIndex(R_COMMENT)).AddComment "Wysokoœæ pó³ki"
'FORMAT
    Cells(BoQRow, BoQIndex(R_REINFORCEMENT_INDICATOR)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    Cells(BoQRow, BoQIndex(R_FORMWORK_HEIGHT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    Cells(BoQRow, BoQIndex(R_COMMENT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Private Sub BrickWall()
'
'Œciana murowana
'
    Add_val_B R_5D4D_CODE, "SCM"
    
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&IF(" & colB(R_MATERIAL_TYPE) & "="""","""","" ""&" & colB(R_MATERIAL_TYPE) & ")&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, B=""&TEXT(" & colB(R_WIDTH) & ",""0,00"")&""m"""
    Add_val_B R_INTERVAL_HEIGHT, intervalHeight
    Add_val_B R_FORMWORK_HEIGHT, "=" & colB(R_UNCONNECT_HEIGHT)
End Sub

Public Sub BrickWall_Porotherm()
'
'Œciana z Porothermu
'
    BrickWall
    Add_val_B R_NAME, "Œciana z Porothermu"
End Sub

Public Sub BrickWall_Ytong()
'
'Œciana z bloczków Ytong
'
    BrickWall
    Add_val_B R_NAME, "Œciana z bloczków Ytong"
End Sub

Public Sub BrickWall_SilkaSmall()
'
'Œciana z silikatów drobnowymiarowych, np Silka E
'
    BrickWall
    Add_val_B R_NAME, "Œciana z silikatów drobnowymiarowych"
End Sub

Public Sub BrickWall_SilkaBig()
'
'Œciana z silikatów wielkowymiarowych, np Silka Tempo
'
    BrickWall
    Add_val_B R_NAME, "Œciana z silikatów wielkowymiarowych"
End Sub

Public Sub BrickWall_SilkaSzacht()
'
'Obudowa szachtu z silikatów
'
    BrickWall
    Add_val_B R_NAME, "Obudowa szachtu z silikatów"
End Sub

Public Sub BrickWall_TeknoAmber()
'
'Œciana z TeknoAmerBlok
'
    BrickWall
    Add_val_B R_NAME, "Œciana z TeknoAmberBlok"
End Sub

Public Sub BrickWall_Tang()
'
'Trzpieñ œciany murowanej
'
    Add_val_B R_5D4D_CODE, "TRZ"

    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_THICKNESS, tiefe_DE
    Add_val_B R_NAME, "Trzpieñ œciany murowanej"
    Add_val_B R_FORMWORK, "=2*(" & colB(R_THICKNESS) & "+2*0.15)*" & colB(R_LENGTH)
    Add_val_B R_FORMWORK_HEIGHT, "=" & colB(R_LENGTH)
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&IF(" & colB(R_MATERIAL_TYPE) & "="""","""","" ""&" & colB(R_MATERIAL_TYPE) & ")&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&"", B=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
    Add_val_B R_INTERVAL_HEIGHT, "=""do ""&IF(" & colB(R_LENGTH) & "<=3.5,""3,50"",TEXT(CEILING(" & colB(R_LENGTH) & ",0.25),""0,00""))"
    Add_val_B R_REINFORCEMENT, reinforcement
End Sub

Public Sub BrickWall_Grommet()
'
'Wieniec œciany murowanej
'
    Add_val_B R_5D4D_CODE, "TRZ"
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
    Add_val_B R_NAME, "Wieniec œciany murowanej"
    Add_val_B R_FORMWORK, "=2*(" & colB(R_UNCONNECT_HEIGHT) & "+2*0.15)*" & colB(R_LENGTH)
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&IF(" & colB(R_MATERIAL_TYPE) & "="""","""","" ""&" & colB(R_MATERIAL_TYPE) & ")&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", H=""&TEXT(" & colB(R_UNCONNECT_HEIGHT) & ",""0,00"")&""m"""
    Add_val_B R_REINFORCEMENT, reinforcement
End Sub

Public Sub LeanConcrete()
'
'Chudy beton
'
    Add_val_B R_5D4D_CODE, "CHB"
    
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Foundation Slab"
            Add_val_B R_THICKNESS, foundThickness
            
        Case "5dki Sauberkeitsschicht unter Vouten"
            Add_val_B R_THICKNESS, "=" & colB(R_DICKE_DE)
    End Select

    Add_val_B R_NAME, "Chudy beton"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")"
End Sub

Public Sub LeanConcrete_Fulfil()
'
'Chudy beton wype³niaj¹cy
'
    Add_val_B R_5D4D_CODE, "CHB"
    LeanConcrete
    Add_val_B R_NAME, "Chudy beton wype³niaj¹cy"
End Sub

Public Sub SlantingEdges()
'
'Skoœna krawêdŸ p³yty
'
    Add_val_B R_5D4D_CODE, "KRA"
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_MATERIAL_STRUCT, "=" & colB(R_MATERIAL)
    Add_val_B R_NAME, "Skoœna krawêdŸ p³yty"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")"
End Sub

Public Sub colBumn()
'
'S³up
'
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_INTERVAL_HEIGHT, "=""do ""&IF(" & colB(R_LENGTH) & "<=3.5,""3,50"",TEXT(CEILING(" & colB(R_LENGTH) & ",0.25),""0,00""))"
    Add_val_B R_FORMWORK_HEIGHT, "=" & colB(R_LENGTH)
End Sub

Public Sub Column_Rec()
'
'S³up ¿elbetowy prostok¹tny
'
    colBumn
    Add_val_B R_5D4D_CODE, "SLP"
    Add_val_B R_THICKNESS, tiefe_DE
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_FORMWORK, "=" & colB(R_LENGTH) & "*(" & colB(R_THICKNESS) & "+" & colB(R_WIDTH) & ")*2"
    Add_val_B R_AREA, "=" & colB(R_THICKNESS) & "*" & colB(R_WIDTH)
    Add_val_B R_NAME, "S³up ¿elbetowy prostok¹tny"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, ""&TEXT(MIN(" & colB(R_THICKNESS) & "," & colB(R_WIDTH) & "),""0,00"")&""x""&TEXT(MAX(" & colB(R_THICKNESS) & "," & colB(R_WIDTH) & "),""0,00"")&""m"""
End Sub

Public Sub Column_Round()
'
'S³up ¿elbetowy okr¹g³y
'
    colBumn
    Add_val_B R_5D4D_CODE, "SLO"
    Add_val_B R_THICKNESS, "=" & colB(R_DIAMETER)
    Add_val_B R_FORMWORK, "=" & colB(R_LENGTH) & "*2*PI()*" & colB(R_DIAMETER) & "/2"
    Add_val_B R_AREA, "=PI()*(" & colB(R_DIAMETER) & "/2)^2"
    Add_val_B R_NAME, "S³up ¿elbetowy okr¹g³y"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, " & ChrW(216) & "=""&TEXT(" & colB(R_DIAMETER) & ",""0,00"")&""m"""
End Sub

Public Sub Column_Rec_Pilaser()
'
'Pilaster ¿elbetowy
'
    Column_Rec
    Add_val_B R_5D4D_CODE, "SLP"
    Add_val_B R_NAME, "Pilaster ¿elbetowy"
End Sub

Public Sub FlaringHeads()
'
'G³owica ¿elbetowa
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "Floor"
            Add_val_B R_FORMWORK, "=" & colB(R_AREA)
            Add_val_B R_VOLUME2, "=" & colB(R_VOLUME) & "+" & colB(R_COMMENT) & "*" & colB(R_AREA)
            Add_val_B R_FORMWORK2, "=" & colB(R_PERIMETER) & "*" & colB(R_THICKNESS)
            
            Cells(BoQRow, BoQIndex(R_FORMWORK2)).ClearComments
            Cells(BoQRow, BoQIndex(R_FORMWORK2)).AddComment "Powierzchnia szalunku bocznego g³owicy"
    
        Case "5dki STB-Stützenkopf rund"
            Add_val_B R_THICKNESS, hohe_DE
            Add_val_B R_AREA, "=PI()*(" & colB(R_DIAMETER) & "/2)^2"
            Add_val_B R_VOLUME2, "=" & colB(R_VOLUME) & "+PI()*(" & colB(R_DIAMETER) & "/2+" & colB(R_THICKNESS) & "/TAN(" & colB(R_SLOPE) & "))^2*" & colB(R_COMMENT)
            Add_val_B R_FORMWORK, "=PI()*(" & colB(R_DIAMETER) & "+" & colB(R_THICKNESS) & "/TAN(" & colB(R_SLOPE) & "))*" & colB(R_THICKNESS) & "/SIN(" & colB(R_SLOPE) & ")"
            Add_val_B R_SLOPE, "=RADIANS(MID(" & colB(R_ANGLE) & ",1,FIND(""" & ChrW(176) & """," & colB(R_ANGLE) & ")-1))"

        Case "5dki STB-Stützenkopf rechteckig"
            Add_val_B R_THICKNESS, hohe_DE
            Add_val_B R_WIDTH, tiefe_DE
            Add_val_B R_LENGTH, breite_DE

            Add_val_B R_AREA, "=" & colB(R_LENGTH) & "*" & colB(R_WIDTH)
            Add_val_B R_VOLUME2, "=" & colB(R_VOLUME) & "+(" & colB(R_WIDTH) & "+2*" & colB(R_THICKNESS) & "/TAN(" & colB(R_SLOPE) & "))*(" & colB(R_LENGTH) & "+2*" & colB(R_THICKNESS) & "/TAN(" & colB(R_SLOPE) & "))*" & colB(R_COMMENT)
            Add_val_B R_PERIMETER, "=2*(4*" & colB(R_THICKNESS) & "/TAN(" & colB(R_SLOPE) & ")+" & colB(R_WIDTH) & "+" & colB(R_LENGTH) & ")"
            Add_val_B R_FORMWORK, "=" & colB(R_PERIMETER) & "*" & colB(R_THICKNESS) & "/SIN(" & colB(R_SLOPE) & ")"
            Add_val_B R_SLOPE, "=RADIANS(MID(" & colB(R_ANGLE) & ",1,FIND(""" & ChrW(176) & """," & colB(R_ANGLE) & ")-1))"
    End Select
    
    Add_val_B R_5D4D_CODE, "STR"
    Add_val_B R_REINFORCEMENT, "=" & colB(R_VOLUME2) & "*0.001*" & colB(R_REINFORCEMENT_INDICATOR)
    Add_val_B R_INTERVAL_HEIGHT, intervalHeight
    Add_val_B R_NAME, "G³owica ¿elbetowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""

    Cells(BoQRow, BoQIndex(R_VOLUME2)).ClearComments
    Cells(BoQRow, BoQIndex(R_VOLUME2)).AddComment "Objetoœæ g³owicy z pó³k¹"
    Cells(BoQRow, BoQIndex(R_COMMENT)).ClearComments
    Cells(BoQRow, BoQIndex(R_COMMENT)).AddComment "Wysokoœæ pó³ki"
'FORMAT
    Cells(BoQRow, BoQIndex(R_REINFORCEMENT_INDICATOR)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    Cells(BoQRow, BoQIndex(R_FORMWORK_HEIGHT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    Cells(BoQRow, BoQIndex(R_COMMENT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Public Sub Landing()
'
'Spocznik ¿elbetowy
'
    Add_val_B R_5D4D_CODE, "SPO"
    Add_val_B R_FORMWORK, "=" & colB(R_AREA)
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_NAME, "Spocznik ¿elbetowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
End Sub

Public Sub Stairs()
'
'Bieg schodowy
'
    Add_val_B R_5D4D_CODE, "SCH"
    Add_val_B R_WIDTH, "=" & colB(R_LAUFBREITE_DE)
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_NAME, "Bieg schodowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")"
'FORMAT
    Cells(BoQRow, BoQIndex(R_FORMWORK_HEIGHT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    Cells(BoQRow, BoQIndex(R_FORMWORK)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Public Sub Prefabric_Stairs()
'
'Prefabrykowany bieg schodowy
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_WIDTH, "=" & colB(R_LAUFBREITE_DE)
    Add_val_B R_NAME, "Prefabrykowany bieg schodowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME)
End Sub
Public Sub Prefabric_Landing()
'
'Prefarykowany spocznik - 5D-KI-STB-FT-TRE-POD
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefabrykowany spocznik"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
End Sub
Public Sub Floor()
'
'Strop ¿elbetowy
'
    Add_val_B R_5D4D_CODE, "STR"
    Add_val_B R_FORMWORK, "=" & colB(R_AREA)
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_INTERVAL_HEIGHT, intervalHeight
    Add_val_B R_FORMWORK2, thickness_M_Perimeter
    Add_val_B R_ANGLE, "=DEGREES(ATAN(" & colB(R_SLOPE) & "))"
    Add_val_B R_NAME, "Strop ¿elbetowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m""&IF(" & colB(R_SLOPE) & "="""","""","", spadek ""&TEXT(" & colB(R_SLOPE) & ",""0%""))"
    
    Cells(BoQRow, BoQIndex(R_FORMWORK2)).ClearComments
    Cells(BoQRow, BoQIndex(R_FORMWORK2)).AddComment "Powierzchnia szalunku bocznego stropu"
'FORMAT
    Cells(BoQRow, BoQIndex(R_FORMWORK_HEIGHT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub
Public Sub Floor_Filigran()
'
'Strop prefabrykowany
'
    Floor
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefabrykowany strop typu Filigran"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&IF(" & colB(R_PRESTRESSED) & "=""Yes"","" sprê¿ony"","""")&"" ""&" & colB(R_INTERVAL_HEIGHT) & "&""m, H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m""&IF(" & colB(R_SLOPE) & "="""","""","", spadek ""&TEXT(" & colB(R_SLOPE) & ",""0%""))"
End Sub

Public Sub Bracket()
'
'Wspornik - 5D-KI-STB-OB-M-KB
'
    Add_val_B R_5D4D_CODE, "WSP"
    Add_val_B R_THICKNESS, tiefe_DE
    Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
    Add_val_B R_LENGTH, "=" & colB(R_VOLUME) & "/" & colB(R_THICKNESS) & "/" & colB(R_UNCONNECT_HEIGHT)
    Add_val_B R_AREA, "=" & colB(R_THICKNESS) & "*" & colB(R_UNCONNECT_HEIGHT)
    Add_val_B R_FORMWORK, "=" & colB(R_LENGTH) & "*(" & colB(R_THICKNESS) & "+" & colB(R_UNCONNECT_HEIGHT) & ")"
    Add_val_B R_REINFORCEMENT, reinforcement
    Add_val_B R_NAME, "Wspornik ¿elbetowy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE) & "&IF(" & colB(R_WATERPROOF) & "=""Yes"","" W8"","""")&"", B=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m, H=""&TEXT(" & colB(R_UNCONNECT_HEIGHT) & ",""0,00"")&""m"""
End Sub

Public Sub Console()
'
'Konsola ¿elbetowa - 5D-KI-STB-OB-M-KON
'
    Bracket
    Add_val_B R_5D4D_CODE, "KON"
    Add_val_B R_LENGTH, breite_DE
    Add_val_B R_NAME, "Konsola ¿elbetowa"
End Sub

Public Sub Foundation_VibrationInsulation()
'
'Mata antywibracyjna pod fundamenty pod maszyny
'
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_THICKNESS, foundThickness
    Add_val_B R_NAME, "Mata antywibracyjna pod fundamenty pod maszyny"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_THICKNESS) & "*100&""cm"""
End Sub

Public Sub PartitionPanel()
'
'Dylatacja ze styropianu
'
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_NAME, "Dylatacja ze styropianu"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_WIDTH) & "*100&""cm"""
End Sub

Public Sub InteriorInsulation()
'
'Izolacja wewnêtrzna
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "5dki Wärmedämmung innen vertikal Linie"
            Add_val_B R_WIDTH, breite_DE
            Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
            Add_val_B R_LENGTH, "=" & colB(R_VOLUME) & "/" & colB(R_WIDTH) & "/" & colB(R_UNCONNECT_HEIGHT)
    End Select
    
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_NAME, "Izolacja wewnêtrzna"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&IF(" & colB(R_THICKNESS) & "," & colB(R_THICKNESS) & "," & colB(R_WIDTH) & ")*100&""cm"""
End Sub

Public Sub WallInsulation()
'
'Izolacja œciany
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "5dki Perimeterdämmung Wand vertikal Linie"
            Add_val_B R_WIDTH, breite_DE
            Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
            Add_val_B R_LENGTH, "=" & colB(R_VOLUME) & "/" & colB(R_WIDTH) & "/" & colB(R_UNCONNECT_HEIGHT)
    End Select
    
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_NAME, "Izolacja œciany"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&IF(" & colB(R_THICKNESS) & "," & colB(R_THICKNESS) & "," & colB(R_WIDTH) & ")*100&""cm"""
End Sub

Public Sub Foundation_Insulation()
'
'Izolacja fundamentu
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "5dki Perimeterdämmung unter Vouten"
            Add_val_B R_THICKNESS, "=" & colB(R_DICKE_DE)
    End Select
    
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_NAME, "Izolacja fundamentu"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&IF(" & colB(R_THICKNESS) & "," & colB(R_THICKNESS) & "," & colB(R_WIDTH) & ")*100&""cm"""
End Sub

Public Sub FoamGlass()
'
'Podsypka keramzytowa
'
    Select Case Cells(BoQRow, BoQIndex(R_FAMILY)).Text
        Case "5dki Schaumglasschotter unter Vouten"
            Add_val_B R_THICKNESS, "=" & colB(R_DICKE_DE)
    End Select
    
    Add_val_B R_5D4D_CODE, "KER"
    Add_val_B R_NAME, "Podsypka keramzytowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_THICKNESS) & "*100&""cm"""
End Sub

Public Sub ElastomericSlidingWasher_Strip()
'
'Podk³adka elastomerowa pasmowa
'
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_NAME, "Podk³adka elastomerowa pasmowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_WIDTH) & "*100&""cm"""
End Sub

Public Sub Elastomeric_Pref_Stairs()
'
'Podk³adka elastomerowa pasmowa boczna pod schody
'
    Add_val_B R_5D4D_CODE, "IZO"
    Add_val_B R_NAME, "Podk³adka elastomerowa pod schody"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME)
End Sub

Public Sub ElastomericSlidingWasher_Point()
'
'Podk³adka elastomerowa punktowa
'
    Add_val_B R_5D4D_CODE, "IZO"
    
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_NAME, "Podk³adka elastomerowa punktowa"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_WIDTH) & "*100&""cm"""
    
'Name depends of Type
    If InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "5dki Punkt-Gleitlager") <> 0 Then
        Add_val_B R_NAME, "Podk³adka elastomerowa punktowa sta³a"

    ElseIf InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "5dki Punkt-Festlager") <> 0 Then
        Add_val_B R_NAME, "Podk³adka elastomerowa punktowa œlizgowa"

    Else
        Add_val_B R_NAME, "-1"
        
    End If
End Sub

Public Sub Comax_T1()
'
'Zbrojenie Comax jednorzêdowe
'
    Comax
    Add_val_B R_NAME, "Zbrojenie Comax jednorzêdowe"
End Sub

Public Sub Comax_T5()
'
'Zbrojenie Comax dwurzêdowe
'
    Comax
    Add_val_B R_NAME, "Zbrojenie Comax dwurzêdowe"
End Sub
Public Sub Comax()
'
'Zbrojenie Comax
'
    Add_val_B R_5D4D_CODE, "CMX"
    Add_val_B R_THICKNESS, hohe_DE
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_LENGTH, "=" & colB(R_VOLUME) & "/" & colB(R_WIDTH) & "/" & colB(R_THICKNESS)
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME)
End Sub

Public Sub Prefabric_Beam()
'
'Prefarykowana belka - 5D-KI-STB-FT-TR-UZ
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_UNCONNECT_HEIGHT, hohe_DE
    Add_val_B R_NAME, "Prefabrykowana belka"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" B=""&TEXT(" & colB(R_WIDTH) & ",""0,00"")&""m, H=""&TEXT(" & colB(R_UNCONNECT_HEIGHT) & ",""0,00"")&""m"""
End Sub

Public Sub Prefabric_Slab()
'
'Prefarykowany strop
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefabrykowany strop"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
End Sub

Public Sub Prefabric_Slab_Hollow()
'
'Prefarykowany strop kana³owy
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefabrykowany strop kana³owy"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" H=""&TEXT(" & colB(R_THICKNESS) & ",""0,00"")&""m"""
End Sub

Public Sub Prefabric_SlabTT()
'
'Prefarykowana p³yta TT z nadbetonem - 5D-KI-STB-FT-DE-TTP
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_NAME, "Prefarykowana p³yta TT z nadbetonem"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&" & colB(R_MATERIAL_TYPE)
'FORMAT
    Cells(BoQRow, BoQIndex(R_FORMWORK_HEIGHT)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
End Sub

Public Sub Prefabric_Column()
'
'Prefarykowany slup - 5D-KI-STB-FT-ST-REC
'
    Add_val_B R_5D4D_CODE, "PREF"
    Add_val_B R_THICKNESS, tiefe_DE
    Add_val_B R_WIDTH, breite_DE
    Add_val_B R_AREA, "=" & colB(R_THICKNESS) & "*" & colB(R_WIDTH)
    Add_val_B R_INTERVAL_HEIGHT, "=""do ""&IF(" & colB(R_LENGTH) & "<=3.5,""3,50"",TEXT(CEILING(" & colB(R_LENGTH) & ",0.25),""0,00""))"
    Add_val_B R_NAME, "Prefarykowany s³up"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME) & "&"" ""&TEXT(MIN(" & colB(R_THICKNESS) & "," & colB(R_WIDTH) & "),""0,00"")&""x""&TEXT(MAX(" & colB(R_THICKNESS) & "," & colB(R_WIDTH) & "),""0,00"")&""m"""
End Sub

Public Sub Steel()
'
'Konstrukcja stalowa
'

'Name depends of Family
    If InStr(Cells(BoQRow, BoQIndex(R_FAMILY)).Text, "Schweißprofil") <> 0 Then
    
        If InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "Verbund") <> 0 Then
            Add_val_B R_NAME, "Konstrukcja stalowa spawana z wype³nieniem betonowym"
        Else
            Add_val_B R_NAME, "Konstrukcja stalowa spawana"
        End If
            
    Else
        If InStr(Cells(BoQRow, BoQIndex(R_TYPE)).Text, "Verbund") <> 0 Then
            Add_val_B R_NAME, "Konstrukcja stalowa z wype³nieniem betonowym"
        Else
            Add_val_B R_NAME, "Konstrukcja stalowa"
        End If
            
    End If

    Add_val_B R_5D4D_CODE, "STA"
    Add_val_B R_MATERIAL_TYPE, "=IFERROR(MID(" & colB(R_TYPE) & ",(IFERROR(FIND(""S355""," & colB(R_TYPE) & "),Find(""S235""," & colB(R_TYPE) & "))),4),-1)"
    Add_val_B R_PROFILE, "=IFERROR(MID(" & colB(R_TYPE) & ",1,find(" & colB(R_MATERIAL_TYPE) & "," & colB(R_TYPE) & ")-2),-1)"
    Add_val_B R_REINFORCEMENT, "=IFERROR(IF(" & colB(R_LENGTH_CUT) & "," & colB(R_LENGTH_CUT) & "," & colB(R_LENGTH) & ")*1.1*VLOOKUP(" & colB(R_TYPE) & ",STA_Profile,5,0),)/1000"
    Add_val_B R_FORMWORK, "=IFERROR(IF(" & colB(R_LENGTH_CUT) & "," & colB(R_LENGTH_CUT) & "," & colB(R_LENGTH) & ")*1.1*VLOOKUP(" & colB(R_TYPE) & ",STA_Profile,6,0),)"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME)
End Sub

Public Sub EarthMoving()
'
'Roboty ziemne
'
    Add_val_B R_5D4D_CODE, "ZIE"
    
    Add_val_B R_NAME, "Roboty ziemne"
    Add_val_B R_NAME_FINAL, "=" & colB(R_NAME)
End Sub

'---------------------------------------------------------------------------------------------
'------------------------------------------FUNCTIONS------------------------------------------
'---------------------------------------------------------------------------------------------

Private Function reinforcement()
    Cells(BoQRow, BoQIndex(R_REINFORCEMENT_INDICATOR)).Interior.ColorIndex = COLOR_BOQ_MUST_HAVE
    reinforcement = "=" & colB(R_VOLUME) & "*0.001*" & colB(R_REINFORCEMENT_INDICATOR)
End Function

Private Function thickness_M_Perimeter()
    thickness_M_Perimeter = "=" & colB(R_THICKNESS) & "*" & colB(R_PERIMETER)
End Function

Private Function foundThickness()
    foundThickness = "=" & colB(R_FOUND_THICKNESS)
End Function

Private Function hohe_DE()
    hohe_DE = "=" & colB(R_HOHE_DE)
End Function

Private Function breite_DE()
    breite_DE = "=" & colB(R_BREITE_DE)
End Function

Private Function tiefe_DE()
    tiefe_DE = "=" & colB(R_TIEFE_DE)
End Function

Private Function intervalHeight()
    intervalHeight = "=""do ""&IF(" & colB(R_FORMWORK_HEIGHT) & "<=3.5,""3,50"",TEXT(CEILING(" & colB(R_FORMWORK_HEIGHT) & ",0.25),""0,00""))"
End Function

