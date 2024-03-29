VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ElementP5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub element(mr As String, material As String, description As String, unit As String, amount As Variant, scope As Variant, Optional comment As Variant)
'
'Add Row in P5
'
    Dim cTable As New C_Table
    Add_value C_MR, mr
    Add_value C_MATERIAL, material
    Add_value C_DESCRIPTION, description
    Add_value C_UNIT, unit
    Add_value C_AMOUNT, amount
    Add_value C_UNITPRICE, "=VLOOKUP(" & cell(CalcRow, CalcIndex(C_PRICE)) & "," & z_nazwa & scope & "_ceny,HLOOKUP(" & cell(Range(SCOPE_START).Offset(scope, 1).row, Range(SCOPE_START).Offset(scope, 1).column, True) & "," & z_nazwa & scope & ",2,0),0)"
    Add_value C_SCOPE, "zakres " & scope
    Add_value C_COMMENTS, comment
    
    cBase
    Cells(CalcRow, CalcIndex(C_UNITPRICE)).Interior.COLOR = S_COLOR_P5(scope)
    Cells(CalcRow, CalcIndex(C_SCOPE)).Interior.COLOR = S_COLOR_P5(scope)
    
    CalcRow = CalcRow + 1
End Sub

Private Sub cBase()
'
'Level P5
'Poziom P5
'
    'POZ
    'OPT
    Add_LP
    RowUp C_AREA
    RowUp C_LVL
    RowUp C_ELEMENT
    'MR
    'MATERIAL
    Add_value C_PRICE, "=" & col(C_ELEMENT) & "&"".""&" & col(C_MR) & "&"".""&" & col(C_MATERIAL)
    'DESCRIPTION
    'UNIT
    'AMOUNT
    Add_Amount2
    'SCOPE
    'UNITPRICE
    Add_ValueK
    'NETVALUE
    'COMMENTS
    
'LASTS COLUMNS
    Add_WorkValue
    Add_Element2

'FORMAT
    cP_Format
    With Selection.Interior
        .ThemeColor = COLOR_P5
        .TintAndShade = COLOR_P5v
    End With
End Sub
