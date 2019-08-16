VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_elementP4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ElementP4(description As String, unit As String, amount As Variant, scope As Variant, Optional comment As Variant)
'
'Add Row in P4
'
    Dim cTable As New C_Table
    
    Add_value C_POZ, "P4"
    Add_value C_DESCRIPTION, description
    Add_value C_UNIT, unit
    Add_value C_AMOUNT, amount
    Add_value C_SCOPE, "zakres " & scope
    Add_UnitPrice
    Add_NetValue
    Add_value C_COMMENTS, comment
    
    cBase
    Cells(CalcRow, CalcIndex(C_UNITPRICE)).Interior.COLOR = S_COLOR_P4(scope)
    Cells(CalcRow, CalcIndex(C_SCOPE)).Interior.COLOR = S_COLOR_P4(scope)
    CalcRow = CalcRow + 1
End Sub

Private Sub cBase()
'
'Level P4
'Poziom P4
'
    'POZ
    'OPT
    Add_LP
    RowUp C_AREA
    RowUp C_LVL
    RowUp C_ELEMENT
    'MR
    'MATERIAL
    'PRICE
    'DESCRIPTION
    'UNIT
    'AMOUNT
    Add_Amount2
    'SCOPE
    Add_UnitPrice
    'VALUE
    'NETVALUE
    'COMMENTS
    
'LASTS COLUMNS
    Add_WorkValue
    Add_Element2

'FORMAT
    cP_Format
    With Selection.Interior
        .ThemeColor = COLOR_P4
        .TintAndShade = COLOR_P4v
    End With
End Sub
