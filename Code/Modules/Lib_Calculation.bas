Attribute VB_Name = "Lib_Calculation"
Public Function col(column)
'
'Shorter way to give cells into formulas
'
        col = Cells(CalcRow, CalcIndex(column)).Address(0, 0)
End Function

Public Function colB(column)
'
'Shorter way to give cells into formulas
'
        colB = Cells(BoQRow, BoQIndex(column)).Address(0, 0)
End Function

Public Function colT(column)
'
'Shorter way to give cells into formulas
'
        colT = Cells(CalcRow_Table, CalcIndex(column)).Address(0, 0)
End Function

Public Sub CalcRange()
'
'Range of Main Table
'
    Range(col(CALC_COLUMNS(WhereInArray(CALC_COLUMNS, CALC_COLUMNS_1(LBound(CALC_COLUMNS_1))) + 2)), col(CALC_COLUMNS(WhereInArray(CALC_COLUMNS, CALC_COLUMNS_1(UBound(CALC_COLUMNS_1))) - 1))).Select
End Sub

Public Sub cP_Format()
'
'Format Main Table
'
    CalcRange
    lib.Borders

    Range(Cells(CalcRow, CalcIndex(C_AMOUNT)), Cells(CalcRow, CalcIndex(C_NETVALUE))).NumberFormat = "#,##0.00"
    Cells(CalcRow, CalcIndex(C_COMMENTS)).Font.ColorIndex = 3

    Cells(CalcRow, CalcIndex(C_SCOPE)).HorizontalAlignment = xlCenter
    Range(Cells(CalcRow, CalcIndex(C_LP)), Cells(CalcRow, CalcIndex(C_PRICE))).HorizontalAlignment = xlCenter
    Cells(CalcRow, CalcIndex(C_UNIT)).HorizontalAlignment = xlCenter
End Sub

Public Sub Add_LP()
'
'LP formula
'
    Add_value C_LP, "=SWITCH(" & cell(CalcRow, CalcIndex(C_POZ)) & "," & _
        """P0"",IF(" & cell(CalcRow - 1, CalcIndex(C_LP)) & "="""",""1."",REPLACE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",1,LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & "),MID(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",1,FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),1))-1)+1&"".""))," & _
        """P1"",IF(FIND("".""," & cell(CalcRow - 1, CalcIndex(C_LP)) & ")=LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")," & cell(CalcRow - 1, CalcIndex(C_LP)) & "&""1."",REPLACE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),1))+1,LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),1)),MID(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),1))+1,FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2))-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),1))-1)+1&"".""))," & _
        """P2"",IF(FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2))=LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")," & cell(CalcRow - 1, CalcIndex(C_LP)) & "&""1."",REPLACE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2))+1,LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2)),MID(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2))+1,FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3))-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),2))-1)+1&"".""))," & _
        """P3"",IF(FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3))=LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")," & cell(CalcRow - 1, CalcIndex(C_LP)) & "&""1."",REPLACE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3))+1,LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3)),MID(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3))+1,FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4))-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),3))-1)+1&"".""))," & _
        """P4"",IF(FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4))=LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")," & cell(CalcRow - 1, CalcIndex(C_LP)) & "&""1."",REPLACE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4))+1,LEN(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4)),MID(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4))+1,FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),5))-FIND(CHAR(1),SUBSTITUTE(" & cell(CalcRow - 1, CalcIndex(C_LP)) & ",""."",CHAR(1),4))-1)+1&"".""))," & _
        "" & cell(CalcRow - 1, CalcIndex(C_LP)) & ")"
End Sub

Public Sub Add_Element2()
'
'ELEMENT do P4
'
    Add_value C_ELEMENT2, "=IF(" & cell(CalcRow - 1, CalcIndex(C_POZ)) & "=""P3""," & cell(CalcRow - 1, CalcIndex(C_DESCRIPTION)) & "," & cell(CalcRow - 1, CalcIndex(C_ELEMENT2)) & ")"
End Sub

Public Sub Add_NetValue()
'
'NETVALUE
'
    Add_value C_NETVALUE, "=SUMIFS(" & column(C_VALUE) & "," & column(C_LP) & "," & col(C_LP) & "&""*"")"
End Sub

Public Sub Add_WorkValue()
'
'WORKVALUE in P4
'Wartosc prac w P4
'
    Add_value C_WORKVALUE, "=" & col(C_AMOUNT) & "*" & col(C_UNITPRICE)
End Sub

Public Sub Add_ValueK()
'
'VALUE in P4
'
    Add_value C_VALUE, "=" & col(C_AMOUNT2) & "*" & col(C_UNITPRICE)
End Sub

Public Sub Add_UnitPrice()
'
'UNITPRICE
'
    Add_value C_UNITPRICE, "=IFERROR(" & col(C_NETVALUE) & "/" & col(C_AMOUNT2) & ",0)"
End Sub

Public Sub Add_Amount2()
'
'AMOUNT2
'
    Add_value C_AMOUNT2, "=IF(" & col(C_OPT) & "=""""," & col(C_AMOUNT) & ",IF(ISERROR(VLOOKUP(" & col(C_OPT) & ",optymalizacje,1,0)),0," & col(C_AMOUNT) & "))"
End Sub

Public Sub RowUp(column)
'
'Formula from upper cell
'Formula do komórki wy¿ej
'
    Add_value column, "=" & cell(CalcRow - 1, CalcIndex(column))
End Sub

Public Sub Add_value(column, value)
'
'Shorter way to add value int ocell
'
        Cells(CalcRow, CalcIndex(column)).Formula = value
End Sub

Private Function column(columnName)
'
'Shorter way to give columns into formulas
'
        column = Columns(CalcIndex(columnName)).Address(0, 0)
End Function
