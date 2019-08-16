Attribute VB_Name = "Main"
Sub Main()
    On Error Resume Next

    Application.Volatile False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ImportBIM.Main
    BoQ.Main
    Table.Main
    Calculation.Main
    lib.Main
    
    Application.Calculate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
'
'Sub Test()
''    On Error Resume Next
'
'    Application.Volatile False
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.DisplayAlerts = False
'
'    ImportBIM.Test
'    BoQ.Main
'    Table.Main
'    Calculation.Main
'    lib.Test
'
'    Application.Calculate
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
'End Sub
'
'Sub Reset()
'
'    Application.Volatile False
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    Renaming A_PRICE_LIST, A_CALCULATION, A_CALCULATION2
'    lib.UsuniecieKalkulacja
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'End Sub
'
'Sub TurnOnCalc()
'    Application.Calculate
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
'End Sub
