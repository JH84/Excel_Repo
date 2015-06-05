Attribute VB_Name = "ActivCLFormula"
Sub FindcellFormula()

    ActiveCell.formula = "=IFERROR(INDEX('New PL Data'!R7C3:R103C40,MATCH('Plant data'!RC1,'New PL Data'!R7C1:R103C1,0),MATCH('Plant data'!R1C,'New PL Data'!R6C3:R6C40,0)),""ND"")"
    
    ActiveCell.Select
    Selection.Copy
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(97, 38)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
End Sub

Sub FindcellFormulaKWH()

    ActiveCell.FormulaR1C1 = _
        "=IFERROR(SUMIFS('New kWh data'!R2C15:R385C15,'New kWh data'!R2C1:R385C1,[@[Date & Time T]],'New kWh data'!R2C6:R385C6,Table26[[#Headers],[Bottling Plant]:[Site Total]]),""ND"")"
    
    ActiveCell.Select
    Selection.Copy
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(47, 6)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
End Sub

