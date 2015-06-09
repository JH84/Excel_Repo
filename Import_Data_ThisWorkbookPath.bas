Attribute VB_Name = "ImportsData"
Sub All_data_Import()

Dim LastCell As Range
Dim InputFname As String
Dim InputFnameCSV As String
Dim Shtname As String

Application.ScreenUpdating = False

Windows("case_room_summary.xlsm").Activate
Worksheets("Imported").Select

    Application.CutCopyMode = False
    Range("d5:eis344").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear

SelectedFname = Sheets("graphs").Range("b1")
InputFname = SelectedFname ' & " checked"
InputFnameCSV = InputFname & ".csv"

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\" & InputFname & ".csv", local:=True)

    wbook.Activate   ' change to csv file name
    Sheets(1).Name = "Data"

'Call InsertTextCol_Answers
'Call Add_range_names_Import_ALL

    Range("b1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'Set LastCell = ActiveCell.UsedRange.SpecialCells(xlLastCell)
'Set Data_range = Range("c355", LastCell)
   
Windows("case_room_summary.xlsm").Activate
Worksheets("Imported").Select
   
    Range("d5").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("Imported").Select
Workbooks(InputFnameCSV).Close savechanges:=False

End Sub
Sub Answers_ranges()

Dim LastCell As Range
Dim InputFname As String
Dim InputFnameCSV As String
Dim Shtname As String

Application.ScreenUpdating = False

SelectedFname = Sheets("graphs").Range("b1")
InputFname = SelectedFname & " checked"
InputFnameCSV = InputFname & ".csv"

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\completed\" & InputFname & ".csv", local:=True)

wbook.Activate   ' change to csv file name
Sheets(1).Name = "Data"

Call InsertTextCol_Answers
Call Add_range_names

'Set LastCell = ActiveCell.UsedRange.SpecialCells(xlLastCell)
'Set Data_range = Range("c355", LastCell)

Windows("case_room_summary.xlsm").Activate
Worksheets("Imported").Select

    Application.CutCopyMode = False
    Range("d355").Select
    
    ActiveCell.Formula = _
        "=IFERROR(INDEX('[" & InputFnameCSV & "]data'!DataRange,MATCH(RC3,'[" & InputFnameCSV & "]data'!DateCol,0),MATCH(R354C,'[" & InputFnameCSV & "]data'!DeviceList,0)),""ND"")"
    
    Range("d355").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("imported").Select
Workbooks(InputFnameCSV).Close savechanges:=False

End Sub


'# Unused...







Sub ImportWommDirect()

Application.ScreenUpdating = False

'Filename:=ThisWorkbook.Path & "\Graphs\" & myFileName, Filtername:="PNG"

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\mountain mill.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("mountain mill").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("Womm").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('Mountain Mill.csv'!R6C3:R342C610,MATCH(RC3,'Mountain Mill.csv'!R6C1:R342C1,0),MATCH(R2C,'Mountain Mill.csv'!R1C3:R1C610,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:me344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("womm").Select
'MsgBox "Import Complete"

Workbooks("mountain mill.csv").Close savechanges:=False

End Sub
Sub ImportPMBCDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\pmb central.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("pmb central").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("pmbc").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('pmb central.csv'!R6C3:R342C598,MATCH(RC3,'pmb central.csv'!R6C1:R342C1,0),MATCH(R2C,'pmb central.csv'!R1C3:R1C598,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:ka344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("pmbc").Select
'MsgBox "Import Complete"

Workbooks("pmb central.csv").Close savechanges:=False

End Sub
Sub ImportSowetoDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\Soweto Maponya.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("Soweto Maponya").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("soweto").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('Soweto Maponya.csv'!R6C3:R342C1020,MATCH(RC3,'Soweto Maponya.csv'!R6C1:R342C1,0),MATCH(R2C,'Soweto Maponya.csv'!R1C3:R1C1020,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:sn344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("soweto").Select
'MsgBox "Import Complete"

Workbooks("Soweto Maponya.csv").Close savechanges:=False

End Sub
Sub ImportCenturionDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\Centurion Hyper.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("Centurion Hyper").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("centurion").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('Centurion Hyper.csv'!R6C3:R342C1047,MATCH(RC3,'Centurion Hyper.csv'!R6C1:R342C1,0),MATCH(R2C,'Centurion Hyper.csv'!R1C3:R1C1047,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:sw344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("centurion").Select
'MsgBox "Import Complete"

Workbooks("Centurion Hyper.csv").Close savechanges:=False

End Sub
Sub ImportBluffDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\the Bluff.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("The Bluff").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("Bluff").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('The Bluff.csv'!R6C3:R342C535,MATCH(RC3,'The Bluff.csv'!R6C1:R342C1,0),MATCH(R2C,'The Bluff.csv'!R1C3:R1C535,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:fw344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("bluff").Select
'MsgBox "Import Complete"

Workbooks("The Bluff.csv").Close savechanges:=False

End Sub
Sub ImportLibertyDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\Liberty Mall.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("Liberty Mall").Select  ' change to first tab name

Call InsertTextCol

Windows("case_room_summary.xlsm").Activate

Worksheets("Liberty").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX('Liberty Mall.csv'!R6C3:R342C761,MATCH(RC3,'Liberty Mall.csv'!R6C1:R342C1,0),MATCH(R2C,'Liberty Mall.csv'!R1C3:R1C761,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:ot344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("Liberty").Select
'MsgBox "Import Complete"

Workbooks("Liberty Mall.csv").Close savechanges:=False

End Sub
Sub ImportHemmingwaysDirect()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\hemmingways.csv", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("hemmingways").Select  ' change to first tab name

Windows("case_room_summary.xlsm").Activate

Worksheets("hemmingways").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX([Hemmingways.csv]hemmingways!R4C3:R340C2401,MATCH(RC3,[Hemmingways.csv]hemmingways!R4C1:R339C1,0),MATCH(R2C,[Hemmingways.csv]hemmingways!R2C3:R2C2401,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:or344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("hemmingways").Select
'MsgBox "Import Complete"

Workbooks("hemmingways.csv").Close savechanges:=False

End Sub
Sub ImportFVCPinehurst()

Application.ScreenUpdating = False

Set wbook = Application.Workbooks.Open(ThisWorkbook.Path & "\data\input data\hemmingways.xlsx", local:=True)

wbook.Activate   ' change to csv file name
Worksheets("hemm").Select  ' change to first tab name

Windows("case_room_summary.xlsm").Activate

Worksheets("hemmingways").Select

    Application.CutCopyMode = False
    Range("d9").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(INDEX([Hemmingways.xlsx]Hemm!R4C3:R340C2401,MATCH(RC3,[Hemmingways.xlsx]Hemm!R4C1:R339C1,0),MATCH(R2C,[Hemmingways.xlsx]Hemm!R1C3:R1C2401,0)),""ND"")"
    
    Range("d9").Select
    Selection.Copy
    Range("d9:or344").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Worksheets("hemmingways").Select
'MsgBox "Import Complete"

Workbooks("hemmingways.xlsx").Close savechanges:=False

End Sub

