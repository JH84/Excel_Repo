Attribute VB_Name = "InsertTextColumns"

Public noMsgBox As Boolean

Sub InsertTextCol()
Attribute InsertTextCol.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Inserts Date & Time Text Col in Col A

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[1],""yyyy/mm/dd""&"" ""&""hh:mm"")"
    Range("A7").Select
    Selection.Copy
    Range("A7:A343").Select
    ActiveSheet.Paste
    Range("C4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C&"" ""&R[2]C"
    Range("C4").Select
    Selection.Copy
    Range("C5").Select
    Selection.End(xlToRight).Offset(-1).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveSheet.Paste
    
End Sub
Sub InsertTextColBlimp()
'
' Inserts Date & Time Text Col in Col A

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[1],""yyyy/mm/dd""&"" ""&""hh:mm"")"
    Range("A4").Select
    Selection.Copy
    Range("A4:A342").Select
    ActiveSheet.Paste
    Range("b1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C&"" ""&R[3]C"
    Range("b1").Select
    Selection.Copy
    Range("b2").Select
    Selection.End(xlToRight).Offset(-1).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveSheet.Paste
    
End Sub
Sub InsertTextColRPRO()
'
' Inserts Date & Time Text Col in Col A

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[1],""yyyy/mm/dd""&"" ""&""hh:mm"")"
    Range("A4").Select
    Selection.Copy
    Range("A4:A342").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("c1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[2]C&"" ""&R[3]C"
    Range("c1").Select
    Selection.Copy
    Range("c2").Select
    Selection.End(xlToRight).Offset(-1).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveSheet.Paste
    
End Sub
Sub InsTxtColBlimpIDs()
'
' Inserts Date & Time Text Col in Col A

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[1],""yyyy/mm/dd""&"" ""&""hh:mm"")"
    Range("A4").Select
    Selection.Copy
    Range("A4:A342").Select
    ActiveSheet.Paste
    Range("b1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[2]C&"" ""&R[4]C"
    Range("b1").Select
    Selection.Copy
    Range("b2").Select
    Selection.End(xlToRight).Offset(-1).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    ActiveSheet.Paste
    
End Sub
Sub InsertTextCol_XLS_CSV()
'
' Inserts Date & Time Text Col in Col A

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[1],""yyyy/mm/dd""&"" ""&""hh:mm"")"
    Range("A5").Select
    Selection.Copy
    Range("A5:A341").Select
    ActiveSheet.Paste
    
End Sub
Sub Find_Data_Pilot()

noMsgBox = True
On Error Resume Next

Call CSVHillcrest
Call BlimpdataHillcrest

Call CopydataAllGalloManor
Call BlimpdataGalloManor

Call CopydataAllBedfordview
Call BlimpDataBedfordview

Call CopydataAllKensington
Call BlimpdataKensington

Call CopydataAllPinelands
Call BlimpdataPinelands

Call CSVTyger

Call CSVCW

Call CSVWonderboom

noMsgBox = False

MsgBox "All Imports Complete"

End Sub
