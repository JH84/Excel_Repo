Attribute VB_Name = "FindCells"
Sub FindPlusOffsetPLANT()

Worksheets("plant data").Select

Dim rng As Range
Dim cl As Range
Dim sFind As String
 
sFind = Range("a1").Value
 
Set rng = Range("a2:a12000")
Set cl = rng.Find(sFind, LookIn:=xlValues)
If Not cl Is Nothing Then cl.Offset(0, 2).Activate

Call FindcellFormula

End Sub

Sub FindPlusOffsetKWH()

Worksheets("2014 kWh").Select

Dim rng As Range
Dim cl As Range
Dim sFind As String
 
sFind = Range("b1").Value
 
Set rng = Range("b2:b17524")
Set cl = rng.Find(sFind, LookIn:=xlValues)
If Not cl Is Nothing Then cl.Offset(0, 1).Activate

Call FindcellFormulaKWH

End Sub
