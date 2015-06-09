Attribute VB_Name = "CFpersite"
Sub PageOneConditions() ' All sheets

On Error GoTo Errhandler
    
    With Range("b2,n64:n65,r64,s6:u7")
        .FormatConditions.Delete
    End With
    
    Dim SavingsTargetGreen As Range
    Dim SavingsTargetOrange As Range
    Dim SavingstargetRed As Range
    
        Set SavingsTargetGreen = Range("m68")
        Set SavingsTargetOrange = Range("m69")
        Set SavingstargetRed = Range("m70")
    
    Range("n64:n65").Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    Select Case Range("n71").Value
        Case 1
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 3
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select

Range("s6:u7,b2,r64").Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    Select Case Range("n71").Value
        Case 1
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 3
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
                       
    End Select

Errhandler:

    Select Case Err
    
        Case 13: MsgBox "One of the sites has missing energy data, please make sure all the kWh values are copied in and try again" & vbNewLine & "The Script will now stop."
        End
    End Select
    
End Sub
Sub FormatRange(ByVal target As String, ByVal formula As String)

    Range(target).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub FormatRangeGreen(ByVal target As String, ByVal formula As String)

    Range(target).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Interior.Color = 3407718
        .Font.Color = vbBlack
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub FormatRangeRed(ByVal target As String, ByVal formula As String)

    Range(target).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Interior.Color = 255
        .Font.Color = vbWhite
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub FormatRangeGray(ByVal target As String, ByVal formula As String)

    Range(target).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:=formula
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Sub CFPinehurstFV()

' page 1

    Application.ScreenUpdating = False
    On Error GoTo ErrhandlerCF

    Worksheets("pine p1").Select

    With Range("b2,r52:u64")
        .FormatConditions.Delete
    End With

    Call FormatRange("r52:u52", "=(r69:u69)=1")   ' LT & MT rack colour
    Call FormatRangeRed("r52:u52", "=(r69:u69)=2")   ' LT & MT rack colour
    Call FormatRange("r55:u55", "=(r70:u70)=1")   'room colour check
    Call FormatRangeRed("r55:u55", "=(r70:u70)=2")   ' room colour check
    Call FormatRange("r58:u58", "=(r71:u71)=1")   ' case colour check
    Call FormatRangeRed("r58:u58", "=(r71:u71)=2")   ' case colour check
    Call FormatRange("r61", "=(r72)=1")   ' PRC colour check
    Call FormatRangeRed("r61", "=(r72)=2")   ' PRC colour check
    Call PageOneConditions

' page 2

    Worksheets("pine db").Select
    
    With Range("e5:s9")
        .FormatConditions.Delete
    End With

    Call FormatRange("e5:s7", "=(e76:r79)=1")   ' check for orange
    Call FormatRangeRed("e5:s7", "=(e76:r79)=2")   ' check for red
    
    Range("a1").Select
    
    Select Case Range("a72").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
    Selection.FormatConditions(1).StopIfTrue = False
    
' page 3
   
    With Range("x5:al13")
        .FormatConditions.Delete
    End With

    Call FormatRange("x5:ak13", "=(x76:ak83)=1")   ' check for orange
    Call FormatRangeRed("x5:ak13", "=(x76:ak83)=2")   ' check for red
    
    Range("t1").Select
    
    Select Case Range("t72").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
    Selection.FormatConditions(1).StopIfTrue = False
    
' page 4
    
    With Range("Am1:bc50")
        .FormatConditions.Delete
    End With
    
    Call FormatRange("ap6:bc50", "=(ap97:bc142)=1")   ' check for orange
    Call FormatRangeRed("ap6:bc50", "=(ap97:bc142)=2")   ' check for red
        
    Range("am1").Select
    
    Select Case Range("am143").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
' page 5
    
    With Range("bd1:bq51")
        .FormatConditions.Delete
    End With
    
    Call FormatRange("be5:bq51", "=(be160:bq206)=1")   ' check for orange
    Call FormatRangeRed("be5:bq51", "=(be160:bq206)=2")   ' check for red
        
    Range("bd1").Select
    
    Select Case Range("bd157").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
        
ErrhandlerCF:

    Select Case Err
    
        Case 13: MsgBox "One of the pages has an error when calculating the CF value..." & vbNewLine & "Please investigate."
        End
    End Select

End Sub
Sub CFBrackenfellFV()

' page 1

'    Application.ScreenUpdating = False
'    On Error GoTo ErrhandlerCF
'
'    Worksheets("bracken p1").Select
'
'    With Range("b2,r52:u64")
'        .FormatConditions.Delete
'    End With
'
'    Call FormatRange("r52:u52", "=(r69:u69)=1")   ' LT & MT rack colour
'    Call FormatRangeRed("r52:u52", "=(r69:u69)=2")   ' LT & MT rack colour
'    Call FormatRange("r55:u55", "=(r70:u70)=1")   'room colour check
'    Call FormatRangeRed("r55:u55", "=(r70:u70)=2")   ' room colour check
'    Call FormatRange("r58:u58", "=(r71:u71)=1")   ' case colour check
'    Call FormatRangeRed("r58:u58", "=(r71:u71)=2")   ' case colour check
'    Call FormatRange("r61", "=(r72)=1")   ' PRC colour check
'    Call FormatRangeRed("r61", "=(r72)=2")   ' PRC colour check
'    Call PageOneConditions

' page 2

    Worksheets("bracken db").Select
    
    With Range("e5:s9")
        .FormatConditions.Delete
    End With

    Call FormatRange("e5:s7", "=(e76:r79)=1")   ' check for orange
    Call FormatRangeRed("e5:s7", "=(e76:r79)=2")   ' check for red
    
    Range("a1").Select
    
    Select Case Range("a72").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
    Selection.FormatConditions(1).StopIfTrue = False
    
' page 3
   
    With Range("x5:al13")
        .FormatConditions.Delete
    End With

    Call FormatRange("x5:ak13", "=(x76:ak83)=1")   ' check for orange
    Call FormatRangeRed("x5:ak13", "=(x76:ak83)=2")   ' check for red
    
    Range("t1").Select
    
    Select Case Range("t72").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
    Selection.FormatConditions(1).StopIfTrue = False
    
' page 4
    
    With Range("Am1:bc50")
        .FormatConditions.Delete
    End With
    
    Call FormatRange("ap6:bc50", "=(ap97:bc142)=1")   ' check for orange
    Call FormatRangeRed("ap6:bc50", "=(ap97:bc142)=2")   ' check for red
        
    Range("am1").Select
    
    Select Case Range("am143").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
' page 5
    
    With Range("bd1:bq51")
        .FormatConditions.Delete
    End With
    
    Call FormatRange("be5:bq51", "=(be160:bq206)=1")   ' check for orange
    Call FormatRangeRed("be5:bq51", "=(be160:bq206)=2")   ' check for red
        
    Range("bd1").Select
    
    Select Case Range("bd157").Value
        Case 0
            With Selection.Interior
                .Color = 3407718
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 1
            With Selection.Interior
                .Color = 49407
            End With
            With Selection.Font
                 .Color = vbBlack
            End With
        Case 2
            With Selection.Interior
                .Color = 255
            End With
            With Selection.Font
                 .Color = vbWhite
            End With
    End Select
    
    
' page 6
'
'    With Range("br1:cj6")
'        .FormatConditions.Delete
'    End With
'
'    Call FormatRange("br5:cj6", "=(br72:cj73)=1")   ' check for orange
'    Call FormatRangeRed("br5:cj6", "=(br72:cj73)=2")   ' check for red
'
'    Range("br1").Select
'
'    Select Case Range("br70").Value
'        Case 0
'            With Selection.Interior
'                .Color = 3407718
'            End With
'        Case 1
'            With Selection.Interior
'                .Color = 49407
'            End With
'        Case 2
'            With Selection.Interior
'                .Color = 255
'            End With
'            With Selection.Font
'                 .Color = vbWhite
'            End With
'        Case 9
'            With Selection.Interior
'             .Pattern = xlSolid
'             .PatternColorIndex = xlAutomatic
'             .ThemeColor = xlThemeColorDark1
'             .TintAndShade = -0.249977111117893
'             .PatternTintAndShade = 0
'            End With
'            With Selection.Font
'                .Color = vbBlack
'            End With
'    End Select
    
ErrhandlerCF:

    Select Case Err
    
        Case 13: MsgBox "One of the pages has an error when calculating the CF value..." & vbNewLine & "Please investigate."
        End
    End Select

End Sub

