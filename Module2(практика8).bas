Attribute VB_Name = "Module2"
Option Explicit
Option Compare Text

Sub ¬рем€√ода()

    Dim varMonth As Variant
    Dim strResult As String
    
    varMonth = InputBox("¬ведите номер или название мес€ца")
    
    Select Case varMonth
        Case 12, 1, 2, "декабрь", "€нварь", "февраль": strResult = "«има"
        Case 3 To 5, "март", "апрель", "май": strResult = "¬есна"
        Case 6 To 8, "июнь", "июль", "август": strResult = "Ћето"
        Case 9 To 11, "сент€брь", "окт€брь", "но€брь": strResult = "ќсень"
        Case Else: strResult = "Ќекорректно введЄн мес€ц"
    End Select
    
    MsgBox strResult

End Sub
