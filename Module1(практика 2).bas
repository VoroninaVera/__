Attribute VB_Name = "Module1"
Sub абоба()
Attribute абоба.VB_ProcData.VB_Invoke_Func = "й\n14"
'
' абоба Макрос
'
' Сочетание клавиш: Ctrl+й
'
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Повторение мать ученье"
    Range("C4").Select
End Sub


Sub абоба2()

    Range("C3").Value = "Повторение мать ученье"
    
End Sub

