Attribute VB_Name = "Module2"
Sub МояГеройскаяАкадемия()
Attribute МояГеройскаяАкадемия.VB_ProcData.VB_Invoke_Func = "ц\n14"
'
' МояГеройскаяАкадемия Макрос
'
' Сочетание клавиш: Ctrl+ц
'
    Range("A1:F22").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("A1:F22").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("A1:F22").Select
    ActiveWindow.SmallScroll Down:=-15
End Sub
Sub СписокСтудентов()
Attribute СписокСтудентов.VB_ProcData.VB_Invoke_Func = "ф\n14"
'
' СписокСтудентов Макрос
'
' Сочетание клавиш: Ctrl+ф
'
    ActiveCell.Offset(17, 2).Range("A1").Select
    ActiveWindow.SmallScroll Down:=-36
    ActiveCell.Offset(-17, -2).Range("A1:F22").Select
End Sub
