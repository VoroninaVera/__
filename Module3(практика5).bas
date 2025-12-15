Attribute VB_Name = "Module3"
' пример скорости работы макроса с необъявленными переменными

Sub Text()

    t = Timer
    
    For i = 1 To 200000000
        m = m + i
    Next i
    
    t = Timer - t
    
    MsgBox "Процедура отработала за " & t & " сек "
    
End Sub


'пример скорости работы макроса с объявленными переменными

Sub text2()

    Dim t As Double, i As Long, m As LongLong
    
    t = Timer
    
    For i = 1 To 200000000
        m = m + i
    Next i
    
    t = Timer - t
    
    MsgBox "Процедура отработала за " & t & " сек "
    
End Sub
