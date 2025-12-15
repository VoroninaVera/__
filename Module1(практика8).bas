Attribute VB_Name = "Module1"
Option Explicit


Sub ДеньНедели()

    Dim varDate As Variant
    Dim strText As String
    Dim intResult As Integer
    
    varDate = InputBox("Введите дату")
    
    If Not IsDate(varDate) Then GoTo 2
    
    intResult = Weekday(varDate, vbMonday)
    
    Select Case intResult
        Case 1: strText = "Понедельник"
        Case 2: strText = "Вторник"
        Case 3: strText = "Среда"
        Case 4: strText = "Четверг"
        Case 5: strText = "Пятница"
        Case 6: strText = "Суббота"
        Case 7: strText = "Воскресенье"
    End Select
    
    strText = varDate & " - это " & strText
    
    MsgBox strText
    
    Exit Sub
    
2:
    MsgBox "Данные введены не корректно"

End Sub
