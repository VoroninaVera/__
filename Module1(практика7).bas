Attribute VB_Name = "Module1"
Option Explicit


Sub РазмерСкидки()

    Dim varCost As Variant 'стоимость покупки
    Dim dbSale As Double 'размер скидки
    Dim strText As String 'сообщение
    
    varCost = InputBox("Введите стоимость покупки")
    
    If Not IsNumeric(varCost) Then
        MsgBox "Данные введены не корректно", vbCritical
        Exit Sub
    End If
    
    Select Case varCost
        Case Is >= 15000: dbSale = 0.25
        Case Is >= 10000: dbSale = 0.2
        Case Is >= 5000: dbSale = 0.15
        Case Else: dbSale = 0
    End Select
'        If varCost >= 15000 Then
'            dbSale = 0.25
'        ElseIf varCost >= 10000 Then
'            dbSale = 0.2
'        ElseIf varCost >= 5000 Then
'            dbSale = 0.15
'        Else
'            dbSale = 0
'        End If
    
    varCost = varCost * (1 - dbSale)
    
    strText = "Размер скидки: " & Format(dbSale, "0%") & Chr(13) & "Стоимость покупки: " & varCost
    
    MsgBox strText

End Sub
