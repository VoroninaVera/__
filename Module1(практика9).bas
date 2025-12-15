Attribute VB_Name = "Module1"
Option Explicit


Sub test()

'    MsgBox Workbooks("Практика9.xlsm").Worksheets("Лист1").Name
'    MsgBox Workbooks("Практика9.xlsm").Worksheets(1).Range("C3:G10").Count
    
    Dim X As Integer
    X = 10
    
    Dim shtX As Worksheet
    Set shtX = Workbooks("Практика9.xlsm").Worksheets("Лист1")
    
    MsgBox shtX.Name
    
    Dim rngX As Range
    Set rngX = Selection
    
    MsgBox rngX.Address

End Sub
