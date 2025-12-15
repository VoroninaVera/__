Attribute VB_Name = "Module1"
Option Explicit

Sub test()

    Dim d As Date
'    d = #mm/dd/yyyy hh:nn:ss#
    d = #10/22/2023 10:31:15 PM#
    MsgBox Format(d, "dd/mm/yy hh:nn")

End Sub
