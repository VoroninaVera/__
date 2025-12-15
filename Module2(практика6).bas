Attribute VB_Name = "Module2"
Option Explicit


Sub ВремениОсталось()

   Dim dtTime As Date
   Dim dtNow As Date
   Dim intMin As Integer
   Dim strText As String
   Dim dtInterval As Date
   
   
   dtTime = #6:00:00 PM#
   dtNow = Time
   intMin = (dtTime - dtNow) * 24 * 60
   dtInterval = dtTime - dtNow
   strText = "Сейчас: " & dtNow & Chr(13) & "До конца рабочего дня осталось: " & Format(dtInterval, "h часов n минут")
   MsgBox strText

End Sub
