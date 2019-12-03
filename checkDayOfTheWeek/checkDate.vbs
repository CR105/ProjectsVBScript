Option explicit
Dim objDic
Dim currentDate
Dim objShell

Set objDic = CreateObject("Scripting.Dictionary")
Set objShell = Wscript.CreateObject("WScript.Shell")

currentDate = Date

objDic.Add "1", "Sunday"
objDic.Add "2", "Monday"
objDic.Add "3", "Tuesday"
objDic.Add "4", "Wednesday"
objDic.Add "5", "Thursday"
objDic.Add "6", "Friday"
objDic.Add "7", "Saturday"

' Wscript.Echo("Today is: " & currentDate)
' Wscript.Echo("Day of week is: " & objDic.Item(Cstr(Weekday(currentDate, 0))))
If Weekday(currentDate, 6) = 1 then
    objShell.Run "script.vbs"
End If

Set objShell = Nothing
Set objDic = Nothing