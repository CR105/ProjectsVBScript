Option explicit

Dim objDic
Dim currentDate
Dim objShell
Dim objPopup
Dim mess

Set objDic = CreateObject("Scripting.Dictionary")
Set objPopup = Wscript.CreateObject("WScript.Shell")
Set objShell = Wscript.CreateObject("WScript.Shell")

currentDate = Date

objDic.Add "1", "Sunday"
objDic.Add "2", "Monday"
objDic.Add "3", "Tuesday"
objDic.Add "4", "Wednesday"
objDic.Add "5", "Thursday"
objDic.Add "6", "Friday"
objDic.Add "7", "Saturday"

Wscript.Echo("Today is: " & currentDate & ", Day of week: " & Weekday(currentDate, 0))

If Weekday(currentDate, 0) = 2 then
               ' Wscript.Echo("Today is : " & objDic.Item(Cstr(Weekday(currentDate, 0))) & ". Don't forget send email")
               WScript.Sleep setTime("M", 1)
               ' mess = objPopup.Popup("Today is : " & objDic.Item(Cstr(Weekday(currentDate, 0))) & ". Are you sure send email?", 7, "Mensaje importante", 4 + 32)
               ' If mess = 6 Then
               objShell.Run "path/file.vbs"
               ' End If
End If

Private Function setTime(strTime, numTime)
               dim timeMin
               If UCase(strTime) = "H" Then
                              timeMin = 60* 60 * 1000 * numTime
               ElseIf UCase(strTime) = "M" Then
                              timeMin = 60 * 1000 * numTime
               ElseIf UCase(strTime) = "S" Then
                              timeMin = 1000 * numTime
               End If 
               setTime = timeMin
End Function

Set objShell = Nothing
Set objPopup = Nothing
Set objDic = Nothing
