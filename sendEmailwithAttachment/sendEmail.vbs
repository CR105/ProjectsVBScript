Option Explicit
Dim outobj, mailobj

Set outobj = CreateObject("Outlook.Application")
Set mailobj = outobj.CreateItem(0)

With mailobj
	.To = "mail@mail.com"
	.CC = "anothermai.@mail.com"
	'.BCC = "bccmail@mail.com"
	.Subject = "Send email"
	.HTMLBody = "<html><body><p style='font-family:Calibri (Cuerpo);font-size:11pt'>" &_
				"Hi Team," &_
				"<br><br>I send this email for you." &_
				"<br><br>Regards.</p>" &_
				"<p style='font-family:Calibri (Cuerpo);color: rgb(0, 32, 96); font-size: 10pt;'>" &_
				"-----------------------------------------------------------------" &_
				"<br><b>Name and Last  name</b> " &_
				"<br>Your title" &_
				"<br>personal o professional email</p></body></html>"
	.Attachments.Add("\somefileattachment.file")
	.Display
	.Send
End With

Set outobj = Nothing
Set mailobj = Nothing
