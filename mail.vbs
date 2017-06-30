Set args = Wscript.Arguments

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "scriptresult@kp.org"
objEmail.To = Wscript.Arguments.Item(0)

objEmail.Subject =  "Result of Input File " + Wscript.Arguments.Item(1)

objEmail.Textbody  =  "Click this link to download your file: " + Wscript.Arguments.Item(2) + vbCrLf + vbCrLf

objEmail.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailhub.kp.org" 
objEmail.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
