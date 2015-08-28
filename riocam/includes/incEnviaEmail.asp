<%

Function enviaEmail(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, emailSender)
' Variaveis para enviar e-mails
DIM		servidorSMTP
CONST	maxStrSubject = 40


	'Pre definidos
	nomeFrom		= "JOMP"
	emailFrom		= "rodrigolsr@gmail.com"
'	nomeTo			= "Rodrigo Lessa"
'	emailTo			= "rodrigolsr@gmail.com"
'	strSubject		= "E-mail automático da JOMP!"
	servidorSMTP	= "smtp.ism.com.br" 'request.serverVariables("LOCAL_ADDR") '"smtp.rodrigolessa.com"
	emailSender		= "CDO"

	Select Case emailSender
		Case "CDO"
			'envio de email via ASPEmail
			Call enviaEmailWin2k(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		Case "CDONTS"
			'envia e-mail via CDONTS
			Call enviaEmailCDONTS(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		Case "ASPEMAIL"
			'envio de email via ASPEmail
			Call enviaEmailAspEMail(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		Case "ASPEMAILSVG"
			'envio de email via aspMailSVG
			Call enviaEmailAspMailSVG(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		Case Else
			'envia e-mail via CDONTS
			Call enviaEmailCDONTS(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
	End Select
		
End Function



'********************************************************************************************************************
'FUNÇÕES PARA ENVIO DE E-MAILS
'********************************************************************************************************************

	'ENVIA E-MAIL VIA CDONTS
	Function enviaEmailCDONTS(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		On Error Resume Next
		DIM objCDO
		SET objCDO = Server.CreateObject("CDONTS.NewMail")
		
		objCDO.To = emailTo
		objCDO.From = emailFrom
		objCDO.Subject = strSubject
		objCDO.BodyFormat = 0 '0 = html, 1 = Texto (default)
		objCDO.MailFormat = 0 
		objCDO.Body = strBody
		objCDO.Send 'Send off the email!
		
		'Cleanup
		SET objCDO = Nothing
		
		If ERR.Number <> 0 then
		   response.write "Erros encontrados: " & Err.Description
		end if
		On Error Goto 0
	End Function



	'ENVIO DE EMAIL VIA CDO
	Function enviaEmailWin2k(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
'		On Error Resume Next
		DIM iMsg
		SET iMsg = CreateObject("CDO.Message")
		
		DIM iConf
		SET iConf = CreateObject("CDO.Configuration")
		
		DIM Flds
		SET Flds = iConf.Fields
		
		'SET the configuration
		With Flds
			'assume constants are defined within script file
		'	.Item("cdoSendUsingMethod")       = 2 ' cdoSendUsingPort
		'	.Item("cdoSMTPServerName")        = "server@scriptorio.com.br"
			.Item("cdoSMTPServer") = servidorSMTP
		'	.Item("cdoSMTPConnectionTimeout") = 10 ' quick timeout
		'	.Item("cdoSMTPAuthenticate")      = cdoBasic
		'	.Item("cdoSendUserName")          = "username"
		'	.Item("cdoSendPassword")          = "password"
		'	.Item("cdoURLProxyServer")        = "server:80"
		'	.Item("cdoURLProxyBypass")        = "<local>"
		'	.Item("cdoURLGetLatestVersion")   = True
			.Update
		End With
		
'		response.write "nomeTo: "&nomeTo & "<br>"
'		response.write "emailTo: "&emailTo & "<br>"
'		response.write "nomeFrom: "&nomeFrom & "<br>"
'		response.write "emailFrom: "&emailFrom & "<br>"
'		response.write "strSubject: "&strSubject & "<br>"
'		response.write "strBody: "&strBody & "<br>"

		With iMsg
		Set	.Configuration = iConf
			.To       = """" & nomeTo & """ <" & emailTo & ">"
			.From     = """" & nomeFrom & """ <" & emailFrom & ">"
			.Subject  = strSubject
			.HTMLBody = strBody
		'	.TextBody = strBody
		'	.CreateMHTMLBody (epag)
		'	.AddAttachment "C:\files\mybook.doc"
			.Send
		End With
		
'		if ERR <> 0 then
'		   response.write "Erros encontrados: " & Err.Description
'		end if
'		On Error Goto 0
	End Function



	'ENVIO DE EMAIL VIA ASPEMAIL
	Function enviaEmailAspEMail(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		DIM Mail
		SET Mail = Server.CreateObject("Persits.MailSender")
		Mail.Host = servidorSMTP ' Specify a valid SMTP server
		Mail.FromName = nomeFrom ' Specify sender's name
		Mail.From = emailFrom ' Specify sender's address
		
		Mail.AddAddress emailTo, nomeTo
	'	Mail.AddAddress "paul@paulscompany.com" ' Name is optional
	'	Mail.AddReplyTo "info@veryhotcakes.com"
	'	Mail.AddAttachment "c:\images\cakes.gif"
		
		Mail.Subject = strSubject
		Mail.Body = strBody
		
		On Error Resume Next
		Mail.Send
		if ERR <> 0 then
		   response.write "Erros encontrados: " & Err.Description
		end if
		On Error Goto 0
		SET Mail = Nothing
	End Function



	'ENVIO DE EMAIL VIA ASPMAILSVG
	Function enviaEmailAspMailSVG(nomeFrom, emailFrom, nomeTo, emailTo, strSubject, strBody, servidorSMTP)
		DIM Mail
		SET Mail = Server.CreateObject("SMTPsvg.Mailer")
		Mail.RemoteHost = servidorSMTP ' Specify a valid SMTP server
		Mail.FromName = nomeFrom ' Specify sender's name
		Mail.FromAddress = emailFrom ' Specify sender's address
		
		Mail.AddRecipient nomeTo, emailTo
		Mail.Subject = strSubject
		Mail.ContentType = "text/html"
		Mail.BodyText = strBody
		
		On Error Resume Next
		Mail.SendMail
		if ERR <> 0 then
		   response.write "Erros encontrados: " & Err.Description
		end if
		On Error Goto 0
		Set Mail = Nothing
	End Function
%>