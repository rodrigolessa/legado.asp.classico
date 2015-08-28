<%
'Enviar e-mails
Dim emailSender, servidorSMTP
Dim nomeFrom, emailFrom, nomeTo, emailTo
Dim strSubject, strBody

Function enviaEmail()
	Dim rs, query
	Set rs = Server.CreateObject("ADODB.RecordSet")
	query = "Select IsNull(emailSender, 'CDO') As emailSender From sed_config"
	rs.open query, conn
	If Not rs.EOF Then emailSender = Trim(rs("emailSender")) Else emailSender = "CDO"
	rs.close
	Set rs = Nothing
	servidorSMTP = Request.ServerVariables("LOCAL_ADDR")
	If emailFrom = "" Then
		nomeFrom = Application("SED_APP_TITLE")
		emailFrom = Application("SED_APP_EMAIL")
	End If
	Select Case emailSender
		Case "CDO"
			'envio de email via aspEmail
			Call enviaEmailWin2k()
		Case "CDONTS"
			'Envia E-mail via CDONTS
			Call enviaEmailCDONTS()
		Case "ASPEMAIL"
			'envio de email via aspEmail
			Call enviaEmailAspEMail()
		Case "ASPEMAILSVG"
			'envio de email via aspMailSVG
			Call enviaEmailAspMailSVG()
		Case Else
			'Envia E-mail via CDONTS
			Call enviaEmailCDONTS()
		End Select
End Function

'Funções para envio de e-mails
	Function enviaEmailCDONTS()
		Dim objCDO
		Set objCDO = Server.CreateObject("CDONTS.NewMail")
		
		objCDO.To = emailTo
		objCDO.From = emailFrom
		objCDO.Subject = strSubject
		objCDO.BodyFormat = 0 '0 = html, 1 = Texto (default)
		objCDO.MailFormat = 0 
		objCDO.Body = strBody
		objCDO.Send 'Send off the email!
		
		'Cleanup
		Set objCDO = Nothing
	End Function

	Function enviaEmailWin2k()
		Dim iMsg
		Set iMsg = CreateObject("CDO.Message")
		
		Dim iConf
		Set iConf = CreateObject("CDO.Configuration")
		
		Dim Flds
		Set Flds = iConf.Fields
		
		' Set the configuration
		With Flds
		'  ' assume constants are defined within script file
		'  .Item("cdoSendUsingMethod")       = 2 ' cdoSendUsingPort
		'  .Item("cdoSMTPServerName")        = "server@scriptorio.com.br"
			.Item("cdoSMTPServer") = servidorSMTP
		'  .Item("cdoSMTPConnectionTimeout") = 10 ' quick timeout
		'  .Item("cdoSMTPAuthenticate")      = cdoBasic
		'  .Item("cdoSendUserName")          = "username"
		'  .Item("cdoSendPassword")          = "password"
		'  .Item("cdoURLProxyServer")        = "server:80"
		'  .Item("cdoURLProxyBypass")        = "<local>"
		'  .Item("cdoURLGetLatestVersion")   = True
		  .Update
		End With
		
		With iMsg
		  Set .Configuration = iConf
		      .To       = """" & nomeTo & """ <" & emailTo & ">"
		      .From     = """" & nomeFrom & """ <" & emailFrom & ">"
		      .Subject  = strSubject
			  .HTMLBody = strBody
	' 		  .TextBody = strBody
	'		  .CreateMHTMLBody (epag)
	'	      .AddAttachment "C:\files\mybook.doc"
		      .Send
		End With
	End Function

	Function enviaEmailAspEMail()
		Dim Mail
		Set Mail = Server.CreateObject("Persits.MailSender")
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
		If Err <> 0 Then
		   Response.Write "Erros encontrados: " & Err.Description
		End If
		On Error Goto 0
		Set Mail = Nothing
	End Function

	Function enviaEmailAspMailSVG()
		Dim Mail
		Set Mail = Server.CreateObject("SMTPsvg.Mailer")
		Mail.RemoteHost = servidorSMTP ' Specify a valid SMTP server
		Mail.FromName = nomeFrom ' Specify sender's name
		Mail.FromAddress = emailFrom ' Specify sender's address
		
		Mail.AddRecipient nomeTo, emailTo
		Mail.Subject = strSubject
		Mail.ContentType = "text/html"
		Mail.BodyText = strBody
		
		On Error Resume Next
		Mail.SendMail
		If Err <> 0 Then
		   Response.Write "Erros encontrados: " & Err.Description
		End If
		On Error Goto 0
		Set Mail = Nothing
	End Function
	
%>