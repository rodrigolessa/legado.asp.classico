<!--#include file="funcoes/verSQL.asp" -->
<!--#include file="funcoes/quantFotos.asp" -->
<!--#include file="funcoes/exFotosGaleria.asp" -->
<!--#include file="funcoes/exVideos.asp" -->

<%
function criaDiretorio(cdCaminho)
	DIM fso, fCaminho
	SET fso			= CreateObject("Scripting.FileSystemObject")
	SET fCaminho	= fso.CreateFolder(cdCaminho)
	criaDiretorio	= fCaminho.Path
	SET fCaminho	= nothing
	SET fso			= nothing
end function


function existeDiretorio(edCaminho)
	DIM fso, f, fTemp
	SET fso			= CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(edCaminho) then
		fTemp = true
	else
		fTemp = false
	end if
	existeDiretorio	= fTemp
end function


Function apagarDiretorio(adCaminho)
	DIM fs
 	SET fs = CreateObject("Scripting.FileSystemObject")
	if fs.FolderExists(adCaminho) then
		fs.DeleteFolder adCaminho
	end if
	SET fs = Nothing
End Function


Function existArq(nomeArquivo)
	DIM fs, mainDir, nomeDir, sSaida
	SET fs = CreateObject("Scripting.FileSystemObject")
	sSaida = False
	if fs.FileExists(nomeArquivo) then
		sSaida = True
	end if
	existArq = sSaida
End Function


Function apagarArquivo(nomeArquivo)
	DIM fs
 	SET fs = CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(nomeArquivo) then
		fs.DeleteFile nomeArquivo
	end if
	SET fs = Nothing
End Function


Function extracNomeArquivo(sPath)
	Dim s 
	'pega a extenção do arquivo
	s = StrReverse(sPath)
	s = Mid(s, 1, Instr(s, "\") - 1)
	s = StrReverse(s)
	extracNomeArquivo = LCase(s)
End Function


Function getExt(nf)
	DIM s 
	'PEGA A EXTENÇÃO DO ARQUIVO
	s = StrReverse(nf)
	s = Mid(s, 1, inStr(s, ".") - 1)
	s = StrReverse(s)
	getExt = uCase(s)
End Function


Function retNomeModelo(rnModelo, rnMaxNome)

	DIM anModelos, temp, strNmModelo
	
	temp = ""
	
	'===== NOME DAS MODELOS PARA EXIBIR =====
	sSQL = "SELECT	modelo, nome " & _
			"FROM	XCA_modelos " & _ 
			"WHERE	modelo = " & rnModelo & " "
			
	anModelos	= getArray(sSQL)
	if isArray(anModelos) then
	
		strNmModelo = left(anModelos(1,0), rnMaxNome)
		if len(anModelos(1,0)) > rnMaxNome then
			while (right(strNmModelo, 1) <> " ") and (len(strNmModelo) > 0)
				strNmModelo = left(strNmModelo, len(strNmModelo) - 1)
			wend
			strNmModelo = strNmModelo & "..."
		end if
		
	end if
	
	retNomeModelo = temp

End Function


Function formatDate(data, dataSaida)
	Dim dia, mes, ano, hora, minuto, segundo
	Dim sSaida
	
	If data <> "" Then
		dia = Day(data)
		mes = Month(data)
		ano = Right(Year(data), 2)
		hora = Hour(data)
		minuto = Minute(data)
		segundo = Second(data)
	
		If Len(dia) = 1 Then dia = "0" & dia 
		If Len(mes) = 1 Then mes = "0" & mes 
		If Len(hora) = 1 Then hora = "0" & hora
		If Len(minuto) = 1 Then minuto = "0" & minuto
		If Len(segundo) = 1 Then segundo = "0" & segundo 
	
		dataSaida = Replace(dataSaida, "DD", dia)
		dataSaida = Replace(dataSaida, "MM", mes)
		dataSaida = Replace(dataSaida, "AA", ano)
		dataSaida = Replace(dataSaida, "hh", hora)
		dataSaida = Replace(dataSaida, "mm", minuto)
		dataSaida = Replace(dataSaida, "ss", segundo)
	Else
		dataSaida = ""
	End If

	formatDate = dataSaida	
End Function


Function mountLead(lead)
	Dim s, pos
	s = lead
	pos = InstrRev(s, " ")
	If pos > 0 Then
		s = Left(s, pos - 1)
	End If
	s = s & "..."
	mountLead = s
End Function


Sub criaComboNumerico(nome, sel, ini, fim, passo)
	Dim s, i
	s = "<Select name='" & nome & "'>" & vbCrLf
	s = s & "<option value=''>"
	For i = ini To fim Step passo
		s = s & "<option value='" & zeros(i, 2) & "'"
		If Cstr(sel) = Cstr(i) Then s = s & " SELECTED "
		s = s & ">" & zeros(i, 2) & vbCrLf
	Next
	Response.Write s & "</select>" & vbCrLf
End Sub


Function zeros(num, t)
'funcao que completa com zeros a esquerda o numero fornecido
'num -> numero original
't -> tamanho final do num
	Dim i, numSaida, numZeros
	numZeros = t - Len(num)
	If numZeros > 0 Then
		zeros = String(numZeros, "0") & num
	Else
		zeros = num
	End If
End Function


Sub criaComboStatus(status)
		s = "<select name='status' class='cx'>" & vbCrLf
		s = s & "<option value='A' " 
		If status = "A" Or  status = "a" Then s = s & " SELECTED "
		s = s & ">Ativo" & vbCrLf
		s = s & "<option value='I' " 
		If status = "I" Or  status = "i" Then s = s & " SELECTED "
		s = s & ">Inativo" & vbCrLf
		s = s & "</select>" & vbCrLf
		Response.Write s
End Sub


Sub wLog(str, fp)
  Const ForAppending = 8, TristateFalse = 0
  Dim fs, f
  Dim usuario, pathFile
	If fp = "" Then
		fp = Application("XCA_APP_PATH_LOG")
	End If
	
	pathFile = fp
  If Session("XCA_login") <> "" Then
		str = Session("XCA_login") & " - " & str
	End If

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile(pathFile, ForAppending, True, TristateFalse)
  f.Write Now & " - " & str & Chr(13) & Chr(10) 
  f.Close
	Set f = Nothing
	Set fs = Nothing
End Sub


Function getTipoUpload()
	If Request.ServerVariables("LOCAL_ADDR") = "192.168.0.13" Then
		getTipoUpload = 1
	Else
		getTipoUpload = 2
	End If
End Function


'Rotina para pesquisa de palavras acentuadas
'Por Rubens Farias (rubensf@bigfoot.com), Fev/2001
'[Preserve a referência ao autor!]
Function FormatLikeSearch( Texto )
 Dim n, NovoTexto, valorASC
 NovoTexto = ""
 For n = 1 To Len( Texto )
     valorASC = asc( mid( Texto, n, 1 ) )
     Select Case valorASC
        case  39: NovoTexto = NovoTexto & "''"
        case  65: NovoTexto = NovoTexto & "[ÁÀÂÄÃA]"
        case  67: NovoTexto = NovoTexto & "[ÇC]"
        case  69: NovoTexto = NovoTexto & "[ÉÈÊËE]"
        case  73: NovoTexto = NovoTexto & "[ÍÌÎÏI]"
        case  79: NovoTexto = NovoTexto & "[ÓÒÔÖÕO]"
        case  85: NovoTexto = NovoTexto & "[ÚÙÛÜU]"
        case  97: NovoTexto = NovoTexto & "[áàâäãa]"
        case  99: NovoTexto = NovoTexto & "[çc]"
        case 101: NovoTexto = NovoTexto & "[éèêëe]"
        case 105: NovoTexto = NovoTexto & "[íìîïi]"
        case 111: NovoTexto = NovoTexto & "[óòôöõo]"
        case 117: NovoTexto = NovoTexto & "[úùûüu]"
        case else
           if valorASC > 31 and valorASC < 127 then
              NovoTexto = NovoTexto & chr( valorASC )
           else
              NovoTexto = NovoTexto & "_"
           end if
     end select
 next
 FormatLikeSearch = "'%" & NovoTexto & "%'"
End Function


Function GeraCaracter(Tipo)
	Dim numOk, tmpNum
	
	numOk = False
	While Not numOk
		Randomize
		tmpNum = Int((122 - 48 + 1) * Rnd + 48)
			
		Select Case Tipo
			Case 0 'gera senha com numeros e letras
				If (tmpNum > 47) AND (tmpNum < 58) Then
					numOk = true
				Else
					If (tmpNum > 64) AND (tmpNum < 91) Then
						numOk = true
					Else
						If (tmpNum > 96) AND (tmpNum < 123) Then
							numOk = true
						Else
							numOk = false
						End If
					End If
				End If
			Case 1 'gera senha apenas com numeros
				If (tmpNum > 47) AND (tmpNum < 58) Then
					numOk = true
				End If			
			Case 2 'gera senha apenas com letras
				If (tmpNum > 64) AND (tmpNum < 91) Then
					numOk = true
				Else
					If (tmpNum > 96) AND (tmpNum < 123) Then
						numOk = true
					Else
						numOk = false
					End If
				End If
		End Select
	Wend
	GeraCaracter = chr(tmpNum)
End Function


Function GeraSenha(intNum, intTipo)
	Dim strSenha, i
	strSenha = ""
	For i=1 To intNum
		strSenha = strSenha & GeraCaracter(intTipo)
	Next
	GeraSenha = strSenha
End Function


Function impImgI(secao, nuImgI)
	Dim temp, dsSecao, i, dirImg, imgOK, ckRadio
	i = 1
	erro = ""
	imgOK = true
	dsSecao = "fig_imgi"
'	Select Case secao
'		Case "1"
'			dsSecao = "fig_evento"
'		Case "2"
'			dsSecao = "fig_novidade"
'		Case "3"
'			dsSecao = "fig_movimento"
'		Case "4"
'			dsSecao = "fig_faq"
'		Case "5"
'			dsSecao = "fig_refletir"
'		Case Else
'			erro = "seção não encontrada!"
'	End Select
	
	If Len(erro) < 1 Then
		temp = temp & "<div id=divBg style='position:relative; z-index:10; width:600px; left:0px; height:40px; clip:rect(0px 10px 10px 0px); visibility:hidden;'>"
		temp = temp & "<div id=divMenu style='position:absolute; z-index:11; left:11px; top:1px; color:#333333; font-size:13px; font-family:verdana,arial,helvetica,sans-serif; visibility:inherit;'>"
		temp = temp & "<table width=600 height=40 cellpadding=0 cellspacing=5 border=0>"
		temp = temp & "<tr>"
		While imgOK
			If i = nuImgI Then ckRadio = " CHECKED " Else ckRadio = ""
			If existArq(Application("XCA_APP_SERVERDIR_IMGS")&dsSecao&i&".jpg") Then
				temp = temp & "<td align=right><input type=radio name='imgI' value="&i&""&ckRadio&"></td><td align=left><img src='"&Application("XCA_APP_HOME")&Application("XCA_APP_DIR_IMGS")&dsSecao&i&".jpg' width=25 height=25 border=0></td>"
			Else
				If existArq(Application("XCA_APP_SERVERDIR_IMGS")&dsSecao&i&".gif") Then
					temp = temp & "<td align=right><input type=radio name='imgI' value="&i&""&ckRadio&"></td><td align=left><img src='"&Application("XCA_APP_HOME")&Application("XCA_APP_DIR_IMGS")&dsSecao&i&".gif' width=25 height=25 border=0></td>"
				Else
					imgOK = false
				End If
			End If
			i = i + 1
		Wend
		temp = temp & "</tr>"
		temp = temp & "</table>"
		temp = temp & "</div>"
		temp = temp & "<div id='divArrowLeft' style='position:absolute; z-index:12; width:11px; height:45px; left:0px; top:0px; visibility:inherit;'><a href=# onmouseover='noScroll=false; mLeft()' onmouseout='noMove()' onclick='sScrollPx-=sScrollExtra; return false' onfocus='if(this.blur)this.blur()' onmousedown='sScrollPx+=sScrollExtra'><img src='"&Application("XCA_APP_HOME")&Application("XCA_APP_DIR_IMGS")&"seta_esq.gif' width=11 height=40 border=0></a></div>"
		temp = temp & "<div id='divArrowRight' style='position:absolute; z-index:12; width:11px; height:45px; top:0px; visibility:inherit;'><a href=# onmouseover='noScroll=false; mRight()' onmouseout='noMove()' onclick='sScrollPx-=sScrollExtra; return false' onfocus='if(this.blur)this.blur()' onmousedown='sScrollPx+=sScrollExtra'><img src='"&Application("XCA_APP_HOME")&Application("XCA_APP_DIR_IMGS")&"seta_dir.gif' width=11 height=40 border=0></a></div>"
		temp = temp & "</div>"
	Else
		temp = erro&"<br>"
	End If
	impImgI = temp
End Function

'-------------------------------------------------------------------------
'PREENCHE VARIAVEL COM 2 DIGITOS COM ZERO A ESQUERDA
function  d2(sValor)
  if (len(sValor) = 1) and (isNumeric(sValor)) then d2 = "0" & sValor else d2 = sValor
end function

'-------------------------------------------------------------------------
'PREENCHE VARIAVEL COM DIGITOS A ESQUERDA
function digEsq(deDig, deTam, deValor)

	DIM retDigEsq
	'VERIFICA O VALOR DO DIGITO A SER COLOCADO A ESQUERDA
	if len(trim(deDig)) > 0 then
		deDig	= cStr(trim(deDig))
	else
		deDig	= "0"
	end if
	'VERIFICA O TAMANHO DETERMINADO PARA A STRING
	if len(trim(deTam)) > 0 then
		if isNumeric(trim(deTam)) then
			deTam	= cInt(deTam)
		else
			deTam	= 2
		end if
	else
		deTam	= 2
	end if
	'VERIFICA O VALOR DA DIREITA PASSADO
	if len(trim(deValor)) > 0 then
		deValor	= cStr(trim(deValor))
	else
		deValor	= "0"
	end if
	'ADICIONA OS CARACTERES A ESQUERDA
	retDigEsq = string(deTam - len(deValor), deDig) & deValor
	'RETORNA O VALOR PARA A VARIAVEL
	digEsq	= retDigEsq
	
end function
'-------------------------------------------------------------------------


'----------------------------------------------------------------------------
'CORTA O TEXTO E EXIBE SÓ UMA PARTE COM '...' PARA CABER NO ESPAÇO DESTINADO
function cortaTxt(ctTexto, ctTam)

	DIM	temp
	
	temp = left(ctTexto, ctTam)
	
	if len(ctTexto) > ctTam then
		while (right(temp, 1) <> " ") and (len(temp) > 0)
			temp = left(temp, len(temp) - 1)
		wend
		temp = temp & "..."
	end if
	
	cortaTxt = temp
	
end function
'-------------------------------------------------------------------------
%>