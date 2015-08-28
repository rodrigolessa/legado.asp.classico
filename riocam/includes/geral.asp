<!--#include file="funcoes/impGaleria.asp" -->
<!--#include file="funcoes/impNoticias.asp" -->
<!--#include file="funcoes/impModelos.asp" -->
<!--#include file="funcoes/impUltModelos.asp" -->
<!--#include file="funcoes/impUltNoticias.asp" -->
<!--#include file="funcoes/impModHome.asp" -->
<!--#include file="funcoes/impVideos.asp" -->
<!--#include file="funcoes/exibeBanner.asp" -->
<!--#include file="funcoes/addAcessoModelo.asp" -->
<!--#include file="funcoes/addAcessoUsuario.asp" -->
<!--#include file="funcoes/limpaLogin.asp" -->
<!--#include file="funcoes/maiuscula.asp" -->
<!--#include file="funcoes/validarMail.asp" -->

<!--#include file="multilng/funcoes.asp" -->

<!--#include file="funcoes/verSQL.asp" -->
<!--#include file="funcoes/valDatHor.asp" -->
<!--#include file="funcoes/compDataHora.asp" -->

<%

Function criaArrayBD(consulta)
	criaArrayBD = conn.Execute(consulta).GetRows()
End Function


Function formatDate(data, dataSaida)
	Dim dia, mes, ano, hora, minuto, segundo
	Dim sSaida
	dia = Day(data)
	mes = Month(data)
	ano = Year(data)
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
	dataSaida = Replace(dataSaida, "AAAA", ano)
	dataSaida = Replace(dataSaida, "AA", Right(ano, 2))
	dataSaida = Replace(dataSaida, "hh", hora)
	dataSaida = Replace(dataSaida, "mm", minuto)
	dataSaida = Replace(dataSaida, "ss", segundo)

	formatDate = dataSaida
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


'ROTINA PARA PESQUISA DE PALAVRAS ACENTUADAS
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


Sub addCountUsuario()
	Dim rs2
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	query = "Update XCA_config Set contador = contador + 1"
	conn.Execute query
	
	query = "SELECT contador FROM XCA_config"
	rs2.Open query, conn
	If Not rs2.EOF Then Session("contUsuario") = rs2("contador")
	rs2.Close
	Set rs2 = Nothing
End Sub

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

'-------------------------------------------------------------------------
'PREENCHE VARIAVEL COM DIGITOS A ESQUERDA
function converteBytes(numBytes)
'Byte		- 8 bits (1 byte)
'Kilobyte	- 1.024 bytes (1 KB)
'Megabyte	- 1.024 * 1 KB [1.048.576 bytes] (1 MB)
'Gigabyte	- 1.024 * 1 MB [1.073.741.824 bytes] (1 GB)
'Terabyte	- 1.024 * 1 GB [1.099.511.627.776 bytes] (1 TB)
'Petabyte	- 1.024 * 1 TB [1.125.899.906.842.624 bytes] (1 PB)
'Exabyte	- 1.024 * 1 PB [1.152.921.504.606.846.976 bytes] (1 EB)
'Zettabyte	- 1.024 * 1 EB [1.180.591.620.717.411.303.424 bytes] (1 ZB)
'Yottabyte	- 1.024 * 1 ZB [1.208.925.819.614.629.174.706.176 bytes] (1 YB)

DIM cbTipo, cbTam

	if len(numBytes) > 0 then

		cbTam		= cDbl(numBytes)

		if	cbTam >= 1024 then
			cbTam	= Round(cbTam / 1024)
			cbTipo	= "KB"
		end if

		if	cbTam >= 1024 then
			cbTam	= FormatNumber(cbTam / 1024, 2)
			cbTipo	= "MB"
		end if

		if	cbTam >= 1024 then
			cbTam	= FormatNumber(cbTam / 1024, 2)
			cbTipo	= "GB"
		end if
	
	end if
	
	converteBytes	= cbTam & " " & cbTipo

end function
'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'RETORNA SE O USUÁRIO ESTA LOGADO
Function verLogado()
	if session("XCA_LOGADO_SITE") <> "TRUE" then
		verLogado = false
	else
		verLogado = true
	end if
End Function
'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'RETORNA SE A MODELO ESTA ON-LINE, LOGADA NO SISTEMA
Function verModeloOn(vmModelo)
DIM strVMM, iV, vmTemp
	strVMM	= trim(cStr(vmModelo))
	vmTemp	= false

	If IsArray(Session.Contents("XCA_USUARIO_MODELO")) Then
		For iV = 0 to uBound(Session.Contents("XCA_USUARIO_MODELO"))
			if strVMM = cStr(Session.Contents("XCA_USUARIO_MODELO")(iV)) then
				vmTemp = true
			end if
		Next
	Else
		If IsObject(Session.Contents("XCA_USUARIO_MODELO")) Then
			vmTemp = false
		Else
			if strVMM = cStr(Session.Contents("XCA_USUARIO_MODELO")) then
				vmTemp = true
			end if
		End If
	End If

	verModeloOn = vmTemp
End Function
'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'RETORNA O LINK DA PÁGINA DA VIDEO CONFERENCIA SE A MODELO ESTA ON-LINE
Function retLinkConfer(rlcModelo)
DIM rlCaModelo
	if verModeloOn(rlcModelo) then
		'***** RECUPERA O NOME DA SALA DA MODELO SELECIONADA *****
		sSQL =	("SELECT	m.modelo, m.nome, m.sala " & _
				"FROM		XCA_modelos m " & _
				"WHERE	m.modelo = " & rlcModelo & " ")
		rlCaModelo	= getArray(sSQL)
		if isArray(rlCaModelo) then
			if len(trim(rlCaModelo(2,0))) > 0 then
				retLinkConfer	= "<a href=""camx.asp?oper=U&modelo="&rlCaModelo(0,0)&"&sala="&rlCaModelo(2,0)&""">on-line</a>"
			else
				retLinkConfer	= "off-line"
			end if
		else
			retLinkConfer	= "off-line"
		end if
	else
		retLinkConfer = "off-line"
	end if
End Function
'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'RETORNA Os dados dA MODELO selecionada
Function retDadosModelo(rdmModelo, rdmTipo)
DIM rdMaModelo
	if len(trim(rdmModelo)) > 0 then
		sSQL =	("SELECT	m.modelo, m.nome, m.sala " & _
				"FROM		XCA_modelos m " & _
				"WHERE	m.modelo = " & rdmModelo & " ")
		rdMaModelo	= getArray(sSQL)
		if isArray(rdMaModelo) then
		
			select case rdmTipo
			case "N"
				retDadosModelo = trim(rdMaModelo(1,0))
			case "S"
				retDadosModelo = trim(rdMaModelo(2,0))
			case else
				retDadosModelo = ""
			end select
		
		else
			retDadosModelo = ""
		end if
	else
		retDadosModelo = ""
	end if
End Function
'-------------------------------------------------------------------------

%>