<%

DIM	id_lang
	id_lang = 0

'SIGLAS DOS IDIOMAS
DIM	id_siglas(3)
	id_siglas(1) = "pt,pt-br"
	id_siglas(2) = "en,en-us,en-uk"
	id_siglas(3) = "es"

'NOMES DOS IDIOMAS
DIM	id_nomes(3)
	id_nomes(1) = "Português"
	id_nomes(2) = "English"
	id_nomes(3) = "Español"
	
'DIRETÓRIO NO SERVIDOR DO ARQUIVO XML DOS IDIOMAS
DIM DIR_XMLLNG
	DIR_XMLLNG = Application("XCA_APP_SERVERDIR_BASE") & "\includes\multilng\"

'IDIOMA PADRÃO
CONST ID_DEFAULT = 1
CONST ID_SESSION = "id_language"

'ARQUIVO XML DOS TEXTOS
'CONST ID_XMLFILE = "multilng/texto.xml"
CONST ID_XMLFILE = "texto.xml"

'OBJETO XML
DIM xmlIdioma
SET xmlIdioma = CreateObject("Microsoft.XMLDOM")
	xmlIdioma.async = false
	xmlIdioma.Load(DIR_XMLLNG & ID_XMLFILE)
	
	



'OBTÉM A LINGUAGEM SELECIONADA
function id_getLang(sig)
    id_getLang = "---"
    for i = 1 to ubound(id_siglas)
        if InStr(("," & id_siglas(i) & ",") , ("," & sig & ",")) > 0 then
        	id_getLang = i
        end if
    next
end function



'OBTÉM A LINGUAGEM SELECIONADA.
function id_getLanguage()

	DIM nLang, aLang, id_Accept

    'TENTA OBTER DA QUERYSTRING
    nLang = request.querystring("id_lang")

    if not IsNumeric(nLang) or nLang = ""  then

        'TENTA OBTER DA SESSION
        nLang = session(ID_SESSION)

        if not IsNumeric(nLang) or nLang = ""  then

            'TENTA OBTER DO NAVEGADOR
            id_Accept = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")

            if id_Accept <> "" then

				id_Accept = split(id_Accept,",")
				for i = 0 to Ubound(id_Accept)
					aLang = split(id_Accept(i),";")
					nLang = id_getLang(aLang(0))
					if isNumeric(nLang) then exit for
				next

            end if

        end if

    end if

    if not IsNumeric(nLang) or nLang = "" then
        'OBTEM A DEFAULT
        nLang = ID_DEFAULT
    end if

    session(ID_SESSION) = nLang
    id_getLanguage = nLang

end function



'OBTÉM O TEXTO DE ID N NO IDIOMA ATUAL
function id_getText(n)

	DIM tsIdioma
	
	if id_lang = 0 then
		id_lang = id_getLanguage()
    end if
    
    SET tsIdioma = xmlIdioma.selectNodes("//texto[@id='" & n & "']/idioma[@id='" & id_lang & "']")
    if tsIdioma.length > 0 then
    	id_getText = tsIdioma(0).text
    end if
    
end function

%>