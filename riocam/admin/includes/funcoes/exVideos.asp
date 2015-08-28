<%
function exVideos(exCodVideo)

	DIM aExFoto, exModelo, exMainDir, exStrFoto, exStrFoto2, exStatus
	
	exStatus = false
	
	'===== RETORNA NOME DAS FOTOS - GALERIAS =====
	sSQL = ("SELECT	v.modelo, gf.nomeArquivo, gf.nomeArquivo2  " & _
			"FROM			XCA_modelos_videos v " & _
			"WHERE	v.modelo_video = " & exCodVideo & " ")
	aExFoto = getArray(sSQL)
	if isArray(aExFoto) then
		exModelo	= trim(aExFoto(0,0))
		exStrFoto	= trim(aExFoto(1,0))
		exStrFoto2	= trim(aExFoto(2,0))
		exMainDir	= Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & exModelo & "\galeria\"
		
		if existArq(exMainDir & exStrFoto) then
			apagarArquivo(exMainDir & exStrFoto)
		end if
		
		if existArq(exMainDir & exStrFoto2) then
			apagarArquivo(exMainDir & exStrFoto2)
		end if
		
		if existArq(exMainDir & exStrFoto) then
			exVideos = false
		else
			if existArq(exMainDir & exStrFoto2) then
				exVideos = false
			else
				exVideos = true
			end if
		end if
	else
		exVideos = false
	end if
	
end function
%>