<%
function exFotosGaleria(exGaleriaFoto)

	DIM aExFoto, exModelo, exMainDir, exStrFoto, exStrFoto2, exStatus
	
	exStatus = false
	
	'===== RETORNA NOME DAS FOTOS - GALERIAS =====
	sSQL = ("SELECT	mg.modelo, mg.modelo_galeria, gf.modelo_galeria_foto, gf.nomeArquivo, gf.nomeArquivo2  " & _
			"FROM			XCA_modelos_galerias mg " & _
			"	INNER JOIN	XCA_modelos_galerias_fotos gf ON mg.modelo_galeria = gf.modelo_galeria " & _
			"WHERE	gf.modelo_galeria_foto = " & exGaleriaFoto & " ")
	aExFoto = getArray(sSQL)
	if isArray(aExFoto) then
		exModelo	= trim(aExFoto(0,0))
		exStrFoto	= trim(aExFoto(3,0))
		exStrFoto2	= trim(aExFoto(4,0))
		exMainDir	= Application("XCA_APP_SERVERDIR_BASE") & "\modelos\" & exModelo & "\galeria\"
		
		if existArq(exMainDir & exStrFoto) then
			apagarArquivo(exMainDir & exStrFoto)
		end if
		
		if existArq(exMainDir & exStrFoto2) then
			apagarArquivo(exMainDir & exStrFoto2)
		end if
		
		if existArq(exMainDir & exStrFoto) then
			exStatus = false
		else
			if existArq(exMainDir & exStrFoto2) then
				exStatus = false
			else
				exStatus = true
			end if
		end if
		
		if exStatus then
			sSQL = "DELETE FROM XCA_modelos_galerias_fotos WHERE modelo_galeria_foto  = " & exGaleriaFoto & " "
			conn.Execute sSQL
		
			exFotosGaleria = true
		else
			exFotosGaleria = false
		end if
	else
		exFotosGaleria = false
	end if
	
end function
%>