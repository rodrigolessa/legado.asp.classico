//CRIA DINAMICAMENTE UMA CHAMADA PARA UM ARQUIVO SWF
function criaFlash(a, w, h) {
	document.writeln('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" ');
	document.write	(' codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" ');
	document.write	(' width="'+w+'" height="'+h+'">');
	document.writeln('<param name="movie" value="'+a+'">');
	document.writeln('<param name="quality" value="high">');
	document.writeln('<param name="menu" value="false" />');
	document.writeln('<embed src="'+a+'" quality="high" ');
	document.write	(' pluginspage="http://www.macromedia.com/go/getflashplayer" ');
	document.write	(' type="application/x-shockwave-flash" ');
	document.write	(' width="'+w+'" height="'+h+'"></embed>');
	document.writeln('</object>');
}

function criaXCam(a, w, h, n) {
	document.writeln('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" ');
	document.write	(' codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" ');
	document.write	(' width="'+w+'" height="'+h+'" id="'+n+'" align="middle">');
	document.writeln('<param name="allowScriptAccess" value="sameDomain" />');
	document.writeln('<param name="menu" value="false" />');
	document.writeln('<param name="bgcolor" value="#333333" />');
	document.writeln('<param name="movie" value="'+a+'" />');
	document.writeln('<param name="quality" value="high" />');
	document.writeln('<embed src="'+a+'" menu="false" quality="high" bgcolor="#333333" ');
	document.write	(' pluginspage="http://www.macromedia.com/go/getflashplayer" ');
	document.write	(' type="application/x-shockwave-flash" ');
	document.write	(' width="'+w+'" height="'+h+'" name="'+n+'" align="middle" allowScriptAccess="sameDomain" />');
	document.writeln('</object>');
}

function criaFlashVideo(a, p, w, h) {
	document.writeln('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="'+w+'" height="'+h+'" id="vplayer8" align="middle">');
	document.writeln('<param name="allowScriptAccess" value="sameDomain" />');
	document.writeln('<param name="movie" value="'+a+'?'+p+'" />');
	document.writeln('<param name="menu" value="false" />');
	document.writeln('<param name="quality" value="medium" />');
	document.writeln('<param name="wmode" value="transparent" />');
	document.writeln('<param name="bgcolor" value="#111111" />');
	document.writeln('<param name="FlashVars" value="'+p+'" />');
	document.writeln('<embed src="'+a+'?'+p+'" quality="high" bgcolor="#111111" width="'+w+'" height="'+h+'" name="vplayer8" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
	document.writeln('</object>');
}

//FUNÇÃO PARA LOGAR NO SITE
function logarSite() {
	document.forms[0].oper.value="G";
	document.forms[0].submit();
}

function abreJanela(a,w,h,s) {
	wnd=window.open(a,"wnd","toolbar=no,location=no,status=no,menubar=no,scrollbars=" + s + ",resizable=no,width=" + w + ",height=" + h + ",top=" + ((screen.height / 2) - (h / 2)) + ",left=" + ((screen.width / 2) - (w / 2)));
	wnd.focus();
}