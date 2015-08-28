function lib_bwcheck(){ //Browsercheck (needed)
	this.ver=navigator.appVersion
	this.agent=navigator.userAgent
	this.dom=document.getElementById?1:0
	this.opera5=this.agent.indexOf("Opera 5")>-1
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom && !this.opera5)?1:0; 
	this.ie6=(this.ver.indexOf("MSIE 6")>-1 && this.dom && !this.opera5)?1:0;
	this.ie4=(document.all && !this.dom && !this.opera5)?1:0;
	this.ie=this.ie4||this.ie5||this.ie6
	this.mac=this.agent.indexOf("Mac")>-1
	this.ns6=(this.dom && parseInt(this.ver) >= 5) ?1:0; 
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie6 || this.ie5 || this.ie4 || this.ns4 || this.ns6 || this.opera5)
	return this
}
var bw=new lib_bwcheck()

/**************************************************************************
Variables to set.
***************************************************************************/
sLeft = 0			//The left placement of the menu
sTop = 0			//The top placement of the menu
sMenuheight = 40	//The height of the menu
sArrowwidth = 11	//Width of the arrows
sScrollspeed = 20	//Scroll speed: (in milliseconds, change this one and the next variable to change the speed)
sScrollPx = 8		//Pixels to scroll per timeout.
sScrollExtra = 15	//Extra speed to scroll onmousedown (pixels)

/**************************************************************************
Scrolling functions
***************************************************************************/
var tim = 0
var noScroll = true
function mLeft(){
	if (!noScroll && oMenu.x<sArrowwidth){
		oMenu.moveBy(sScrollPx,0)
		tim = setTimeout("mLeft()",sScrollspeed)
	}
}
function mRight(){
	if (!noScroll && oMenu.x>-(oMenu.scrollWidth-(pageWidth))-sArrowwidth){
		oMenu.moveBy(-sScrollPx,0)
		tim = setTimeout("mRight()",sScrollspeed)
	}
}
function noMove(){
	clearTimeout(tim);
	noScroll = true;
	sScrollPx = sScrollPxOriginal;
}
/**************************************************************************
Object part
***************************************************************************/
function makeObj(obj,nest,menu){
	nest = (!nest) ? "":'document.'+nest+'.';
	this.elm = bw.ns4?eval(nest+"document.layers." +obj):bw.ie4?document.all[obj]:document.getElementById(obj);
   	this.css = bw.ns4?this.elm:this.elm.style;
	this.scrollWidth = bw.ns4?this.css.document.width:this.elm.offsetWidth;
	this.x = bw.ns4?this.css.left:this.elm.offsetLeft;
	this.y = bw.ns4?this.css.top:this.elm.offsetTop;
	this.moveBy = b_moveBy;
	this.moveIt = b_moveIt;
	this.clipTo = b_clipTo;
	return this;
}
var px = bw.ns4||window.opera?"":"px";
function b_moveIt(x,y){
	if (x!=null){this.x=x; this.css.left=this.x+px;}
	if (y!=null){this.y=y; this.css.top=this.y+px;}
}
function b_moveBy(x,y){this.x=this.x+x; this.y=this.y+y; this.css.left=this.x+px; this.css.top=this.y+px;}
function b_clipTo(t,r,b,l){
	if(bw.ns4){this.css.clip.top=t; this.css.clip.right=r; this.css.clip.bottom=b; this.css.clip.left=l;}
	else this.css.clip="rect("+t+"px "+r+"px "+b+"px "+l+"px)";
}
/**************************************************************************
Object part end
***************************************************************************/

/**************************************************************************
Init function. Set the placements of the objects here.
***************************************************************************/
var sScrollPxOriginal = sScrollPx;
function sideInit(){
	//Width of the menu, Currently set to the width of the document.
	//If you want the menu to be 500px wide for instance, just 
	//set the pageWidth=500 in stead.
	pageWidth = 600//(bw.ns4 || bw.ns6 || window.opera)?innerWidth:document.body.clientWidth;
	
	//Making the objects...
	oBg = new makeObj('divBg')
	oMenu = new makeObj('divMenu','divBg',1)
	oArrowRight = new makeObj('divArrowRight','divBg')
	
	//Placing the menucontainer, the layer with links, and the right arrow.
	oBg.moveIt(sLeft,sTop) //Main div, holds all the other divs.
	oMenu.moveIt(sArrowwidth,null)
	oArrowRight.css.width = sArrowwidth;
	oArrowRight.moveIt(pageWidth-sArrowwidth,null)
	
	//Setting the width and the visible area of the links.
	if (!bw.ns4) oBg.css.overflow = "hidden";
	if (bw.ns6) oMenu.css.position = "relative";
	oBg.css.width = pageWidth+px;
	oBg.clipTo(0,pageWidth,sMenuheight,0)
	oBg.css.visibility = "visible";
}

//executing the init function on pageload if the browser is ok.
//Copiar esta linha para a página que contem o scroll
//Iniciar o scroll:
//if (bw.bw) onload = sideInit;


function openDlgCalendario(c, d) {
	var f = document.forms[0];
	var sURL = "dlgCalendario.asp?fData=" + c + "&data=" + d;
	var sName = "wDlgCalendario";
	var features = "left=400,top=150,width=200,height=172";
	_WCal=window.open(sURL, sName, features);
	_WCal.focus();
}


// load htmlarea
_editor_url = ""; // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
	document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
	document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { 
	document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); 
}


	function validarEmail(valor) {
	  if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(valor)){
		return true
	  } else {
		return false;
	  }
	}
	
//script para mudar cores da tabela de lista
	function onColor1(blah)
	{
		blah.style.backgroundColor='#F3CE81';
	}

	function offColor1(blah)
	{
		blah.style.backgroundColor='#CBCBCB';
	}
	function onColor2(blah)
	{
		blah.style.backgroundColor='#F3CE81';
	}

	function offColor2(blah)
	{
		blah.style.backgroundColor='#DFDFDF';
	}
	
	
	
// **** FUNÇÕES DE VALIDAÇÃO DE TECLAS ****
	
//script para validacao de digitacao e salto de campo automatico
// Colocar o foco em determinado campo
function SetarFoco(ind) {
	InicializarIndices();
	if ( isNaN(ind) && document.forms[0].elements[ind].type!="hidden" && !document.forms[0].elements[ind].disabled)
		document.forms[0].elements[ind].focus();
	else
		for (;ind<document.forms[0].elements.length;ind++)
			if (document.forms[0].elements[ind].type!="hidden" && !document.forms[0].elements[ind].disabled)
				break;
		if (ind<=document.forms[0].elements.length)
			document.forms[0].elements[ind].focus();
	}

// Verificar qual navegador
function QualNavegador() {
	var s = navigator.appName
	if(s == "Microsoft Internet Explorer") return "IE";
	else if ( s == "Netscape" ) return "NE";
	else return "";
}

// Verificar qual a versão do navegador
function QualVersao() {
	var s = navigator.appVersion;
	if ( QualNavegador() == "IE" ) {
		var i = s.search("MSIE");
		s=s.substring(i+5);
		i=s.search(".");
		return parseInt(s.substring(0,i+1));
	}
	else if ( QualNavegador() == "NE" )	return parseInt(s.substring(0,1));
	else return 0;
}

function InicializarIndices() {
	if (document.CargaInicial==null) {
		document.CargaInicial=false; // Seta para só fazer uma vez por documento
		var ctrlAnterior=null;
		var IndAnt=0;
		for ( var i=0; i<document.forms[0].elements.length;i++)	{
			var e=document.forms[0].elements[i];
			if ( e.type!="hidden" && e.type!="image" ) {
				if (ctrlAnterior != null) ctrlAnterior.IndicePosterior=i;
				ctrlAnterior=e;
				e.Indice=i;
				e.IndiceAnterior=IndAnt;
			}
		}
	}
}

// Setar o evento
function SetarEvento(ctrl, Tam, Tipo, AutoSkip) { // Filtra navegadores conhecidos
	var s = QualNavegador();
	if (s.length==0) return;
	if (s=="IE" && QualVersao()>6) return;
	if (s=="NE" && QualVersao()>4) return;
	if (ctrl.onkeypress==null) {
		if (AutoSkip==null) AutoSkip=true;
		if (Tipo!=null)	Tipo.toUpperCase();
		ctrl.Tam=Tam;
		ctrl.Tipo=Tipo;
		ctrl.AutoSkip=true;
		ctrl.Saltar=false;
		InicializarIndices();
		ctrl.onkeypress=ValidarTecla;
		if (QualNavegador()=="IE" && QualVersao()>=5) ctrl.onkeyup=SaltarCampo;
	}
}

function SaltarCampo(ctrl) {
	if (ctrl==null)	ctrl=this;
	if ( ctrl.AutoSkip && ctrl.Saltar)
		if (ctrl.Saltar) {
			ctrl.Saltar=false;
			if ( ctrl.IndicePosterior != null ) SetarFoco(ctrl.IndicePosterior);
		}
}

// Fazer o salto de campo
function ValidarTecla(evnt) {
	var tk;
    var c;
	// Recebe a tela pressionada
	tk = ( (QualNavegador()=="IE") ? event.keyCode : evnt.which);
    c=String.fromCharCode(tk);
	c=c.toUpperCase();
	// Só aceita teclas alfanuméricas. Não aceita teclas de controle
    if(tk<32) return true;
	if (tk>255)	return false;

	switch (this.Tipo) {
	case "I":
		if (c!="") {
       		return false;
       	}
		break;
	case "D":
		if ((c<"0" || c>"9") && (c!="/")) {
       		return false;
       	}
		break;
	case "P":
		var vlcep = this.value
		if (vlcep.length == 5) this.value = vlcep + ".";
		//return false;
		break;
	case "F":
		var vlcpf = this.value
		if (vlcpf.length == 3) this.value = vlcpf + ".";
		if (vlcpf.length == 7) this.value = vlcpf + ".";
		if (vlcpf.length == 11) this.value = vlcpf + "-";
		//return false;
		break;
	case "N":
		if ((c<"0" || c>"9") && (c!="." && c!=","))
			return false;
		if ((c==",") && ((this.value.search(",")>-1) || (this.value.length==0)))
			return false;
		if ((c==".") && (this.value.length==0))
			return false;
		break;
	case "C":
		if ( c<"A" || c>"Z" ) return false;
		break;
	default:
		break;
	}
	this.Saltar=(this.value.length==this.Tam-1);
	if(((QualNavegador()=="IE") && QualVersao()<5) || (QualNavegador()!="IE")) SaltarCampo(this);
	return true;
}