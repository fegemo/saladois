
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html dir=ltr>

<head>
<style>
a:link			{font:8pt/11pt verdana; color:FF0000}
a:visited		{font:8pt/11pt verdana; color:#4e4e4e}
</style>

<META NAME="ROBOTS" CONTENT="NOINDEX">

<title>N�o � poss�vel exibir a p�gina</title>

<META HTTP-EQUIV="Content-Type" Content="text-html; charset=Windows-1252">
</head>

<script> 
function Homepage(){
<!--
// in real bits, urls get returned to our script like this:
// res://shdocvw.dll/http_404.htm#http://www.DocURL.com/bar.htm 

	//For testing use DocURL = "res://shdocvw.dll/http_404.htm#https://www.microsoft.com/bar.htm"
	DocURL=document.URL;
	
	//this is where the http or https will be, as found by searching for :// but skipping the res://
	protocolIndex=DocURL.indexOf("://",4);
	
	//this finds the ending slash for the domain server 
	serverIndex=DocURL.indexOf("/",protocolIndex + 3);

	//for the href, we need a valid URL to the domain. We search for the # symbol to find the begining 
	//of the true URL, and add 1 to skip it - this is the BeginURL value. We use serverIndex as the end marker.
	//urlresult=DocURL.substring(protocolIndex - 4,serverIndex);
	BeginURL=DocURL.indexOf("#",1) + 1;
	urlresult=DocURL.substring(BeginURL,serverIndex);
		
	//for display, we need to skip after http://, and go to the next slash
	displayresult=DocURL.substring(protocolIndex + 3 ,serverIndex);
	InsertElementAnchor(urlresult, displayresult);
}

function HtmlEncode(text)
{
    return text.replace(/&/g, '&amp').replace(/'/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function TagAttrib(name, value)
{
    return ' '+name+'="'+HtmlEncode(value)+'"';
}

function PrintTag(tagName, needCloseTag, attrib, inner){
    document.write( '<' + tagName + attrib + '>' + HtmlEncode(inner) );
    if (needCloseTag) document.write( '</' + tagName +'>' );
}

function URI(href)
{
    IEVer = window.navigator.appVersion;
    IEVer = IEVer.substr( IEVer.indexOf('MSIE') + 5, 3 );

    return (IEVer.charAt(1)=='.' && IEVer >= '5.5') ?
        encodeURI(href) :
        escape(href).replace(/%3A/g, ':').replace(/%3B/g, ';');
}

function InsertElementAnchor(href, text)
{
    PrintTag('A', true, TagAttrib('HREF', URI(href)), text);
}

//-->
</script>

<body bgcolor="FFFFFF">

<table width="410" cellpadding="3" cellspacing="5">

  <tr>    
    <td align="left" valign="middle" width="360">
	<h1 style="COLOR:000000; FONT: 13pt/15pt verdana"><!--Problem-->N�o � poss�vel exibir a p�gina</h1>
    </td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 8pt/11pt verdana">Ocorreu um problema com a p�gina que voc� est� tentando acessar e n�o � poss�vel exibi-la.</font></td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 8pt/11pt verdana">

	<hr color="#C0C0C0" noshade>
	
    <p>Experimente o seguinte:</p>

	<ul>
      <li id="instructionsText1">Clique no bot�o 
      <a href="javascript:location.reload()">
      Atualizar</a> ou tente novamente mais tarde.<br>
      </li>
	  
      <li>Abra a 
	  
	  <script>
	  <!--
	  if (!((window.navigator.userAgent.indexOf("MSIE") > 0) && (window.navigator.appVersion.charAt(0) == "2")))
	  {
	  	 Homepage();
	  }
	  //-->
	  </script>

	  home page e procure os links para as informa��es desejadas. </li>
    </ul>
	
    <h2 style="font:8pt/11pt verdana; color:000000">HTTP 500.100 - Servidor interno
    Erro - erro do ASP<br>
    Internet Information Services</h2>

	<hr color="#C0C0C0" noshade>
	
	<p>Informa��es t�cnicas (para a equipe de suporte)</p>

<ul>
<li>Tipo de erro:<br>
Microsoft VBScript compilation  (0x800A03F6)<br>Expected 'End'<br><b>/saladois/amigo oculto/salvar.asp, line 6</b><br>
</li>
<p>
<li>Tipo de navegador: <br>
Mozilla/4.5 (compatible; HTTrack 3.0x; Windows 98)
</li>
<p>
<li>P�gina: <br>
GET /saladois/amigo oculto/salvar.asp</li>
<p>
<li>Hora: <br>
Friday, December 30, 2016, 12:06:19 PM
</li>
</p>
<p>
<li>Mais informa��es: <br>
 
<a href="http://www.microsoft.com/ContentRedirect.asp?prd=iis&sbp=&pver=5.0&ID=500;100&cat=Microsoft+VBScript+compilation+&os=&over=&hrd=&Opt1=&Opt2=%2D2146827274&Opt3=Expected+%27End%27">Suporte da Microsoft</a>
</li>
</p>

    </font></td>
  </tr>
  
</table>
</body>
</html>












