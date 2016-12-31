<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Fotos</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF" text="#000000">
<table width="780" border="0" cellspacing="0" cellpadding="0">
  <!--#include file="topo.htm" -->
  <tr> 
    <td> 
      <!--#include file="banner.asp" -->
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr valign="top"> 
          <td width="126" height="100%"> 
            <!--#include file="esquerda.asp" -->
          </td>
          <td width="509" height="100%"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
              <tr> 
                <td height="17" background="imgs/41.gif" class="text3mm" align="center">.: 
                  Fotos :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text3m">
                  <p>Churrasc&atilde;o Sala&ETH;ois 2007 (22/12/07)  - P&aacute;gina 2/2</p>
                  <table width="100%" border="0" cellspacing="0" cellpadding="4">
                    <tr> 
                      <td> 
                        <div align="right"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200720.jpg','Ridículas!!! 2')"><img src="imgs/fotos/churrasco2007/tchurrasco200720.JPG" width="140" height="105" border="0"></a></div>                      </td>
                      <td width="10"> 
                        <div align="center"></div>                      </td>
                      <td width="140"> 
                        <div align="center"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200722.jpg','Bolinho padrão rox (pré-avacalhação)')"><img src="imgs/fotos/churrasco2007/tchurrasco200722.jpg" width="140" height="105" border="0"></a></div>                      </td>
                      <td width="10"> 
                        <div align="center"></div>                      </td>
                      <td> 
                        <div align="left"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200721.jpg','Bolinha padrão rox (pós-avacalhação)')"><img src="imgs/fotos/churrasco2007/tchurrasco200721.JPG" width="140" height="105" border="0"></a></div>                      </td>
                    </tr>
                    <tr> 
                      <td height="10"> 
                        <div align="right"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200723.jpg','É disso aqui q eu gosto, ó')"><img src="imgs/fotos/churrasco2007/tchurrasco200723.JPG" width="140" height="105" border="0"></a></div>                      </td>
                      <td height="10" width="10"> 
                        <div align="center"></div>                      </td>
                      <td height="10"> 
                        <div align="center"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200724.jpg','Arnaldão, ou sr. pêlos! O salvador do churrasco!! :D ')"><img src="imgs/fotos/churrasco2007/tchurrasco200724.JPG" width="140" height="105" border="0"></a></div>                      </td>
                      <td height="10" width="10"> 
                        <div align="center"></div>                      </td>
                      <td height="10"> 
                        <div align="left"><a href="javascript:Foto('960','1280','imgs\\fotos\\churrasco2007\\churrasco200725.jpg','I like food, man')"><img src="imgs/fotos/churrasco2007/tchurrasco200725.JPG" width="140" height="187" border="0"></a></div>                      </td>
                    </tr>
                    <tr> 
                      <td> 
                        <div align="right"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200726.jpg','Sessão wiizin')"><img src="imgs/fotos/churrasco2007/tchurrasco200726.JPG" width="140" height="105" border="0"></a></div>                      </td>
                      <td width="10"> 
                        <div align="center"></div>                      </td>
                      <td> 
                        <div align="center"><a href="javascript:Foto('960','720','imgs\\fotos\\churrasco2007\\churrasco200711.jpg','Friends again =)')"><img src="imgs/fotos/churrasco2007/tchurrasco200711.jpg" width="140" height="105" border="0"></a></div>                      </td>
                      <td width="10"> 
                        <div align="center"></div>                      </td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="5"> 
                        <p class="text2m" align="center">Download do pacote <a href="http://www.dcc.ufmg.br/~flavioro/saladois/churrasco2007.zip">churrasco2007.zip </a></p>
                      </td>
                    </tr>
                  </table>
                  <p>P&aacute;gina<a href="fchurrasco2007a.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';"> 1 </a>..<a href="fchurrasco2007b.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';"> 2</a></p>
                </td>
              </tr>
              <tr> 
                <td height="23" bgcolor="13AB13">&nbsp;</td>
              </tr>
            </table>
          </td>
          <td height="100%"> 
            <!--#include file="direita.asp" -->
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
<script language="JavaScript">
function Foto(largura,altura,nome,texto)
{
    window.open("foto.asp?largura=" + largura + "&altura=" + altura + "&nome=" + nome + "&texto=" + texto + "","foto","scrollbars=yes,toolbar=no,directories=no,status=no,menubar=no,resizable=yes,width=" + largura + ",height=" + eval(parseInt(altura)+132) + "");
}
function Enquete()
{
    window.open("espera.htm","enquete","toolbar=no,scrollbars=no,directories=no,status=no,menubar=no,resizable=yes,width=400,height=280");
}
</script>
</html>
