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
                  <p>Churrasc&atilde;o Sala&ETH;ois 2006 (23/12/06)  - P&aacute;gina 6/6</p>
                  <table width="100%" border="0" cellspacing="0" cellpadding="4">
                    <tr> 
                      <td> 
                        <div align="right"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200687.jpg','Olha o cofre!!!')"><img src="imgs/fotos/churrasco2006/tchurrasco200687.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td width="140"> 
                        <div align="center"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200688.jpg','Foto da galera')"><img src="imgs/fotos/churrasco2006/tchurrasco200688.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="left"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200689.jpg','Foto da galera 2')"><img src="imgs/fotos/churrasco2006/tchurrasco200689.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="10"> 
                        <div align="right"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200690.jpg','Mostra pra todo mundo!!')"><img src="imgs/fotos/churrasco2006/tchurrasco200690.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                      <td height="10" width="10"> 
                        <div align="center"></div>
                      </td>
                      <td height="10"> 
                        <div align="center"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200691.jpg','Foto do bolinho antiplay no Fegemo')"><img src="imgs/fotos/churrasco2006/tchurrasco200691.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                      <td height="10" width="10"> 
                        <div align="center"></div>
                      </td>
                      <td height="10"> 
                        <div align="left"><a href="javascript:Foto('1024','768','imgs\\fotos\\churrasco2006\\churrasco200692.jpg','Tira o saco!!!')"><img src="imgs/fotos/churrasco2006/tchurrasco200692.JPG" width="140" height="105" border="0"></a></div>
                      </td>
                    </tr>
                    <tr> 
                      <td> 
                        <div align="right"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="center"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="left"></div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="10"> 
                        <div align="right"></div>
                      </td>
                      <td width="10" height="10"> 
                        <div align="center"></div>
                      </td>
                      <td height="10"> 
                        <div align="center"></div>
                      </td>
                      <td width="10" height="10"> 
                        <div align="center"></div>
                      </td>
                      <td height="10"> 
                        <div align="left"></div>
                      </td>
                    </tr>
                    <tr> 
                      <td> 
                        <div align="right"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="center"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="left"></div>
                      </td>
                    </tr>
                    <tr> 
                      <td height="10"> 
                        <div align="right"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="center"></div>
                      </td>
                      <td width="10"> 
                        <div align="center"></div>
                      </td>
                      <td> 
                        <div align="left"></div>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="2"> 
                        <div align="center"></div>
                        <div align="right"></div>
                      </td>
                      <td> 
                        <div align="center"></div>
                      </td>
                      <td colspan="2"> 
                        <div align="left"></div>
                        <div align="center"></div>
                      </td>
                    </tr>
                  </table>
                  <p>P&aacute;gina<a href="fchurrasco2006a.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';"> 1 </a>..<a href="fchurrasco2006b.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';"> 2</a> .. <a href="fchurrasco2006c.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">3</a> .. <a href="fchurrasco2006d.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">4</a> .. <a href="fchurrasco2006e.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">5</a> .. <a href="fchurrasco2006f.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">6</a></p>
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
