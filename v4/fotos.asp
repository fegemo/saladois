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
                <td align="center" valign="top" class="text4"><p><a href="fchurrasco2007a.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Churrasc&atilde;o Sala&ETH;ois 2007 (22/12/07) </a><a href="http://www.dcc.ufmg.br/~flavioro/saladois/churrasco2007.zip"><img src="imgs/iconrar.gif" alt="Download do pacote inteiro" width="20" height="17" border="0" align="absmiddle"></a></p>
                  <p><a href="fchurrasco2006a.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Churrasc&atilde;o Sala&ETH;ois 2006 (23/12/06) </a><a href="http://www.dcc.ufmg.br/~flavioro/saladois/churrasco2006.zip"><img src="imgs/iconrar.gif" alt="Download do pacote inteiro" width="20" height="17" border="0" align="absmiddle"></a></p>
                  <p><a href="fbomburgao1a.asp" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Bomburg&atilde;o 
                  1</a> <a href="http://www.saladois.hpg.com.br/downloads/fotos1.zip"><img src="imgs/iconrar.gif" alt="Download do pacote inteiro" width="20" height="17" border="0" align="absmiddle"></a></p></td>
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
function Enquete()
{
    window.open("espera.htm","enquete","toolbar=no,scrollbars=no,directories=no,status=no,menubar=no,resizable=yes,width=400,height=280");
}
</script>
</html>
