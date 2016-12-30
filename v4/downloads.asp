<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Downloads</title>
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
                  Downloads :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text3"> 
                  <table width="90%" border="0" cellspacing="2" cellpadding="0">
                    <tr bgcolor="#CCCCCC"> 
                      <td rowspan="2" class="text3" bgcolor="#DEDEDE" width="20%" valign="middle"> 
                        <div align="center"><span class="text4">Fight 3</span><br>
                          (nota: 10)</div>
                      </td>
                      <td width="63%" bgcolor="#DEDEDE" class="text3"><span class="text4">Descri&ccedil;&atilde;o</span>: 
                        bonequinhos de pauzinho rolando o espanco geral! Vale 
                        a pena! </td>
                      <td width="17%" bgcolor="#DEDEDE" class="text3"> 
                        <div align="center" class="text4"><span class="text4">Tamanho</span>: 
                          <span class="text3"> 1,62MB</span></div>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#DEDEDE" class="text3" width="63%"> 
                        <div align="center"> 
                          <p><br>
                            <a href="http://www.saladois.hpg.com.br/downloads/fight3.EXE">fight3.EXE</a></p>
                        </div>
                      </td>
                      <td bgcolor="#DEDEDE" class="text3">
                        <div align="center" class="text2"><img src="imgs/iconflash.gif" width="40" height="40"></div>
                      </td>
                    </tr>
                  </table>
                  <p>&nbsp;</p>
                  <table width="90%" border="0" cellspacing="2" cellpadding="0">
                    <tr bgcolor="#CCCCCC"> 
                      <td rowspan="2" class="text3" bgcolor="#DEDEDE" width="20%" valign="middle"> 
                        <div align="center"><span class="text4">Bill Clinton</span><br>
                          (nota: 10)</div>
                      </td>
                      <td width="63%" bgcolor="#DEDEDE" class="text3"><span class="text4">Descri&ccedil;&atilde;o</span>: 
                        videozinho MPG muito bom sobre as falcatruas sexuais do 
                        ex-presidente norte-americano<br>
                        (meninas, tampem os olhos!) </td>
                      <td width="17%" bgcolor="#DEDEDE" class="text3"> 
                        <div align="center"><span class="text4">Tamanho:</span> 
                          3,213 Kb</div>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#DEDEDE" class="text3" width="63%">
                        <div align="center"><a href="http://www.saladois.hpg.com.br/downloads/billclinton.zip">billclinton.zip</a></div>
                      </td>
                      <td bgcolor="#DEDEDE" class="text3"> 
                        <div align="center" class="text2"><img src="imgs/iconcfucker.GIF" width="40" height="40"></div>
                      </td>
                    </tr>
                  </table>
                  <p>&nbsp;</p>
                  <table width="90%" border="0" cellspacing="2" cellpadding="0">
                    <tr bgcolor="#CCCCCC"> 
                      <td rowspan="2" class="text3" bgcolor="#DEDEDE" width="20%" valign="middle"> 
                        <div align="center"><span class="text4">Cavalinhos</span><br>
                          (nota: 10)</div>
                      </td>
                      <td width="63%" bgcolor="#DEDEDE" class="text3"><span class="text4">Descri&ccedil;&atilde;o</span>: 
                        se voc&ecirc; est&aacute; depressivo, pegue este arquivo 
                        e ensoberbe empolgantemente de alegria. Converse com os 
                        cavalinhos e pe&ccedil;a conselhos a eles, que s&atilde;o 
                        os nossos her&oacute;is. Be happy too!</td>
                      <td width="17%" bgcolor="#DEDEDE" class="text3"> 
                        <div align="center"><span class="text4">Tamanho:</span> 
                          567 Kb</div>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#DEDEDE" class="text3" width="63%"> 
                        <div align="center"><a href="http://www.saladois.hpg.com.br/downloads/cavalinhos.zip">cavalinhos.zip</a></div>
                      </td>
                      <td bgcolor="#DEDEDE" class="text3"> 
                        <div align="center" class="text2"><img src="imgs/iconhorse4.gif" width="40" height="40"></div>
                      </td>
                    </tr>
                  </table>
                  <p>&nbsp;</p>
                  <p>&nbsp;</p>
                  <p>&nbsp;</p>
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
function Enquete()
{
    window.open("espera.htm","enquete","toolbar=no,scrollbars=no,directories=no,status=no,menubar=no,resizable=yes,width=400,height=280");
}
</script>
</html>
