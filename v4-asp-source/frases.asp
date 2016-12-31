<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Frases</title>
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
                <td height="17" background="imgs/41.gif" class="text3mm" align="center">
                  .: Frases :.<br>
                  </td></tr><tr><td valign="top" align="center"><br><br><br>
                  <form name="" method="post" action="salvarsl.asp?func=frases">
                    <p align="left" class="text3">Este espa&ccedil;o &eacute; 
                      reservado para o visitante SalaDois que pretende expressar 
                      seu ponto de vista ou apenas um pensamento sabiamente pensado:</p>
                    <p align="left"><font class="text3">Autor: 
                      <input type="text" name="Autor" maxlength="100" class="textbox1">
                      (m&aacute;ximo de caracteres: 100)<br>
                      Frase: 
                      <input type="text" name="Texto" size="70" class="textbox1" maxlength="150">
                      <br>
                      Senha SalaDois: 
                      <input type="password" name="Senha" class="textbox1" size="6" maxlength="6">
                      </font></p>
                    <p align="center"><font class="text2">&Eacute;, parece que 
                      esses fracassados da oitava s&eacute;rie perderam huahuahauhauahuha</font></p>
                    <p align="left"><font class="text3"><br>
                      Obs.: N&atilde;o utilize par&ecirc;nteses nem &aacute;spas. 
                      Eles ser&atilde;o automaticamente adicionados &agrave; sua 
                      postagem</font></p>
                    <p>
                      <input type="submit" value="Enviar" class="botao1">
                      <input type="reset" value="Limpar" name="Reset" class="botao1">
                      </p>
                  </form>
                  <br><br>
                  <br><%
					Rs.Source = "SELECT * FROM TblFrases ORDER BY CodigoFrase DESC"
					Rs.Open()
					Do Until Rs.EOF
					  %><font class="text4">"<%=Rs("TextoFrase")%>"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(<%=Rs("AutorFrase")%>)<br><br></font><%
					  Rs.MoveNext
					Loop
					Rs.Close()
					%><br><br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text4"> </td>
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
