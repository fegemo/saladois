<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Bomburg&atilde;o</title>
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
                  Bomburg&atilde;o :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text4"> 
                  <p> Para os leigos no assunto, o BOMBURG&Atilde;O &eacute; um 
                    recanto SalaDoisiano, ou seja, &eacute; um lugar onde todos 
                    os indiv&iacute;duos que se dizem SalaDois se re&uacute;nem 
                    para praticarem perip&eacute;cias e (&eacute; claro!) comerem 
                    um ou dois sandu&iacute;ches.<br>
                    O cidad&atilde;o genuinamente SalaDoisiano visita o recinto 
                    pelo menos uma vez a cada final de semana, isso devido &agrave; 
                    um pacto de sangue que concretizamos com o dono do santu&aacute;rio. 
                    O pacto consiste no seguinte: temos que visitar o local todos 
                    os fins de semana e comer, pelo menos, 40 Bomburg&otilde;es 
                    a cada m&ecirc;s. Voc&ecirc; deve estar se perguntando: &quot;E 
                    o que os SalaDoisianos ganham com isso???&quot;. Eles ganham 
                    o melhor de tudo: Ganham inteiramente de gr&aacute;tis uma 
                    carteirinha da filia&ccedil;&atilde;o BOMBURG&Atilde;O-SALADOIS 
                    com descontos no conjunto de duas fatias de p&atilde;o (sandu&iacute;che! 
                    D&atilde;h...).<br>
                    E voc&ecirc; pensa que acabou??? N&atilde;o acabou n&atilde;o!!! 
                    Sempre que voc&ecirc; for pagar a sua conta, ganhar&aacute; 
                    totalmente de gra&ccedil;a e sem nenhum custo adicional balas 
                    de gr&aacute;tis! Balas por conta da multinacional Bomburg&atilde;o! 
                    Mas, o que voc&ecirc; far&aacute; com essas balas??? Bem, 
                    a&iacute; a vontade &eacute; do cliente! Se quiser jog&aacute;-las 
                    no telhado, (n&eacute; P&atilde;o??) chup&aacute;-las... voc&ecirc; 
                    &eacute; que se decide!<br>
                    Agora que j&aacute; sabe o que &eacute; o Bomburg&atilde;o, 
                    entre e vejas os recordes!<br>
                  </p>
                  <table width="400" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="middle"> 
                      <td width="200"> 
                        <div align="center">
						<%
						Rs.Source = "SELECT * FROM TblCategorias ORDER BY NomeCategoria ASC"
						Rs.Open()
						Do Until Rs.EOF
						  %><a href="bombs.asp?cat=<%=Rs("CodigoCategoria")%>" target="bombs" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">:: <%=Rs("NomeCategoria")%> ::</a><br><br><%
						  Rs.MoveNext
						Loop
						Rs.Close()
						%>
						</div>
                      </td>
                      <td width="200"><iframe name="bombs" src="bombs.asp" width="200" frameborder=0 marginwidth=0 marginheight=0 scrolling="auto"></iframe></td>
                    </tr>
                  </table>
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
