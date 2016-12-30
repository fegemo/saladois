<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Integrantes</title>
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
                  Integrantes :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text4"> 
                  <p>Conhe&ccedil;a o perfil de cada integrante da SalaDois</p>
                  <p>&nbsp;</p>
                  <%
				  If Request.QueryString("cod") = "" Then
					Rs.Source = "SELECT * FROM TblIntegrantes ORDER BY NomeIntegrante ASC"
					Rs.Open()
					If Not Rs.EOF Then
					  Do Until Rs.EOF
						%>
                  <a href="integrantes.asp?cod=<%=Rs("CodigoIntegrante")%>" class="text2" onMouseOver="this.className='text3';" OnMouseOut="this.className='text2';"><%=Rs("NomeIntegrante")%></a><br>
                  <%
						Rs.MoveNext
					  Loop
					Else
					  %>
                  <br>
                  <br>
                  Nenhum integrante cadastrado<br>
                  <br>
                  <%
					End If
					Rs.Close()
				  Else
					Rs.Source = "SELECT * FROM TblIntegrantes WHERE CodigoIntegrante=" & Request.QueryString("cod")
					Rs.Open()
					If Not Rs.EOF Then
  					  %>
                  <br>
                  <table width="95%" border="0" cellspacing="3" cellpadding="0" class="text4">
                    <tr> 
                      <td bgcolor="#DEDEDE" colspan="2">Nome: <font class="text2"><%=Rs("NomeIntegrante")%></font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#DEDEDE" width="52%">Apelido: <font class="text2"><%=Rs("ApelidoIntegrante")%></font></td>
                      <td bgcolor="#DEDEDE" width="48%">Nascimento: <font class="text2"><%=Rs("NascimentoIntegrante")%></font></td>
                    </tr>
                    <tr> 
                      <td colspan="2"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="text4">
                          <tr bgcolor="#DEDEDE"> 
                            <td>ICQ: <font class="text2"><%=Rs("ICQIntegrante")%></font></td>
                            <td bgcolor="#FFFFFF" width="3">&nbsp;</td>
                            <td>E-mail: <font class="text2"><%=Rs("EmailIntegrante")%></font></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr bgcolor="#DEDEDE"> 
                      <td colspan="2"> 
                        <div align="center">
                          <% If Not Rs("FotoIntegrante") = "" Then %>
                          <a href="<%=Rs("FotoIntegrante")%>" target="_blank"><img border="0" src="<%=Rs("FotoIntegrante")%>" width="64" height="48"></a>
                          <%
							Else
							  Response.Write("<img src='imgs\xis.gif' width='40' height='40' alt='Foto não disponível'>")
							End If %>
                        </div>
                      </td>
                    </tr>
                  </table>
                  <br>
                  <%
					Else
					  %>
                  <br>
                  <br>
                  Integrante não cadastrado<br>
                  <br>
                  <%
					End If
					Rs.Close()
				  End If
				  %>
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
