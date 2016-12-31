<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>Jornais</title>
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
                  Jornais :.<br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td align="center" valign="top" class="text4"> 
                  <%
					If Request.QueryString("cod") = "" Then
					  Rs.Source = "SELECT * FROM TblJornais"
					  Rs.Open()
					  cJornais = 0
					  Do Until Rs.EOF
						%>
                  <a href="jornais.asp?cod=<%=Rs("CodigoJornal")%>" class="text2" onMouseOver="this.className='text3';" onMouseOut="this.className='text2';"><%=Rs("EdicaoJornal")%></a><br>
                  <%
						cJornais = cJornais + 1
						Rs.MoveNext
					  Loop
					%><br><br>
                  Total de edi&ccedil;&otilde;es -> <%=cJornais%><%
					Else
					  Rs.Source = "SELECT * FROM TblJornais WHERE CodigoJornal=" & Request.QueryString("cod")
					  Rs.Open()
					  If Not Rs.EOF Then
						Dim J1, J2, J3, J4, J5, J6, J7, J8
j1 = Rs("N1Jornal")
j2 = Rs("N2Jornal")
j3 = Rs("N3Jornal")
j4 = Rs("N4Jornal")
j5 = Rs("N5Jornal")
j6 = Rs("N6Jornal")
j7 = Rs("N7Jornal")
j8 = Rs("N8Jornal")
						%><br>
                  <table cellpadding=0 cellspacing=0 width="95%">
                    <tr align="center">
                      <td bgcolor="#DEDEDE" width="482"><font face="Verdana, Arial, Helvetica" size="5"><b><%=Rs("EdicaoJornal")%></b></font><br></td></tr><tr class="text3m">
                      <td width="482"><br>
                        <table width="100%" border="0" cellspacing="2" cellpadding="2" class="text3" bgcolor="#FFFFFF">
                          <tr valign="top">
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J1,1,1) & "</b></font>" & Mid(J1,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J2,1,1) & "</b></font>" & Mid(J2,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J3,1,1) & "</b></font>" & Mid(J3,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J4,1,1) & "</b></font>" & Mid(J4,2)%></td>
                          </tr>
                          <tr valign="top">
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J5,1,1) & "</b></font>" & Mid(J5,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J6,1,1) & "</b></font>" & Mid(J6,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J7,1,1) & "</b></font>" & Mid(J7,2)%></td>
                            <td bgcolor="#DEDEDE"><%="<font size='4'><b>" & Mid(J8,1,1) & "</b></font>" & Mid(J8,2)%></td>
                          </tr>
                        </table>
                        <br>
                      </td></tr><tr>
                      <td align="left" class="text3" bgcolor="#DEDEDE" width="482"><%=Rs("AutorJornal")%></td></tr></table>
                  <%
					  Else
						Response.Redirect "jornais.asp"
					  End If
					End If
					Rs.Close()
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
