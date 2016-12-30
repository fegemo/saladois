<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Rs.Source = "SELECT * FROM TblPaginas WHERE CodigoPagina=2"
Rs.Open()
%>
<html>
<head>
<title><%=Rs("TituloPagina")%></title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="#DEDEDE" text="#000000">
<%
'Votar
If Request.QueryString("action") = "votar" Then
  If Not Request.Form("radiobutton") = 0 Then
    Rs.Close()
    Rs.LockType = 3
    Rs.Source = "SELECT * FROM TblOpcoes WHERE CodigoOpcao=" & Request.Form("radiobutton")
    Rs.Open()
    Rs("VotosOpcao") = Rs("VotosOpcao") + 1
    Rs.Update()
  End If
End If

Rs.Close()
If Not Request.QueryString("cod") = "" Then
Rs.Source = "SELECT TblEnquetes.*, TblOpcoes.* FROM TblEnquetes INNER JOIN TblOpcoes ON TblEnquetes.CodigoEnquete = TblOpcoes.EnqueteOpcao WHERE (((TblEnquetes.CodigoEnquete)=" & Request.QueryString("cod") & "));"
Rs.Open()
Dim TotalVotos
Do Until Rs.EOF
  TotalVotos = TotalVotos + Rs("VotosOpcao")
  Rs.MoveNext
Loop
If Not Rs.BOF Then
  Rs.MoveFirst
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td valign="middle" align="center"> 
      <table width="400" border="0" cellspacing="0" cellpadding="0" height="280">
        <tr> 
          <td valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="2" class="text2m"> 
                  <div align="center"><b><img src="imgs/int.gif" width="16" height="16"><img src="imgs/int.gif" width="16" height="16">-:<span class="text3mm"><%=Rs("TituloEnquete")%></span> :-<img src="imgs/int.gif" width="16" height="16"><img src="imgs/int.gif" width="16" height="16"></b></div>
                  <br>
                </td>
              </tr>
              <tr> 
                <td colspan="2" class="text2m" valign="middle"><%=Rs("PerguntaEnquete")%>&nbsp;<span class="text3m">[<%=TotalVotos%> voto(s)]</span><br>
                  <br>
                </td>
              </tr>
              <%
			  Dim Contador
			  Do Until Rs.EOF
				Contador = Contador + 1
				If Contador = 5 Then Contador = 1
			  %>
              <tr bgcolor="#A6A6A6"> 
                <td class="text4"><%=Rs("TextoOpcao")%>&nbsp;<span class="text3">[<%=Rs("VotosOpcao")%> voto(s)]</span></td>
                <% If Not TotalVotos = 0 Then %>
                <td class="text5"><img src="imgs/vot<%=Contador%>.gif" height="12" width="<%=(Int((100 * Rs("VotosOpcao")) / TotalVotos) * 2) + 6%>">&nbsp;<%=Int((100 * Rs("VotosOpcao")) / TotalVotos)%>%</td>
                <% Else %>
                <td class="text5" width="40%"><img src="imgs/vot<%=Contador%>.gif" height="12" width="6">&nbsp;0%</td>
                <% End If %>
              </tr>
              <tr bgcolor="#DEDEDE"> 
                <td bgcolor="#DEDEDE" height="1" colspan="2"><img src="imgs/falso.gif" width="1" height="4"></td>
              </tr>
              <%
				Rs.MoveNext
			  Loop
			  Rs.Close()
			  Rs.Source = "SELECT * FROM TblEnquetes WHERE CodigoEnquete=" & Request.QueryString("cod")
			  Rs.Open()
			  %>
              <tr> 
                <td colspan="2" class="text3" align="center"><br><br><br>Enquete ativa desde: <%=Rs("DataInicioEnquete")%></td>
              </tr>
			  <%
			  Rs.Close()
			  %>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
End If
End If
%>
</body>
</html>
