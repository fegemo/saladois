<%@LANGUAGE="VBSCRIPT"%> 
<%
If Session("Logged")="" Then
  Response.Redirect"login.htm"
End If
%>
<!--#include file="Connections/Conneccao.asp" -->
<!--#include file="functions.asp" -->
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Set lConsulta = Server.CreateObject("ADODB.Recordset")
lConsulta.ActiveConnection = MM_Conneccao_STRING
%>
<html>
<head>
<title>-=Gerenciamento.:.SalaDois=-</title>
<link rel="stylesheet" href="csss/sal.css" type="text/css">
</head>
<body bgcolor="#DEDEDE" text="#000000">
<table width="100%" border="1" cellspacing="3" cellpadding="0" height="100%">
  <tr valign="top" bgcolor="#FFFFFF"> 
    <td width="149" height="106"><a href="gerenc.asp"><img src="imgs/logo.gif" width="149" height="106" border="0"></a></td>
    <td height="106" class="text3m">Nome: <%=Session("User")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IP: <%=Request.ServerVariables("REMOTE_HOST")%><br>
      Logs: <%=Session("Logs")%><br>
      Status: Logado | <a href="login.asp?cod=logout" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Log 
      Off</a></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td valign="top" align="center" class="text2m"> 
      <a href="gerenc.asp?func=artigos" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Artigos</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=banners" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Banners</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=bomburgoes" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Bomburg&atilde;o</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=dicionarios" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Dicionário</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=enquetes" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Enquetes</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=frases" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Frases</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=integrantes" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Integrantes</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=jornais" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Jornais</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=links" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Links</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=menus" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Menus</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=newspro" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Newspro</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=usuarios" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Usuarios</a><br>
      <img src="imgs/int.gif" width="16" height="16"><br>
      <a href="gerenc.asp?func=vacilos" class="text2m" onMouseOver="this.className='text3m';" onMouseOut="this.className='text2m';">Vacilos</a><br>
    </td>
    <td valign="top" class="text3" align="center"> 
      <p>
        <%
Select Case Request.QueryString("func") 
Case "artigos"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblArtigos"
	Rs.Open()
	If Not Rs.EOF Then %>
        Artigos:<br>
        <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Titulo</td>
          <td>Autor</td>
          <td>Data</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=artigos&cod=<%=Rs("CodigoArtigo")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TituloArtigo")%> ::</a></td>
          <td><%=Rs("AutorArtigo")%></td>
          <td><%=Rs("DataArtigo")%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=artigos&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhum artigo encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=artigos&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo artigo
	%>Novo artigo:
      <form method="post" action="salvar.asp?func=artigos&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE" class="text3">
          <tr align="center"> 
          	<td>Titulo</td>
          	<td>Autor</td>
          	<td>Data</td>
          	<td>Texto</td>
          	<td>Preview</td>
          </tr>
          <tr align="center"> 
            <td class="text3"> 
              <input type="text" name="Titulo" maxlength="50" class="textbox1" size="30">
            </td>
            <td class="text3"> 
              <input type="text" name="Autor" maxlength="50" size="20" class="textbox1">
            </td>
            <td class="text3"> 
              <input type="text" name="Data" value="<%=Data()%>" maxlength="28" size="20" class="textbox1">
            </td>
            <td class="text3"> 
              <input type="text" name="Texto" size="20" class="textbox1">
            </td>
			<td>
			  <input type="checkbox" value="1" name="preview" CHECKED>
			</td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br><input type="hidden" name="Codigo" value="0">
      </form>
      <%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblArtigos WHERE CodigoArtigo=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Artigo:<br><br><br>
      <form method="post" action="salvar.asp?func=artigos&cod=<%=Rs("CodigoArtigo")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
          	<td>Titulo</td>
          	<td>Autor</td>
          	<td>Data</td>
          	<td>Texto</td>
            <td>Op&ccedil;&otilde;es</td>
            <td>Preview</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Titulo" value="<%=Rs("TituloArtigo")%>" size="30" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Autor" value="<%=Rs("AutorArtigo")%>" maxlength="50" size="20" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Data" size="20" class="textbox1" value="<%=Rs("DataArtigo")%>">
            </td>
            <td> 
              <input type="text" name="Texto" value="<%=Rs("TextoArtigo")%>" size="20" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delartigo" value="1">
               deletar</td>
			<td>
			  <input type="checkbox" value="1" name="preview" CHECKED>
			</td>
          </tr>
        </table>
        <br><input type="hidden" name="Codigo" value="<%=Rs("CodigoArtigo")%>">
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O artigo não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "banners"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblBanners"
	Rs.Open()
	If Not Rs.EOF Then %>
      Links:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Endereço</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=banners&cod=<%=Rs("CodigoBanner")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("EnderecoBanner")%> ::</a></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=banners&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhum banner encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=banners&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo banner
	%>Novo banner:
      <form method="post" action="salvar.asp?func=banners&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center"> 
            <td class="text4">Endereço</td>
          </tr>
          <tr align="center"> 
            <td class="text3"> 
              <input type="text" name="Endereco" maxlength="50" size="50" class="textbox1">
            </td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblBanners WHERE CodigoBanner=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Link:<br><br><br>
      <form method="post" action="salvar.asp?func=banners&cod=<%=Rs("CodigoBanner")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center" class="text2"> 
            <td>Endereço</td>
            <td width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Endereco" value="<%=Rs("EnderecoBanner")%>" maxlength="50" size="50" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delbanner" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O banner não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "bomburgoes"
Case "dicionarios"
  If Request.QueryString("cod") = "" Then
	'Mostrar todas
	Rs.Source = "SELECT * FROM TblPalavras"
	Rs.Open()
	If Not Rs.EOF Then %>
      Dicionário:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Palavra</td>
          <td>Significado</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=dicionarios&cod=<%=Rs("CodigoPalavra")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TextoPalavra")%> ::</a></td>
          <td width="30%"><%=Rs("SignificadoPalavra")%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=dicionarios&cod=0">
        <input type="submit" value="Nova" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhuma palavra encontrada no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=dicionarios&cod=0">
        <input type="submit" value="Nova" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Nova palavra
	%>Nova palavra:
      <form method="post" action="salvar.asp?func=dicionarios&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center"> 
            <td class="text4">Palavra</td>
            <td class="text4">Significado</td>
          </tr>
          <tr align="center"> 
            <td class="text3"> 
              <input type="text" name="Texto" maxlength="50" class="textbox1" size="30">
            </td>
            <td class="text3"> 
              <input type="text" name="Significado" maxlength="200" size="50" class="textbox1">
            </td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Mostrar/Editar apenas uma
	Rs.Source = "SELECT * FROM TblPalavras WHERE CodigoPalavra=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Dicionário:<br><br><br>
      <form method="post" action="salvar.asp?func=dicionarios&cod=<%=Rs("CodigoPalavra")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
            <td class="text2">Palavra</td>
            <td class="text2">Significado</td>
            <td class="text2" width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Texto" value="<%=Rs("TextoPalavra")%>" size="30" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Significado" value="<%=Rs("SignificadoPalavra")%>" maxlength="200" size="50" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delpalavra" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>A palavra não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "enquetes"
  If Request.QueryString("cod") = "" Then
	'Se não houver enquete selecionada, mostrar todas
	Rs.Source = "SELECT * FROM TblEnquetes"
	Rs.Open()
	If Not Rs.EOF Then
	%>
      Enquetes:<br>
      <br>
      <table width="80%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
        <tr class="text4"> 
          <td align="center">Titulo</td>
          <td align="center">Pergunta</td>
          <td align="center">Status</td>
          <td align="center">Número Votos</td>
        </tr>
        <%
	  Dim Contador
	  Do Until Rs.EOF
		lConsulta.Source = "SELECT TblOpcoes.* FROM TblEnquetes INNER JOIN TblOpcoes ON TblEnquetes.CodigoEnquete = TblOpcoes.EnqueteOpcao WHERE EnqueteOpcao=" & Rs("CodigoEnquete")
		lConsulta.Open()
		Contador = 0
		Do Until lConsulta.EOF
		  Contador = Contador + lConsulta("VotosOpcao")
		  lConsulta.MoveNext
		Loop
	  %>
        <tr bgcolor="#009900" class="text5" align="center"> 
          <td><a href="gerenc.asp?func=enquetes&cod=<%=Rs("CodigoEnquete")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TituloEnquete")%> ::</a></td>
          <td><%=Rs("PerguntaEnquete")%></td>
          <td>
            <% If Rs("StatusEnquete") = 1 Then
								  Response.Write("Em andamento")
								Else
								  Response.Write("Desativada")
								End If %>
          </td>
          <td><%=Contador%></td>
          <td class="text1"></td>
        </tr>
        <%
		If Not lConsulta.BOF Then
		  lConsulta.MoveFirst
		  %>
        <tr>
          <td colspan=4 align="right"> 
            <table border=0 width="80%" cellspacing="1" cellpadding="0" class="text2">
              <%
		  Do Until lConsulta.EOF
			%>
              <tr bgcolor="#FFFFFF"> 
                <td><%=lConsulta("TextoOpcao")%></td>
                <td width="80"><%=lConsulta("VotosOpcao")%> votos</td>
              </tr>
              <%
			lConsulta.MoveNext
		  Loop
		  lConsulta.Close() %>
            </table>
          </td>
        </tr>
        <%
		End If
	    Rs.MoveNext
	  Loop
	  Rs.Close()
	  %>
      </table>
      <br><br>
      <form action="gerenc.asp?func=enquetes&cod=0" method="post">
        <input type="submit" value="Nova" class="botao1"></form>
      <br>
      <br>
      <%
	Else
	  %>
      Nenhuma enquete consta no banco de dados<br><br>
      <form action="gerenc.asp?func=enquetes&cod=0" method="post">
        <input type="submit" value="Nova" class="botao1"></form>
      <br>
      <br>
      <%
	End If
  ElseIf Request.QueryString("cod") = "0" Then
	'Nova Enquete
	%>Nova Enquete:
      <form name="" method="post" action="salvar.asp?func=enquetes&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center" class="text4"> 
            <td>Titulo</td>
            <td>Pergunta</td>
            <td>Status</td>
            <td>Data In&iacute;cio</td>
          </tr>
          <tr align="center" class="text3"> 
            <td> 
              <input type="text" name="Titulo2" maxlength="50" class="textbox1" size="36">
            </td>
            <td> 
              <input type="text" name="Pergunta2" maxlength="150" size="36" class="textbox1">
            </td>
            <td> 
              <input type="checkbox" name="Status2" value="1" checked>
              Ativa</td>
            <td> 
              <input type="text" name="DataInicio2" maxlength="50" size="22" class="textbox1" value="<%=Data()%>">
            </td>
          </tr>
          <tr> 
            <td colspan="4" align="right"> 
              <table width="80%" border="0" cellspacing="1" cellpadding="0" class="text3">
                <%
				For i = 1 To 6 %>
                <tr bgcolor="#FFFFFF"> 
                  <td width="60">Opção <%=i%></td>
                  <td> 
                    <input type="text" name="0<%=i%>" size="70" class="textbox1" maxlength="150">
                  </td>
                  <td width="80"> 
                    <input type="checkbox" name="addopcao<%=i%>" value="1">
                    adicionar </td>
                </tr>
                <%
				Next %>
              </table>
            </td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Alguma enquete foi selecionada
	Rs.Source = "SELECT * FROM TblEnquetes WHERE CodigoEnquete=" & Request.QueryString("cod")
	Rs.Open()
	If Rs.EOF Then
	  %>
      <br><br>A enquete foi removida
      <%
	Else
	  %>
      Enquete -> <%=Rs("TituloEnquete")%>:<br>
      <br>
      <form method="post" action="salvar.asp?func=enquetes&cod=<%=Rs("CodigoEnquete")%>">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr class="text4" align="center"> 
            <td>Titulo</td>
            <td>Pergunta</td>
            <td>Status</td>
            <td>Data In&iacute;cio</td>
            <td>Opções</td>
          </tr>
          <tr align="center" class="text3"> 
            <td> 
              <input type="hidden" name="hCodigoEnquete" value="<%=Rs("CodigoEnquete")%>">
              <input type="text" name="Titulo" maxlength="50" value="<%=Rs("TituloEnquete")%>" class="textbox1" size="36">
            </td>
            <td> 
              <input type="text" name="Pergunta" maxlength="150" value="<%=Rs("PerguntaEnquete")%>" size="36" class="textbox1">
            </td>
            <td> 
              <%  If Rs("StatusEnquete") = 0 Then
					Response.Write("<input type='checkbox' name='Status' value='1' class='textbox1'>")
				  Else
					Response.Write("<input type='checkbox' name='Status' value='1' class='textbox1' checked>")
				  End If
			   %>
            </td>
            <td> 
              <input type="text" name="DataInicio" maxlength="50" value="<%=Rs("DataInicioEnquete")%>" size="22" class="textbox1">
            </td>
            <td width="60"> 
              <input type="checkbox" name="delenquete" value="1">
              deletar</td>
          </tr>
          <tr> 
            <td colspan="5" align="right"> 
              <table width="80%" cellpadding="0" cellspacing="1" class="text3">
                <%
				lConsulta.Source = "SELECT * FROM TblOpcoes WHERE EnqueteOpcao=" & Rs("CodigoEnquete")
				lConsulta.Open()
				If lConsulta.EOF Then
				  Response.Write("<tr><td></td></tr>")
				Else
				  Contador = 0
				  Do Until lConsulta.EOF
					Contador = Contador + 1
				  %>
                <tr bgcolor="#FFFFFF"> 
                  <td width="60"> 
                    <input type="hidden" name="hCodigoOpcao<%=Contador%>" value="<%=lConsulta("CodigoOpcao")%>">
                    Opção <%=Contador%> </td>
                  <td> 
                    <input type="text" name="Opcao<%=Contador%>" maxlength="150" size="75" class="textbox1" value="<%=lConsulta("TextoOpcao")%>">
                  </td>
                  <td width="60"> 
                    <input type="checkbox" name="delopcao<%=Contador%>" value="1">
                    deletar</td>
                </tr>
                <%
					lConsulta.MoveNext
				  Loop
				  lConsulta.Close() %>
                <tr>
                  <td>Opção <%=Contador+1%></td>
                  <td>
                    <input type="text" name="0" maxlenght="150" size="75" class="textbox1" maxlength="150">
                  </td>
                  <td>
                    <input type="checkbox" name="addopcao" value="1">
                    adicionar</td>
                </tr>
                <%
				End If
				%>
              </table>
            </td>
          </tr>
        </table>
        <br>
        <input type="hidden" name="nOpcoes" value="<%=Contador%>">
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1" name="Reset">
        <br><br>
        Obs: Ao excluir a enquete, o us&aacute;rio estar&aacute; excluindo todas 
        as op&ccedil;&otilde;es automaticamente<br><br>
      </form>
      <br>
      <%
	End If
	Rs.Close()
  End If
Case "frases"
  If Request.QueryString("cod") = "" Then
	'Mostrar todas
	Rs.Source = "SELECT * FROM TblFrases"
	Rs.Open()
	If Not Rs.EOF Then %>
      Frases:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Frase</td>
          <td>Autor</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=frases&cod=<%=Rs("CodigoFrase")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TextoFrase")%> ::</a></td>
          <td width="30%"><%=Rs("AutorFrase")%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table>
	<%
	End If
	Rs.Close()
  Else
	'Mostrar/Editar apenas uma
	Rs.Source = "SELECT * FROM TblFrases WHERE CodigoFrase=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Frase:<br><br><br>
      <form method="post" action="salvar.asp?func=frases&cod=<%=Rs("CodigoFrase")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center" class="text2"> 
            <td>Frase</td>
            <td>Autor</td>
            <td width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Frase" value="<%=Rs("TextoFrase")%>" size="60" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Autor" value="<%=Rs("AutorFrase")%>" maxlength="50" size="35" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delfrase" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>A frase não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "integrantes"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblIntegrantes"
	Rs.Open()
	If Not Rs.EOF Then %>
      Integrantes:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Nome</td>
          <td>Apelido</td>
		  <td>ICQ</td>
		  <td>E-Mail</td>
		  <td>Sala</td>
		  <td>Nascimento</td>
		  <td>Descrição</td>
		  <td>Cerimonial</td>
		  <td>Foto</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=integrantes&cod=<%=Rs("CodigoIntegrante")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("NomeIntegrante")%> ::</a></td>
          <td><%=Rs("ApelidoIntegrante")%></td>
          <td><%=Rs("ICQIntegrante")%></td>
          <td><%=Rs("EmailIntegrante")%></td>
          <td><%=Rs("SalaIntegrante")%></td>
          <td><%=Rs("NascimentoIntegrante")%></td>
          <td><%=Rs("DescricaoIntegrante")%></td>
          <td><%=Rs("CerimonialIntegrante")%></td>
          <td><% If Rs("FotoIntegrante") <> "" Then
				   Response.Write("<img src='" & Rs("FotoIntegrante") & "'>")
				 Else
				   Response.Write("<img src='imgs\xis.gif' width='20' height='20'>")
				 End If %></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=integrantes&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhum integrante encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=integrantes&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo integrante
	%>Novo integrante:
      <form method="post" action="salvar.asp?func=integrantes&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center" class="text2"> 
          	<td>Nome</td>
          	<td>Apelido</td>
		  	<td>ICQ</td>
		  	<td>E-Mail</td>
		  	<td>Sala</td>
		  	<td>Nascimento</td>
		  	<td>Descrição</td>
		  	<td>Cerimonial</td>
		  	<td>Foto</td>
          </tr>
          <tr align="center" class="text3"> 
            <td> 
              <input type="text" name="Nome" maxlength="100" class="textbox1" size="18">
            </td>
            <td> 
              <input type="text" name="Apelido" maxlength="50" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="ICQ" maxlength="12" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Email" maxlength="50" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Sala" maxlength="50" size="15" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Nascimento" maxlength="26" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Descricao" size="15" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Cerimonial" maxlength="100" size="15" class="textbox1">
            </td>
            <td > 
              <input type="text" name="Foto" maxlength="50" size="10" class="textbox1">
            </td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblIntegrantes WHERE CodigoIntegrante=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Integrante:<br><br><br>
      <form method="post" action="salvar.asp?func=integrantes&cod=<%=Rs("CodigoIntegrante")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
          	<td>Nome</td>
          	<td>Apelido</td>
		  	<td>ICQ</td>
		  	<td>E-Mail</td>
		  	<td>Sala</td>
		  	<td>Nascimento</td>
		  	<td>Descrição</td>
		  	<td>Cerimonial</td>
		  	<td>Foto</td>
            <td class="text2" width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Nome" value="<%=Rs("NomeIntegrante")%>" size="18" class="textbox1" maxlength="100">
            </td>
            <td> 
              <input type="text" name="Apelido" value="<%=Rs("ApelidoIntegrante")%>" maxlength="50" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="ICQ" value="<%=Rs("ICQIntegrante")%>" maxlength="12" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Email" value="<%=Rs("EmailIntegrante")%>" maxlength="50" size="10" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Sala" value="<%=Rs("SalaIntegrante")%>" maxlength="50" size="15" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Nascimento" value="<%=Rs("NascimentoIntegrante")%>" maxlength="26" size="12" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Descricao" value="<%=Rs("DescricaoIntegrante")%>" size="15" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Cerimonial" value="<%=Rs("CerimonialIntegrante")%>" maxlength="100" size="15" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Foto" value="<%=Rs("FotoIntegrante")%>" maxlength="50" size="10" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delintegrante" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O integrante não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "jornais"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblJornais"
	Rs.Open()
	If Not Rs.EOF Then %>
      Jornais:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Edição</td>
          <td>Autor</td>
		  <td>Reportagens</td>
        </tr>
        <%
	  Do Until Rs.EOF
		Contador = 1
		cReportagens = 0
		For Contador= 1 to 8
		  If Not Trim(Rs("N" & Contador & "Jornal")) = "" Then
			cReportagens = cReportagens + 1
		  End If
		Next
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=jornais&cod=<%=Rs("CodigoJornal")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("EdicaoJornal")%> ::</a></td>
          <td><%=Rs("AutorJornal")%></td>
		  <td><%=cReportagens%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=jornais&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhum jornal encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=jornais&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo jornal
	%>Novo jornal:
      <form method="post" action="salvar.asp?func=jornais&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center" class="text3"> 
          	<td>Edição</td>
          	<td>Autor</td>
		  	<td>Preview</td>
          </tr>
          <tr align="center"> 
            <td class="text3"> 
              <input type="text" name="Edicao" maxlength="50" class="textbox1" size="30">
            </td>
            <td class="text3"> 
              <input type="text" name="Autor" maxlength="50" size="30" class="textbox1">
            </td>
			<td>
			  <input type="checkbox" value="1" name="preview" CHECKED>
			</td>
          </tr>
		  <tr>
			<td colspan="3" align="right">
              <table width="80%" border="0" cellspacing="1" cellpadding="0" class="text3">
                <%
				For Contador = 1 To 8
				%> 
                <tr> 
                  <td bgcolor="#FFFFFF" width="100">Reportagem <%=Contador%></td>
                  <td bgcolor="#FFFFFF"> 
                    <input type="text" name="N<%=Contador%>" class="textbox1" maxlength="200" size="100">
                  </td>
                </tr>
				<%
				Next
				%>
              </table>
            </td>
		  </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblJornais WHERE CodigoJornal=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Jornal:<br><br><br>
      <form method="post" action="salvar.asp?func=jornais&cod=<%=Rs("CodigoJornal")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center" class="text2"> 
            <td>Edição</td>
            <td>Endereço</td>
            <td width="80">Op&ccedil;&otilde;es</td>
            <td>Preview</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Edicao" value="<%=Rs("EdicaoJornal")%>" size="30" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Autor" value="<%=Rs("AutorJornal")%>" maxlength="50" size="30" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="deljornal" value="1">
               deletar</td>
            <td>
              <input type="checkbox" CHECKED name="preview" value="1">
			</td>
		  </tr>
		  <tr>
			<td colspan="4" align="right">
              <table width="80%" border="0" cellspacing="1" cellpadding="0" class="text3">
                <%
				For Contador = 1 To 8
				%> 
                <tr> 
                  <td bgcolor="#FFFFFF" width="100">Reportagem <%=Contador%></td>
                  <td bgcolor="#FFFFFF"> 
                    <input type="text" name="N<%=Contador%>" value="<%=Rs("N" & Contador & "Jornal")%>" class="textbox1" maxlength="200" size="100">
                  </td>
                </tr>
				<%
				Next
				%>
              </table>
            </td>
		  </tr>
        </table>
        <input type="hidden" name="Codigo" value="<%=Rs("CodigoJornal")%>">
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O jornal não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "links"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblLinks"
	Rs.Open()
	If Not Rs.EOF Then %>
      Links:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Nome</td>
          <td>Endereço</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=links&cod=<%=Rs("CodigoLink")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("NomeLink")%> ::</a></td>
          <td width="30%"><%=Rs("EnderecoLink")%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=links&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	<%
	Else
	  %>Nenhum link encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=links&cod=0">
        <input type="submit" value="Novo" class="botao1">
      </form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo link
	%>Novo Link:
      <form method="post" action="salvar.asp?func=links&cod=0">
        <table width="90%" border="1" cellspacing="2" cellpadding="0" bgcolor="#DEDEDE">
          <tr align="center"> 
            <td class="text4">Nome</td>
            <td class="text4">Endereco</td>
          </tr>
          <tr align="center"> 
            <td class="text3"> 
              <input type="text" name="Nome" maxlength="50" class="textbox1" size="36">
            </td>
            <td class="text3"> 
              <input type="text" name="Endereco" maxlength="150" size="36" class="textbox1">
            </td>
          </tr>
        </table><br><br>
        <input type="submit" value="Salvar" class="botao1">
        <br>
      </form>
      <%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblLinks WHERE CodigoLink=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Link:<br><br><br>
      <form method="post" action="salvar.asp?func=links&cod=<%=Rs("CodigoLink")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
            <td class="text2">Nome</td>
            <td class="text2">Endereço</td>
            <td class="text2" width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Nome" value="<%=Rs("NomeLink")%>" size="60" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Endereco" value="<%=Rs("EnderecoLink")%>" maxlength="50" size="35" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="dellink" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O link não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "menus"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblMenus ORDER BY NomeMenu ASC"
	Rs.Open()
	If Not Rs.EOF Then %>
      Menus:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Nome</td>
          <td>Endereço</td>
		  <td>Status</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=menus&cod=<%=Rs("CodigoMenu")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("NomeMenu")%> ::</a></td>
          <td width="30%"><%=Rs("EnderecoMenu")%></td>
		  <td><%
		If Rs("StatusMenu") = 0 Then
		  Response.Write("Desativado")
		Else
		  Response.Write("Ativado")
		End If
		%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table><br><br><form method="POST" action="gerenc.asp?func=menus&cod=0"><input type="submit" class="botao1" value="Novo"></form><br><br>
	<%
	Else
	  %>Nenhum menu encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=menus&cod=0"><input type="submit" class="botao1" value="Novo"></form><br><br>
	  <%
	End If
	Rs.Close()
  ElseIf Request.QueryString("cod") = "0" Then %>Novo Menu:<br><br><form method="POST" action="salvar.asp?func=menus&cod=0">
		<table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
            <td class="text2">Nome</td>
            <td class="text2">Endereço</td>
			<td class="text2">Status</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Nome" size="50" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Endereco" maxlength="50" size="35" class="textbox1">
            </td>
			<td><input type='checkbox' name='Status' value='1'>
            </td>
          </tr>
        </table><br><br><input type="submit" value="Salvar" class="botao1"></form><%
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblMenus WHERE CodigoMenu=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Menu:<br><br><br>
      <form method="post" action="salvar.asp?func=menus&cod=<%=Rs("CodigoMenu")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
            <td class="text2">Nome</td>
            <td class="text2">Endereço</td>
			<td class="text2">Status</td>
            <td class="text2" width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Nome" value="<%=Rs("NomeMenu")%>" size="50" class="textbox1" maxlength="50">
            </td>
            <td> 
              <input type="text" name="Endereco" value="<%=Rs("EnderecoMenu")%>" maxlength="50" size="35" class="textbox1">
            </td>
			<td><% If Rs("StatusMenu") = 1 Then 
			          Response.Write("<input type='checkbox' name='Status' value='1' checked>")
					Else
					  Response.Write("<input type='checkbox' name='Status' value='1'>")
					End If %>
            </td>
            <td>
              <input type="checkbox" name="delmenu" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form><br><br>
      <%
	Else
	  %><br><br>O menu não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case "newspro"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblNews"
	Rs.Open()
	If Not Rs.EOF Then
	  %>
	  News:<br><br>
	  <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Titulo</td>
          <td>Autor</td>
		  <td>Avatar</td>
		  <td>Data</td>
		  <td>Conteúdo</td>
        </tr><% 
		Do Until Rs.EOF
		  %>
		<tr>
            <td><a href="gerenc.asp?func=newspro&cod=<%=Rs("CodigoNews")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TituloNews")%> ::</a></td>
		    <td><%=Rs("AutorNews")%></td>
            <td><% If Rs("AvatarNews") <> "" Then
				   Response.Write("<img src='" & Rs("AvatarNews") & "'>")
				 Else
				   Response.Write("<img src='imgs\xis.gif' width='20' height='20'>")
				 End If %></td>
			<td><%=Rs("DataNews")%></td>
			<td><%=Rs("TextoNews")%></td>
		</tr>
		  <%
		  Rs.MoveNext
		Loop
	  %></table><br><br><form method="POST" action="gerenc.asp?func=newspro&cod=0"><input type="submit" class="botao1" value="Novo"></form><br><br>
	  <%
	Else
	  'Se não há nenhum news
	  %>Nenhum news encontrado no banco de dados
	  <br><br><form method="POST" action="gerenc.asp?func=newspro&cod=0"><input type="submit" class="botao1" value="Novo"></form><br><br>
	  <%
	End If
  ElseIf Request.QueryString("cod") = "0" Then
	'Novo news
	%>
	  Novo news:<br><br><br><form method="POST" action="salvar.asp?func=newspro&cod=0">
	  <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Titulo</td>
          <td>Autor</td>
		  <td>Avatar</td>
		  <td>Data</td>
		  <td>Conteúdo</td>
		  <td>Preview</td>
        </tr> 
		<tr align="center">
            <td><input type="text" name="Titulo" size="20" class="textbox1" maxlength="50"></td>
		    <td><input type="text" name="Autor" value="<%=Session("user")%>" size="15" class="textbox1" maxlength="50"></td>
            <td><input type="text" name="Avatar" size="20" class="textbox1" maxlength="50"></td>
			<td><input type="text" name="Data" value="<%=Data()%>" size="18" class="textbox1" maxlength="50"></td>
			<td><input type="text" name="Texto" size="50" class="textbox1"></td>
			<td><input type="checkbox" value="1" name="preview" CHECKED></td>
		</tr>
	  </table><br><br><input type="hidden" name="Codigo" value="0"><input type="submit" value="Salvar" class="botao1"></form><br><br>
	<%
  Else
	'Mostrar/Editar apenas um
	If Not Request.QueryString("cod") = "800" Then
	  'Mostrar/Editar
	  Rs.Source = "SELECT * FROM TblNews WHERE CodigoNews=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
	  %>
	  News:<br><br><br><form method="POST" action="salvar.asp?func=newspro&cod=<%=Rs("CodigoNews")%>">
	  <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Titulo</td>
          <td>Autor</td>
		  <td>Avatar</td>
		  <td>Data</td>
		  <td>Conteúdo</td>
		  <td>Opções</td>
		  <td>Preview</td>
        </tr> 
		<tr align="center">
            <td><input type="text" value="<%=Rs("TituloNews")%>" name="Titulo" size="20" class="textbox1" maxlength="50"></td>
		    <td><input type="text" value="<%=Rs("AutorNews")%>" name="Autor" size="15" class="textbox1" maxlength="50"></td>
            <td><input type="text" value="<%=Rs("AvatarNews")%>" name="Avatar" size="20" class="textbox1" maxlength="50"></td>
			<td><input type="text" value="<%=Rs("DataNews")%>" name="Data" size="18" class="textbox1" maxlength="50"></td>
			<td><input type="text" value="<%=Rs("TextoNews")%>" name="Texto" size="50" class="textbox1"></td>
			<td><input type="checkbox" value="1" name="delnews"></td>
			<td><input type="checkbox" value="1" name="preview" CHECKED></td>
		</tr>
	  </table><br><br><input type="hidden" name="Codigo" value="<%=Rs("CodigoNews")%>"><input type="submit" value="Salvar" class="botao1"></form><br><br>
	  <%
	  End If
	  Rs.Close()
	End If
  End If
Case "usuarios"
  'Validar usuário-> Tem q ser o boi
  If lCase(Session("User")) = "boi" Then
	If Request.QueryString("cod") = "" Then
	  'Mostrar todos
	  Rs.Source = "SELECT * FROM TblUsuarios ORDER BY NomeUsuario ASC"
	  Rs.Open()
	  If Not Rs.EOF Then %>
	    Usuários:<br><br>
        <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
          <tr align="center" class="text2"> 
            <td>Usuário</td>
		    <td>Senha</td>
            <td>Titulo</td>
			<td>Logs</td>
          </tr><%
		  Do Until Rs.EOF  %>
          <tr align="center">
            <td><a href="gerenc.asp?func=usuarios&cod=<%=Rs("CodigoUsuario")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("NomeUsuario")%> ::</a></td>
		    <td><%=Rs("SenhaUsuario")%></td>
            <td><%=Rs("TituloUsuario")%></td>
			<td><%=Rs("LogsUsuario")%></td>
          </tr>
	    <%
		    Rs.MoveNext
		  Loop
	    %></table><br><br><form method="POST" action="gerenc.asp?func=usuarios&cod=0"><input type="submit" value="Novo" class="botao1"></form><br><br><%
	  Else
		'Se não há nenhum usuário (impossível), colocar botão para adicionar
		%>
		Nenhum usuário encontrado no banco de dados<br><br><form method="POST" action="gerenc.asp?func=usuarios&cod=0"><input type="submit" value="Novo" class="botao1"></form><br><br>
		<%
	  End If
	  Rs.Close()
	ElseIf Request.QueryString("cod") = "0" Then
	  'Novo usuário
		%>
		Novo usuário:<br><br><br>
        <form method="post" action="salvar.asp?func=usuarios&cod=0">
          <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
            <tr align="center"> 
              <td class="text2">Nome</td>
              <td class="text2">Senha</td>
			  <td class="text2">Titulo</td>
			  <td class="text2">Biografia</td>
			  <td class="text2">Logs</td>
            </tr>
            <tr align="center"> 
              <td> 
                <input type="text" name="Nome" maxlength="150" size="60" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Senha" maxlength="10" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Titulo" maxlength="50" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Biografia" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Logs" size="3" value="0" class="textbox1">
              </td>
            </tr>
          </table>
          <br>
          <input type="submit" value="Salvar" class="botao1"><br>
        </form><%
	Else
	  'Mostrar/Editar apenas um
	  Rs.Source = "SELECT * FROM TblUsuarios WHERE CodigoUsuario=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
	    %>Usuário:<br><br><br>
        <form method="post" action="salvar.asp?func=usuarios&cod=<%=Rs("CodigoUsuario")%>">
          <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
            <tr align="center"> 
              <td class="text2">Nome</td>
              <td class="text2">Senha</td>
			  <td class="text2">Titulo</td>
			  <td class="text2">Biografia</td>
			  <td class="text2">Logs</td>
              <td class="text2" width="80">Op&ccedil;&otilde;es</td>
            </tr>
            <tr align="center"> 
              <td> 
                <input type="text" name="Nome" value="<%=Rs("NomeUsuario")%>" maxlength="150" size="60" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Senha" value="<%=Rs("SenhaUsuario")%>" maxlength="10" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Titulo" value="<%=Rs("TituloUsuario")%>" maxlength="50" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Biografia" value="<%=Rs("BiografiaUsuario")%>" size="10" class="textbox1">
              </td>
              <td> 
                <input type="text" name="Logs" value="<%=Rs("LogsUsuario")%>" size="3" class="textbox1">
              </td>
              <td>
                <input type="checkbox" name="delusuario" value="1">
                 deletar</td>
            </tr>
          </table>
          <br>
          <input type="submit" value="Salvar" class="botao1">
          <input type="reset" value="Limpar" class="botao1">
          <br>
        </form>
        <%
	  Else
	    %><br><br>O usuário não existe mais<br><br><%
	  End If
	  Rs.Close()
	End If
  Else
	%>Não é permitido a visualização deste item, uma vez que a senha geral não foi utilizada<br><br>
	<%
  End If
Case "vacilos"
  If Request.QueryString("cod") = "" Then
	'Mostrar todos
	Rs.Source = "SELECT * FROM TblVacilos"
	Rs.Open()
	If Not Rs.EOF Then %>
      Vacilos:<br>
      <br>
      <table width="90%" cellpadding="0" cellspacing="2" bgcolor="#DEDEDE" class="text4" border="1">
        <tr align="center" class="text2"> 
          <td>Vacilo</td>
          <td>Autor</td>
        </tr>
        <%
	  Do Until Rs.EOF
	%>
        <tr align="center">
          <td><a href="gerenc.asp?func=vacilos&cod=<%=Rs("CodigoVacilo")%>" class="text4" onMouseOver="this.className='text3';" onMouseOut="this.className='text4';">:: <%=Rs("TextoVacilo")%> ::</a></td>
          <td width="30%"><%=Rs("AutorVacilo")%></td>
        </tr>
        <%
		Rs.MoveNext
	  Loop %>
      </table>
	<%
	End If
	Rs.Close()
  Else
	'Mostrar/Editar apenas um
	Rs.Source = "SELECT * FROM TblVacilos WHERE CodigoVacilo=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  %>Vacilo:<br><br><br>
      <form method="post" action="salvar.asp?func=vacilos&cod=<%=Rs("CodigoVacilo")%>">
        <table cellspacing="2" cellpadding="0" border="1" bgcolor="#DEDEDE" width="90%" class="text3">
          <tr align="center"> 
            <td class="text2">Vacilo</td>
            <td class="text2">Autor</td>
            <td class="text2" width="80">Op&ccedil;&otilde;es</td>
          </tr>
          <tr align="center"> 
            <td> 
              <input type="text" name="Vacilo" value="<%=Rs("TextoVacilo")%>" size="60" class="textbox1">
            </td>
            <td> 
              <input type="text" name="Autor" value="<%=Rs("AutorVacilo")%>" maxlength="50" size="35" class="textbox1">
            </td>
            <td>
              <input type="checkbox" name="delvacilo" value="1">
               deletar</td>
          </tr>
        </table>
        <br>
        <input type="submit" value="Salvar" class="botao1">
        <input type="reset" value="Limpar" class="botao1">
        <br>
      </form>
      <%
	Else
	  %><br><br>O vacilo não existe mais<br><br><%
	End If
	Rs.Close()
  End If
Case Else
  %>
        <br>
        <p align="center"><font class=text2>Bem vindo Webmaster SalaDois</font></p>
      <p align="justify">Para facillitar a autaliza&ccedil;&atilde;o constante 
        do site, o sistema foi trocado, e agora, est&aacute; escrito em <font class=text2>ASP</font>, 
        quer dizer, voc&ecirc; n&atilde;o precisa saber linguagem de programa&ccedil;&atilde;o 
        pra poder atualizar o site. Vale ressaltar que em alguns casos, &eacute; 
        necess&aacute;rio um m&iacute;nimo de conhecimento sobre a linguagem HTML.</p>
      <p align="justify">A linguagem <font class=text2>HTML</font> &eacute; composta 
        por <font class=text2>>tags</font>, que s&atilde;o comandos expressos 
        entre sinais de maior e menor, como por exemplo, o c&oacute;digo &lt;hr&gt; 
        ou o c&oacute;digo &lt;br&gt;. A maioria das tags devem ser abertas e 
        depois fechadas, para determinar quando seu efeito come&ccedil;a e quando 
        ele termina. Um bom exemplo &eacute; o &lt;b&gt;, logo abaixo. Algumas 
        tags necessitam par&acirc;metros, como a tag &lt;font color='cor_aqui'&gt;, 
        usada para trocar a cor do texto. O(s) par&acirc;metros deve(m) estar 
        contido(s) dentro de seus (respectivos) ap&oacute;strofos: &lt;font color='blue'size='1'&gt;&lt;/font&gt;.</p>
      <p align="justify">Aqui seguem os c&oacute;digos mais importantes:</p>
      <p align="justify"><font class=text2>&lt;br&gt;</font> - quebra de linha. Exemplo:</p>
      <p align="center"><i>C&oacute;digo</i>: O Demon Hunter &eacute; um cara<font class=text2>&lt;br&gt;</font>legal<br>
        <i>Resultado</i>: O Demon Hunter &eacute; um cara<br>
        legal</p>
      <p align="justify"><br>
        <font class=text2>&lt;hr&gt;</font> - linha horizontal<br>
        <font class=text2>&lt;b&gt;...&lt;/b&gt;</font> - negrito (necessita fechamento). Exemplo:</p>
      <p align="center"><i>C&oacute;digo</i>: O <font class=text2>&lt;b&gt;</font>cavalinho<font class=text2>&lt;/b&gt;</font> 
        n&uacute;mero quatro &eacute; o mais <font class=text2>&lt;b&gt;</font>lesado<font class=text2>&lt;/b&gt;</font><br>
        <i>Resultado:</i> O <font color="#000000">cavalinho</font> n&uacute;mero quatro &eacute; o mais 
        <font color="#000000">lesado</font></p>
      <p align="justify"><font class=text2>&lt;i&gt;...&lt;/i&gt;</font> - it&aacute;lico 
        (necessita fechamento)<br>
        <font class=text2>&lt;center&gt;...&lt;/center&gt;</font> - centralizar (necessita fechamento).<br>
        <font class=text2>&lt;font color='cor_em_ingles' face='fonte_do_texto' 
        size='tamanho'&gt;...&lt;/font&gt;</font> - Muda cor, fonte e tamanho 
        do texto (necessita fechamento e par&acirc;metro). N&atilde;o &eacute; 
        obrigat&oacute;rio determinar todos os par&acirc;metros. Se quiser mudar 
        s&oacute; a cor, coloque &lt;font color=&quot;cor_em_ingles&quot;&gt;...&lt;/font&gt;. 
        Exemplo:</p>
      <p align="center"><i>C&oacute;digo</i>: O <font class=text2>&lt;font color='red'&gt;</font>Prego<font class=text2>&lt;/font&gt;</font> 
        &eacute; meu <font class=text2>&lt;br&gt;</font>amigo<br>
        <i>Resultado</i>: O <font color=red>Prego</font> é meu<br>
        amigo</p>
      <p align="center">&nbsp;</p>
      <p align="center">Nunca se esque&ccedil;a de fechar uma tag que precise 
        ser fechada!!</p>
      <p align="justify">&nbsp;</p>
      <p align="justify">Obs: N&atilde;o use, em hip&oacute;tese alguma, &aacute;spas(&quot;) 
        em seu texto. Se o fizer, h&aacute; uma grande chance de haver um curto-circuito 
        deixando seu computador em risco de explos&atilde;o. Se, por algum acaso 
        voc&ecirc; se esquecer, colocar &aacute;spas em algum texto e salvar, 
        desligue o computador imediatamente antes que ocorra uma explos&atilde;o 
        :D</p>
      <p align="justify">Obs(2):Quando for especificar cor, em &lt;font color='cor_aqui'&gt; 
        voc&ecirc; pode coloc&aacute;-la em hexadecimal, como por exemplo o preto: 
        black ou #000000.</p>
      <p align="left"><font class=text2>Pad&otilde;res:</font></p>
      <p align="left">Data: <%=Data()%> (dia_do_mes + " de " + mês_com_primeira_maiúscula + " de 
        " + ano_em_4_digitos)<br>
        Cores: Cinza (#333333), <font color="green">Verde</font> (green), <font color="#CCCCCC">Branco</font> 
        (white ou #FFFFFF), <font color="#000000">Preto</font> (black ou #000000)<br>
        <br>
      </p>
      <br><br>
      <%
End Select
%>
    </td>
  </tr>
</table>
</body>
</html>
