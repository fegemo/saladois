<%@LANGUAGE="VBSCRIPT"%> 
<%
If Session("Logged")="" Then
  Response.Redirect"login.htm"
End If
%>
<!--#include file="Connections/Conneccao.asp" -->
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Rs.LockType = 3
Set lConsulta = Server.CreateObject("ADODB.Recordset")
lConsulta.ActiveConnection = MM_Conneccao_STRING
lConsulta.LockType = 3

Select Case Request.QueryString("func")
Case ""
  Response.Redirect "gerenc.asp"
Case "artigos"
  If Request.Form("preview") = "1" Then
	'Preview
	Response.Redirect "artigopreview.asp?codigo=" & Request.Form("Codigo") & "&titulo=" & Request.Form("Titulo") & "&autor=" & Request.Form("Autor") & "&data=" & Request.Form("Data") & "&texto=" & Request.Form("Texto") & "&delartigo=" & Request.Form("delartigo")
  Else
  	If Not Request.QueryString("cod") = "0" Then
	  'Se o artigo já existir, editá-lo...
	  Rs.Source = "SELECT * FROM TblArtigos WHERE CodigoArtigo=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
	  	'Editar...
	  	Rs("TituloArtigo") = Request.Form("Titulo")
	  	Rs("AutorArtigo") = Request.Form("Autor")
	  	Rs("DataArtigo") = Request.Form("Data")
	  	Rs("TextoArtigo") = Request.Form("Texto")
	  	Rs.Update
	  Else
	  	'Se o artigo não existir, criá-lo
	  	Rs.AddNew()
	  	Rs("TituloArtigo") = Request.Form("Titulo")
	  	Rs("AutorArtigo") = Request.Form("Autor")
	  	Rs("DataArtigo") = Request.Form("Data")
	  	Rs("TextoArtigo") = Request.Form("Texto")
	  	Rs.Update
	  End If
	  'Verificar se é para ser deletado
	  If Request.Form("delartigo") = "1" Then
	  	Rs.Delete
	  End If
		Rs.Close()
  	Else
	  'Se for um novo artigo
	  Rs.Source = "SELECT * FROM TblArtigos"
	  Rs.Open()
	  Rs.AddNew
	  Rs("TituloArtigo") = Request.Form("Titulo")
	  Rs("AutorArtigo") = Request.Form("Autor")
	  Rs("DataArtigo") = Request.Form("Data")
	  Rs("TextoArtigo") = Request.Form("Texto")
	  Rs.Update
	  Rs.Close
  	End If
  End If
  Response.Redirect "gerenc.asp?func=artigos"
Case "banners"
  If Not Request.QueryString("cod") = "0" Then
	'Se o banner já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblBanners WHERE CodigoBanner=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("EnderecoBanner") = Request.Form("Endereco")
	  Rs.Update
	Else
	  'Se o banner não existir, criá-lo
	  Rs.AddNew()
	  Rs("EnderecoBanner") = Request.Form("Endereco")
	  Rs.Update
	End If
	'Verificar se é para ser deletado
	If Request.Form("delbanner") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  Else
	'Se for um novo banner
	Rs.Source = "SELECT * FROM TblBanners"
	Rs.Open()
	Rs.AddNew
    Rs("EnderecoBanner") = Request.Form("Endereco")
	Rs.Update
	Rs.Close
  End If
  Response.Redirect "gerenc.asp?func=banners"
Case "bomburgoes"
Case "dicionarios"
  If Not Request.QueryString("cod") = "0" Then
	'Se a palavra já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblPalavras WHERE CodigoPalavra=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("TextoPalavra") = Request.Form("Texto")
	  Rs("SignificadoPalavra") = Request.Form("Significado")
	  Rs.Update
	Else
	  'Se a palavra não existir, criá-lo
	  Rs.AddNew()
	  Rs("TextoPalavra") = Request.Form("Texto")
	  Rs("SignificadoPalavra") = Request.Form("Significado")
	  Rs.Update
	End If
	'Verificar se é para ser deletada
	If Request.Form("delpalavra") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  Else
	'Se for uma nova palavra
	Rs.Source = "SELECT * FROM TblPalavras"
	Rs.Open()
	Rs.AddNew
	Rs("TextoPalavra") = Request.Form("Texto")
	Rs("SignificadoPalavra") = Request.Form("Significado")
	Rs.Update
	Rs.Close
  End If
  Response.Redirect "gerenc.asp?func=dicionarios"
Case "enquetes"
  If Not Request.QueryString("cod") = 0 Then
	'Enquete já existente (edição)
	Rs.Source = "SELECT * FROM TblEnquetes WHERE CodigoEnquete=" & Request.QueryString("cod")
	Rs.Open()
	If Rs.EOF Then
	  'Se a enquete não existe...
	  Rs.Close()
	  Response.Redirect "gerenc.asp?func=enquetes"
	Else
	  'Se a enquete existe, atualizá-la:
	  Rs("TituloEnquete") = Request.Form("Titulo")
	  Rs("PerguntaEnquete") = Request.Form("Pergunta")
	  Rs("StatusEnquete") = Request.Form("Status")
	  If Request.Form("Status") = "" Then Rs("StatusEnquete") = 0
	  Rs("DataInicioEnquete") = Request.Form("DataInicio")
	  Rs.Update
	  lConsulta.Source = "SELECT * FROM TblOpcoes WHERE EnqueteOpcao=" & Rs("CodigoEnquete")
	  lConsulta.Open()
	  If Not lConsulta.EOF Then
		'Se a enquete tem no mínimo 1 opção, atualizá-la(s):
		'Encontrar os registros das opções para atualizá-los, um por um...
		Contador = 0
		Do Until lConsulta.EOF
		  Contador = Contador + 1
		  lConsulta("TextoOpcao") = Request.Form("Opcao" & Contador)
		  lConsulta.Update
		  lConsulta.MoveNext
		Loop
		'Verificar se tem novas opções...
		If Request.Form("addopcao") = 1 Then
		  'Adicionar nova opção
		  lConsulta.AddNew
		  lConsulta("EnqueteOpcao") = Request.Form("hCodigoEnquete")
		  lConsulta("TextoOpcao") = Request.Form("0")
		  lConsulta.Update
		End If
		lConsulta.Close()
	    'Verificar se alguma deleção foi solicitada
	    For Contador = 1 To Request.Form("nOpcoes")
		  If Request.Form("delopcao" & Contador) = 1 Then
		    'Se a opção é para ser deletada...
		    lConsulta.Source = "SELECT * FROM TblOpcoes WHERE EnqueteOpcao=" & Rs("CodigoEnquete") & " AND CodigoOpcao=" & Request.Form("hCodigoOpcao" & Contador)
		    lConsulta.Open()
		    lConsulta.Delete
		    lConsulta.Close()
		  End If
	    Next
	  End If

	  'Verificar se a deleção da enquete foi solicitada
	  If Request.Form("delenquete") = 1 Then
		Rs.Delete
	  End If
	  Rs.Close()
	End If
  Else
	'Nova enquete
	Rs.Source = "SELECT * FROM TblEnquetes"
	Rs.Open()
	Rs.AddNew
	Rs("TituloEnquete") = Request.Form("Titulo2")
	Rs("PerguntaEnquete") = Request.Form("Pergunta2")
	Rs("StatusEnquete") = Request.Form("Status2")
	If Request.Form("Status2") = "" Then Rs("StatusEnquete") = 0
	Rs("DataInicioEnquete") = Request.Form("DataInicio2")
	Rs.Update
	Rs.Close()
	Rs.Source = "SELECT * FROM TblEnquetes ORDER BY CodigoEnquete DESC"
	Rs.Open()
	lConsulta.Source = "SELECT * FROM TblOpcoes"
	lConsulta.Open()
	For i = 1 to 6
	  If Request.Form("addopcao" & i) = 1 Then
		lConsulta.AddNew
		lConsulta("EnqueteOpcao") = Rs("CodigoEnquete")
		lConsulta("TextoOpcao") = Request.Form("0" & i)
		lConsulta.Update
	  End If
	Next
	lConsulta.Close()
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=enquetes"
Case "frases"
  If Not Request.QueryString("cod") = "0" Then
	'Se a frase já existir, editá-la...
	Rs.Source = "SELECT * FROM TblFrases WHERE CodigoFrase=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("TextoFrase") = Request.Form("Frase")
	  Rs("AutorFrase") = Request.Form("Autor")
	  Rs.Update
	Else
	  'Se a frase não existir, criá-la
	  Rs.AddNew()
	  Rs("TextoFrase") = Request.Form("Frase")
	  Rs("AutorFrase") = Request.Form("Autor")
	  Rs.Update
	End If
	'Verificar se é para ser deletada
	If Request.Form("delfrase") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=frases"
Case "integrantes"
  If Not Request.QueryString("cod") = "0" Then
	'Se o integrante já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblIntegrantes WHERE CodigoIntegrante=" & Request.QueryString("cod")
	Rs.Open()
	If Rs.EOF Then
	  'Se o integrante não existir, criá-lo
	  Rs.AddNew()
	End If
	Rs("NomeIntegrante") = Request.Form("Nome")
	Rs("ApelidoIntegrante") = Request.Form("Apelido")
	Rs("ICQIntegrante") = Request.Form("ICQ")
	Rs("EmailIntegrante") = Request.Form("Email")
	Rs("SalaIntegrante") = Request.Form("Sala")
	Rs("NascimentoIntegrante") = Request.Form("Nascimento")
	Rs("DescricaoIntegrante") = Request.Form("Descricao")
	Rs("CerimonialIntegrante") = Request.Form("Cerimonial")
	Rs("FotoIntegrante") = Request.Form("Foto")
	Rs.Update
	'Verificar se é para ser deletado
	If Request.Form("delintegrante") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  Else
	'Se for um novo integrante
	Rs.Source = "SELECT * FROM TblIntegrantes"
	Rs.Open()
	Rs.AddNew
	Rs("NomeIntegrante") = Request.Form("Nome")
	Rs("ApelidoIntegrante") = Request.Form("Apelido")
	Rs("ICQIntegrante") = Request.Form("ICQ")
	Rs("EmailIntegrante") = Request.Form("Email")
	Rs("SalaIntegrante") = Request.Form("Sala")
	Rs("NascimentoIntegrante") = Request.Form("Nascimento")
	Rs("DescricaoIntegrante") = Request.Form("Descricao")
	Rs("CerimonialIntegrante") = Request.Form("Cerimonial")
	Rs("FotoIntegrante") = Request.Form("Foto")
	Rs.Update
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=integrantes"
Case "jornais"
  If Request.Form("preview") = "1" Then
	'Preview
	Response.Redirect "jornalpreview.asp?codigo=" & Request.Form("Codigo") & "&edicao=" & Request.Form("Edicao") & "&autor=" & Request.Form("Autor") & "&n1=" & Request.Form("N1") & "&n2=" & Request.Form("N2") & "&n3=" & Request.Form("N3") & "&n4=" & Request.Form("N4") & "&n5=" & Request.Form("N5") & "&n6=" & Request.Form("N6") & "&n7=" & Request.Form("N7") & "&n8=" & Request.Form("N8") & "&deljornal=" & Request.Form("deljornal")
  Else
  	If Not Request.QueryString("cod") = "0" Then
	  'Se o jornal já existir, editá-lo...
	  Rs.Source = "SELECT * FROM TblJornais WHERE CodigoJornal=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
	  	'Editar...
	  	Rs("EdicaoJornal") = Request.Form("Edicao")
	  	Rs("AutorJornal") = Request.Form("Autor")
	  	Rs("N1Jornal") = Request.Form("N1")
	  	Rs("N2Jornal") = Request.Form("N2")
	  	Rs("N3Jornal") = Request.Form("N3")
	  	Rs("N4Jornal") = Request.Form("N4")
	  	Rs("N5Jornal") = Request.Form("N5")
	  	Rs("N6Jornal") = Request.Form("N6")
	  	Rs("N7Jornal") = Request.Form("N7")
	  	Rs("N8Jornal") = Request.Form("N8")
	  	Rs.Update
	  Else
	  	'Se o jornal não existir, criá-lo
	  	Rs.AddNew()
	  	Rs("EdicaoJornal") = Request.Form("Edicao")
	  	Rs("AutorJornal") = Request.Form("Autor")
	  	Rs("N1Jornal") = Request.Form("N1")
	  	Rs("N2Jornal") = Request.Form("N2")
	  	Rs("N3Jornal") = Request.Form("N3")
	  	Rs("N4Jornal") = Request.Form("N4")
	  	Rs("N5Jornal") = Request.Form("N5")
	  	Rs("N6Jornal") = Request.Form("N6")
	  	Rs("N7Jornal") = Request.Form("N7")
	  	Rs("N8Jornal") = Request.Form("N8")
	  	Rs.Update
	  End If
	  'Verificar se é para ser deletado
	  If Request.Form("deljornal") = "1" Then
	  	Rs.Delete
	  End If
		Rs.Close()
  	Else
	  'Se for um novo jornal
	  Rs.Source = "SELECT * FROM TblJornais"
	  Rs.Open()
	  Rs.AddNew
	  	Rs("EdicaoJornal") = Request.Form("Edicao")
	  	Rs("AutorJornal") = Request.Form("Autor")
	  	Rs("N1Jornal") = Request.Form("N1")
	  	Rs("N2Jornal") = Request.Form("N2")
	  	Rs("N3Jornal") = Request.Form("N3")
	  	Rs("N4Jornal") = Request.Form("N4")
	  	Rs("N5Jornal") = Request.Form("N5")
	  	Rs("N6Jornal") = Request.Form("N6")
	  	Rs("N7Jornal") = Request.Form("N7")
	  	Rs("N8Jornal") = Request.Form("N8")
	  Rs.Update
	  Rs.Close
  	End If
  End If
  Response.Redirect "gerenc.asp?func=jornais"
Case "links"
  If Not Request.QueryString("cod") = "0" Then
	'Se o link já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblLinks WHERE CodigoLink=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("NomeLink") = Request.Form("Nome")
	  Rs("EnderecoLink") = Request.Form("Endereco")
	  Rs.Update
	Else
	  'Se o link não existir, criá-lo
	  Rs.AddNew()
	  Rs("NomeLink") = Request.Form("Nome")
	  Rs("EnderecoLink") = Request.Form("Endereco")
	  Rs.Update
	End If
	'Verificar se é para ser deletado
	If Request.Form("dellink") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  Else
	'Se for um novo link
	Rs.Source = "SELECT * FROM TblLinks"
	Rs.Open()
	Rs.AddNew
	Rs("NomeLink") = Request.Form("Nome")
	Rs("EnderecoLink") = Request.Form("Endereco")
	Rs.Update
	Rs.Close
  End If
  Response.Redirect "gerenc.asp?func=links"
Case "menus"
  If Not Request.QueryString("cod") = "0" Then
	'Se o menu já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblMenus WHERE CodigoMenu=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("NomeMenu") = Request.Form("Nome")
	  Rs("EnderecoMenu") = Request.Form("Endereco")
	  Rs("StatusMenu") = Request.Form("Status")
	  If Request.Form("Status") = "" Then Rs("StatusMenu") = 0
	  Rs.Update
	Else
	  'Se o link não existir, criá-lo
	  Rs.AddNew()
	  Rs("NomeMenu") = Request.Form("Nome")
	  Rs("EnderecoMenu") = Request.Form("Endereco")
	  Rs("StatusMenu") = Request.Form("Status")
	  If Request.Form("Status") = "" Then Rs("StatusMenu") = 0
	  Rs.Update
	End If
	'Verificar se é para ser deletado
	If Request.Form("delmenu") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  Else
	Rs.Source = "SELECT * FROM TblMenus"
	Rs.Open()
	Rs.AddNew
	Rs("NomeMenu") = Request.Form("Nome")
	Rs("EnderecoMenu") = Request.Form("Endereco")
	Rs("StatusMenu") = Request.Form("Status")
	If Request.Form("Status") = "" Then Rs("StatusMenu") = 0
	Rs.Update
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=menus"
Case "newspro"
  'Verificar se é para vizualizar ou salvar
  If Request.Form("preview") = "1" Then
	'Preview
	Response.Redirect "newspreview.asp?codigo=" & Request.Form("Codigo") & "&titulo=" & Request.Form("Titulo") & "&autor=" & Request.Form("Autor")& "&avatar=" & Request.Form("Avatar") & "&data=" & Request.Form("Data") & "&texto=" & Request.Form("Texto") & "&delnews=" & Request.Form("delnews")
  Else
	'Salvar
	If Not Request.QueryString("cod") = "0" Then
	  'Editar
	  Rs.Source = "SELECT * FROM TblNews WHERE CodigoNews=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
		'Se o registro existir, editar
		Rs("TituloNews") = Request.Form("Titulo")
		Rs("AutorNews") = Request.Form("Autor")
		Rs("AvatarNews") = Request.Form("Avatar")
		Rs("DataNews") = Request.Form("Data")
		Rs("TextoNews") = Request.Form("Texto")
		Rs.Update
	  	If Request.Form("delnews") = "1" Then
		  Rs.Delete
	  	End If
	  End If
	  Rs.Close()
	Else
	  'Novo
	  Rs.Source = "SELECT * FROM TblNews"
	  Rs.Open()
	  Rs.AddNew
	  Rs("TituloNews") = Request.Form("Titulo")
	  Rs("AutorNews") = Request.Form("Autor")
	  Rs("AvatarNews") = Request.Form("Avatar")
	  Rs("DataNews") = Request.Form("Data")
	  Rs("TextoNews") = Request.Form("Texto")
	  Rs.Update
	  Rs.Close()
	End If
  End If
  Response.Redirect "gerenc.asp?func=newspro"
Case "usuarios"
  'Verificar a validade da senha (tem q ser a geral)
  If lCase(Session("user")) = "boi" Then
	If Not Request.QueryString("cod") = "0" Then
	  'Editar...
	  Rs.Source = "SELECT * FROM TblUsuarios WHERE CodigoUsuario=" & Request.QueryString("cod")
	  Rs.Open()
	  If Not Rs.EOF Then
		Rs("NomeUsuario") = Request.Form("Nome")
		Rs("SenhaUsuario") = Request.Form("Senha")
		Rs("TituloUsuario") = Request.Form("Titulo")
		Rs("BiografiaUsuario") = Request.Form("Biografia")
		Rs("LogsUsuario") = Request.Form("Logs")
		Rs.Update
	  End If
	Else
	  'Se o usuário não exixtir, criá-lo
	  Rs.Source = "SELECT * FROM TblUsuarios"
	  Rs.Open()
	  Rs.AddNew
	  Rs("NomeUsuario") = Request.Form("Nome")
	  Rs("SenhaUsuario") = Request.Form("Senha")
	  Rs("TituloUsuario") = Request.Form("Titulo")
	  Rs("BiografiaUsuario") = Request.Form("Biografia")
	  Rs("LogsUsuario") = 0
	  If Not Trim(Request.Form("Logs")) = "" Then Rs("LogsUsuario") = Request.Form("Logs")
	  Rs.Update
	End If
	'Verificar se é para ser deletado
	If Request.Form("delusuario") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=usuarios"
Case "vacilos"
  If Not Request.QueryString("cod") = "0" Then
	'Se o vacilo já existir, editá-lo...
	Rs.Source = "SELECT * FROM TblVacilos WHERE CodigoVacilo=" & Request.QueryString("cod")
	Rs.Open()
	If Not Rs.EOF Then
	  'Editar...
	  Rs("TextoVacilo") = Request.Form("Vacilo")
	  Rs("AutorVacilo") = Request.Form("Autor")
	  Rs.Update
	Else
	  'Se o vacilo não existir, criá-lo
	  Rs.AddNew()
	  Rs("TextoVacilo") = Request.Form("Vacilo")
	  Rs("AutorVacilo") = Request.Form("Autor")
	  Rs.Update
	End If
	'Verificar se é para ser deletado
	If Request.Form("delvacilo") = 1 Then
	  Rs.Delete
	End If
	Rs.Close()
  End If
  Response.Redirect "gerenc.asp?func=vacilos"
End Select
Response.Redirect "gerenc.asp"
%>