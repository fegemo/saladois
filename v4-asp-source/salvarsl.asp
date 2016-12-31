 <%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/Conneccao.asp" -->
<!--#include file="functions.asp" -->
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MM_Conneccao_STRING
Rs.LockType = 3

Erro = ""
'Filtro anti-adonai e fracassados da oitava série
If Not lCase(Request.Form("Senha")) = "salasd" Then
  Erro = "Hauhauhuhauhuhauhaau fracassados, estão sem senha! Vão ficar com o rabinho entre as pernas agora, hein! Hauahauhauhauhauahhaua!!!"
  Response.Redirect "rsltd.asp?rsltd=" & Erro
End If
'Filtro anti-galera do mal
If Not InStr(1, Request.Form("Autor"), ">", 1) = 0 Or Not InStr(1, Request.Form("Autor"), "<", 1) = 0 Then
  Erro = "É, seu fracassado... agora eu te peguei hehehehehe<br><br>Espertinho, agora eu tenho seu IP :P"
  Rs.Source = "SELECT * FROM TblFracassados"
  Rs.Open()
  Rs.AddNew
  Rs("AutorFracassado") = Request.Form("Autor")
  Rs("TextoFracassado") = Request.Form("Texto")
  Rs("IPFracassado") = Request.ServerVariables("REMOTE_HOST")
  Rs("DataFracassado") = Data()
  Rs.Update
  Rs.Close()
  Response.Redirect "rsltd.asp?rsltd=" & Erro
End If
If Not InStr(1, Request.Form("Texto"), ">", 1) = 0 Or Not InStr(1, Request.Form("Texto"), "<", 1) = 0 Then
  Erro = "É, seu fracassado... agora eu te peguei hehehehehe<br><br>Espertinho, agora eu tenho seu IP :P"
  Rs.Source = "SELECT * FROM TblFracassados"
  Rs.Open()
  Rs.AddNew
  Rs("AutorFracassado") = Request.Form("Autor")
  Rs("TextoFracassado") = Request.Form("Texto")
  Rs("IPFracassado") = Request.ServerVariables("REMOTE_HOST")
  Rs("DataFracassado") = Data()
  Rs.Update
  Rs.Close()
  Response.Redirect "rsltd.asp?rsltd=" & Erro
End If

Select Case Request.QueryString("func")
Case "frases"
  If Trim(Request.Form("Texto")) = "" Then
	Erro = "Você se esqueceu de escrever a frase! Mas que lesão!!"
  Else
	Rs.Source = "SELECT * FROM TblFrases"
  	Rs.Open()
  	Rs.AddNew
  	Rs("AutorFrase") = Request.Form("Autor")
  	If Trim(Request.Form("Autor")) = "" Then
	  Rs("AutorFrase") = "Anônimo"
  	End If
  	Rs("TextoFrase") = Request.Form("Texto")
  	Rs.Update
  	Rs.Close()
	Response.Redirect "frases.asp"
  End If
Case "vacilos"
  If Trim(Request.Form("Texto")) = "" Then
	Erro = "Você se esqueceu de escrever o vacilo! Mas que lesão!!"
  Else
	Rs.Source = "SELECT * FROM TblVacilos"
  	Rs.Open()
  	Rs.AddNew
  	Rs("AutorVacilo") = Request.Form("Autor")
  	If Trim(Request.Form("Autor")) = "" Then
	  Rs("AutorVacilo") = "Anônimo"
  	End If
  	Rs("TextoVacilo") = Request.Form("Texto")
  	Rs.Update
  	Rs.Close()
	Response.Redirect "vacilos.asp"
  End If
End Select
If Erro = "" Then
  Response.Redirect "index.asp"
Else
  Response.Redirect "rsltd.asp?rsltd=" & Erro
End If
%>