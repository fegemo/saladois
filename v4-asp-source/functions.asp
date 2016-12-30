<%
Function Data()
  Select Case Month(Date)
  Case 1
	Mes = "Janeiro"
  Case 2
	Mes = "Fevereiro"
  Case 3
	Mes = "Março"
  Case 4
	Mes = "Abril"
  Case 5
	Mes = "Maio"
  Case 6
	Mes = "Junho"
  Case 7
	Mes = "Julho"
  Case 8
	Mes = "Agosto"
  Case 9
	Mes = "Setembro"
  Case 10
	Mes = "Outubro"
  Case 11
	Mes = "Novembro"
  Case 12
	Mes = "Dezembro"
  End Select
  Data = Day(Date) & " de " & Mes & " de " & Year(Date)
End Function
%>
<script language="JavaScript">
function fPrompt(pergunta, valorpadrao) {
  fPrompt = prompt(pergunta,valorpadrao);
};
</script>