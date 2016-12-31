<%
Set MM_Conneccao_STRING = Server.CreateObject("ADODB.Connection")
MM_Conneccao_STRING.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("\saladois\db\sd.mdb")

%>