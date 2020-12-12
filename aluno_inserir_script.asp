<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<%

sql = "INSERT INTO lslAlunos (Nome, DN, Telefone, Data_cadastro, Cod_curso) VALUES ('"& Replace(Request("nome_aluno"), "'", """") &"', '"& Request("mes")&"/"&Request("dia")&"/"&Request("ano") &"', '"& Replace(Request("Telefone"), "'", """") &"', getdate(), "& Request("curso") &");" 
'Response.Write(sql)
'Response.End
Set Reg = Server.CreateObject("adodb.recordset")
Reg.Open sql, MM_Conexao_STRING
Set Reg = Nothing

Response.Redirect("default.asp")
%>
