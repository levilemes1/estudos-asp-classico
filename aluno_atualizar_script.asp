<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<%

sql = "UPDATE lslAlunos SET Nome = '"& Request("nome_aluno")&"', DN = '"& Request("mes")&"/"&Request("dia")&"/"&Request("ano") &"', Telefone = '"& Request("telefone")&"', Cod_curso = "& Request("Cod_curso")&" WHERE Cod_aluno = "& Request("Cod_aluno") &";"
'Response.Write(sql)
'Response.End
Set Reg = Server.CreateObject("adodb.recordset")
Reg.Open sql, MM_Conexao_STRING
Set Reg = Nothing

Response.Redirect("default.asp")
%>
