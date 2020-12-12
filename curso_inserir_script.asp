<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<%
sql = "INSERT INTO lslCurso (Nome) VALUES ('"& Replace(Request("nome_curso"), "'", """") &"');" 
'Response.Write(sql)
'Response.End
Set Reg = Server.CreateObject("adodb.recordset")
Reg.Open sql, MM_Conexao_STRING
Set Reg = Nothing

Response.Redirect("default.asp")

%>
