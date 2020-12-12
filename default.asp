<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<% 
sql = ""
if Request("buscar_aluno") <> "" then
  sql = " AND lslAlunos.Nome LIKE '%" & Replace(Request("buscar_aluno"),"'","""") &"%'"
End If

if Request("buscar_curso") <> 0 then
  sql = sql & " AND lslAlunos.Cod_curso = "& Request("buscar_curso") 
End If


Dim Reg_curso
Dim Reg_curso_cmd
Dim Reg_curso_numRows

Set Reg_curso_cmd = Server.CreateObject ("ADODB.Command")
Reg_curso_cmd.ActiveConnection = MM_Conexao_STRING
Reg_curso_cmd.CommandText = "SELECT Cod_curso, Nome FROM dbo.lslCurso" 
Reg_curso_cmd.Prepared = true

Set Reg_curso = Reg_curso_cmd.Execute
Reg_curso_numRows = 0


Dim Reg_alunos
Dim Reg_alunos_cmd
Dim Reg_alunos_numRows

Set Reg_alunos_cmd = Server.CreateObject ("ADODB.Command")
Reg_alunos_cmd.ActiveConnection = MM_Conexao_STRING
Reg_alunos_cmd.CommandText = "SELECT lslAlunos.Cod_aluno, lslAlunos.Nome, lslAlunos.DN, lslAlunos.Telefone, lslAlunos.Data_cadastro, lslAlunos.Cod_curso AS Cod_curso_aluno, lslCurso.Cod_curso AS Cod_curso, lslCurso.Nome AS Nome_curso FROM lslAlunos INNER JOIN lslCurso ON lslAlunos.Cod_curso = lslCurso.Cod_curso where 1 = 1 "& sql &"" 
'Response.Write(Reg_alunos_cmd.CommandText)
Reg_alunos_cmd.Prepared = true

Set Reg_alunos = Reg_alunos_cmd.Execute
Reg_alunos_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Sistema Zica</title>
<link rel="stylesheet" type="text/css" href="estilo.css"/>

</head>

<body>
<table width="100%" border="0">
  <tr>
    <td width="10%"><a href="aluno_inserir.asp">Cadastrar Aluno</a></td>
    <td width="90%"><a href="curso_inserir.asp">Cadastrar Curso</a></td>
  </tr>
</table>
<hr>
<font face="Verdana, Geneva, sans-serif" size="+3">Listar alunos</font>
<br>
<form action="default.asp" method="get" id="form_buscar">
<table width="100%" border="0" bgcolor="#CCCCCC">
  <tr>
    <td width="5%" align="right">Nome:</td>
    <td width="10%"><input type="text" name="buscar_aluno" id="buscar_aluno" maxlength="50" size="20" value="<%= Request("buscar_aluno") %>" /></td>
    <td width="5%" align="right">Curso:</td>
    <td width="12%">
    <select id="buscar_curso" name="buscar_curso">
            <option value="0">Selecione</option>
        <% While Not Reg_curso.Eof %>
          <option <% If cInt(Request("buscar_curso")) = cInt(Reg_curso.Fields.Item("Cod_curso").Value) Then %> selected="selected" <% End If %> value="<%=(Reg_curso.Fields.Item("Cod_curso").Value)%>"><%=(Reg_curso.Fields.Item("Nome").Value)%></option>
        <% 
            Reg_curso.MoveNext()
            Wend
        %>
    </select>
    </td>
    <td width="68%"><input type="submit" id="buscar" value="Buscar" /></td>
  </tr>
</table>
</form>
<table width="100%" border="0">
  <tr bgcolor="#33CCFF" align="center">
    <td width="35%">Aluno</td>
    <td width="20%">Telefone</td>
    <td width="35%">Curso</td>
    <td width="10%">Ação</td>
  </tr>
  <% While Not Reg_alunos.Eof %>
  <% If cor = "#E1E1E1" Then cor = "#AEAEAE" Else cor = "#E1E1E1" End If %>
  <tr bgcolor="<%= cor %>">
    <td><%=(Reg_alunos.Fields.Item("Nome").Value)%></td>
    <td align="center"><%=(Reg_alunos.Fields.Item("Telefone").Value)%></td>
    <td><%=(Reg_alunos.Fields.Item("Nome_curso").Value)%></td>
    <td align="center"><a href="aluno_atualizar.asp?Cod_aluno=<%=(Reg_alunos.Fields.Item("Cod_aluno").Value)%>"><img src="Imagens/ico_atualizar.gif" border="0"></a></td>
  </tr>
  <% 
   Reg_alunos.MoveNext()
   Wend
  %>
</table>
</body>
</html>
<%
Reg_curso.Close()
Set Reg_curso = Nothing
%>
<%
Reg_alunos.Close()
Set Reg_alunos = Nothing
%>
