<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<%
Dim Reg_curso
Dim Reg_curso_cmd
Dim Reg_curso_numRows

Set Reg_curso_cmd = Server.CreateObject ("ADODB.Command")
Reg_curso_cmd.ActiveConnection = MM_Conexao_STRING
Reg_curso_cmd.CommandText = "SELECT Cod_curso, Nome FROM dbo.lslCurso" 
Reg_curso_cmd.Prepared = true

Set Reg_curso = Reg_curso_cmd.Execute
Reg_curso_numRows = 0


Dim Reg_atualizar
Dim Reg_atualizar_cmd
Dim Reg_atualizar_numRows

Set Reg_atualizar_cmd = Server.CreateObject ("ADODB.Command")
Reg_atualizar_cmd.ActiveConnection = MM_Conexao_STRING
Reg_atualizar_cmd.CommandText = "SELECT lslAlunos.Cod_aluno, lslAlunos.Nome, lslAlunos.DN, lslAlunos.Telefone, lslAlunos.Data_cadastro, lslAlunos.Cod_curso, lslCurso.Cod_curso AS Cod_curso, lslCurso.Nome AS Nome_curso FROM lslAlunos INNER JOIN lslCurso ON lslAlunos.Cod_curso = lslCurso.Cod_curso WHERE Cod_aluno = "& Request("Cod_aluno") &";"
'Response.Write(Reg_atualizar_cmd.CommandText)
Reg_atualizar_cmd.Prepared = true

Set Reg_atualizar = Reg_atualizar_cmd.Execute
Reg_atualizar_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Sistema Zica</title>
<link rel="stylesheet" type="text/css" href="estilo.css"/>
</head>
<script language="javascript" type="text/javascript">

function Verificar(){
	if ((document.form_aluno.nome_aluno.value.length <= 4) && (document.form_aluno.curso.value != "Selecione")){
		alert('Todos os campos são obrigatórios');
		return false;
	}
	
}

</script>
<body>
<table width="100%" border="0">
  <tr>
    <td width="100%"><a href="default.asp">Voltar</a></td>
  </tr>
</table>
<hr>
<font face="Verdana, Geneva, sans-serif" size="+3">Atualizar Aluno</font>
<br>
<form action="aluno_atualizar_script.asp?Cod_aluno=<%= Request("Cod_aluno")%>" method="post" id="form_aluno" name="form_aluno">
<table width="100%" border="0" align="center">
  <tr>
    <td width="40%" align="right">Aluno:</td>
    <td width="60%"><input type="text" id="nome_aluno" name="nome_aluno" maxlength="50" size="50" value="<%=(Reg_atualizar.Fields.Item("Nome").Value)%>" /></td>
  </tr>
  <tr>
    <td width="40%" align="right">Curso:</td>
    <td width="60%">
        <select id="Cod_curso" name="Cod_curso">
                <option value="0">Selecione</option>
			<% While Not Reg_curso.Eof %>
                <option <% If cInt(Reg_atualizar.Fields.Item("Cod_curso").Value) = cInt(Reg_curso.Fields.Item("Cod_curso").Value) Then %> selected="selected" <% End If %> value="<%=(Reg_curso.Fields.Item("Cod_curso").Value)%>"><%=(Reg_curso.Fields.Item("Nome").Value)%></option>
            <% 
				Reg_curso.MoveNext()
				Wend
            %>
        </select>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right">Data Nasc:</td>
    <td width="60%">
    <select id="dia" name="dia">
      <% i = 1
		while i <= 31 %>
            <option <% If CINT(Day(Reg_atualizar.Fields.Item("DN").Value)) = CINT(i) Then %> selected="selected" <% End If %> value="<%= Day(Reg_atualizar.Fields.Item("DN").Value)%>"><%= Day(Reg_atualizar.Fields.Item("DN").Value) %></option>
		<% i = i + 1
		wend %>
      </select>
        <select id="mes" name="mes">
		<% for i = 1 to 12 %>
          <option <% If Month(Reg_atualizar.Fields.Item("DN").Value) = i Then %> selected="selected" <% End If %> value="<%= Month(Reg_atualizar.Fields.Item("DN").Value)%>"><%= Month(Reg_atualizar.Fields.Item("DN").Value) %></option>
        <% next %>      
        </select>
        <select id="ano" name="ano">
		<% for i = 1950 to 2013 %>
          <option <% If Year(Reg_atualizar.Fields.Item("DN").Value) = i Then %> selected="selected" <% End If %> value="<%= Year(Reg_atualizar.Fields.Item("DN").Value)%>"><%= Year(Reg_atualizar.Fields.Item("DN").Value) %></option>
        <% next %>      
        </select>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right">Telefone:</td>
    <td width="60%"><input type="text" id="telefone" name="telefone" maxlength="14" size="50" value="<%=(Reg_atualizar.Fields.Item("Telefone").Value)%>" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Atualizar" onclick="javascript:Verificar();" /></td>
  </tr>
</table>
</form>
</body>
</html>
<%
Reg_curso.Close()
Set Reg_curso = Nothing
%>
<%
Reg_atualizar.Close()
Set Reg_atualizar = Nothing
%>
