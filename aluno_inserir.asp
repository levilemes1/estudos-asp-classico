<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/Conexao.asp" -->
<%
Dim Reg_escolher_curso
Dim Reg_escolher_curso_cmd
Dim Reg_escolher_curso_numRows

Set Reg_escolher_curso_cmd = Server.CreateObject ("ADODB.Command")
Reg_escolher_curso_cmd.ActiveConnection = MM_Conexao_STRING
Reg_escolher_curso_cmd.CommandText = "SELECT Cod_curso, Nome FROM dbo.lslCurso" 
Reg_escolher_curso_cmd.Prepared = true

Set Reg_escolher_curso = Reg_escolher_curso_cmd.Execute
Reg_escolher_curso_numRows = 0
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
<font face="Verdana, Geneva, sans-serif" size="+3">Cadastrar Aluno</font>
<br>
<form action="aluno_inserir_script.asp" method="post" id="form_aluno" name="form_aluno">
<table width="100%" border="0" align="center">
  <tr>
    <td width="40%" align="right">Aluno:</td>
    <td width="60%"><input type="text" id="nome_aluno" name="nome_aluno" maxlength="50" size="50" /></td>
  </tr>
  <tr>
    <td width="40%" align="right">Curso:</td>
    <td width="60%">
        <select id="curso" name="curso">
                <option value="Selecione">Selecione</option>
			<% While Not Reg_escolher_curso.Eof %>
                <option value="<%=(Reg_escolher_curso.Fields.Item("Cod_curso").Value)%>"><%=(Reg_escolher_curso.Fields.Item("Nome").Value)%></option>
            <% 
				Reg_escolher_curso.MoveNext()
				Wend
            %>
        </select>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right">Data Nasc:</td>
    <td width="60%">
        <select id="dia" name="dia" >
        <% i = 1
		while i <= 31 %>
            <option value="<%= i %>"><%= i %></option>
		<% i = i + 1
		wend %>
        </select>
        <select id="mes" name="mes">
		<% for i = 1 to 12 %>
        	<option value="<%= i %>"><%= i %></option>
        <% next %>      
        </select>
        <select id="ano" name="ano">
		<% for i = 1950 to 2013 %>
                <option value="<%= i %>"><%= i %></option>
        <% next %>      
        </select>
    </td>
  </tr>
  <tr>
    <td width="40%" align="right">Telefone:</td>
    <td width="60%"><input type="text" id="telefone" name="telefone" maxlength="14" size="50" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Cadastrar" onclick="javascript:Verificar();" /></td>
  </tr>
</table>
</form>
</body>
</html>
<%
Reg_escolher_curso.Close()
Set Reg_escolher_curso = Nothing
%>
