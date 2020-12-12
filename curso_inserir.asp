<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Sistema Zica</title>
<link rel="stylesheet" type="text/css" href="estilo.css"/>
</head>
<script language="javascript" type="text/javascript">

function Verificar(){
	if (document.form_curso.nome_curso.value.length <= 4){
		alert('O nome do curso é invalido');
		document.form_curso.nome_curso.select();
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
<font face="Verdana, Geneva, sans-serif" size="+3">Cadastrar Curso</font>
<br>
<form action="curso_inserir_script.asp" method="post" id="form_curso" name="form_curso">
<table width="100%" border="0" align="center">
  <tr>
    <td width="40%" align="right">Curso:</td>
    <td width="60%"><input type="text" id="nome_curso" name="nome_curso" maxlength="50" size="50" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Cadastrar" onclick="javascript:return Verificar();" /></td>
  </tr>
</table>
</form>
</body>
</html>
