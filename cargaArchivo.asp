<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title> CONVERSOR ARCHIVOS SANCOR RAFAELA PAQUETERIA </title>

</HEAD>

<body onload="maximizar()">


<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SUC. OCA RAFAELA - PAQUETERIA (Oper. 288140 )</h1>
<br>

<form action="procesa.asp" method="post" TARGET="_self">
<table width=30%" align="center">
<tr align="center"><td><input type="File" name="NOMBREARCHIVO" accept=".txt" MAXLENGTH="60" autofocus required></td></tr>
<tr height="50"></tr>
<tr align="center"><td><input type="submit" value="ENVIAR" name="manda" class="button" > </td></tr>
</table>
</form> 


</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>