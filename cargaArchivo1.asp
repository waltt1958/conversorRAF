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

He leído decenas de tutoriales sobre subir archivos al servidor usando ASP, en todos los casos lo que encontré fueron copias de un mismo procedimiento, con variaciones, o archivos que se pueden bajar, sin ninguna explicación.
Si bien la mayoría sirve, ante cualquier variación que queremos realizar sobre esos archivos no se encuentran explicaciones detalladas de cómo se realizan.
El presente instructivo, es una explicación detallada de cómo subir un archivo a un servidor, de forma sencilla y sin comprobaciones, como usan la mayoría de los ejemplos, que hacen que no se comprenda bien el proceso.
Luego, el lector podrá agregarle a este procedimiento las complejidades necesarias para realizar esta acción.
El presente instructivo es exclusivamente para subir un archivo, no comprobará el tamaño del archivo, ni otras consideraciones, pero partiendo de este código se podrá ampliar la programación, tan compleja como se quiera.
Entonces, partimos de la hipótesis que necesitamos subir un archivo a desde nuestra máquina a un servidor remoto.
Para ello armamos un formulario que contenga el campo y el botón examinar, que nos permite ver la carpeta local donde buscaremos el archivo. Esto sería:
En código:
<form id="form1" name="form1" method="POST" action="">
<table>
<tr>
<td width="59%"><INPUT NAME="File1" SIZE=30 TYPE="file"></td>
</tr>
<tr><td> <input type="text" name="button" id="button" value="Enviar" /></td></tr>
</table>
</form>
 
Esto ubicará un campo, con un botón de examinar a la derecha. Y otro botón abajo, enviar, para iniciar la acción de subir al servidor el archivo
 
Cuando se presione el botón examinar, abre una ventana que muestra el contenido de las carpetas del servidor local.
 
Al seleccionar un archivo, quedará en el campo la dirección física donde se encuentra el archivo.
Por ejemplo algo como esto: C:\imagenes\nombre.jpg
Si el camino a la imagen es más largo, sería algo como esto: C:\carpeta\imagenes\nombre.jpg
 
Para aislar el nombre del archivo, en cualquier caso, tendría que identificar lo que haya hacia la derecha de la barra, \.
En algunos foros, he visto que realizan una especie de función recursiva que rastrea cada \, hasta que llega a la última y ahí encuentra el nombre del archivo.
 
En mi ejemplo haré algo que considero más sencillo: invertiré todo el nombre, buscaré una sola vez la barra, extraeré el string que quede desde el comienzo hasta esa barra y volveré a invertir el archivo, con ello obtendré el nombre del archivo. Esto es mucho más rápido que una función que se ejecute varias veces para encontrar todas las barras.
 
Suponemos que el nombre del campo es: nombreArchivo.
Invierto el campo:

<%

Dim Invertida
Invertida= strReverse(nombreArchivo)

%>

Esto daría como resultado en el contenido de invertir esto: “gpj.erbmon\senegami\ateprac\:C”
Ahora, busco la primera barra:

<%

Dim dondeBarra
dondeBarra= instr(Invertida,"\")
 
%>
 
Ahora en dondeBarra hay un número que indica en dónde encontró la barra, dentro de la variable Invertida, contando desde 1.
Ahora extraigo la cadena de caracteres desde 1, hasta la barra, y le resto 1 carácter para extraer la misma barra.
Para el ejemplo el valor que contiene dondeBarra es: 11, contando desde el principio e incluyendo la barra.
 
<%

Dim Extraer
Extraer= mid(Invertida,1,dondeBarra-1)

%>
 
Ahora  Extraer contiene esto: gpj.erbmon
 
Sencillamente volvemos invertir y obtenemos el nombre del archivo:

<% 

Dim nombreFinal
nombreFinal= strReverse(Extraer)

%>
 
En este momento, en nombreFinal contiene “nombre.jpg”
 
Ahora, utilizamos un objeto que permitirá pasar el archivo al servidor remoto.
 
<% 

Dim ForWriting, FileName
ForWriting = 2
                FileName=nombreFinal
                Set fso = CreateObject("Scripting.FileSystemObject")
                set f = fso.OpenTextFile("DireccionDeDestino\" & FileName, ForWriting, True)
                f.Write FileName
                Set f = nothing
                Set fso = nothing

				%>
 
Esto es todo.
ForWriting es un parámetro del objeto fso que implica un GET, un tomar el archivo, y ForWriting es un parámetro que indica que se va a escribir en una carpeta.
Direcciondedestino: es el lugar físico, cuidado, no confundir con una dirección http, que no lo es.
 
DirecciónDeDestino puede ser algo asi: “C:\inetpub\wwwroot\imagen\” y lo que hará es colocar el archivo copiado en ese sitio.
Si en el servidor remoto se desconoce la dirección física en donde se desea copiar al archivo, se puede suplantar la línea con la siguiente:
 
Set f = fso.OpenTextFile(server.mappath("..") & "\imagen\" &  FileName, ForWriting, True)
 
Donde el objeto server.mapmath contiene la dirección física.
 
El archivo completo implicaría que primero debería verse el formulario, que nos permite elegir el archivo y luego de presionar el botón enviar, se ejecute la parte en que sube el archivo. Por ello, nuestra página se ejecuta, dos veces, una cuando abre el formulario, otra cuando sube, entonces iniciamos la página con una condición:
 
Esta función devuelve un 1 o un 2, cuando arranca es 1. En el formulario indicamos que en caso que sea 1, se muestre el formulario, y cuando sea 2, se ejecute la acción de subir archivo al servidor.
 
El archivo se llamara: upload.asp
 
<%

Func = Request("Func")
if isempty(Func) Then
Func = 1
End if
Select Case Func
Case 1

%>

<form id="form1" name="form1" method="POST" action="upload.asp">
<table>
<tr>
<td width="59%"><INPUT NAME="File1" SIZE=30 TYPE="file"></td>
</tr>
<tr><td> <input type="text" name="button" id="button" value="Enviar" /></td></tr>
</table>
</form>

<%

Case 2

Dim Invertida, dondeBarra, Extraer, nombreFinal
Invertida= strReverse(nombreArchivo)
dondeBarra= instr(Invertida,"\")
Extraer= mid(Invertida,1,dondeBarra-1)
nombreFinal= strReverse(Extraer)
Dim ForWriting, FileName
ForWriting = 2
FileName=nombreFinal
Set fso = CreateObject("Scripting.FileSystemObject")
set f = fso.OpenTextFile("DireccionDeDestino\" & FileName, ForWriting, True)
f.Write FileName
Set f = nothing
Set fso = nothing

%>

</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>