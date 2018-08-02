<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title> CONVERSOR ARCHIVOS SANCOR RAFAELA PAQUETERIA </title>

</HEAD>

<body onload="maximizar()">

<!--#include virtual="/conectar.asp"-->

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SUC. OCA RAFAELA - PAQUETERIA (Oper. 288140 )</h1>
<br>
<br>
<br>
<br><br>
<br><br>
<br>
<br>
<br>
<br>
<br>

<%
recupera= request.form("NOMBREARCHIVO")
archivo= "c:\inetpub\wwwroot\conversorRAF\" & recupera

sqlLIMPIA = "DELETE * from sancor"
conectarOEP.execute sqlLIMPIA

sqlBORRA= "DELETE * from copiaSANCOR"
conectarOEP.execute sqlBORRA

Set objFSO = Server.CreateObject ("Scripting.FileSystemObject")

Set varArchivo = objFSO.OpenTextFile (archivo,1)

Do while not varArchivo.AtEndOfStream

	 arrayLinea = split (varArchivo.ReadLine, "|", - 1,1)

	sqlinsert= "INSERT INTO sancor (Apellido, Calle, CP, Localidad,Provincia, Operativa, Guia) VALUES ( '" & arrayLinea(0) & "','" & arrayLinea(1) & "','" & arrayLinea(2) & "', '"& arrayLinea(3) & "','" & arrayLinea(4) & "','" & arrayLinea(5) & "','" & arrayLinea(6) & "')"
	 
	conectarOEP.execute (sqlinsert)
 
loop 
	
Set varArchivo = Nothing
Set objFSO = Nothing



set rsCLONAR= Server.CreateObject("ADODB.recordset")
sqlCLONAR = "select * from sancor"
rsCLONAR.open sqlCLONAR, conectarOEP


do while not rsCLONAR.EOF

	rsCLONAR.movefirst
	sqlINSERT = "
	conectarOEP.sqlINSERT (poner toda lasentencia más los datos fijos)
	
	rsCLONAR.movenext	

loop


set rsARCHIVO= nothing
rsARCHIVO.close






'sqlCLONAR= "select * from copiaSANCOR from sancor"
'sqlCLONAR= "INSERT INTO copiaSANCOR select * from sancor"
'conectarOEP.execute sqlCLONAR

'sqlALTERA = "ALTER TABLE copiaSANCOR ADD COLUMN DESTnombre TEXT(30), COLUMN DESTnumero text(5),COLUMN DESTpiso text(2),COLUMN DESTdepto text (4), COLUMN DESTtelefono text(15), column DESTemail text(50),column RETIdomicilio text(60), column RETInumero text(5),column RETIpiso text(2), column RETIdepto text(4), column RETItelefono text(15),column RETIcp text(8), column RETIlocalidad text(30), column RETIprov text(30), column RETIcontacto text(30), column PAQpeso integer, column PAQalto integer, column PAQlargo integer, column PAQancho integer, column PAQvalor integer, column NROremito text(13),column IMPremito integer,column NROproducto text(30), column RETIemail text(50), column observaciones text(200)"
'conectarOEP.execute sqlALTERA

'sqlACTUALIZA ="UPDATE copiaSANCOR SET copiaSANCOR.RETIdomicilio = 'Independencia', copiaSANCOR.RETInumero = '333', copiaSANCOR.RETIpiso ='0', copiaSANCOR.RETIdepto ='0', copiaSANCOR.RETIcp ='2322', copiaSANCOR.RETIlocalidad = 'Sunchales', copiaSANCOR.RETIprov = 'Santa Fe'"
'conectarOEP.execute sqlACTUALIZA






Set rsARCHIVO = Server.CreateObject("ADODB.recordset")

sqlARCHIVO= "select * from copiaSANCOR"

rsARCHIVO.open sqlARCHIVO, conectarOEP

Set fso = Server.CreateObject ("Scripting.FileSystemObject")

Set arcTEXTO = fso.CreateTextFile(server.mappath("Archivo_Pauta_Meridiano.txt"), true)

do while not rsARCHIVO.EOF

texto= rsARCHIVO.Fields("Apellido") & "|" & rsARCHIVO("Calle") & "|" & rsARCHIVO("CP") & "|" & rsARCHIVO("Localidad") & "|" & rsARCHIVO("Provincia") & "|" & _
rsARCHIVO("Operativa")  & "|" & rsARCHIVO("Guia") & "|" & rsARCHIVO("DESTnombre") & "|" & rsARCHIVO("DESTnumero") & "|" & rsARCHIVO("DESTpiso") & "|" & _
rsARCHIVO("DESTdepto") & "|" & rsARCHIVO("DESTtelefono") & "|" & rsARCHIVO("DESTemail") & "|" & rsARCHIVO("RETIdomicilio") & "|" & rsARCHIVO("RETInumero") _
& "|" & rsARCHIVO("RETIpiso") & "|" & rsARCHIVO("RETIdepto") & "|" & rsARCHIVO("RETItelefono") & "|" & rsARCHIVO("RETIcp") & "|" & rsARCHIVO("RETIlocalidad") _
& "|" & rsARCHIVO("RETIprov") & "|" & rsARCHIVO("RETIcontacto") & "|" & rsARCHIVO("PAQpeso") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQlargo") _
& "|" & rsARCHIVO("PAQancho") & "|" & rsARCHIVO("PAQvalor") & "|" & rsARCHIVO("NROremito") & "|" & rsARCHIVO("IMPremito") & "|" & rsARCHIVO("NROproducto") _
& "|" & rsARCHIVO("RETIemail") & "|" & rsARCHIVO("observaciones")

arcTEXTO.WriteLine(texto)

rsARCHIVO.MoveNext

loop



Set rsARCHIVO= nothing
rsARCHIVO.close	
Set fso = nothing
Set arcTEXTO = nothing

'sqlBORRA="DROP table copiaSANCOR"
'conectarOEP.execute sqlBORRA


sqlLIMPIA = "DELETE * from sancor"
conectarOEP.execute sqlLIMPIA

sqlBORRA= "DELETE * from copiaSANCOR"
conectarOEP.execute sqlBORRA



' FileName= request.form("NOMBREARCHIVO")
' Response.Clear 
' Response.ContentType="application / octet-stream"  
' Response.AddHeader "content-disposition", "attachment; filename=" & FileName


' Set stream = Server.CreateObject("ADODB.stream") 
' stream.type = adTypeBinary 
' stream.open

' stream.LoadFromFile Server.MapPath(FileName)
' While Not stream.EOS 
' response.BinaryWrite stream.Read(1024 * 64)
' Wend

' stream.Close
' Set stream= Nothing
' Response.Flush
' Response.End


'exportacion = "SANCOR" & format$(now, "dd-mm-yyyy hh-mm-ss")
' texto = ".txt"
' salida = exportacion & texto

' if request.form("Enviar")<> " " then
' response.write (salida)
	' if request.form("NOMBREARCHIVO")<>" " then
         ' response.write (request.form("nombrearchivo"))
		 ' response.write ("viene el archivo")
  ' Else
  
			' response.write ("xxx")
	' End If
' else

' redireccionar a la pagina de carga con venta aviso de que no había archivo

' end if
%>


<!--#include virtual="/desconectar.asp"-->

</script>

<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>