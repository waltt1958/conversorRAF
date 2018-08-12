<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title> CONVERSOR ARCHIVOS SANCOR RAFAELA PAQUETERIA </title>

</HEAD>

<body onload="maximizar()">


ARCHIVO</a>
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
' recupera= request.form("NOMBREARCHIVO")
' archivo= "c:\inetpub\wwwroot\conversorRAF\" & recupera

' sqlLIMPIA = "DELETE * from sancor"
' conectarOEP.execute sqlLIMPIA

' sqlBORRA= "DELETE * from copiaSANCOR"
' conectarOEP.execute sqlBORRA

' Set objFSO = Server.CreateObject ("Scripting.FileSystemObject")

' Set varArchivo = objFSO.OpenTextFile (archivo,1)

' Do while not varArchivo.AtEndOfStream

	 ' arrayLinea = split (varArchivo.ReadLine, "|", - 1,1)

	' sqlinsert= "INSERT INTO sancor (Apellido, Calle, CP, Localidad,Provincia, Operativa, Guia) VALUES ( '" & arrayLinea(0) & "','" & arrayLinea(1) & "','" & arrayLinea(2) & "', '"& arrayLinea(3) & "','" & arrayLinea(4) & "','" & arrayLinea(5) & "','" & arrayLinea(6) & "')"
	 
	' conectarOEP.execute (sqlinsert)
 
' loop 
	
' Set varArchivo = Nothing
' Set objFSO = Nothing

' sqlINSERT="INSERT INTO copiaSANCOR select * from sancor"
' conectarOEP.execute sqlINSERT

' sqlACTUALIZA ="UPDATE copiaSANCOR SET copiaSANCOR.RETIdomicilio = 'Independencia', copiaSANCOR.RETInumero = '333', copiaSANCOR.RETIpiso ='0', copiaSANCOR.RETIdepto ='0', copiaSANCOR.RETIcp ='2322', copiaSANCOR.RETIlocalidad = 'Sunchales', copiaSANCOR.RETIprov = 'Santa Fe'"
' conectarOEP.execute sqlACTUALIZA

 ' Set rsARCHIVO = Server.CreateObject("ADODB.recordset")

 ' sqlARCHIVO= "select * from copiaSANCOR"

 ' rsARCHIVO.open sqlARCHIVO, conectarOEP

   ' Set fso = Server.CreateObject ("Scripting.FileSystemObject")

  ' Set arcTEXTO = fso.CreateTextFile(server.mappath("bajaSANCOR.txt"), true)

  ' texto1 = rsARCHIVO.Fields(0).name & "|" & rsARCHIVO.Fields(1).name & "|" & rsARCHIVO.Fields(2).name & "|" & rsARCHIVO.Fields(3).name & "|" & rsARCHIVO.Fields(04).name & "|" & _
  ' rsARCHIVO.Fields(5).name & "|" & rsARCHIVO.Fields(6).name & "|" & rsARCHIVO.Fields(7).name & "|" & rsARCHIVO.Fields(8).name & "|" & rsARCHIVO.Fields(9).name & "|" & _
  ' rsARCHIVO.Fields(10).name & "|" & rsARCHIVO.Fields(11).name & "|" & rsARCHIVO.Fields(12).name & "|" & rsARCHIVO.Fields(13).name & "|" & rsARCHIVO.Fields(14).name _
  ' & "|" & rsARCHIVO.Fields(15).name & "|" & rsARCHIVO.Fields(16).name & "|" & rsARCHIVO.Fields(17).name & "|" & rsARCHIVO.Fields(18).name & "|" & _
  ' rsARCHIVO.Fields(19).name & "|" & rsARCHIVO.Fields(20).name & "|" & rsARCHIVO.Fields(21).name & "|" & rsARCHIVO.Fields(22).name & "|" & _
  ' rsARCHIVO.Fields(23).name & "|" & rsARCHIVO.Fields(24).name & "|" & rsARCHIVO.Fields(25).name & "|" & rsARCHIVO.Fields(26).name & "|" & rsARCHIVO.Fields(27).name _
  ' & "|" & rsARCHIVO.Fields(28).name & "|" & rsARCHIVO.Fields(29).name & "|" & rsARCHIVO.Fields(30).name & "|" & rsARCHIVO.Fields(31).name
  
  ' arcTEXTO.WriteLine(texto1)
 
  ' do while not rsARCHIVO.EOF

  ' texto= rsARCHIVO.Fields("Apellido") & "|" & rsARCHIVO("Calle") & "|" & rsARCHIVO("CP") & "|" & rsARCHIVO("Localidad") & "|" & rsARCHIVO("Provincia") & "|" & _
  ' rsARCHIVO("Operativa")  & "|" & rsARCHIVO("Guia") & "|" & rsARCHIVO("DESTnombre") & "|" & rsARCHIVO("DESTnumero") & "|" & rsARCHIVO("DESTpiso") & "|" & _
  ' rsARCHIVO("DESTdepto") & "|" & rsARCHIVO("DESTtelefono") & "|" & rsARCHIVO("DESTemail") & "|" & rsARCHIVO("RETIdomicilio") & "|" & rsARCHIVO("RETInumero") _
  ' & "|" & rsARCHIVO("RETIpiso") & "|" & rsARCHIVO("RETIdepto") & "|" & rsARCHIVO("RETItelefono") & "|" & rsARCHIVO("RETIcp") & "|" & rsARCHIVO("RETIlocalidad") _
  ' & "|" & rsARCHIVO("RETIprov") & "|" & rsARCHIVO("RETIcontacto") & "|" & rsARCHIVO("PAQpeso") & "|" & rsARCHIVO("PAQalto") & "|" & rsARCHIVO("PAQlargo") _
  ' & "|" & rsARCHIVO("PAQancho") & "|" & rsARCHIVO("PAQvalor") & "|" & rsARCHIVO("NROremito") & "|" & rsARCHIVO("IMPremito") & "|" & rsARCHIVO("NROproducto") _
  ' & "|" & rsARCHIVO("RETIemail") & "|" & rsARCHIVO("observaciones")

  ' arcTEXTO.WriteLine(texto)

  ' rsARCHIVO.MoveNext

  ' loop

 ' rsARCHIVO.close
 ' Set rsARCHIVO= nothing
	
  ' Set fso = nothing
  ' Set arcTEXTO = nothing

' sqlLIMPIA = "DELETE * from sancor"
' conectarOEP.execute sqlLIMPIA

' sqlBORRA= "DELETE * from copiaSANCOR"
' conectarOEP.execute sqlBORRA

FileName= "bajaSANCOR.txt"
Response.Clear 
Response.ContentType="application / octet-stream"  
Response.AddHeader "content-disposition", "attachment; filename=" & FileName


Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type= 1
Response.CharSet = "UTF-8"
objStream.Open
objStream.LoadFromFile Server.MapPath(FileName)
Response.AddHeader "Content-Disposition", "attachment; filename=" & FileName
Response.ContentType = "application/octet-stream"
while not objStream.EOS
	Response.BinaryWrite objStream.Read(1024 * 64)
Wend
objStream.Close
Set stream= Nothing
Response.Flush
Response.End

' Set objStream = Nothing


' FileName= "bajaSANCOR.txt"
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


<!-- <a href="../bajaSANCOR.txt" download="bajaSANCOR.txt">DESCARGAR 

<!--#include virtual="/desconectar.asp"-->





</body>

</HTML>