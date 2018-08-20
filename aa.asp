<%
texto= "txt"
bbdd= "mdb"
clasico= "asp"
forma= "css"
imagen = "png"

Set objFSO = server.CreateObject ("Scripting.FileSystemObject")
set objFolder=objFSO.GetFolder("c:\inetpub\wwwroot\conversorRAF\")

for each objFile in objFolder.files

Select case objFSO.GetExtensionName(objFile)
case bbdd
case clasico
case forma
case imagen
case else

objFile.delete
end select



next

%>