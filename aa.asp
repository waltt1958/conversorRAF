<%
texto= "txt"
bbdd= "mdb"
clasico= "asp"
forma= "css"
imagen = "png"

Set objFSO = server.CreateObject ("Scripting.FileSystemObject")
set objFolder=objFSO.GetFolder("c:\inetpub\wwwroot\conversorRAF\")

for each objFile in objFolder.files


if (objFSO.GetExtensionName(objFile)) <> bbdd then
objFile.delete
end if

if (objFSO.GetExtensionName(objFile)) <> clasico then
objFile.delete
end if

if (objFSO.GetExtensionName(objFile)) <> forma then
objFile.delete
end if

if (objFSO.GetExtensionName(objFile)) <> imagen then
objFile.delete
end if

next

%>