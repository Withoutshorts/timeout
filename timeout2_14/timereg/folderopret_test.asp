<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
    <br /><br /><br />
    Test af holder oprettelse hos Dencker:<br /><br /><br />

    <%
        
         
                                            'Set objFSO = server.createobject("Scripting.FileSystemObject")
                                            'folderPath = "\\195.189.130.210\" 'timeout_xp\wwwroot\"
                                            'Set objnewFolder = objFSO.CreateFolder(folderPath & "DER1")
                                            'Set objnewFolder = nothing


                                            'ServerShare = "\\outzource.dk\dencker_job"
                                            'ServerShare = "\\remote.dencker.net\JOB_Timeout"
                                            ServerShare = "\\remote.dencker.net\"
                                            
                                            
                                            
                                            'UserName = "Administrator"
                                            UserName = "ad"
                                            'UserName = "TimeOut"
                                            'Password = "Sok!2637"
                                            Password = "ad1996"
                                            'Password = "betina00"
                                            'Password = "timeout"
                                            'Password = ""

                                            Set NetworkObject = CreateObject("WScript.Network")
                                            Set FSO = CreateObject("Scripting.FileSystemObject")

                                            NetworkObject.MapNetworkDrive "", ServerShare, False, UserName, Password

                                            Set objnewFolder = FSO.CreateFolder(ServerShare &"/"& strNavn &"_"& strjnr)

                                            Set Directory = FSO.GetFolder(ServerShare)
                                            For Each FileName In Directory.Files
                                                response.write FileName.Name & "<br>"
                                                response.write fso.GetParentFolderName(FileName.Name) & "<br>"
                                            Next

                                            Set FileName = Nothing
                                            Set Directory = Nothing
                                            Set FSO = Nothing
                                            Set objnewFolder = nothing

                                            response.write "Folder opRettet!"
                                            response.end
         %>


</body>
</html>
