<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--"http://www.w3.org/TR/html4/loose.dtd" -->
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />

</head>
<body>

<!-- #include file = "../cuteeditor_2013/CuteEditor_Files/include_CuteEditor.asp" -->

<%



	                    
	                    dim content
	                    content = strBeskrivelse
            			
			            
			            Set editorK = New CuteEditor
            					
			            editorK.ID = "FM_beskrivelse"
			            editorK.Text = content
			            editorK.FilesPath = "CuteEditor_Files"
			            editorK.AutoConfigure = "Default"
            			
			            editorK.Width = 700
			            editorK.Height = 280
			            editorK.Draw()
		                %>
		                <br />
            &nbsp;
		
</body>
