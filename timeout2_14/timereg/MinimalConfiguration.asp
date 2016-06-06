<!-- #include file = "CuteEditor_Files/include_CuteEditor.asp" --> 
<html>	
    <head>
		<link rel="stylesheet" href="asp.css"  type="text/css" />
	</head>
    <body>
		
		<form name="theForm" action="Get_HtmlContent.asp" method="post" ID="Form1">
					
		<h1>Minimal Configuration</h1>
		<hr>
		<div class="infobox">
			This example shows you the Minimal Configuration of editor.
		</div>
						
		<br />
		<%
			
			dim content
					
			content ="<STRONG><EM><U>Easy to Install, Easy to Use</U></EM></STRONG><BR>"
			
			Dim editor
			Set editor = New CuteEditor
					
			editor.ID = "Editor1"
			editor.Text = content
			editor.FilesPath = "CuteEditor_Files"
			editor.AutoConfigure = "Minimal"
			
			editor.Width = 600
			'editor.Height = 240
			editor.Draw()
		%>
		
		<br>
		<input type="submit" value="Submit"  ID="Submit2" NAME="Submit1">
						
		</form>
	</body>
</html>

