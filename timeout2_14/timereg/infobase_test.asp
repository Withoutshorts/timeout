<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<textarea cols="5" rows="5">dfgdgdfg</textarea>
<%
x = 0
if x = 0 then 
			vzb = "visible"
			dsp = ""
			%>
			<input type="hidden" name="lastredigervalue" id="lastredigervalue" value="<%=infobase_record_ID%>">
			<%
			else
			vzb = "hidden"
			dsp = "none"
			end if
			%>
			<input type="text" name="navn_<%=infobase_record_ID%>" id="navn_<%=infobase_record_ID%>" value="<%=infobase_record_navn%>" style="border:1px #003399 solid; font-size:9px; width:200px;">
			<br>
			<textarea  cols="50" rows="20" name="FM_besk_<%=infobase_record_ID%>" id="FM_besk_<%=infobase_record_ID%>" style="border:1px #003399 solid;">
			dfgdgdfg
			</textarea>
			<img src="ill/blank.gif" width="150" height="1" border="0"><input type="image" src="../ill/opdaterpil.gif">
			


</body>
</html>
