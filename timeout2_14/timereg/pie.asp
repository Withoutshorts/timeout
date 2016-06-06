<%@ LANGUAGE=VBSCRIPT %>
<% Server.ScriptTimeOut = 600 %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<% 
blnIE4 = False
strUA = Request.ServerVariables("HTTP_USER_AGENT")
intMSIE = Instr(strUA, "MSIE ") + 5
If intMSIE > 5 Then
  If CInt(Mid(strUA, intMSIE, Instr(intMSIE, strUA, ".") - intMSIE)) >= 4 Then 
    blnIE4 = True
  End If 
End If
%>

<HTML>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Results of your traffic query</title>
<style type="text/css">  
BODY {font-family:Tahoma,Verdana,Arial,sans-serif; font-size:12px; font-weight:normal}
SPAN {font-family:Tahoma,Verdana,Arial,sans-serif; font-size:14px; font-weight:bold} 
TD {font-family:Tahoma,Verdana,Arial,sans-serif; font-size:12px; font-weight:normal}
H1 {font-family:Arial,sans-serif; font-size:16px; color:darkgray}
</style>
</head>
<body BGCOLOR="#FFFFFF">
<H1>Timeregistrering for <%=session("user")%></H1>
<SPAN>Fordelt på eksterne og interne timeregistreringer:</SPAN><BR><P>
<%
QUOT = Chr(34)
CRLF = Chr(13) & Chr(10)

intTotalEks = 0

strSQL = "SELECT timer AS antTimerEks FROM timer WHERE Tmnavn='"& Session("user") &"' AND Tfaktim = 1"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalEks = intTotalEks + Cint(oRec("antTimerEks"))

oRec.movenext
wend

oRec.close

intTotalEksNotfakbar = 0

strSQL = "SELECT timer AS antTimerEks FROM timer WHERE Tmnavn='"& Session("user") &"' AND Tfaktim = 2"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalEksNotfakbar = intTotalEksNotfakbar + Cint(oRec("antTimerEks"))

oRec.movenext
wend

oRec.close

strSQL = "SELECT timer AS antTimerInt FROM timer WHERE Tmnavn='"& Session("user") &"' AND Tfaktim = 0"
oRec.open strSQL, oConn, 3

while not oRec.EOF 
intTotalInt = intTotalInt + Cint(oRec("antTimerInt"))

oRec.movenext
wend

oRec.close
Set oConn = Nothing

intTotal = intTotalEks + intTotalInt + intTotalEksNotfakbar

'Response.write "intTotalEks: " & intTotalEks & " Timer<br>"
'Response.write "intTotalInt: " & intTotalInt & " Timer<br>"
%>

<%
If blnIE4 Then

'*** create the pie chart ***
iLineCnt = 3
iDegPos = 0
iRectTop = -110
%>

<OBJECT ID=sgPie STYLE="position:absolute; height:230; width:300; top:120; left:20" 
  CLASSID="CLSID:369303C2-D7AC-11D0-89D5-00A0C90833E6">
<PARAM NAME="Line0001" VALUE="SetLineColor(0, 0, 0)"> 
<PARAM NAME="Line0002" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0003" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0004" VALUE="Oval(-130, -90, 200, 200, 0)">

<PARAM NAME="Line0005" VALUE="Rect(100, <%= CStr(iRectTop) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0006" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0007" VALUE="SetFillColor(255, 51, 0)">
<PARAM NAME="Line0008" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPos) %>, <%= CStr(Round(intTotalInt / intTotal * 360)) %>, 0)">
<PARAM NAME="Line0009" VALUE="Rect(95, <%= CStr(iRectTop) %>, 20, 15, 0)">
<%
iDegPos = iDegPos + CInt(intTotalInt / intTotal * 360)
iRectTop = iRectTop + 25 
%>
<PARAM NAME="Line0010" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0011" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0012" VALUE="Rect(100, <%= CStr(iRectTop) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0013" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0014" VALUE="SetFillColor(0, 51, 153)">
<PARAM NAME="Line0015" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPos) %>, <%= CStr(Round(intTotalEks/intTotal*360)) %>, 0)">
<PARAM NAME="Line0016" VALUE="Rect(95, <%= CStr(iRectTop) %>, 20, 15, 0)">
<%
iDegPos = iDegPos + CInt(intTotalEksNotfakbar / intTotal * 360)
iRectTop = iRectTop + 25 
%>
<!--Procent indeks forklaring -->
<PARAM NAME="Line0017" VALUE="SetFillColor(192, 192, 192">
<PARAM NAME="Line0018" VALUE="SetLineStyle(0)">
<PARAM NAME="Line0019" VALUE="Rect(100, <%= CStr(iRectTop) + 5 %>, 20, 15, 0)">
<PARAM NAME="Line0023" VALUE="Rect(95, <%= CStr(iRectTop) %>, 20, 15, 0)">
<!--Pie -->
<PARAM NAME="Line0020" VALUE="SetLineStyle(1)">
<PARAM NAME="Line0021" VALUE="SetFillColor(102, 204, 51)">
<PARAM NAME="Line0022" VALUE="Pie(-140, -100, 200, 200, <%= CStr(iDegPos) %>, <%= CStr(Round(intTotalEksNotfakbar/intTotal*360)) %>, 10)">
</OBJECT>

<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:126; left:295">
<B>Interne Job</B>  &nbsp; Count: <B><%= intNav2 %></B> &nbsp;(<B><%= FormatPercent(intTotalInt / intTotal, 1) %></B>)
</DIV>
<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:151; left:295">
<B>Eksterne fakturerbare Job</B>  &nbsp; Count: <B><%= intNav3 %></B> &nbsp;(<B><%= FormatPercent(intTotalEks / intTotal, 1) %></B>)
</DIV>
<DIV STYLE="font-family:Tahoma,Verdana,Arial,sans-serif; position:absolute; top:176; left:295">
<B>Eksterne ikke fakturerbare Job</B>  &nbsp; Count: <B><%= intNav4 %></B> &nbsp;(<B><%= FormatPercent(intTotalEksNotfakbar / intTotal, 1) %></B>)
</DIV>



<%
Else
%>

<CENTER><TABLE>
<%
'*** create the table of values ***
Response.Write "<TR><TD>Netscape Navigator 2.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;" & intNav2 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;" & FormatPercent(intNav2 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Netscape Navigator 3.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;" & intNav3 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;" & FormatPercent(intNav3 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Netscape Navigator 4.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;" & intNav4 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;" & FormatPercent(intNav4 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Netscape Navigator 5.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;" & intNav5 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;" & FormatPercent(intNav5 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Internet Explorer 3.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;"  & intIE3 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;"  & FormatPercent(intIE3 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Internet Explorer 4.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;"  & intIE4 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;"  & FormatPercent(intIE4 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Internet Explorer 5.x &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;"  & intIE5 & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;"  & FormatPercent(intIE5 / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Opera (all versions) &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;"  & intOpera & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;"  & FormatPercent(intOpera / intTotal, 1) & "</TD></TR>" & CRLF
Response.Write "<TR><TD>Other User Agents &nbsp; &nbsp;</TD><TD ALIGN=RIGHT><B> &nbsp; &nbsp; &nbsp; &nbsp;" & intOther & "</B></TD><TD ALIGN=RIGHT> &nbsp; &nbsp; &nbsp; &nbsp;" & FormatPercent(intOther / intTotal, 1) & "</TD></TR>" & CRLF
%>
</TABLE></CENTER>

<%
End If
%>


</body>
</HTML>





