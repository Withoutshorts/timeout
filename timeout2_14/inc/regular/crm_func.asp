<%
function crmtopmenu()
%><br>
<a href="crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal&status=0&id=0&emner=0" class='vmenu'>Kalender</a>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href="kunder.asp?menu=crm&shokselector=1&ketype=e&selpkt=osigt" class="vmenu">Kontakter</a>
&nbsp;&nbsp;|&nbsp;&nbsp;<a href='crmhistorik.asp?menu=crm&ketype=e&func=hist&id=0&selpkt=hist' class='vmenu'>Aktions historik</a>


<%
end function




'**********************************************'
'**** Funktioner til CRM aktions historik *****'
'**********************************************'


Function DropDownMinute(FormNavn, KlokkesletField)
if func = "red" then
strSQL = "SELECT " & KlokkesletField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSelVal = Minute(FormatDateTime(oRec(KlokkesletField)))
oRec.Close
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
        strSelVal = "00"
        end if

        Dim arrMinutes
        arrMinutes = split("00,15,30,45", ",")
        for i = 0 to 3
        if arrMinutes(i) = strSelVal then
        strSelOpt = " selected"
        else
        strSelOpt = Null
        end if
        %>
        <option value="<%=arrMinutes(i)%>" <%=strSelOpt%>><%=arrMinutes(i)%></option>
        <%Next
End Function



Function DropDownHour(FormNavn, KlokkesletField)
if func = "red" then
strSQL = "SELECT " & KlokkesletField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSelVal = Hour(FormatDateTime(oRec(KlokkesletField)))
oRec.close
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
		 strSelVal = Hour(Now)
		 end if
        Response.Write "strSelVal="&strSelVal
		 for i = 0 to 23
		 if i = cint(strSelVal) then
		 strSelOpt = " selected"
		 else
		 strSelOpt = Null
		 end if%>
		<option value="<%=i%>" <%=strSelOpt%>><%=i%></option>
		<%Next
End Function




Function DropDownDay(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(0)
else
        strSelVal = Request.Form(FormNavn)
end if
        if Not strSelVal <> "" then
		 strSelVal = Day(Now)
		end if

        for i = 1 to 31
		if cint(strSelVal) = i then
		strSelOpt = " selected"
		else
		strSelOpt = Null
		end if %>
		<option value="<%=i%>" <%=strSelOpt%>><%=i%></option>
        <%Next
End Function


Function DropDownMonth(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(1)
else
		strSelVal = Request.Form(FormNavn)
end if
		if Not strSelVal <> "" then
		strSelVal = Month(Now())
		end if

		dim ArrMonth
		ArrMonth = split("0, jan, feb, mar, apr, maj, jun, jul, aug, sep, okt, nov, dec",",")
		for i = 1 to 12
		if cint(strSelVal) = i then
		strSelOpt = " selected"
		else
		strSelOpt = Null
		end if
		%>
		<option value="<%=i%>"<%=strSelOpt%>><%=ArrMonth(i)%></option>
<%Next 
End Function
Function DropDownYear(FormNavn, DatoField)
if func = "red" then
strSQL = "SELECT " & DatoField & " FROM crmhistorik Where id = " & request.QueryString("crmaktion")
			oRec.open strSQL, oConn, 3
strSplitDato = split(oRec(DatoField), "-")
oRec.Close
strSelVal = strSplitDato(2)
else
        strSelVal = Request.Form(FormNavn)
end if
		 if Not strSelVal <> "" then
		 strSelVal = Year(Now())
		 end if

		 for i = Year(Now())-5 to Year(Now())+5
		 if i = cint(strSelVal) then
		 strSelOpt = " selected"
		 else
		 strSelOpt = Null
		 end if
		 %>
		<option value="<%=i%>"<%=strSelOpt%>><%=i%></option>
		<%Next
End Function
	

%>