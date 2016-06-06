	<%response.buffer = true
	%>
	<!--#include file="../inc/connection/conn_db_inc.asp"-->
	<!--#include file="../inc/errors/error_inc.asp"-->
	<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->

	<%
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	func = request("func")
	thisfile = "crmstat"
	%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			
	<%call menu_2014() %>
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:90px; top:82px; visibility:visible;">
	
	
    <h4>CRM stat</h4>
    
    <%call filterheader(0,0,600,pTxt)%>
	<table cellspacing=0 width=100% cellpadding=4 border=0>
	<tr><form action="#" method="post"><td>
		<b>Kontakter:</b>&nbsp;<select name="kundeid" size="1" style="width:225px;">
		<option value="0">Ingen Kontakt valgt</option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE Kid <> 0 ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>
		&nbsp;&nbsp;<input type="submit" value="Vælg Kontakt">
		</form>
	</td></tr>
	<tr>
	    <td><b>Medarbejder:</b>&nbsp;<select id="Select1">
                <option></option>
            </select></td>
	</tr>
	<tr>
	    <td><b>Emner:</b>&nbsp;<select id="Select2">
                <option></option>
            </select></td>
	</tr>
	<tr>
	    <td><b>Status:</b>&nbsp;<select id="Select3">
                <option></option>
            </select></td>
	</tr>
	
	</table>
    <!-- filter header -->
	</td></tr></table>
	</div>
	<br /><br />
    
    
    <%
	
	tTop = 0
	tLeft = 0
	tWdth = 400
	
	
	call tableDiv(tTop,tLeft,tWdth)
	
	%>
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#ffffff" width=100%>
	<tr><td align=right>
    
    Antal aktioner:<br />
    Fordelt pr. kontakt:<br />
    Fordelt pr. medarb.:<br />
    Antal hotte kontakter:<br />
    </td>
    <td>
    <b>231</b><br />
    <b>41</b><br />
    <b>11</b><br />
    <b>31</b><br />
    
    </td>
    </tr>
    </table>
     </div>
    
    </div>
	
<%
end if
%>


<!--#include file="../inc/regular/footer_inc.asp"-->
 



