




	
	
	

    
	
	<%
	'**********************************************************
	'**************** Orpet / red Faktura Step 1 **************
	'**********************************************************
	 %>
	 
	 <!--
	    <table cellspacing=2 cellpadding=0 border=0>
	    <tr><td>
	    <a href="erp_opr_faktura_kontakter.asp" target="erp2_1" class="rmenu"><u>1) - Vælg kontakt (debitor)</u></a>
	    </td></tr>
	    <tr><td>
	    <a href="#" target="erp2_1" class="erp_gray">2) - Vælg job / aftale og datointerval</a>
	    </td></tr>
	    </table>
	    -->
	 

	 <form action="erp_opr_faktura_fs.asp?formsubmitted=1" method="POST">
        <table style="width:280px;">
	 
	 <tr>
	   <td bgcolor="#ffffff" style="padding:10px 10px 10px 10px; border:0px #8caae6 solid;">
           <b>Søg på kunde:</b><br />
            <input id="FM_sog" name="FM_sog" type="text" value="<%=sogKri %>" style="width:205px; font-size:11px; border:2px yellowgreen solid;">&nbsp;<input id="Submit0" type="submit" value=">>" style="font-size:9px;" />
           <br /><span style="color:#999999; font-size:9px;">(% wildcard)</span></td></tr> 
	 </table>
         </form>
	    
        <form action="erp_opr_faktura_fs.asp?formsubmitted=1&visjobogaftaler=1" method="POST">
            <input type="hidden" name="FM_sog" value="<%=sogKri %>" />
	 <table style="width:280px;">
	 
	  <tr>
	   <td bgcolor="#FFFFFF" style="padding:10px 10px 10px 10px; border:0px #8caae6 solid;">
	
	 <b>Kunde(r):</b><br />
      <select name="FM_kunde" id="FM_kunde" size="1" style="width:215px; font-size:11px;">
		<%
		
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE ketype <> 'e' AND (useasfak <= 2) "& kSQLkri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
				<option value="0">Ingen</option>
		</select>&nbsp;<input type="submit" value=">>" style="font-size: 9px;"  />
           <!--<input id="Button1" type="image" src="../ill/pilstorxp.gif" onclick="nextstep1()" />-->

		<br />
		  
		 <input id="FM_jobonoff" name="FM_jobonoff" type="checkbox" value="j" <%=jobonoffCHK %> /> Vis lukkede job og aft. 
		
		</td></tr>
	 </table>
        </form>


    <!--
	</div>
	-->