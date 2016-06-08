<%
function mattopmenu()
%><br>


<a href='materialer.asp' class='rmenu'>Materialer</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href='materiale_grp.asp' class='rmenu'>Materialegrupper</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<!--<a href='materiale_stat.asp?menu=mat' class='rmenu'>Materiale forbrug (statistik)</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->
<%if ((level < 8 AND lto <> "bowe") OR (level < 7)) then %>
<a href="materialer_indtast.asp?id=0&fromsdsk=0&aftid=0" class=rmenu target="_blank"><%=tsa_txt_191 %></a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%end if %>	
<a href='materialer_ordrer.asp' class='rmenu'>Materialeordrer</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href='produktioner.asp' class='rmenu'>Standard Produktioner</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href='leverand.asp' class='rmenu'>Leverandør- og Service -partnere</a>
<!--
&nbsp;&nbsp;|&nbsp;&nbsp;
<a href='materiale_stat.asp?menu=mat' class='rmenu'>Materialeforbrug /udlæg i periode</a>-->



<%
end function

function matregmenu(mtop,mleft,Knap1_bg,Knap2_bg,Knap3_bg,Knap4_bg)

%>
<div id="matregmenu" style="position:absolute; top:<%=mtop+20%>px; left:<%=mleft%>px; border:0px #000000 solid;">
    
    

    <div id=knap1 style="float:left; width:180px; background-color:<%=Knap1_bg %>; padding:4px; border:1px #D6DFf5 solid;">
	<a href="materialer_indtast.asp?vis=otf&id=<%=id %>&aftid=<%=aftid %>&fromdesk=<%=fromdesk %>&lastid=<%=lastid %>" class=alt><%=tsa_txt_188 %></a>
	</div>

    <%select case lto
     case "epi", "epi_no", "epi_no", "epi_sta", "epi_ab", "epi_uk", "xintranet - local" 
       
       case else%>
	<div id=knap2 style="float:left; width:100px; background-color:<%=Knap2_bg %>; padding:4px; border:1px #D6DFf5 solid;">
	<a href="materialer_indtast.asp?vis=lag&id=<%=id %>&aftid=<%=aftid %>&fromdesk=<%=fromdesk %>&lastid=<%=lastid %>" class=alt><%=tsa_txt_189 %></a>
	</div>
    <%end select %>
	
	
	<div id=knap4 style="float:left; width:130px; background-color:<%=Knap4_bg %>; padding:4px; border:1px #D6DFf5 solid;">
	<a href="materiale_stat.asp?hidemenu=1&id=<%=id %>&FM_visprjob_ell_sum=2" class=alt><%=tsa_txt_257 %></a>
	</div>
	
    <!-- 
    <%if level = 1 OR lto = "bowe" then%>
	<div id=knap3 style="position:absolute; left:390px; top:20px; width:80px; background-color:<%=Knap3_bg %>; padding:4px; border:1px #D6DFf5 solid;">
	<a href="materialer.asp?menu=0" class=alt target="_blank"><%=tsa_txt_190 %></a>
	</div>
        <%end if %>
    -->
	
	
	
	</div>

<%
end function	


function flytlager(id, antal, nytLager, opretkun)



	             
					
	'*** Opdaterer valgte materiale på valgte gruppe **
	strSQL = "SELECT antal, id, varenr FROM materialer WHERE id = " & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		
		if opretkun <> "1" then
		nytAntalglLager = (oRec("antal")) - (antal)
		strSQLupd = "UPDATE materialer SET antal = " & nytAntalglLager & " WHERE id =" & oRec("id")
		
		'Response.Write "opretkun" & opretkun &"<br>"
		'Response.Write strSQLupd & "<br>"
        'Response.end 
		
		oConn.execute(strSQLupd)
		end if
	
	varenr = oRec("varenr")
				
	end if
	oRec.close
	
	'*** Opdaterer valgte materiale på den nye gruppe **
	strSQL = "SELECT antal, id, varenr FROM materialer WHERE varenr = '" & varenr & "' AND matgrp = " & nytLager
	oRec.open strSQL, oConn, 3
	n = 0
	if not oRec.EOF then
		
		nytAntalnytLager = (oRec("antal")) + (antal)
		strSQLupd2 = "UPDATE materialer SET antal = " & nytAntalnytLager & " WHERE id =" & oRec("id")
		oConn.execute(strSQLupd2)
	
	n = 1			
	end if
	
	oRec.close
	
	
	if cint(n) = 0 then
	
					
					'*** Indsætter valgte materiale i valgte gruppe **
					strSQL = "SELECT editor, dato, navn, varenr, indkobspris, salgspris, "_
					&" arrestordredato, ptid, ptid_arr, enhed, pic, leva, levb, sera, serb, minlager "_
					&" FROM materialer WHERE id = " & id
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
						
						editor = session("user")
						strDato = day(now)&"-"&month(now)&"-"&year(now)
						matNavn = oRec("navn")
						varenr = oRec("varenr")
						
						if opretkun <> "1" then
						ikpris = replace(oRec("indkobspris"), ",",".")
						spris = replace(oRec("salgspris"), ",",".")
						else
						ikpris = dblKobsPris
						spris = dblSalgsPris
						end if
						
						restordredato = oRec("arrestordredato")
						ptid = oRec("ptid")
						ptid_arr = oRec("ptid_arr")
						enhed = oRec("enhed")
						picId = oRec("pic")
						leva = oRec("leva")
						levb = oRec("levb")
						sera = oRec("sera")
						serb = oRec("serb")
						minlager = oRec("minlager")
						
						strSQLins = "INSERT INTO materialer (editor, dato, navn, varenr, matgrp, antal, indkobspris, salgspris, "_
						&" arrestordredato, ptid, ptid_arr, enhed, pic, leva, levb, sera, serb, minlager) VALUES "_
						&"('"& editor &"', '"& strDato &"', "_
						&" '"& matNavn &"', '"& varenr &"',"_
						&""& nytLager &", "& antal &", "_
						&""& ikpris &", "& spris  &","_
						&" '"& restordredato &"', "& oRec("ptid") &", "_
						&""& ptid_arr &", '"& enhed &"', "_
						&""& picId &", "_
						&""& leva &", "& levb &", "_
						&""& sera &", "_
						&""& serb &", "& minlager &")"
						
						'Response.write strSQLins
						'Response.flush
						
						oConn.execute(strSQLins)
					
					end if
					oRec.close
					
					
	else
	
	
	end if




end function	
	
	
	%>