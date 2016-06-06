<%
'*** Viser timeregistreringe for valgte uge ****'
       

        if print <> "j" then

        itop = 0
        ileft = 10
        iwdt = 340

        call sideinfo(itop,ileft,iwdt)

        %>
        <b><%=tsa_txt_061 %></b><br>
         <%=tsa_txt_062 %><br />
          <img src="../ill/pile_left.gif" /><br /><br />
        <%=tsa_txt_063 %>
        <br />
         <img src="../ill/pile_left.gif" />
         <!--- side info -->
        </td></tr>
        </table>
        </div>

        <%end if %>
        
        
      
	</div><!-- sidediv -->
        
        
        <%
	
	oimg = "ikon_timereg_48.png"
	oleft = 15
	otop =  100
	owdt = 500
	oskrift = tsa_txt_064 &" "& datepart("ww", tjekdag(1), 2 ,2) & " "& datepart("yyyy", tjekdag(1), 2 ,3) 
	
	call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	
	
            
            if print <> "j" then
            tTop = 150
	        tLeft = 15
	        tWdth = 700
	        else
	        tTop = 100
	        tLeft = 40
	        tWdth = 700
	        end if
        	
        	
	        call tableDiv(tTop,tLeft,tWdth)
            %>
            <table cellpadding=2 cellspacing=1 border=0 id="inputTable" background-color="#d6dff5" width=100%>
	        <%if print <> "j" then %>
	        <tr>
	            <td colspan=4><a href="timereg_akt_2006.asp?showakt=1&fromsdsk=<%=fromsdsk%>" class=vmenu><%=tsa_txt_258%>..</td>
	        </tr>
	        <%end if%>
	        
	        <tr>
		        <td colspan=4 valign=top style="padding-top:5px;">
        		
        		
        		
        	
	        <%
	        varTjDatoUS_man = convertDateYMD(tjekdag(1))
	        varTjDatoUS_tir = convertDateYMD(tjekdag(2))
	        varTjDatoUS_ons = convertDateYMD(tjekdag(3)) 
	        varTjDatoUS_tor = convertDateYMD(tjekdag(4)) 
	        varTjDatoUS_fre = convertDateYMD(tjekdag(5))
	        varTjDatoUS_lor = convertDateYMD(tjekdag(6))
	        varTjDatoUS_son = convertDateYMD(tjekdag(7))
	        %>
        	
        	
        	
        	
	        <%
	        'vzb0 = "visible"
	        'dsp0 = ""
        	
	        '*** Aktiviteter og Timer SQL MAIN **** 
	        strSQL = "SELECT t.tid, t.taktivitetnavn, t.timer, t.tdato, t.timerkom, t.offentlig, tjobnavn, tjobnr, "_
	        &" k.kkundenavn, k.kkundenr, t.godkendtstatus, "_
	        &" t.tmnavn, j.id AS jid, t.tfaktim, t.timepris, t.valuta, v.valutakode "_
	        &" FROM timer t "_
	        &" LEFT JOIN job j ON j.jobnr = t.tjobnr "_
	        &" LEFT JOIN kunder k ON k.kid = j.jobknr "_
	        &" LEFT JOIN valutaer v ON (v.id = t.valuta)"_
	        &" WHERE "_
	        &" (t.tmnr = "& usemrn &" AND (Tdato = '"& varTjDatoUS_son &"'"_
	        &" OR Tdato = '"& varTjDatoUS_man &"'"_
	        &" OR Tdato = '"& varTjDatoUS_tir &"'"_
	        &" OR Tdato = '"& varTjDatoUS_ons &"'"_
	        &" OR Tdato = '"& varTjDatoUS_tor &"'"_
	        &" OR Tdato = '"& varTjDatoUS_fre &"'"_
	        &" OR Tdato = '"& varTjDatoUS_lor &"'))"_
	        &" ORDER BY Tdato, t.tjobnavn, t.taktivitetnavn"
        	
	        'tfaktim <> 5 AND

	        'Response.write strSQL
	        'Response.flush
	        lastdivval = 0
	        timerLastDato = 0
	        lastDato = ""
	        x = 0
	        oRec.open strSQL, oConn, 3
	        while not oRec.EOF
        		
		        if lastDato <> oRec("tdato") then%>
			        <%if x > 0 then%>
			        <tr>
				        <td colspan=5 align=right><%=tsa_txt_065 %>: <b><%=formatnumber(timerLastDato, 2)%></b></td>
			        </tr>
			        </table>
			        <!--</div>-->
			        <%
			        timerugeTot = timerugeTot + timerLastDato
			        timerLastDato = 0
			        end if
        		
        		
		       
		        %>
        		
		        <!--<div name="tr_<%=weekday(oRec("tdato"))%>" id="tr_<%=weekday(oRec("tdato"))%>" style="position:absolute; left:0px; top:73px; display:<%=dsp%>; visibility:<%=vzb%>; background-color:#d6dff5; width:700;">-->
		        <br /><br />
        		
        		
        	
        		
        		
		        <table cellpadding=4 cellspacing=0 border=0 width=100%>
        		
		        <%
		        diviswrt = diviswrt & ",#"& weekday(oRec("tdato")) &"#"
		        end if
        		
		        if lastDato <> oRec("tdato") then%>
		        <tr>
				        <td colspan=5><b><%=weekdayname(weekday(oRec("tdato"))) &"&nbsp;"& formatdatetime(oRec("tdato"), 1)%></b></td>
		        </tr>
		        <tr bgcolor="#D6Dff5">
			        <td style="border-top:1px #8caae6 solid; border-left:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_066 %></b></td>
			        <td style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_067 %></b></td>
			        <td style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_068 %></b> (<%=tsa_txt_069 %>)</td>
			        <td align=right style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid;"><b><%=tsa_txt_070 %></b></td>
			        <td align=right style="border-top:1px #8caae6 solid; border-bottom:1px #8caae6 solid; border-right:1px #8caae6 solid;"><b><%=tsa_txt_186 %></b></td>
			    </tr>
		        <%
        		
	            timerugeTot = timerugeTot + timerLastDato
	            'Response.Write timerugeTot & "<br>"
        		
		        end if%>
        		
		        <tr bgcolor="#ffffff">
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;">
		        <%if len(oRec("kkundenavn")) > 25 then%>
		        <%=left(oRec("kkundenavn"), 25)%>..
		        <%else%>
		        <%=oRec("kkundenavn")%>
		        <%end if%> 
		        (<%=oRec("kkundenr")%>)</td>
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;">
		        
		        <%if print <> "j" then%>
		        <a href="jobs.asp?menu=job&func=red&id=<%=oRec("jid")%>&int=1&rdir=treg" class=vmenu target="_top">
		        <%end if %>
		        
		        <%if len(oRec("tjobnavn")) > 25 then%>
		        <%=left(oRec("tjobnavn"), 25)%>..
		        <%else%>
		        <%=oRec("tjobnavn")%>
		        <%end if%> 
	 	        (<%=oRec("tjobnr")%>)
	 	        
	 	        <%if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
	 	        </a>
	 	        <%end if%>
	 	        
	 	        </td>
		        <td valign=top style="border-bottom:1px #c4c4c4 dashed;"><b><%=oRec("taktivitetnavn")%></b>
        		
		        <%
		        call akttyper(oRec("tfaktim"), 4) 
		        Response.Write "&nbsp;<font class=megetlillesort>("& akttypenavn &")"%></td>
		        <td align=right valign=top style="border-bottom:1px #c4c4c4 dashed;">
        		
		        <%
		        '******* Skal det være muligt at redigere indtastning ****
		        '*** Lastfakdato
		        lastfakdato = "1/1/2001"
        		
		        strSQLFAK = "SELECT f.fakdato FROM fakturaer f WHERE f.jobid = "& oRec("jid") &" AND f.fakdato >= '"& varTjDatoUS_man &"' AND faktype = 0 ORDER BY f.fakdato DESC"
		        oRec2.open strSQLFAK, oConn, 3
		        if not oRec2.EOF then
			        if len(trim(oRec2("fakdato"))) <> 0 then
			        lastfakdato = oRec2("fakdato")
			        end if
		        end if
		        oRec2.close
        		
		        call fakfarver(lastfakdato, tjekdag(1), oRec("tdato"))
        		
		        'if maxl <> 0 then
		        if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
		        <!--a href="rediger_tastede_dage.asp?id=<=oRec("Tid")%>&medarb=<=oRec("Tmnavn")%>&jobnr=<=intJobnr%>&eks=<=request("eks")%>&lastFakdag=<=lastFakdag%>&selmedarb=<=selmedarb%>&selaktid=<=selaktid%>&FM_job=<=request("FM_job")%>&FM_medarb=<=request("FM_medarb")%>&FM_start_dag=<=strDag%>&FM_start_mrd=%=strMrd%>&FM_start_aar=<=strAar%>&FM_slut_dag=<=strDag_slut%>&FM_slut_mrd=<=strMrd_slut%>&FM_slut_aar=<=strAar_slut%>" class="vmenu">-->
		        <a href="javascript:popUp('rediger_tastede_dage_2006.asp?id=<%=oRec("Tid")%>','600','500','250','120');" target="_self" class=vmenu>
		        <%end if%>
        		
		        <%=formatnumber(oRec("timer"), 2)%>
        		
		        <%if oRec("godkendtstatus") = 0 AND maxl <> 0 AND print <> "j" then%>
		        </a>
		        <%end if%></td>
		        <td style="border-bottom:1px #c4c4c4 dashed;" align=right><%=formatnumber(oRec("timepris"), 2) &" "& oRec("valutakode") %></td>
        		
	        </tr>
	        <%
	        x = x + 1
	        select case oRec("tfaktim")
	        case 1,2,6,13,14,20,21
	        timerLastDato = timerLastDato + oRec("timer")
            case else
	        timerLastDato = timerLastDato
	        end select
        	
	        lastDato = oRec("tdato") 
        	
	        oRec.movenext
	        wend
        	
	        oRec.close

	        timerugeTot = timerugeTot + timerLastDato
        	
	        %>
        	
	        <%if x > 0 then%>
	        <tr>
		        <td colspan=5 align=right><%=tsa_txt_065 %>: <b><%=formatnumber(timerLastDato, 2)%></b></td>
	        </tr>
	        <tr bgcolor="#FFDFDF">
		        <td colspan=5 align=right><%=tsa_txt_065 %>&nbsp;<%=tsa_txt_071 %>&nbsp;<%= datepart("ww", tjekdag(1), 2,2)%>, <%= datepart("yyyy", tjekdag(1), 2 ,3)%>: <b><%=formatnumber(timerugeTot, 2)%></b></td>
	        </tr>
	        </table>
	        
	        <%if print <> "j" then %>
	        <br />
	        <a href="timereg_akt_2006.asp?print=j" target="_blank" class=vmenu><%=tsa_txt_072 %>..</a> 
	        <%end if%>
        	
	        <!--</div>-->
	        <%else %>
	        <tr>
		        <td colspan=5 style="padding-top:20px;"><b><%=tsa_txt_124 %></b></td>
	        </tr>
	        </table>
	        <%end if%>
        	
        	
	        </td>
	        </tr>
        	</table>
        	
        	
        	
	        <%
	        timerLastDato = 0
	        %>
	        </div>
	        
	        <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                &nbsp;
        	
	        <form><input type="hidden" name="lastdivid" id="lastdivid" value="<%=lastdivval%>"></form>
        	
	        <%
        	
         	
	        '*** Smiley ***
	        useYear = year(now)
	        %>
	        <div name="smileydiv" id="smileydiv" style="position:absolute; left:60px; top:60px; display:none; visibility:hidden; background-color:#8caae6; padding:5px; border:1px #003399 solid; border-right:2px #003399 solid; border-bottom:2px #003399 solid;">
	        <a href="#" class=red onClick="gemSmileystatus();">[x]</a><br>
	        <%
	        call smileystatus(usemrn)
	        %>
	        </div>
        	
	
