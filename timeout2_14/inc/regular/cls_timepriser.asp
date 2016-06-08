 
 <%

 sub timepriser_akt
 %>

                <table cellspacing=1 cellpadding=2 border=0 width=100% bgcolor="#8CAAe6" id="timepristable">
				<tr bgcolor=Gainsboro><td class=lille bgcolor="#8caae6" style="width:100px;" valign=bottom >Medarbejder</td>
				<%if Not (func = "opretstam" OR func = "redstam") then%>
				<td class=lille>Nedarv <br />fra job * <br /><input type="button" class="checkAll" value="kol." style="width:30px;" /></td>
				<%end if%>
				<td class=lille valign=bottom>Timepris 1 <br /><input type="button" class="checkAll" value="Kolonne" /></td>
				<td class=lille valign=bottom>Timepris 2 <br /><input type="button" class="checkAll" value="Kolonne" /></td>
                <td class=lille valign=bottom>Timepris 3 <br /><input type="button" class="checkAll" value="Kolonne" /></td>
                <td class=lille valign=bottom>Timepris 4 <br /><input type="button" class="checkAll" value="Kolonne" /></td>
				<td class=lille valign=bottom>Timepris 5 <br /><input type="button" class="checkAll" value="Kolonne" />
                </td><td class=lille valign=bottom bgcolor=#DCF5BD>Valgt Timepris<br />
				(angiv eller vælg<br /> fra kolonne)<br /><input type="button" class="checkAll" value="kolonne" /></td>
                
                <%if lto = "xxxintranet - local" OR lto = "xxxwwf" then %>
                <td class=lille valign=bottom bgcolor="#FFFFFF">Ressource forecast <br />(timer)</td>
                
                <%end if %>
                </tr>
				
				<%
				'**** Hvis der redigeres / oprettes stamakt. ***
				if func = "opretstam" OR func = "redstam" then
				gp1 = 10
				else
					if len(gp1) <> 0 then
					gp1 = gp1
					else
					gp1 = 1
					end if
				end if	
				'****
				
					if len(gp2) <> 0 then
					gp2 = gp2
					else
					gp2 = 1
					end if
					
					if len(gp3) <> 0 then
					gp3 = gp3
					else
					gp3 = 1
					end if
					
					if len(gp4) <> 0 then
					gp4 = gp4
					else
					gp4 = 1
					end if
					
					if len(gp5) <> 0 then
					gp5 = gp5
					else
					gp5 = 1
					end if
					
					if len(gp6) <> 0 then
					gp6 = gp6
					else
					gp6 = 1
					end if
					
					if len(gp7) <> 0 then
					gp7 = gp7
					else
					gp7 = 1
					end if
					
					if len(gp8) <> 0 then
					gp8 = gp8
					else
					gp8 = 1
					end if
					
					if len(gp9) <> 0 then
					gp9 = gp9
					else
					gp9 = 1
					end if
					
					if len(gp10) <> 0 then
					gp10 = gp10
					else
					gp10 = 1
					end if
				
				usedmids = "0#"
				strSQL = "SELECT id, navn FROM projektgrupper WHERE id = "& gp1 &" OR id = "& gp2 &" OR id = "& gp3 &" OR id = "& gp4 &" OR id = "& gp5 &" OR id = "& gp6 &" OR id = "& gp7 &" OR id = "& gp8 &" OR id = "& gp9 &" OR id = "& gp10 &" ORDER BY navn"
				oRec.open strSQL, oConn, 0, 1
				
				while not oRec.EOF
					
					strSQL3 = "SELECT medarbejderid, projektgruppeid, mid, mnavn, mansat, timepris, "_
					&" timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
					&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, "_
					&" v0.valutakode AS valkode_0, "_
					&" v1.valutakode AS valkode_1, "_
					&" v2.valutakode AS valkode_2, "_
					&" v3.valutakode AS valkode_3, "_
					&" v4.valutakode AS valkode_4, "_
					&" v5.valutakode AS valkode_5 "_
					&" FROM progrupperelationer "_
					&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
					&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
					&" LEFT JOIN valutaer v0 ON (v0.id = tp0_valuta) "_
					&" LEFT JOIN valutaer v1 ON (v1.id = tp1_valuta) "_
					&" LEFT JOIN valutaer v2 ON (v2.id = tp2_valuta) "_
					&" LEFT JOIN valutaer v3 ON (v3.id = tp3_valuta) "_
					&" LEFT JOIN valutaer v4 ON (v4.id = tp4_valuta) "_
					&" LEFT JOIN valutaer v5 ON (v5.id = tp5_valuta) "_
					&" WHERE projektgruppeid = "& oRec("id") &" AND mnavn <> '' AND mansat <> 2 AND "& mtypeSQLKri &" ORDER BY mnavn"
					



					oRec5.open strSQL3, oConn, 0, 1
					'this6timepris = 0
					while not oRec5.EOF
					thissel = 0
						
						if instr(usedmids, "#"&oRec5("mid")&"#") = 0 then
						t = 0
						
						
						this6timepris = ""
						this6valuta = 1
						
                        '*** Henter eksisterende timepriser på aktiv hivs de er oprettet / ellers nedarv ***'

						if func <> "opretstam" then
						strSQL = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& tp_jobid &" AND aktid = "& tp_jobaktid &" AND medarbid =  "& oRec5("mid")
						'Response.write strSQL & "<br>"
						'Response.flush
                        oRec2.open strSQL, oConn, 3 
						    
                            if not oRec2.EOF then
						    thissel = oRec2("timeprisalt")
						    this6timepris = formatnumber(oRec2("6timepris"), 2)
						    this6valuta = oRec2("6valuta")
                            
						    end if 
						oRec2.close 
						end if
						
						
						if len(thissel) <> 0 then
						thissel = thissel
						else
						thissel = 0
						end if
						
						
						call meStamdata(oRec5("mid"))

                                        if oRec5("mansat") <> 1 then
                                            mTdBgCol = "#CCCCCC"
                                        else
                                            mTdBgCol = ""
                                        end if

						%>
						
						<tr bgcolor=#ffffff><td class=lille style="background-color:<%=mTdBgCol%>;"><%=left(meNavn, 10)%> (<%=meNr %>)
                        <%if len(trim(meInit)) <> 0 then %>
                        - <b><%=meInit %></b>
                        <%end if %>
                        </td>
						<%
						v = 6
						
						if func = "opretstam" OR func = "redstam" then
						start = 1
						else
						start = 0
						end if
						
						
						
						for t = start to v
							
                            select case t
							case 0 '** Spriges over ved stamakt. **'
							val = oRec5("timepris")
							valutaId = oRec5("tp0_valuta")
                            mtimer = 0
							case 6 
							val = this6timepris
							valutaId = this6valuta
                            mtimer = 0 'formatnumber(oRec5("mtimer"))
							case else
							val = formatnumber(oRec5("timepris_a"&t&""), 2)
							valutaId = oRec5("tp"&t&"_valuta")
							    
							    if func = "opretstam" AND t = 1 then
							    this6timepris = val
							    this6valuta = oRec5("tp1_valuta")
							    end if
							
							end select



							if func = "opretstam" then
								if cint(t) = 1 then
								chk = "CHECKED" 
								else
								chk = ""
								end if
							else
                                
                                if len(trim(this6timepris)) <> 0 then
                                this6timeprisTjk = formatnumber(this6timepris)
                                this6valutaTjk = this6valuta
                                else
                                this6timeprisTjk = 0
                                this6valutaTjk = 1
                                end if

                                if len(trim(val)) <> 0 then
                                valTjk = formatnumber(val)
                                valutaIdTjk = valutaId
                                else
                                valTjk = ""
                                valutaIdTjk = 1
                                end if

								if valTjk = this6timeprisTjk AND cint(valutaIdTjk) = cint(this6valutaTjk) then
								chk = "CHECKED" 
								    
								    'if t = 0 then '**** ?
								    
								    'end if
								    
								else
                                    if t = 0 AND len(trim(this6timepris)) = 0 then
                                    chk = "CHECKED"
                                    else
								    chk = ""
                                    end if
								end if
							end if
							
							
							
							%>              
                                            <%if t = 6 then%>
											<td class=lille style="white-space:nowrap; background-color:#DCF5BD;">
                                            <%else %>
                                            <td class=lille style="white-space:nowrap;">
                                            <%end if %>

											<%if t = 6 then%>
											<input type="text" name="FM_6timepris_<%=oRec5("mid")%>" id="FM_6timepris_<%=oRec5("mid")%>" value="<%=val%>" style="width:45px;">
											<%call valutakoder(600&oRec5("mid"), valutaId, 0) %>
                                            
                                            <%if lto = "xxintranet - local" OR lto = "xxxwwf" then %>
                                            </td><td class=lille align=right style="white-space:nowrap;">
                                            <input type="text" name="FM_mtimer_<%=oRec5("mid")%>" id="FM_mtimer_<%=oRec5("mid")%>" value="<%=mtimer%>" style="width:50px;">
                                            <%end if %>
											
                                            <%else%>

                                           
                                            <input type="radio" name="FM_timepris_<%=oRec5("mid")%>" id="FM_timepris_<%=oRec5("mid")%>" value="<%=t%>" <%=chk%> onclick="overfortp('<%=oRec5("mid")%>', '<%=t %>');">
											
                                          

											<% if t = 0 then
											hdval = ""
											else
											%>
											<%=formatnumber(val, 2) &" "& oRec5("valkode_"&t&"") %>
											<%
											hdval = formatnumber(val,2)
											end if%>
											
											
											<input type="hidden" id="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" value="<%=hdval%>">
                                            <input type="hidden" id="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" value="<%=valutaId%>">
											
											<%end if
											
											%></td><%
							
						next 
						%>
						</tr>
						<input type="hidden" name="FM_use_medarb_tpris" id="Hidden3" value="<%=oRec5("mid")%>">
						<%
						usedmids = usedmids &oRec5("mid")&"#"
						
						end if
					oRec5.movenext
					wend
					oRec5.close
					
				oRec.movenext
				wend
				oRec.close
				%>
				</table>


 
 <%
 end sub



 sub timepriser_job

    %>
    <table cellspacing=1 cellpadding=2 border=0 width=800 bgcolor="#8CAAe6" id="timepristable">
                                <tr bgcolor=Gainsboro>
                                <td bgcolor=#d6dff5 class=lille style="width:100px;">Medarbejder</td>
								<td class=lille>Generel timepris<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 1<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 2<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 3<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 4<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td class=lille>Timepris 5<br /><input type="button" class="checkAll" value="kolonne" /></td>
								<td bgcolor="#DCF5BD" class=lille>Valgt timepris<br /><input type="button" class="checkAll" value="kolonne" /></td></tr>
								<%
								usedmids = "0#"
								strSQL = "SELECT id, navn FROM projektgrupper WHERE id = "& gp1 &" OR id = "& gp2 &" OR id = "& gp3 &" OR id = "& gp4 &" OR id = "& gp5 &" OR id = "& gp6 &" OR id = "& gp7 &" OR id = "& gp8 &" OR id = "& gp9 &" OR id = "& gp10 &" ORDER BY navn"
								oRec.open strSQL, oConn, 0, 1
								
								'Response.write strSQL
								while not oRec.EOF
									
									strSQL3 = "SELECT medarbejderid, projektgruppeid, mid, mnavn, mansat, timepris, timepris_a1, "_
									&" timepris_a2, timepris_a3, timepris_a4, timepris_a5, "_
									&" tp0_valuta, tp1_valuta, tp2_valuta, tp3_valuta, tp4_valuta, tp5_valuta, "_
									&" v0.valutakode AS valkode_0, "_
									&" v1.valutakode AS valkode_1, "_
									&" v2.valutakode AS valkode_2, "_
									&" v3.valutakode AS valkode_3, "_
									&" v4.valutakode AS valkode_4, "_
									&" v5.valutakode AS valkode_5 "_
									&" FROM progrupperelationer "_
									&" LEFT JOIN medarbejdere ON (mid = progrupperelationer.medarbejderid) "_
									&" LEFT JOIN medarbejdertyper ON (medarbejdertyper.id = medarbejdertype) "_
									&" LEFT JOIN valutaer v0 ON (v0.id = tp0_valuta) "_
									&" LEFT JOIN valutaer v1 ON (v1.id = tp1_valuta) "_
									&" LEFT JOIN valutaer v2 ON (v2.id = tp2_valuta) "_
									&" LEFT JOIN valutaer v3 ON (v3.id = tp3_valuta) "_
									&" LEFT JOIN valutaer v4 ON (v4.id = tp4_valuta) "_
									&" LEFT JOIN valutaer v5 ON (v5.id = tp5_valuta) "_
									&" WHERE projektgruppeid = "& oRec("id") &" AND mnavn <> '' AND mansat <> 2 AND "& mtypeSQLKri &" ORDER BY mnavn"
									
									'Response.Write strSQL3
									'Response.flush
									
									oRec5.open strSQL3, oConn, 0, 1
									this6timepris = "-1"
									while not oRec5.EOF
									thissel = 0
									
										if instr(usedmids, "#"&oRec5("mid")&"#") = 0 then
										t = 0
										strSQL = "SELECT jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE jobid = "& id &" AND aktid = 0 AND medarbid =  "& oRec5("mid")
										oRec4.open strSQL, oConn, 3 
										
										if not oRec4.EOF then
										
										thissel = oRec4("timeprisalt")
										this6timepris = oRec4("6timepris")
										this6valuta = oRec4("6valuta")
										
										else
										
										thissel = 0
										this6timepris = "-1"
										this6valuta = 1
										
										end if 
										oRec4.close 
										
										
                                        call meStamdata(oRec5("mid"))

                                        if oRec5("mansat") <> 1 then
                                            mTdBgCol = "#CCCCCC"
                                        else
                                            mTdBgCol = ""
                                        end if
										%>
										
										<tr bgcolor=#ffffff><td class=lille style="background-color:<%=mTdBgCol%>;"><%=left(meNavn, 10)%> (<%=meNr %>)
                                        <%if len(trim(meInit)) <> 0 then %>
                                        - <b><%=meInit %></b>
                                        <%end if %>
										</td>
										<%
										v = 6
										for t = 0 to v
											
											
											
											if cint(thissel) = cint(t) then
											chk = "CHECKED" 
											else
											chk = ""
											end if
											
											
											
											select case t
											case 0
											val = oRec5("timepris")
											valutaId = oRec5("tp"&t&"_valuta")
											
											case 6 
											
											if this6timepris = "-1" OR thissel = 0 then 
											val = oRec5("timepris")
                                            valutaId = oRec5("tp0_valuta")
                                            else
											val = this6timepris
											valutaId = this6valuta
											end if
											
											case else
											val = oRec5("timepris_a"&t&"")
											valutaId = oRec5("tp"&t&"_valuta")
											end select
											
											%> 
											
											
											<%if t = 6 then%>
                                            <td class=lille style="background-color:#DCF5BD;">
											<input type="text" name="FM_6timepris_<%=oRec5("mid")%>" id="FM_6timepris_<%=oRec5("mid")%>" value="<%=formatnumber(val,2)%>" style="width:60px;">
											<%call valutakoder(600&oRec5("mid"), valutaId, 0) %>
											<%else%>
                                            <td class=lille><input type="radio" name="FM_timepris_<%=oRec5("mid")%>" id="FM_timepris_<%=oRec5("mid")%>_<%=t%>" value="<%=t%>" <%=chk%> onclick="overfortp('<%=oRec5("mid")%>', '<%=t %>');"><%=formatnumber(val, 2) & " " & oRec5("valkode_"&t&"") %>
											<input type="hidden" id="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_timepris_<%=oRec5("mid")%>_<%=t%>" value="<%=formatnumber(val,2)%>">
											<input type="hidden" id="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" name="FM_hd_valuta_<%=oRec5("mid")%>_<%=t%>" value="<%=valutaId%>">
											
											<%end if
											
											%></td><%
										next 
										%>
										</tr>
										<input type="hidden" name="FM_use_medarb_tpris" id="FM_use_medarb_tpris" value="<%=oRec5("mid")%>">
										<%
										usedmids = usedmids &oRec5("mid")&"#"
										this6timepris = "-1"
										end if
									oRec5.movenext
									wend
									oRec5.close
									
								oRec.movenext
								wend
								oRec.close

                                %></table>
<%								
 end sub



function opd_aktfasttp(jobid,opdekstp,valuta)

   
    'response.write "jobid " & jobid
    'response.end
    '***** Opdateret aktiviteter ****'
    strSQlaktupdatetp = "UPDATE aktiviteter SET brug_fasttp = 1, fasttp = aktbudget, fasttp_val = "& valuta &" WHERE job = "& jobid
    oConn.execute(strSQlaktupdatetp) 


    '**** Opdaterer eksisterende timereg. *****'
    strSQLa = "SELECT a.id AS aid, fasttp, fasttp_val, v.kurs AS kurs FROM aktiviteter AS a LEFT JOIN valutaer AS v ON (v.id = fasttp_val) WHERE job = "& jobid

    'if session("mid") = 1 then
    'response.write strSQLa & "<br><br>"
    'end if

    oRec6.open strSQLa, oConn, 3
    while not oRec6.EOF 
    
        fasttp = replace(oRec6("fasttp"), ",", ".")
        kurs = replace(oRec6("kurs"), ",", ".")

        strSQLtimer = "UPDATE timer SET timepris = "& fasttp &", valuta = "& oRec6("fasttp_val") &", kurs = "& kurs &" WHERE taktivitetid = "& oRec6("aid") 
        
        'if session("mid") = 1 then
        'response.write strSQLtimer
        'response.flush
        'end if    

        oConn.execute(strSQLtimer)    

    oRec6.movenext
    wend
    oRec6.close
      
    'if session("mid") = 1 then
    'response.end
    'end if


end function
   %>