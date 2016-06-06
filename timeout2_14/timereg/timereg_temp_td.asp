


<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" id="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'man'), tjektimer('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%else%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%end if%>
				<%call showcommentfunc(2)%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>















<td align="center">
				<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" id="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'man'), tjektimer('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%else%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%end if%>
				<%call showcommentfunc(2)%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>

				
				
				
				
				<td align="center">
				<%
				strMandag = ""
				strMandag = strMandag & "<input type='hidden' name='FM_man_opr_"& iRowLoop &"' id='FM_man_opr_"& iRowLoop &"' value='"&manTimerVal(manRLoop)&"'>"
				if aTable1Values(16,iRowLoop) <> 2 then
				strMandag = strMandag & "<input type='text' name='Timer_man_"&iRowLoop&"' id='Timer_man_"&iRowLoop&"' value='"&SQLBless2(manTimerVal(manRLoop))&"' size='3' maxlength='"&maxl_man&"' style='background-color: "&fmbg_man&"' onKeyUp='setTimerTot("&iRowLoop&", 'man'), tjektimer('man',"&iRowLoop&");'>"
				else
				strMandag = strMandag & "<input type='text' name='Timer_man_"&iRowLoop&"' id='Timer_man_"&iRowLoop&"' value='"&SQLBless2(manTimerVal(manRLoop))&"' size='3' maxlength='"&maxl_man&"' style='background-color: "&fmbg_man&"' onkeyup='tjekkm('man',"&iRowLoop&");'>"
				end if
				
				Response.write strMandag
				call showcommentfunc(2)
				%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				
				
				
				
				
				
				
				
				
				
				<!--- ---->
				
				
				<td align="center">
				<input type="hidden" name="FM_man_opr_<%=iRowLoop%>" id="FM_man_opr_<%=iRowLoop%>" value="<%=manTimerVal(manRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'man'), tjektimer('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%else%>
				<input type="Text" name="Timer_man_<%=iRowLoop%>" id="Timer_man_<%=iRowLoop%>" maxlength="<%=maxl_man%>" value="<%=SQLBless2(manTimerVal(manRLoop))%>" onkeyup="tjekkm('man',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_man%>">
				<%end if%>
				<%call showcommentfunc(2)%>
				<a href="#" onclick="expandkomm('man',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_tir_opr_<%=iRowLoop%>" id="FM_tir_opr_<%=iRowLoop%>" value="<%=tirTimerVal(tirRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tir'), tjektimer('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%else%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="tjekkm('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%end if%>
				<%call showcommentfunc(3)%>
				<a href="#" onclick="expandkomm('tir',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_ons_opr_<%=iRowLoop%>" id="FM_ons_opr_<%=iRowLoop%>" value="<%=onsTimerVal(onsRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'ons'), tjektimer('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%else%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="tjekkm('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%end if%>
				<%call showcommentfunc(4)%>
				<a href="#" onclick="expandkomm('ons',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_tor_opr_<%=iRowLoop%>" id="FM_tor_opr_<%=iRowLoop%>" value="<%=torTimerVal(torRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tor'), tjektimer('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%else%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="tjekkm('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%end if%>
				<%call showcommentfunc(5)%>
				<a href="#" onclick="expandkomm('tor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_fre_opr_<%=iRowLoop%>" id="FM_fre_opr_<%=iRowLoop%>" value="<%=freTimerVal(freRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'fre'), tjektimer('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%else%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="tjekkm('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%end if%>
				<%call showcommentfunc(6)%>
				<a href="#" onclick="expandkomm('fre',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_lor_opr_<%=iRowLoop%>" id="FM_lor_opr_<%=iRowLoop%>" value="<%=lorTimerVal(lorRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'lor'), tjektimer('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%else%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="tjekkm('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%end if%>
				<%call showcommentfunc(7)%>
				<a href="#" onclick="expandkomm('lor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				<td align="center">
				<input type="hidden" name="FM_son_opr_<%=iRowLoop%>" id="FM_son_opr_<%=iRowLoop%>" value="<%=sonTimerVal(sonRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then ' 2 = Kørsel (2 i akt.fakbar 5 i timer.tfaktim)/ alm aktivitet %>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'son'), tjektimer('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%else%>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="tjekkm('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%end if%>
				<%call showcommentfunc(1)%>
				<a href="#" onclick="expandkomm('son',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				<---->
				<td align="center">
				<input type="hidden" name="FM_tir_opr_<%=iRowLoop%>" id="FM_tir_opr_<%=iRowLoop%>" value="<%=tirTimerVal(tirRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tir'), tjektimer('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%else%>
				<input type="Text" name="Timer_tir_<%=iRowLoop%>" id="Timer_tir_<%=iRowLoop%>" maxlength="<%=maxl_tir%>" value="<%=SQLBless2(tirTimerVal(tirRLoop))%>" onkeyup="tjekkm('tir',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tir%>">
				<%end if%>
				<%call showcommentfunc(3)%>
				<a href="#" onclick="expandkomm('tir',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_ons_opr_<%=iRowLoop%>" id="FM_ons_opr_<%=iRowLoop%>" value="<%=onsTimerVal(onsRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'ons'), tjektimer('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%else%>
				<input type="Text" name="Timer_ons_<%=iRowLoop%>" id="Timer_ons_<%=iRowLoop%>" maxlength="<%=maxl_ons%>" value="<%=SQLBless2(onsTimerVal(onsRLoop))%>" onkeyup="tjekkm('ons',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_ons%>">
				<%end if%>
				<%call showcommentfunc(4)%>
				<a href="#" onclick="expandkomm('ons',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_tor_opr_<%=iRowLoop%>" id="FM_tor_opr_<%=iRowLoop%>" value="<%=torTimerVal(torRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'tor'), tjektimer('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%else%>
				<input type="Text" name="Timer_tor_<%=iRowLoop%>" id="Timer_tor_<%=iRowLoop%>" maxlength="<%=maxl_tor%>" value="<%=SQLBless2(torTimerVal(torRLoop))%>" onkeyup="tjekkm('tor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_tor%>">
				<%end if%>
				<%call showcommentfunc(5)%>
				<a href="#" onclick="expandkomm('tor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_fre_opr_<%=iRowLoop%>" id="FM_fre_opr_<%=iRowLoop%>" value="<%=freTimerVal(freRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'fre'), tjektimer('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%else%>
				<input type="Text" name="Timer_fre_<%=iRowLoop%>" id="Timer_fre_<%=iRowLoop%>" maxlength="<%=maxl_fre%>" value="<%=SQLBless2(freTimerVal(freRLoop))%>" onkeyup="tjekkm('fre',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_fre%>">
				<%end if%>
				<%call showcommentfunc(6)%>
				<a href="#" onclick="expandkomm('fre',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				<td align="center">
				<input type="hidden" name="FM_lor_opr_<%=iRowLoop%>" id="FM_lor_opr_<%=iRowLoop%>" value="<%=lorTimerVal(lorRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'lor'), tjektimer('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%else%>
				<input type="Text" name="Timer_lor_<%=iRowLoop%>" id="Timer_lor_<%=iRowLoop%>" maxlength="<%=maxl_lor%>" value="<%=SQLBless2(lorTimerVal(lorRLoop))%>" onkeyup="tjekkm('lor',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_lor%>">
				<%end if%>
				<%call showcommentfunc(7)%>
				<a href="#" onclick="expandkomm('lor',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				
				<td align="center">
				<input type="hidden" name="FM_son_opr_<%=iRowLoop%>" id="FM_son_opr_<%=iRowLoop%>" value="<%=sonTimerVal(sonRLoop)%>">
				<%if aTable1Values(16,iRowLoop) <> 2 then ' 2 = Kørsel (2 i akt.fakbar 5 i timer.tfaktim)/ alm aktivitet %>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="setTimerTot(<%=iRowLoop%>, 'son'), tjektimer('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%else%>
				<input type="Text" name="Timer_son_<%=iRowLoop%>" id="Timer_son_<%=iRowLoop%>" maxlength="<%=maxl_son%>" value="<%=SQLBless2(sonTimerVal(sonRLoop))%>" onkeyup="tjekkm('son',<%=iRowLoop%>);" size=3 style="background-color: <%=fmbg_son%>">
				<%end if%>
				<%call showcommentfunc(1)%>
				<a href="#" onclick="expandkomm('son',<%=iRowLoop%>);"><%=kommtrue%></a>
				</td>
				