  

  <table cellspacing=0 cellpadding=2 border=0 width=100%>
  <%if (lto <> "cst" AND lto <> "tec") AND instr(akttype_sel, "#13#") <> 0 then 

    '** Ferie fri
    if normTimerDag(x) <> 0 then
	 fefriVal = fefriTimer(x)/normTimerDag(x)
	 else
	 fefriVal = 0
	 end if
	 

     if normTimerDag(x) <> 0 then
	 fefriplVal = fefriplTimer(x)/normTimerDag(x)
	 else
	 fefriplVal = 0
	 end if

     if normTimerDag(x) <> 0 then
	 fefriBrVal = fefriTimerBr(x)/normTimerDag(x)
	 else
	 fefriBrVal = 0
	 end if

     if normTimerDag(x) <> 0 then
	 fefriUdbVal = fefriTimerUdb(x)/normTimerDag(x)
	 else
	 fefriUdbVal = 0
	 end if
	  

	   fefriBal = 0
	 fefriBal  = (fefriTimer(x) - (fefriTimerBr(x) + fefriTimerUdb(x)))
	 
	 if normTimerDag(x) <> 0 then
	 fefriBalVal = fefriBal/normTimerDag(x)
	 else
	 fefriBalVal = 0
	 end if


     


  
   if media <> "export" then%>
	   <!-- Ferie fridage -->

    <tr>
    <td colspan=6>
	 <br /><br /><span style="border-bottom:2px #FFFF99 solid; padding:2px;"><b>Feriefridage</b> (<%=ferieFriaarStart%>)</span><br /><br />&nbsp;
	 </td>
     </tr>
	  <tr>
	 
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_174 &" "& tsa_txt_164%></b><br />~ dage</td>
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b>Planlagt <br> >> dd.</b><br />~ dage</td>
	 
	  <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_165%></b><br />~ dage</td>
        <td style="border-bottom:1px silver dashed;">&nbsp;</td>
	   <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b><br />~ dage</td>
	  <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_282 &" "& tsa_txt_280 %></b><br />~ dage</td>
    
	 </tr>
	 
     
    
     
     <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%=formatnumber(fefriVal,2) %></td>
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 <%=formatnumber(fefriplVal,2) %> </td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 <%=formatnumber(fefriBrVal,2) %> 
     </td>
     <td style="border-bottom:1px silver dashed;">&nbsp;</td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 <%=formatnumber(fefriUdbVal,2) %></td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;"><b><%=formatnumber(fefriBalVal,2) %><br />
	 </td>
     
	 </tr>
	

      <%
      
      else
      
      ekspTxtA = formatnumber(fefriVal,2)&";"&formatnumber(fefriplVal,2)&";"&formatnumber(fefriBrVal,2)&";;"&formatnumber(fefriUdbVal,2)&";"&formatnumber(fefriBalVal,2)&";"

      end if %>

	 <%end if 'cst%>
	 
	 
	 <!-- Ferie -->
	 
     <%if instr(akttype_sel, "#14#") <> 0 then 


     '**** Ferie
     if normTimerDag(x) <> 0 then
	 ferieOptjVal = ferieOptjtimer(x)/normTimerDag(x)
	 else
	 ferieOptjVal = 0
	 end if

     if instr(akttype_sel, "#111#") <> 0 then 
     if normTimerDag(x) <> 0 then
	 ferieOverfortVal = ferieOptjOverforttimer(x)/normTimerDag(x)
	 else
	 ferieOverfortVal = 0
	 end if
     end if
	 
     if instr(akttype_sel, "#112#") <> 0 then 
     if normTimerDag(x) <> 0 then
	 ferieOptjValuLon = ferieOptjUlontimer(x)/normTimerDag(x)
	 else
	 ferieOptjValuLon = 0
	 end if
     end if
      

     if normTimerDag(x) <> 0 then
	 feriePlVal = feriePLTimer(x)/normTimerDag(x)
	 else
	 feriePlVal = 0
	 end if


     if normTimerDag(x) <> 0 then
	 ferieAfVal = ferieAFTimer(x)/normTimerDag(x)
	 else
	 ferieAfVal = 0
	 end if

     if normTimerDag(x) <> 0 then
	 ferieAfulonVal = ferieAFulonTimer(x)/normTimerDag(x)
	 else
	 ferieAfulonVal = 0
	 end if
     

     if normTimerDag(x) <> 0 then
	 ferieUdbVal = ferieUdbTimer(x)/normTimerDag(x)
	 else
	 ferieUdbVal = 0
	 end if


     ferieBal = 0
	 call ferieBal_fn(ferieOptjtimer(x), ferieOptjOverforttimer(x), ferieOptjUlontimer(x), ferieAFTimer(x), ferieAFulonTimer(x), ferieUdbTimer(x))
	 
	 if normTimerDag(x) <> 0 then
	 ferieBalVal = ferieBal/normTimerDag(x)
	 else
	 ferieBalVal = 0
	 end if

                       select case lto
                       case "akelius", "intranet - local"
                       ferieAarTxt = "1.1."& ferieaarStart &" - 31.12."& ferieaarSlut
                       case else
                       ferieAarTxt = "1.5."& ferieaarStart &" - 30.4."& ferieaarSlut
                       end select

     
      if media <> "export" then%>
	  
      <tr><td colspan=6>
        <br /><br /><span style="border-bottom:2px #6CAE1C solid; padding:2px;"><b><%=tsa_txt_281 %></b> (<%=ferieAarTxt %>)</span> <br /><br />&nbsp;
	</td></tr>
	  
	 <tr>
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_152%> Opt.</b><br />~ dage</td>
	 
     <% if instr(akttype_sel, "#111#") <> 0 then  %>
     <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=global_txt_175%></b><br />~ dage</td>
     <%end if %>

     <% if instr(akttype_sel, "#112#") <> 0 then  %>
     <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=global_txt_176%></b><br />~ dage</td>
     <%end if %>

     <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b><%=tsa_txt_317%><br> >> dd.</b><br />~ dage</td>
	 <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b>Afholdt</b><br />~ dage</td>
	   <td valign=bottom align=right class=lille style="border-bottom:1px silver dashed;"><b>Afholdt u. løn</b><br />~ dage</td>
	   <td valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b>Udbetalt</b><br />~ dage</td>
	  <td  valign=bottom align=right style="border-bottom:1px silver dashed;" class=lille><b><%=tsa_txt_281 &" "& tsa_txt_280 %></b><br />~ dage</td>
	 </tr>
	  <tr>
	
	 <td align=right style="border-bottom:1px silver dashed;" class=lille>
	 
	
	 
	 <%=formatnumber(ferieOptjVal,2) %>
	 </td>

     <%
     if instr(akttype_sel, "#111#") <> 0 then   %>
     
     <td align=right class=lille style="border-bottom:1px silver dashed;">
     <%= ferieOptjOverforttimer(x) %>
     </td>
     <%
     end if

     if instr(akttype_sel, "#112#") <> 0 then   %>
     <td align=right class=lille style="border-bottom:1px silver dashed;">
     <%=ferieOptjUlontimer(x) %>
     </td>
     <%end if %>

	 <td align=right class=lille style="border-bottom:1px silver dashed;">

	 <%=formatnumber(feriePlVal,2) %>
	</td>
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	
	 <%=formatnumber(ferieAfVal,2) %>
	</td>
	 
	  <td align=right class=lille style="border-bottom:1px silver dashed;">
	
	 <%=formatnumber(ferieAfulonVal,2) %>
	 </td>
	 
	 <td align=right class=lille style="border-bottom:1px silver dashed;">
	 <%=formatnumber(ferieUdbVal,2) %>
	</td>
	  
	
	  <td align=right class=lille style="border-bottom:1px silver dashed;"><%=formatnumber(ferieBalVal,2) %><br />
	</td>
	 </tr>
	 </table>
	 
	 <% 
     else
       ekspTxtB = formatnumber(ferieOptjVal,2) &";" 

       if instr(akttype_sel, "#111#") <> 0 then  
       ekspTxtB = ekspTxtB & formatnumber(ferieOverfortVal,2) &";"
       end if

       if instr(akttype_sel, "#112#") <> 0 then
       ekspTxtB = ekspTxtB & formatnumber(ferieOptjValuLon,2) &";"
       end if

       ekspTxtB = ekspTxtB & formatnumber(feriePlVal,2)&";"&formatnumber(ferieAfVal,2)&";"&formatnumber(ferieAfulonVal,2)&";"&formatnumber(ferieUdbVal,2)&";"&formatnumber(ferieBalVal,2)&";"
     end if'media
     
     end if 'instr 14
     %>