 
 <%

public valutaKode_CCC 
function valutakode_fn(valid)

     valutaKode_CCC = ""
      strSQL3 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer WHERE id = "& valid
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    if not oRec3.EOF then 

            valutaKode_CCC = oRec3("valutaKode")
            
            end if
            oRec3.close


end function






 function valutakoder(i, valuta, layout)
	
     select case cint(layout) 
     case 0
     cssThis = "width:55px; font-size:9px;" 
     cls = ""
     case 1
     cssThis = ""
     cls = "form-control input-small"
     case else
     cssThis = ""
     cls = ""
     end select

     if instr(i, "freight_price_pc_valuta") <> 0 then
     dsabl = "disabled"
     else
     dsabl = ""
     end if
     
     %>
	<select name="FM_valuta_<%=i %>" id="FM_valuta_<%=i %>" style="<%=cssThis%>" class="<%=cls%>" <%=dsabl %>>
		    
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = oRec3("id") then
		    valGrpCHK = "SELECTED"
		    else
		    valGrpCHK = ""
		    end if
		    
                if cint(layout) = 0 then
                kursTxt = ""
                else
                    if oRec3("kurs") <> 100 then
                    kursTxt = "("& oRec3("kurs")/100 &")"
                    else
                    kursTxt = ""
                    end if
                end if
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode") &" "& kursTxt%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
	<%
	end function
	
	
	public dblKurs
	function valutaKurs(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       dblKurs = replace(oRec("kurs"), ",", ".")
       end if 
       oRec.close
	end function


    public dblKursOR3
	function valutaKursOR3(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKursOR3 = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec3.open strSQL, oConn, 3
       if not oRec3.EOF then
       dblKursOR3 = oRec3("kurs")
       end if 
       oRec3.close
	end function

    public basisValId, basisValISO, basisValKurs, basisValISO_f8
	function basisValutaFN()
	    '**** Finder basisValuta ***'
       dblKurs = 100
       
        
       strSQL = "SELECT id, valutakode, kurs FROM valutaer WHERE grundvaluta = 1"
       
       'oConn.execute(strSQL) 
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       basisValKurs = replace(oRec("kurs"), ",", ".")

       basisValId = oRec("id")
       basisValISO = oRec("valutakode")
       basisValISO_f8 = "<span style='font-size:7px;'>"& basisValISO &"</span>"

       end if 
       oRec.close
	end function
	
	
	public valBelobBeregnet
	function beregnValuta(belob,frakurs,tilkurs)
	        
        'Response.write lastFaknr & " belob: "& belob & "<br>"
        'Response.flush

        if isNull(belob) <> true AND belob <> "-" AND isNull(frakurs) <> true AND len(trim(belob)) <> 0 then
        valBelobBeregnet = belob * (frakurs/tilkurs)
        else
        valBelobBeregnet = 0
        end if
    
    if len(valBelobBeregnet) <> 0 then
    valBelobBeregnet = valBelobBeregnet
    else
    valBelobBeregnet = 0
    end if
    
    
    
    valBelobBeregnet = valBelobBeregnet/1
    'Response.Write valBelobBeregnet & "<br>"
    'Response.flush
	end function

    %>