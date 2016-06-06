<%
    
     function ktyper(name, wdt, sel, autosubmit, lto, level)

        select case sel
        case "-1"
        useasfakCHK01 = "SELECTED"
        case "0"
        useasfakCHK0 = "SELECTED"
        case "1"
        useasfakCHK1 = "SELECTED"
        case "2"
        useasfakCHK2 = "SELECTED"
        case "3"
        useasfakCHK3 = "SELECTED"
        case "4"
        useasfakCHK4 = "SELECTED"
        case "5"
        useasfakCHK5 = "SELECTED"
        case "6"
        useasfakCHK6 = "SELECTED"
        case "7"
        useasfakCHK7 = "SELECTED"
        case else
        useasfakCHK0 = "SELECTED"
        end select


        if cint(autosubmit) = 1 then
        %>
        <select name="FM_useasfak" class="form-control input-small" onchange="submit();">
        <option value=-1 <%=useasfakCHK01 %>>Alle</option>
        <%else %>
        <select name="FM_useasfak" class="form-control input-small">
        <%end if %>

        <%select case lto 
          case "epi", "epi_no", "epi_uk", "epi_ab"
            
            
         if level = 1 then %>

        <option value=1 <%=useasfakCHK1 %>>1. Licensindehaver (forvalgt på fakturaer)</option>
        <option value=2 <%=useasfakCHK2 %>>2. Datterselskab (kan vælges på faktura)</option>
        <option value=3 <%=useasfakCHK3 %>>3. Andet</option>
        <option value=0 <%=useasfakCHK0 %>>4. Kunde</option>
        <option value=5 <%=useasfakCHK5 %>>5. CRM-relation</option>
        <option value=6 <%=useasfakCHK6 %>>6. Leverandør</option>
        <option value=7 <%=useasfakCHK7 %>>7. Underleverandør</option>

        <%else %>

            <option value=5 <%=useasfakCHK5 %>>5. CRM-relation / Ny kunde</option>

            <%end if %>
        
            
           <%case else%>

        
        <option value=1 <%=useasfakCHK1 %>>1. Licensindehaver (forvalgt på fakturaer)</option>
        <option value=2 <%=useasfakCHK2 %>>2. Datterselskab (kan vælges på faktura)</option>
        <option value=3 <%=useasfakCHK3 %>>3. Andet</option>
        <option value=0 <%=useasfakCHK0 %>>4. Kunde</option>
        <option value=5 <%=useasfakCHK5 %>>5. CRM-relation</option>
        <option value=6 <%=useasfakCHK6 %>>6. Leverandør</option>
        <option value=7 <%=useasfakCHK7 %>>7. Underleverandør</option>

        <%end select %>

            </select>
        
        
        <%
      end function   



    '******************* FASTE Variable fra SEARCH mv ************************



  

    '*** Sætter lokal dato/kr format.
    Session.LCID = 1030

    func = request("func")
    
    if len(trim(request("id"))) <> 0 then
    id = request("id")
    else
    id = 0
    end if    

    if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
    end if
    
    menu = request("menu")

    sort = Request("sort")

    if len(trim(request("visikkekunder"))) <> 0 then
    visikkekunder = request("visikkekunder")
    else
    visikkekunder = 0
    end if


    if len(request("hidemenu")) <> 0 then
	hidemenu = request("hidemenu")
	else
	hidemenu = 0
	end if
	
	if len(request("rdir")) <> 0 then
	rdir = request("rdir")
	else
	rdir = 0
	end if
	
	if len(request("visikkekunder")) <> 0 then
	visikkekunder = 1
	else
	visikkekunder = 0
	end if
	
	ketype = request("ketype")
	
	if ketype = "e" then
	ketype = ketype
	else
	ketype = "k"
	end if
	
    if len(trim(request("sogsubmitted"))) <> 0 then
    sogsubmitted = 1
    else
    sogsubmitted = 0
    end if

   
	if cint(sogsubmitted) = 1 then
    ''**AND len(request("ktype")) <> 0 
	'** request("ktype") tjekker at form e submitted **'
	thiskri = request("FM_soeg")
	useKri = 1

    response.cookies("kunder_2015")("kundesogekri") = thiskri
	response.cookies("kunder_2015").expires = date + 1
	    
	else
	        
	        if request.cookies("kunder_2015")("kundesogekri") <> "" then
	        thiskri = request.cookies("kunder_2015")("kundesogekri")
	        useKri = 1
	        visikkekunder = 0
	        else
	        thiskri = ""
	        useKri = 0
	        end if
	        
	end if
	
   

    if len(trim(request("medarb"))) <> 0 then
	    selmedarb = request("medarb")
        response.cookies("kunder")("medarb") = selmedarb
    else

         if request.cookies("kunder")("medarb") <> "" then
         selmedarb = request.cookies("kunder")("medarb")
        else
         selmedarb = ""
        end if
    end if

	medarb = selmedarb
	
	if len(request("ktype")) <> 0 then
	ktype = request("ktype")
	else
	ktype = 0
	end if
	

     if len(request("FM_useasfak")) <> 0 then
		intuseasfak = request("FM_useasfak")

                if intuseasfak <> "-1" then
                useasfakSQL = " AND useasfak = "& intuseasfak
                else
                useasfakSQL = ""
                end if
        
        response.cookies("kunder")("useasfak") = intuseasfak
		else

            if request.cookies("kunder")("useasfak") <> "" then
                intuseasfak = request.cookies("kunder")("useasfak")
                        
                if intuseasfak <> "-1" then
                useasfakSQL = " AND useasfak = "& intuseasfak
                else
                useasfakSQL = ""
                end if
            else
				intuseasfak = "-1"
                useasfakSQL = ""
            end if
	end if




     if len(trim(request("FM_sorter"))) then
            
            sort = request("FM_sorter")

            if cint(sort) = 0 then
            sortBy = " ORDER BY kkundenavn"

            sortSel0 = "SELECTED"
            sortSel1 = ""

            else
            sortBy = " ORDER BY kkundenr"

             sortSel0 = ""
            sortSel1 = "SELECTED"

            end if
    else

            if request.cookies("tsa")("ksort") <> "" then
            sort = request.cookies("tsa")("ksort")

            if cint(sort) = 0 then
            sortBy = " ORDER BY kkundenavn"

            sortSel0 = "SELECTED"
            sortSel1 = ""

            else
            sortBy = " ORDER BY kkundenr"

             sortSel0 = ""
            sortSel1 = "SELECTED"

            end if

            else

            sort = 0
            sortBy = " ORDER BY kkundenavn"

            sortSel0 = "SELECTED"
            sortSel1 = ""

            end if
    end if

    response.cookies("tsa")("ksort") = sort



            if len(request("FM_hotnot")) <> 0 then
            hotnot = request("FM_hotnot")
            response.cookies("tsa")("hotnot") = hotnot
           
            else
                
                if request.cookies("tsa")("hotnot") <> "" then
                hotnot = request.cookies("tsa")("hotnot") 
                else
                hotnot = 99
                end if
            
            end if 

            select case cint(hotnot) 
            case -2
            hotnotSEL1 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case -1
            hotnotSEL2 = "SELECTED"
            hotnotSEL1 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 0
            hotnotSEL3 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL1 = ""
            hotnotSEL4 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 1
            hotnotSEL4 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL1 = ""
            hotnotSEL5 = ""
            hotnotSEL0 = ""
            case 2
            hotnotSEL5 = "SELECTED"
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL1 = ""
            hotnotSEL0 = ""
            case else
            hotnotSEL5 = ""
            hotnotSEL2 = ""
            hotnotSEL3 = ""
            hotnotSEL4 = ""
            hotnotSEL1 = ""
            hotnotSEL0 = "SELECTED"
            end select


	if len(medarb) <> 0 then
	useMedarb = medarb
	else
	useMedarb = 0 'session("Mid")
	medarb = useMedarb
	end if
	
	
	if useMedarb = "0" then
		useopraf_editor = " "
	else
		useopraf_editor = " AND (kunder.kundeans1 = "& useMedarb &" OR kunder.kundeans2 = "& useMedarb &")"
	end if
    
    
     %>
