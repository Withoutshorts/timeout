

<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->



<%      '**** S�gekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_afd"
        
            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if
            strOptions = "<option value=0>Choose..</option>"
            

            strSQL = "SELECT id, afdeling FROM kontaktpers WHERE kundeid = "& kid
            
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
        
              if len(trim(oRec("afdeling"))) <> 0 then
              strOptions = strOptions & "<option value="& oRec("id") &">"& oRec("afdeling") &"</option>"
              end if 
				
			oRec.movenext
			wend
			oRec.close

            '*** ��� **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions

         case "FN_betlev"
        
            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if
            

            strSQL = "SELECT betbetint, levbet FROM kunder WHERE kid = "& kid
            
            'response.Write strSQL
            'response.end

            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
        
            betbetint = oRec("betbetint")

			
			end if
			oRec.close

           

            response.write betbetint
            response.end



       
        case "FN_sup"
        
            if len(trim(request("jq_origin"))) <> 0 then
            origin = request("jq_origin")
            else 
            origin = ""
            end if
            
            lndTxt = origin
        
            lndTxt2SQL = " OR land = 'xx'"

         
            if lcase(lndTxt) = "hongkong" then
            lndTxt2SQL = " OR land = 'Hong Kong'"
            end if

            strOptions = "<option value=0>Choose..</option>"
            strSQL = "SELECT kid, kkundenavn FROM kunder WHERE (land = '"& lndTxt &"' "& lndTxt2SQL &") AND useasfak = 6 AND ketype = 'k' ORDER BY kkundenavn"
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
        
            
              strOptions = strOptions & "<option value="& oRec("kid") &">"& oRec("kkundenavn") &"</option>"
            
				
			oRec.movenext
			wend
			oRec.close

            '*** ��� **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions
            
         case "xFN_kpers" '�ndret til inter sales rep.
        

            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if

   

            strOptions = "<option value=0>Choose..</option>"

            strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
            
            oRec.open strSQL, oConn, 3
            while not oRec.EOF

                if len(trim(oRec("email"))) <> 0 then
                emlTxt = " ("& oRec("email") &")"
                else
                emlTxt = ""
                end if           

              strOptions = strOptions & "<option value="& oRec("mid") &">"& oRec("mnavn") &" "& emlTxt &"</option>" 
				
			oRec.movenext
			wend
			oRec.close

             '*** ��� **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions
            
    
        end select

        response.end
    
        end if %>




<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<%if media = "" then %>
<!--#include file="../inc/regular/header_lysblaa_2014_nt_inc.asp"--> 
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<%end if %>

<%if media = "print" then %>
<!--#include file="../inc/regular/header_hvd_inc.asp"--> 
<%end if %>

<%if media = "exp" then %>
<!--#include file="../inc/regular/header_lysblaa_2014_nt_inc.asp"--> 
<%end if %>


<script src="inc/job_nt_jav.js"></script>


<%
'*** S�tter lokal dato/kr format.
 Session.LCID = 1030






  
    if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)

    response.end

	end if

   

func = request("func")    

 if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
    end if
    




        
    
select case func  
	case "slet"
	'*** Her sp�rges om det er ok at der slettes  ***
	
	
	
    id = request("id")

	slttxt = "<b>Delete order</b><br />"_
	&"You are aboute to delete an order. Are you sure you want do that?"
	slturl = "job_nt.asp?menu=job&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)
	
	
	
	case "sletok"
	
        id = request("id")

	strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& id &"" 
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		
		'*** Inds�tter i delete historik ****'
	    call insertDelhist("job", id, oRec("jobnr"), oRec("jobnavn"), session("mid"), session("user"))
		
	
	end if
	oRec.close
	
	
	
	
	'Response.flush
	
	
	oConn.execute("DELETE FROM job WHERE id = "& id &"")
	
	
	
	Response.redirect "job_nt.asp"
	


case "dbopr", "dbred"

    if len(trim(request("FM_kunde"))) <> 0 AND request("FM_kunde") <> "0" then
    kid = request("FM_kunde")
    else
    kid = 0
    end if

    if cdbl(kid) = 0 then
               
     errortype = 170
    call showError(errortype)
				
    response.End
    end if

    if len(trim(request("FM_kopierordre"))) <> 0 then
    kopier_ordre = 1
    func = "dbopr"
    else
    kopier_ordre = 0
    end if

    if len(trim(request("FM_jobnavn"))) <> 0 then
    jobnavn = replace(request("FM_jobnavn"), "'", "")
    else
    jobnavn = ""
    end if

    if cint(kopier_ordre) = 1 then
    jobnavn = jobnavn & " - COPY" 
    end if


    if len(trim(jobnavn)) = 0 then
               
     errortype = 171
    call showError(errortype)
				
    response.End
    end if


    jobid = request("FM_jobid")

    if len(trim(request("FM_jobnr"))) <> 0 then
    jobnr = request("FM_jobnr")
            
            jobnrFindes = 0        
            strSQL = "SELECT jobnr FROM job WHERE id <> "& jobid &" AND jobnr = "& jobnr &""
            oRec2.open strSQL, oConn, 3
            if not oRec2.EOF then
    
            jobnrFindes = 1

            end if
            oRec2.close

    else
    jobnr = 0
    end if


    if jobnr = 0 OR jobnrFindes = 1 then
               
     errortype = 172
    call showError(errortype)
				
    response.End
    end if


    jobstatus = request("FM_status")
  

    kpers = request("FM_kpers")
   
    department = request("FM_department") 
    origin = request("FM_origin") 

    if len(trim(request("FM_supplier"))) <> 0 then
    supplier = request("FM_supplier")
    else
    supplier = 0
    end if

    fastpris = request("FM_fastpris")

    collection = replace(request("FM_collection"), "'", "")
    composition = replace(request("FM_composition"), "'", "") 

    product_group = request("FM_product_group")

    scsq_note = replace(request("FM_scsq_note"), "'", "") 
    sample_note = replace(request("FM_sample_note"), "'", "") 

    job_internbesk = replace(request("FM_internnote"), "'", "") 

    beskrivelse = replace(request("FM_beskrivelse"), "'", "") 'invoice note

    rekvnr = replace(request("FM_rekvnr"), "'", "")

    if len(trim(request("FM_bruttooms"))) <> 0 then
    bruttooms = replace(request("FM_bruttooms"), ".","")
    bruttooms = replace(bruttooms, ",",".")
    else
    bruttooms = 0
    end if 

    if len(trim(request("FM_jo_udgifter_intern"))) <> 0 then
    jo_udgifter_intern = replace(request("FM_jo_udgifter_intern"), ".","")
    jo_udgifter_intern = replace(jo_udgifter_intern, ",",".")
    else
    jo_udgifter_intern = 0
    end if 


    
    if len(trim(request("FM_orderqty"))) <> 0 then 
    orderqty = request("FM_orderqty")
    else
    orderqty = 0
    end if
   
    
    if len(trim(request("FM_shippedqty"))) <> 0 then 
    shippedqty = request("FM_shippedqty")
    else
    shippedqty = 0
    end if

    supplier_invoiceno = request("FM_supplier_invoiceno") 

    transport = request("FM_transport")
    destination = request("FM_destination")


    '**** ETD Buyer ***'
    

   
    dt_confb_etd = request("FM_dt_confb_etd")
    call fmatDate_fn(dt_confb_etd)
    dt_confb_etd = fmatDate

    
      
    dt_actual_etd = request("FM_dt_actual_etd")
    call fmatDate_fn(dt_actual_etd)
    dt_actual_etd = fmatDate


    '********* Requesred fields p� ORDER jobstatus = 1 ***'


    if cint(jobstatus) <> 3 then 'inquery

         
        if cint(jobstatus) = 1 then 'active

            if cint(supplier) = 0 OR cint(orderqty) = 0 OR len(orderqty) = 0 OR dt_confb_etd = "2010-01-01" then

            errortype = 173
            call showError(errortype)
				
            response.End

            end if

         end if



        if cint(jobstatus) = 2 then 'shipped

            if cint(supplier) = 0 OR cint(orderqty) = 0 OR len(orderqty) = 0 OR dt_confb_etd = "2010-01-01" OR cint(kpers) = 0 OR dt_actual_etd = "2010-01-01" OR len(shippedqty) = 0 or shippedqty = 0 OR len(trim(supplier_invoiceno)) = 0 then

            errortype = 174
            call showError(errortype)
				
            response.End

         end if


        end if
    end if
    '*** Dato felter ****'
    
    '** enquery dates
    
    dt_jobstdato = request("FM_dt_jobstdato") 

    if isDate(dt_jobstdato) then
    dt_jobenddato = dateAdd("m", 4, dt_jobstdato)
    dt_jobstdato = year(dt_jobstdato)&"-"&month(dt_jobstdato)&"-"&day(dt_jobstdato)
    else
    dt_jobstdato = "2010-01-01"
    dt_jobenddato = "2010-01-01"
    end if

    

    dt_enq_st = request("FM_dt_enq_st")
    call fmatDate_fn(dt_enq_st)
    dt_enq_st = fmatDate

   

    dt_enq_end = request("FM_dt_enq_end")
    call fmatDate_fn(dt_enq_end)
    dt_enq_end = fmatDate


    dt_sour_dead = request("FM_dt_sour_dead")
    call fmatDate_fn(dt_sour_dead)
    dt_sour_dead = fmatDate
    
    
 
    dt_proto_dead = request("FM_dt_proto_dead")
    call fmatDate_fn(dt_proto_dead)
    dt_proto_dead = fmatDate
    
   
    
    
    dt_proto_sent = request("FM_dt_proto_sent")
    call fmatDate_fn(dt_proto_sent)
    dt_proto_sent = fmatDate
    
    

    dt_sms_dead = request("FM_dt_sms_dead")
    call fmatDate_fn(dt_sms_dead)
    dt_sms_dead = fmatDate

    dt_sms_sent = request("FM_dt_sms_sent")
    call fmatDate_fn(dt_sms_sent)
    dt_sms_sent = fmatDate

    
    dt_photo_dead = request("FM_dt_photo_dead")
    call fmatDate_fn(dt_photo_dead)
    dt_photo_dead = fmatDate

  
 
   dt_photo_sent = request("FM_dt_photo_sent")
    call fmatDate_fn(dt_photo_sent)
    dt_photo_sent = fmatDate
    
      dt_sup_photo_dead = request("FM_dt_sup_photo_dead")
    call fmatDate_fn(dt_sup_photo_dead)
    dt_sup_photo_dead = fmatDate

    dt_sup_sms_dead = request("FM_dt_sup_sms_dead")
    call fmatDate_fn(dt_sup_sms_dead)
    dt_sup_sms_dead = fmatDate


   call orderAndProdDates
     
    
    '*********** SLUT DATOER *************'


    lastid = 0


    if len(trim(request("FM_freight_pc"))) <> 0 then
    freight_pc = replace(request("FM_freight_pc"), ".","")
    freight_pc = replace(freight_pc, ",",".")
    else
    freight_pc = 0
    end if

    if len(trim(request("FM_tax_pc"))) <> 0 then
    tax_pc = replace(request("FM_tax_pc"), ".","")
    tax_pc = replace(tax_pc, ",",".") 
    else
    tax_pc = 0
    end if

    if len(trim(request("FM_comm_pc"))) <> 0  then
    comm_pc = replace(request("FM_comm_pc"), ".","")
    comm_pc = replace(comm_pc, ",",".")
    else
    comm_pc = 0
    end if


    

   
    if len(trim(request("FM_cost_price_pc"))) <> 0 then
    cost_price_pc = replace(request("FM_cost_price_pc"), ".","")
    cost_price_pc = replace(cost_price_pc, ",",".")
    else
    cost_price_pc = 0
    end if


    if len(trim(request("FM_cost_price_pc_base"))) <> 0 then
    cost_price_pc_base = replace(request("FM_cost_price_pc_base"), ".","")
    cost_price_pc_base = replace(cost_price_pc_base, ",",".")
    else
    cost_price_pc_base = 0
    end if


    
     if len(trim(request("FM_sales_price_pc"))) <> 0 then
    sales_price_pc = replace(request("FM_sales_price_pc"), ".","")
    sales_price_pc = replace(sales_price_pc, ",",".")
    else
    sales_price_pc = 0
    end if
       
     if len(trim(request("FM_tgt_price_pc"))) <> 0 then
     tgt_price_pc = replace(request("FM_tgt_price_pc"), ".","")
     tgt_price_pc = replace(tgt_price_pc, ",",".")
     else
     tgt_price_pc = 0
     end if

    if len(trim(request("FM_jo_dbproc"))) <> 0 then
    jo_dbproc = replace(request("FM_jo_dbproc"), ".","")
    jo_dbproc = replace(jo_dbproc, ",",".")
    else
    jo_dbproc = 0
    end if

    sales_price_pc_valuta = request("FM_valuta_sales_price_pc_valuta")
    cost_price_pc_valuta = request("FM_valuta_cost_price_pc_valuta") 
    tgt_price_pc_valuta = request("FM_valuta_tgt_price_pc_valuta") 

    kunde_betbetint = request("FM_betbetint") 
    kunde_levbetint = request("FM_levbetint")
    lev_betbetint =  request("FM_levbetbetint")
    lev_levbetint = request("FM_levlevbetint")
    'sup_betbetint = request("FM_betbetint")


    if len(trim(request("FM_alert"))) <> 0 then
    alert = 1
    else
    alert = 0
    end if


    dd_dato = day(now) & ". "& left(monthname(month(now)), 3) &" "& year(now)
    editor = session("user")

    valuta = sales_price_pc_valuta

    'freight_price_pc_valuta, tgt_price_pc_valuta


    orderqty = replace(orderqty, ",", ".")


     jfak_moms = 0
     jfak_sprog = 2
     strSQLmomsK = "SELECT kfak_moms, kfak_sprog FROM kunder WHERE kid = "& kid 
     oRec.open strSQLmomsK, oConn, 3
     if not oRec.EOF then
    
        jfak_moms = oRec("kfak_moms")
        jfak_sprog = oRec("kfak_sprog") '2 UK

     end if
     oRec.close 

     


    if func = "dbopr" OR cint(kopier_ordre) = 1 then

        if cint(kopier_ordre) = 1 then

         call lastjobnr_fn()
        jobnr = nextjobnr

        end if


 

    strSQLjob = "INSERT INTO job (dato, editor, jobknr, jobnavn, jobstatus, jobnr, projektgruppe1, "_
    &" department, jobans1, origin, supplier, fastpris, collection, composition, product_group, scsq_note, sample_note, job_internbesk, beskrivelse, "_
    &" dt_enq_st, dt_enq_end, dt_sour_dead, dt_proto_dead, dt_proto_sent, dt_sms_dead, dt_sms_sent, dt_photo_dead, dt_photo_sent, dt_exp_order, "_
    &" dt_confb_etd, dt_confb_eta, dt_confs_etd, dt_confs_eta, dt_actual_etd, dt_actual_eta, rekvnr, jobstartdato, jobslutdato, "_
    &" dt_firstorderc, dt_ldapp, dt_sizeexp, dt_sizeapp, dt_ppexp, dt_ppapp, dt_shsexp, dt_shsapp, orderqty, shippedqty, supplier_invoiceno, transport, destination, jo_bruttooms, "_
    &" jo_udgifter_intern, dt_sup_photo_dead, dt_sup_sms_dead, freight_pc, tax_pc, comm_pc, cost_price_pc, sales_price_pc, tgt_price_pc, jo_dbproc, sales_price_pc_valuta, "_
    &" cost_price_pc_valuta, tgt_price_pc_valuta, cost_price_pc_base, kunde_betbetint, kunde_levbetint, lev_betbetint, lev_levbetint, valuta, jfak_moms, jfak_sprog, alert"_
    &" ) "_
    &" VALUES "_
    &" ('"& dd_dato &"', '"& editor &"', "& kid &", '"& jobnavn &"', "& jobstatus & ", '"& jobnr &"', 10, "_
    &"'"& department &"', "& kpers &", '"& origin &"', "& supplier &", "& fastpris &", '"& collection &"', '"& composition &"', "& product_group &", "_
    &"'"& scsq_note &"', '"& sample_note &"', '"& job_internbesk &"', '"& beskrivelse &"', '"& dt_enq_st &"', '"& dt_enq_end &"', '"& dt_sour_dead &"', "_
    &"'"& dt_proto_dead  &"', '"& dt_proto_sent  &"', '"& dt_sms_dead  &"', '"& dt_sms_sent &"', '"& dt_photo_dead  &"', '"& dt_photo_sent  &"', '"& dt_exp_order &"', "_
    &"'"& dt_confb_etd &"', '"& dt_confb_eta &"','"& dt_confs_etd &"', '"& dt_confs_eta &"', '"& dt_actual_etd &"', '"& dt_actual_eta &"', '"& rekvnr &"', '"& dt_jobstdato &"', '"& dt_jobenddato &"', "_
    &"'"& dt_firstorderc &"', '"& dt_ldapp &"', '"& dt_sizeexp &"', '"& dt_sizeapp &"', '"& dt_ppexp &"', '"& dt_ppapp &"', '"& dt_shsexp &"', '"& dt_shsapp &"', "_
    &""& orderqty &","& shippedqty &",'"& supplier_invoiceno &"', '"& transport &"', '"& destination &"', "& bruttooms &", "_
    &" "& jo_udgifter_intern &", '"& dt_sup_photo_dead &"', '"& dt_sup_sms_dead &"', "_
    &" "& freight_pc &","& tax_pc &","& comm_pc &", "& cost_price_pc &","& sales_price_pc &","& tgt_price_pc &", "& jo_dbproc &", "& sales_price_pc_valuta &","_
    &" "& cost_price_pc_valuta &", "& tgt_price_pc_valuta &", "& cost_price_pc_base &", "& kunde_betbetint &","& kunde_levbetint &","& lev_betbetint &","& lev_levbetint &", "& valuta &", "& jfak_moms &","& jfak_sprog &", "& alert &""_
    &" )" 


    'response.write strSQLjob
    'response.flush 

    oConn.execute(strSQLjob)


     

       strSQL = "SELECT id FROM job WHERE id <> 0 ORDER by id DESC limit 1"
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
        
        lastid = oRec("id")

       end if
       oRec.close

                
                
          
			
        

      '*** Opdater jobnr r�kkef�lge ***'
      strSQL = "UPDATE licens SET jobnr = "& jobnr &" WHERE id = 1"
	  oConn.execute(strSQL)

    else

     strSQLjob = "UPDATE job SET dato = '"& dd_dato &"', editor = '"& editor &"', jobknr = "& kid &", jobnavn = '"& jobnavn &"', "_
     &" jobstatus  = "& jobstatus & ", jobnr = '"& jobnr &"', projektgruppe1 = 10, "_
     &" department = '"& department &"', jobans1 = "& kpers &", origin = '"& origin &"', supplier = "& supplier &", "_
     &" fastpris = "& fastpris &", collection = '"& collection &"', composition = '"& composition &"', product_group = "& product_group &", "_
     &" scsq_note = '"& scsq_note &"', sample_note = '"& sample_note &"', job_internbesk = '"& job_internbesk &"', beskrivelse = '"& beskrivelse &"', "_
     &" dt_enq_st = '"& dt_enq_st &"', dt_enq_end = '"& dt_enq_end &"', dt_sour_dead = '"& dt_sour_dead &"', "_
     &" dt_proto_dead = '"& dt_proto_dead &"', dt_proto_sent = '"& dt_proto_sent &"', dt_sms_dead = '"& dt_sms_dead &"', dt_sms_sent = '"& dt_sms_sent &"', dt_photo_dead = '"& dt_photo_dead &"',"_
     &" dt_photo_sent = '"& dt_photo_sent &"', dt_exp_order = '"& dt_exp_order & "', "_
     &" dt_confb_etd = '"& dt_confb_etd &"', dt_confb_eta = '"& dt_confb_eta &"', dt_confs_etd = '"& dt_confs_etd &"', "_
     &" dt_confs_eta = '"& dt_confs_eta &"', dt_actual_etd = '"& dt_actual_etd &"', dt_actual_eta = '"& dt_actual_eta &"', rekvnr = '"& rekvnr &"', "_
     &" jobstartdato = '"& dt_jobstdato &"', jobslutdato = '"& dt_jobenddato &"', "_
     &" dt_firstorderc = '"& dt_firstorderc &"', dt_ldapp = '"& dt_ldapp &"', dt_sizeexp = '"& dt_sizeexp &"', dt_sizeapp = '"& dt_sizeapp &"', "_
     &" dt_ppexp = '"& dt_ppexp &"', dt_ppapp = '"& dt_ppapp &"', dt_shsexp = '"& dt_shsexp &"', dt_shsapp = '"& dt_shsapp &"', "_
     &" orderqty = "& orderqty &", shippedqty = "& shippedqty &", supplier_invoiceno = '"& supplier_invoiceno &"', transport = '"& transport &"', "_
     &" destination = '"& destination &"', jo_bruttooms = "& bruttooms &", jo_udgifter_intern = "& jo_udgifter_intern &", "_
     &" dt_sup_photo_dead = '"& dt_sup_photo_dead &"', dt_sup_sms_dead = '"& dt_sup_sms_dead &"', "_
     &" freight_pc = "& freight_pc &", tax_pc = "& tax_pc &", comm_pc = "& comm_pc &", "_
     &" cost_price_pc = "& cost_price_pc &", sales_price_pc = "& sales_price_pc &", tgt_price_pc = "& tgt_price_pc &", jo_dbproc = "& jo_dbproc &", "_
     &" sales_price_pc_valuta = "& sales_price_pc_valuta &", cost_price_pc_valuta = "& cost_price_pc_valuta &", tgt_price_pc_valuta = "& tgt_price_pc_valuta &", "_
     &" cost_price_pc_base = "& cost_price_pc_base &", "_
     &" kunde_betbetint = "& kunde_betbetint &", kunde_levbetint = "& kunde_levbetint &", lev_betbetint = "& lev_betbetint &", lev_levbetint = "& lev_levbetint &", valuta = "& valuta &", jfak_moms= "& jfak_moms &", jfak_sprog = "& jfak_sprog &", alert = "& alert &""_
     &" WHERE id = " & jobid
    
    'response.write strSQLjob
    'response.flush 
    
    oConn.execute(strSQLjob)

    lastid = jobid

    end if


    response.redirect "job_nt.asp?lastid="&lastid

case "bulk"


    dd_dato = day(now) & ". "& left(monthname(month(now)), 3) &" "& year(now)
    editor = session("user")

    if len(trim(request("bulk_jobids"))) > 3 then
    bulk_jobids = request("bulk_jobids")
    else
    bulk_jobids = 0
    end if



    bulk_jobidsArr = split(bulk_jobids, ",")


    for j = 0 TO UBOUND(bulk_jobidsArr)

     
    strSQLjobbulk = "UPDATE job SET dato = '"& dd_dato &"', editor = '"& editor &"'"

   

    '** Enq

    call enqDates

    if dt_sms_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sms_dead = '"& dt_sms_dead & "'"
    end if

    if dt_sup_sms_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sup_sms_dead = '"& dt_sup_sms_dead & "'"
    end if

     if dt_sms_sent <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sms_sent = '"& dt_sms_sent & "'"
    end if

     if dt_photo_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_photo_dead = '"& dt_photo_dead & "'"
    end if

      if dt_sup_photo_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sup_photo_dead = '"& dt_sup_photo_dead & "'"
    end if

   if dt_photo_sent <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_photo_sent = '"& dt_photo_sent & "'"
    end if
    
    
    '** Order 
   
      call orderAndProdDates

    if dt_exp_order <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_exp_order = '"& dt_exp_order & "'"
    end if


     if dt_confb_etd  <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confb_etd  = '"& dt_confb_etd  & "'"
    end if

    if dt_confb_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confb_eta = '"& dt_confb_eta & "'"
    end if
    
    if dt_confs_etd <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confs_etd = '"& dt_confs_etd & "'"
    end if

    if dt_confs_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confs_eta = '"& dt_confs_eta & "'"
    end if


    if dt_actual_etd <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_actual_etd = '"& dt_actual_etd & "'"
    end if

      if dt_actual_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_actual_eta = '"& dt_actual_eta & "'"
    end if




     if dt_firstorderc <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_firstorderc = '"& dt_firstorderc & "'"
    end if


     if dt_firstorderc <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_firstorderc = '"& dt_firstorderc & "'"
    end if

    if dt_ldapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ldapp = '"& dt_ldapp & "'"
    end if

    if dt_sizeexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sizeexp = '"& dt_sizeexp & "'"
    end if

     if dt_sizeapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sizeapp = '"& dt_sizeapp & "'"
    end if

    if dt_ppexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppexp = '"& dt_ppexp & "'"
    end if

    if dt_ppexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppexp = '"& dt_ppexp & "'"
    end if

     if dt_ppapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppapp = '"& dt_ppapp & "'"
    end if

    if dt_shsexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_shsexp = '"& dt_shsexp & "'"
    end if

     if dt_shsapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_shsapp = '"& dt_shsapp & "'"
    end if

   


   
    strSQLjobbulk = strSQLjobbulk &" WHERE id = " & bulk_jobidsArr(j)

   
    'response.write strSQLjob
    'response.flush 
    
    oConn.execute(strSQLjobbulk)


    next


    lastid = jobid

    response.redirect "job_nt.asp?lastid="&lastid


case "opret", "red"






if func = "red" then
    dbfunc = "dbred"

    if len(trim(request("jobid"))) <> 0 then
    id = request("jobid")
    else
    id = 0
    end if


    strSQLjob = "SELECT jobnavn, jobnr, jobknr, jobstatus,"_
    &" department, jobans1, origin, supplier, fastpris, collection, composition, product_group, "_
    &" scsq_note, sample_note, job_internbesk, beskrivelse, dt_enq_st, dt_enq_end, "_
    &" dt_proto_dead, dt_proto_sent, dt_sms_dead, dt_sms_sent, dt_photo_dead, dt_photo_sent, dt_exp_order, dt_sour_dead, "_ 
    &" dt_confb_etd, dt_confb_eta, dt_confs_etd, dt_confs_eta, dt_actual_etd, dt_actual_eta, rekvnr, jobstartdato, "_
    &" dt_firstorderc, dt_ldapp, dt_sizeexp, dt_sizeapp, dt_ppexp, dt_ppapp, dt_shsexp, dt_shsapp, orderqty, "_
    &" shippedqty, supplier_invoiceno, transport, destination, jo_bruttooms, jo_udgifter_intern, dt_sup_photo_dead, dt_sup_sms_dead, "_
    &" freight_pc, tax_pc, comm_pc, cost_price_pc, sales_price_pc, tgt_price_pc, jo_dbproc, sales_price_pc_valuta, cost_price_pc_valuta, tgt_price_pc_valuta, cost_price_pc_base, "_
    &" kunde_betbetint, kunde_levbetint, lev_betbetint, lev_levbetint, alert "_
    &" FROM job WHERE id = "& id

    'response.write strSQLjob
    'response.flush

    oRec.open strSQLjob, oConn, 3
    if not oRec.EOF then

    jobnavn = oRec("jobnavn")
    jobnr = oRec("jobnr")
    jobknr = oRec("jobknr")
    jobstatus = oRec("jobstatus")
    department = oRec("department")
    kpers = oRec("jobans1")
    origin = oRec("origin")
    supplier = oRec("supplier")
    fastpris = oRec("fastpris")
    collection = oRec("collection")
    composition = oRec("composition")
    product_group = oRec("product_group")
    scsq_note = oRec("scsq_note") 
    sample_note = oRec("sample_note")
    job_internbesk = oRec("job_internbesk") 
    beskrivelse = oRec("beskrivelse") 'inv. note
    
    dt_enq_st = oRec("dt_enq_st")
    dt_enq_end = oRec("dt_enq_end")

    dt_sour_dead = oRec("dt_sour_dead")

    dt_proto_dead = oRec("dt_proto_dead")
    dt_proto_sent = oRec("dt_proto_sent")
    dt_sms_dead = oRec("dt_sms_dead")
    dt_sms_sent = oRec("dt_sms_sent")
    dt_photo_dead = oRec("dt_photo_dead") 
    dt_photo_sent = oRec("dt_photo_sent")
    dt_exp_order = oRec("dt_exp_order")


    dt_confb_etd = oRec("dt_confb_etd") 
    dt_confb_eta = oRec("dt_confb_eta")
    dt_confs_etd = oRec("dt_confs_etd")
    dt_confs_eta = oRec("dt_confs_eta")
    dt_actual_etd = oRec("dt_actual_etd")
    dt_actual_eta = oRec("dt_actual_eta")


    rekvnr = oRec("rekvnr")

    dt_jobstdato = oRec("jobstartdato")


    dt_firstorderc = oRec("dt_firstorderc") 
    dt_ldapp = oRec("dt_ldapp") 
    dt_sizeexp = oRec("dt_sizeexp") 
    dt_sizeapp = oRec("dt_sizeapp")
    dt_ppexp = oRec("dt_ppexp")
    dt_ppapp = oRec("dt_ppapp")
    dt_shsexp = oRec("dt_shsexp")
    dt_shsapp = oRec("dt_shsapp")

    dt_sup_photo_dead = oRec("dt_sup_photo_dead")
    dt_sup_sms_dead = oRec("dt_sup_sms_dead")



    orderqty = oRec("orderqty")
    shippedqty = oRec("shippedqty")
    if shippedqty = 0 then
    shippedqty = ""
    end if
    
    supplier_invoiceno = oRec("supplier_invoiceno")


    transport = oRec("transport")
    destination = oRec("destination")

    bruttooms = oRec("jo_bruttooms")
    jo_udgifter_intern = oRec("jo_udgifter_intern")



    freight_pc = oRec("freight_pc")
    tax_pc = oRec("tax_pc") 
    comm_pc = oRec("comm_pc")

    cost_price_pc = oRec("cost_price_pc")
    cost_price_pc_base = oRec("cost_price_pc_base")

    sales_price_pc = oRec("sales_price_pc") 
    tgt_price_pc = oRec("tgt_price_pc") 

    jo_dbproc = oRec("jo_dbproc")

    sales_price_pc_valuta = oRec("sales_price_pc_valuta")
    cost_price_pc_valuta = oRec("cost_price_pc_valuta")

    tgt_price_pc_valuta = oRec("tgt_price_pc_valuta")


    kunde_betbetint = oRec("kunde_betbetint")
    kunde_levbetint = oRec("kunde_levbetint")
    lev_betbetint = oRec("lev_betbetint")
    lev_levbetint = oRec("lev_levbetint")

    alert = oRec("alert")

    if cint(alert) = 1 then
    alertCHK = "CHECKED"
    else
    alertCHK = ""
    end if


    end if
    oRec.close


    call alfanumerisk(jobnr)
    jobnr = alfanumeriskTxt
    jobnr = left(jobnr,20)



else

    dbfunc = "dbopr"
    id = 0

    call lastjobnr_fn()
    jobnr = nextjobnr


    '*** Opdater jobnr r�kkef�lge ***'
    strSQL = "UPDATE licens SET jobnr = "& jobnr &" WHERE id = 1"
	oConn.execute(strSQL)


    jobknr = 0
    jobnavn = ""

    jobstatus = 3
    kpers = 0
    origin = ""
    supplier = 0

    fastpris = 3

    collection = ""
    composition = ""

    product_group = 0

    sample_note = ""
    scsq_note = ""
    job_internbesk = ""
    beskrivelse = ""
    dt_enq_st = ""
    dt_enq_end = ""

     dt_sour_dead = ""

     dt_proto_dead = ""
    dt_proto_sent = ""
    dt_sms_dead = ""
    dt_sms_sent = ""
    dt_photo_dead = ""
    dt_photo_sent = ""
    dt_exp_order = ""

      dt_confb_etd = ""
    dt_confb_eta = ""
    dt_confs_etd = ""
    dt_confs_eta = ""
    dt_actual_etd = ""
    dt_actual_eta = ""

    rekvnr = ""

    dt_jobstdato = formatdatetime(now, 2)



     dt_firstorderc = ""
    dt_ldapp = "" 
    dt_sizeexp = "" 
    dt_sizeapp = ""
    dt_ppexp = ""
    dt_ppapp = ""
    dt_shsexp = ""
    dt_shsapp = ""
    dt_sup_photo_dead = ""
    dt_sup_sms_dead = ""


    orderqty = 0
    shippedqty = ""
    supplier_invoiceno = ""
    

    transport = ""
    destination = "Danmark"

    bruttooms = 0
    jo_udgifter_intern = 0

    cost_price_pc = 0
    cost_price_pc_base = 0
    sales_price_pc = 0
    tgt_price_pc = 0

    freight_pc = 0
    tax_pc = 0
    comm_pc = 0

    jo_dbproc = 0

    sales_price_pc_valuta = 3
    cost_price_pc_valuta = 3
    tgt_price_pc_valuta = 3


    kunde_betbetint = 0
    kunde_levbetint = 0
    lev_betbetint = 0
    lev_levbetint = 0

    alert = 0
    alertCHK = ""

    '*** lev & betbet kunde 
    'strSQLlevbetkunde = "SELECT levbet, betbet, betbetint FROM kunder WHERE kid = " & jobknr
    'oRec.open strSQLlevbetkunde, oConn, 3
    'if not oRec.EOF then

    ''kundeLevBet = oRec("levbet")
    'kunde_betbetint = oRec("betbetint")
    ''kundeBetBet = oRec("betbet")

    'end if
    'oRec.close


    
    '*** lev & betbet supplier 
    'strSQLlevbetkunde = "SELECT levbet, betbet, betbetint FROM kunder WHERE kid = " & supplier
    'oRec.open strSQLlevbetkunde, oConn, 3
    'if not oRec.EOF then

    ''supLevBet = oRec("levbet")
    'lev_betbetint = oRec("betbetint")

    'end if
    'oRec.close
  

end if


   

    




%>


<!--#include file="inc/job_nt_inc.asp"-->





<%call menu_2014 %>

<div class="container">
        <form action="job_nt.asp?func=<%=dbfunc %>" method="post">
            <input type="hidden" name="FM_jobid" value="<%=id %>" />

            <%call valutaKurs(1) %>
            <input type="hidden" id="valuta_kurs_1" value="<%=dblKurs %>" />
            <%call valutaKurs(2) %>
            <input type="hidden" id="valuta_kurs_2" value="<%=dblKurs %>" />
            <%call valutaKurs(3) %>
            <input type="hidden" id="valuta_kurs_3" value="<%=dblKurs %>" />
            <%call valutaKurs(4) %>
            <input type="hidden" id="valuta_kurs_4" value="<%=dblKurs %>" />
            <%call valutaKurs(5) %>
            <input type="hidden" id="valuta_kurs_5" value="<%=dblKurs %>" />

        <div class="row">
            <!--BASEDATA START -->
            <div id="basedata" class="col-s-12">
                
                
                <section class="panel">
                    <header class="panel-heading">Basedata</header>
                    <div class="panel-body no-pad-hoz">
                    
                        <div class="form-group col-s-12 col-m-6">
                            <label>Buyer <span style="color:red;">*</span></label>
                                <select id="FM_kunde" name="FM_kunde"><%=strFil_Kon_Txt %></select>

                        </div>

                         <div class="form-group col-s-12 col-m-6">
                            <label>Order No. <span style="color:red;">*</span></label>
                            <input type="text" name="" value="<%=jobnr %>" DISABLED />
                             <input type="hidden" name="FM_jobnr" value="<%=jobnr %>" />
                        </div>

                        
                        <div class="form-group col-s-12 col-m-6">
                            <label>Department</label>
                            <select id="xFM_afd" name="FM_department"><option value="0">Choose..</option>
                                <option value="Men" <%=departmentMenSEL %>>Men</option>
                                <option value="Women" <%=departmentWomenSEL %>>Women</option>
                                <option value="Kids" <%=departmentKidsSEL %>>Kids</option>
                                <option value="Unisex" <%=departmentUnisexSEL %>>Unisex</option>
                            </select>
                        </div>
                        <div class="form-group col-s-12 col-m-6">
                            <label>Place of origin</label>
                             <select id="FM_origin" name="FM_origin"><option value="0">Choose..</option>
                                <option value="Turkey" <%=originTurSel %>>Turkey</option>
                                <option value="China" <%=originChiSel %>>China</option>
                                 <option value="India" <%=originIndSel %>>India</option>
                                <option value="Bangladesh" <%=originBanSel %>>Bangladesh</option>
                                <option value="Vietnam" <%=originVieSel %>>Vietnam</option>
                                 <option value="Hongkong" <%=originHonSel %>>Hongkong</option>
                            </select>
                        </div>
                        <div class="form-group col-s-12 col-m-6">
                            <label>Sales Rep. <span style="color:red;">(*)</span></label>

                            <%call salesreplist(kpers) %>

                             <select id="FM_kpers" name="FM_kpers"><%=strFil_Kpers_Txt %></select>
                        </div>
                        <div class="form-group col-s-12 col-m-6">
                            <label>Supplier <span style="color:red;">(*)</span></label>
                            <select id="FM_supplier" name="FM_supplier"><%=strFil_Sup_Txt %></select>
                        </div>
                         <div class="form-group col-s-12 col-m-6">
                            <label>Status</label>
                            <select name="FM_status" id="status">
                            <option value="3" <%=jobstatus3SEL %>>Enquiry</option>
                            <option value="1" <%=jobstatus1SEL %>>Active</option>
                            
                            <option value="2" <%=jobstatus2SEL %>>Shipped</option>
                            <option value="4" <%=jobstatus4SEL %>>Cancelled</option>
                            <option value="0" <%=jobstatus0SEL %>>Invoiced/Closed</option>
                           
                          </select>
                        </div>

                        <div class="form-group col-s-12 col-m-6">
                            <label>Order type</label>
                            <select name="FM_fastpris" id="fastpris">
                                <option value="2" <%=fastpris2SEL %>>Commission</option>
                                <option value="3" <%=fastpris3SEL %>>Salesorder</option>
                               <option value="1" disabled>Fixed price</option>
                                <option value="0" disabled>Time & Mat.</option>
                            </select>
                        </div>
                       
                        <%if func = "red" then %>
                     <div class="form-group col-s-12 col-m-6">
                         <input type="checkbox" name="FM_kopierordre" value="1" /><label> Copy Order</label><br />
                    </div>
                    <%end if %>

                        

                    <div class="form-group" style="float:right; padding:20px;">
                        
                        <input type="submit" value="Submit >>">
                               </div>
                    

                </div>
                </section>
                </div>

                       

             
                
            
            <!--BASEDATA END -->



            <!--ENQUIRY INFO START -->
            <div id="enquiry-info" class="col-s-12 col-m-6">
                <section class="panel">
                    <header class="panel-heading">Enquiry info</header>
                    <div class="panel-body">
                        <div class="form-group">
                            <label>Style <span style="color:red;">*</span></label>
                            <input type="text" name="FM_jobnavn" value="<%=jobnavn %>" />
                        </div>
                        <div class="form-group">
                            <label>Collection</label>
                             <input type="text" name="FM_collection" value="<%=collection %>" />
                        </div>
                        <div class="form-group col-s-6 no-pad-left">
                            <label>Enquiry startup</label>
                            <input type="text" name="FM_dt_enq_st" value="<%=dt_enq_st %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <!-- <div class="form-group col-s-6 no-pad-left">
                            <label>Result</label>
                            <input type="text" id="result" value="" />
                        </div>-->
                        <div class="form-group col-s-6 no-pad-right">
                            <label>Enquiry end</label>
                            <input type="text" name="FM_dt_enq_end" value="<%=dt_enq_end %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group">
                            <label>Product group</label>
                            <select type="text" name="FM_product_group"><%=strFil_PG_Txt %></select>
                        </div>
                       
                        <div class="form-group">
                            <label>Composition</label>
                            <input type="text" name="FM_composition" value="<%=composition %>" />
                        </div>
                        <div class="form-group">
                            <label>Sourcing deadline</label>
                            <input type="text" name="FM_dt_sour_dead" value="<%=dt_sour_dead %>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-6 no-pad-left">
                            <label>Proto deadline</label>
                            <input type="text" name="FM_dt_proto_dead" value="<%=dt_proto_dead %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6 no-pad-right">
                            <label>Proto sent</label>
                            <input type="text" name="FM_dt_proto_sent" value="<%=dt_proto_sent %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-4 no-pad-left">
                            <label>SMS Buyer deadline</label>
                            <input type="text" name="FM_dt_sms_dead" value="<%=dt_sms_dead %>" placeholder="dd-mm-yyyy" />
                        </div>
                         <div class="form-group col-s-4 no-pad-right">
                            <label>SMS Supplier DL</label>
                            <input type="text" name="FM_dt_sup_sms_dead" value="<%=dt_sup_sms_dead%>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-4 no-pad-right">
                            <label>SMS sent</label>
                            <input type="text" name="FM_dt_sms_sent" value="<%=dt_sms_sent %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-4 no-pad-left">
                            <label>Photo Buyer Deadline</label>
                            <input type="text" name="FM_dt_photo_dead" value="<%=dt_photo_dead %>" placeholder="dd-mm-yyyy" />
                        </div>
                          <div class="form-group col-s-4 no-pad-right">
                            <label>Photo Supplier DL</label>
                            <input type="text" name="FM_dt_sup_photo_dead" value="<%=dt_sup_photo_dead%>" placeholder="dd-mm-yyyy" />
                        </div>
                       
                        <div class="form-group col-s-4 no-pad-right">
                            <label>Photo sent</label>
                            <input type="text" name="FM_dt_photo_sent" value="<%=dt_photo_sent %>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group">
                            <label>Exp order</label>
                            <input type="text" name="FM_dt_exp_order" value="<%=dt_exp_order %>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Sample colour &amp Sample Qty.</label>
                            <textarea name="FM_scsq_note" placeholder="Write your comment here"><%=scsq_note %></textarea>
                        </div>
                        <div class="form-group">
                            <label>Sample Note</label>
                            <textarea name="FM_sample_note" placeholder="Write your comment here"><%=sample_note %></textarea>
                        </div>
                  
                </section>
            </div>
            <!--ENQUIRY INFO END -->
            <!--PRICE TERMS START -->
            <div id="price-terms" class="col-s-12 col-m-6">
                <section class="panel">
                    <header class="panel-heading">Price terms</header>
                    <div class="panel-body">
                        <div class="form-group">
                            <label>Order Qty. <span style="color:red;">(*)</span></label>
                            <input type="text" name="FM_orderqty" id="orderqty" value="<%=orderqty%>" />
                        </div>

                       
                        <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Cost price PC (<label id="cost_price_pc_label"><%=formatnumber(cost_price_pc_base, 4) %></label>) base amount here</label>
                            <span class="col-s-9 no-pad-left">
                                <input type="text" name="FM_cost_price_pc" id="cost_price_pc" value="<%=formatnumber(cost_price_pc, 4) %>" />
                                <input type="hidden" name="FM_cost_price_pc_base" id="cost_price_pc_base" value="<%=formatnumber(cost_price_pc_base, 4) %>" />
                            </span>
                            <span class="col-s-3 no-pad-hoz">
                                <%call valutakoder("cost_price_pc_valuta", cost_price_pc_valuta, 1) %>
                             
                                
                            </span>
                        </div>

                         

                        <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Sales price PC <span style="color:#999999;">(invoice currency)</span> <img src="../img/blank.gif" width="100" height="1" border="0" /><input type="checkbox" id="cc" disabled /> Add sales price manually (and calc. comi.)</label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" name="" id="sales_price_pc_label" value="<%=formatnumber(sales_price_pc, 4) %>" DISABLED/>
                                <input type="text" name="FM_sales_price_pc" id="sales_price_pc" value="<%=formatnumber(sales_price_pc, 4) %>"/>
                            </span>
                            <span class="col-s-3 no-pad-hoz" >
                                <%call valutakoder("sales_price_pc_valuta", sales_price_pc_valuta, 1) %>
                            </span>
                        </div>
                         <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Target price PC</label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" name="FM_tgt_price_pc" id="tgt_price_pc" value="<%=formatnumber(tgt_price_pc, 4) %>"/>
                            </span>
                           
                                <span class="col-s-3 no-pad-hoz" >
                                <%
                                    
                                  call valutakoder("tgt_price_pc_valuta", tgt_price_pc_valuta, 1) %>
                            </span>
                           
                        </div>
                       
                        <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Comission PC %</label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" name="" id="comm_pc_label" value="<%=comm_pc%>" DISABLED/>
                                <input type="text" name="FM_comm_pc" id="comm_pc" value="<%=comm_pc %>" />
                            </span>
                            
                            <span class="col-s-3 no-pad-hoz" >
                                <span class="col-s-3 no-pad-hoz" >
                              &nbsp;<label style="padding-top:8px;">%</label>
                                 </span>
                            </span>
                        </div>
                     
                        <div class="form-group">
                                <label class="col-s-12 no-pad-hoz">Freight PC</label>
                             <span class="col-s-9 no-pad-left" >
                                 <input type="text" name="" id="freight_pc_label" value="0" DISABLED/>
                            <input type="text" name="FM_freight_pc" id="freight_pc" value="<%=freight_pc %>" />
                                 </span>
                             <span class="col-s-3 no-pad-hoz" >

                                 <%freight_price_pc_valuta = cost_price_pc_valuta 'f�lger altid %>
                             <%call valutakoder("freight_price_pc_valuta", freight_price_pc_valuta, 1) %>
                                 </span>
                        </div>
                        <div class="form-group">
                             <label class="col-s-12 no-pad-hoz">TAX % </label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" name="" id="tax_pc_label" value="<%=tax_pc %>" DISABLED/>
                                <input type="text" name="FM_tax_pc" id="tax_pc" value="<%=tax_pc %>" />
                            </span>

                         
                            <span class="col-s-3 no-pad-hoz" >
                                <span class="col-s-3 no-pad-hoz" >
                              &nbsp;<label style="padding-top:8px;">%</label>
                                 </span>
                            </span>

                        </div>
                        <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Total cost price (Order Qty. * Cost. PC * TAX PC) + (Freight PC * Order Qty.)</label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" id="udgifter_intern_label" value="<%=formatnumber(jo_udgifter_intern, 2) %>" DISABLED/>
                                <input type="hidden" name="FM_jo_udgifter_intern" id="udgifter_intern" value="<%=formatnumber(jo_udgifter_intern, 2) %>"/>
                            </span>

                                <span class="col-s-3 no-pad-hoz" >
                              &nbsp;<label style="padding-top:8px;">DKK</label>
                                 </span>
                        </div>
                        <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Total sales price (Order Qty. * Sales PC)</label>
                            <span class="col-s-9 no-pad-left" >
                                <input type="text" id="bruttooms_label" value="<%=formatnumber(bruttooms, 2) %>" DISABLED/>
                                <input type="hidden" name="FM_bruttooms" id="bruttooms" value="<%=formatnumber(bruttooms, 2) %>"/>
                            </span>
                               <span class="col-s-3 no-pad-hoz" >
                              &nbsp;<label style="padding-top:8px;">DKK</label>
                                 </span>

                         
                        </div>

                         <div class="form-group">
                            <label class="col-s-12 no-pad-hoz">Total Profit ~</label>
                            <span class="col-s-3 no-pad-left" >
                                <input type="text" id="jo_dbproc_label" value="<%=jo_dbproc %>" DISABLED />
                                <input type="hidden" name="FM_jo_dbproc" id="jo_dbproc" value="<%=jo_dbproc %>" />
                            </span>
                             <span class="col-s-3 no-pad-hoz" >
                            <label>%</label>
                                 </span>

                            

                             <%jo_dbproc_bel = formatnumber((bruttooms - jo_udgifter_intern), 2) %>

                              

                             <span class="col-s-3 no-pad-left" >
                                <input type="text" name="" id="jo_dbproc_bel" value="<%=jo_dbproc_bel %>" disabled />
                            </span>

                                 <span class="col-s-3 no-pad-hoz" >
                              &nbsp;<label style="padding-top:8px;">DKK</label>
                                 </span>
                            
                        </div>
                        <!--
                        <div class="form-group">
                            <label>Profit/PC</label>
                            <input type="text" value="DKK" disabled />
                        </div>
                        -->
                       

                        <div class="form-group">   <br /><br /><br /><br />
                            <label>Payment Term Buyer</label><input type="hidden" id="FM_t5" value="0" />
                             <%
                            disa = 0
                            lang = 1
                            nameid = "FM_betbetint"
                            call betalingsbetDage(kunde_betbetint, disa, lang, nameid) %>
                        </div>
                        <div class="form-group">
                            
                            <label>Delivery Term Buyer</label> 
                            <select name="FM_levbetint">
                                <option value="0" <%=kunde_levbetint0SEL %>>Choose..</option>
                                <option value="1" <%=kunde_levbetint1SEL %>>FOB</option>
                                <option value="2" <%=kunde_levbetint2SEL %>>DDP</option>

                            </select>
                            
                        </div>

                        
                        <div class="form-group">
                            <label>Payment Term Supplier</label>
                            <%
                            disa = 0
                            lang = 1
                            nameid = "FM_levbetbetint"
                            call betalingsbetDage(lev_betbetint, disa, lang, nameid) %>
                        </div>
                        <div class="form-group">
                            <label>Delivery Term Supplier</label>
                         
                              <select name="FM_levlevbetint">
                                <option value="0" <%=levlevbetint0SEL %>>Choose..</option>
                                <option value="1" <%=levlevbetint1SEL %>>FOB</option>
                                <option value="2" <%=levlevbetint2SEL %>>DDP</option>

                            </select>
                        </div>
                        <div class="form-group">
                            <br /><br /><br />
                             <label>Price or Production Note</label>
                            <textarea name="FM_internnote" placeholder="Write your comment here"><%=job_internbesk %></textarea>
                            <br /> <input type="checkbox" name="FM_alert" value="1" <%=alertCHK %> /><label> Alert (attention needed, add comment to production note)</label>
                        </div>

                          <div class="form-group">
                        <br />&nbsp;
                        </div>
                       
                </section>
            </div>
            <!--PRICE TERMS END -->
            <!--ORDER INFO START -->
            <div id="price-terms" class="col-s-12 col-m-6">
                <section class="panel">
                    <header class="panel-heading">Order info</header>
                    <div class="panel-body">
                        <div class="form-group">
                            <label>Buyers Order No. (PO no.)</label>
                            <input type="text" name="FM_rekvnr" value="<%=rekvnr %>" />
                        </div>
                        <div class="form-group">
                            <label>Order placement</label>
                            <input type="text" name="FM_dt_jobstdato" value="<%=dt_jobstdato%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Destination</label>
                            <select name="FM_destination">
                                <option>Denmark</option>
                                 <option disabled>------------------------</option>
                                 <option value="0" disabled>Choose destination</option>
                                <%if func = "red" then%>
		                        <option SELECTED><%=destination%></option>
                                <%end if %>

                               
                               <!--#include file="inc/inc_option_land.asp"-->
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Transport</label>
                            <select name="FM_transport">
                                <option value="0">Choose transport</option>
                                <option value="By Air" <%=transportFlSel %>>By Air</option>
                                <option value="By Sea" <%=transportShSel %>>By Sea (payment terms +5 weeks)</option>
                                <option value="By Train" <%=transportTrSel%>>By Train</option>
                                <option value="By Truck" <%=transportRdSel %>>By Truck (payment terms +7 days)</option>
                                <option value="By Currier" <%=transportCurSel %>>By Currier</option>
                            </select>
                        </div>
                        <div class="form-group col-s-6 no-pad-left">
                            <label>Confirmed buyer ETD <span style="color:red;">(*)</span></label>
                            <input type="text" name="FM_dt_confb_etd" value="<%=dt_confb_etd%>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6 no-pad-right">
                            <label>Confirmed buyer ETA</label>
                            <input type="text" name="FM_dt_confb_eta" value="<%=dt_confb_eta%>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6 no-pad-left">
                            <label>Confirmed supplier ETD</label>
                            <input type="text" name="FM_dt_confs_etd" value="<%=dt_confs_etd%>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6 no-pad-right">
                            <label>Confirmed supplier ETA</label>
                            <input type="text" name="FM_dt_confs_eta" value="<%=dt_confs_eta%>" placeholder="dd-mm-yyyy" />
                        </div>
                       
                        <div class="form-group col-s-6 no-pad-left">
                            <label>Actual ETD <span style="color:red;">(*)</span></label>
                            <input type="text" name="FM_dt_actual_etd" value="<%=dt_actual_etd%>" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6 no-pad-right">
                            <label>Actual ETA</label>
                            <input type="text"  name="FM_dt_actual_eta" value="<%=dt_actual_eta%>" placeholder="dd-mm-yyyy" />
                        </div>
                        
                        <div class="form-group">
                            <label>Shipped Qty. <span style="color:red;">(*)</span></label>
                            <input type="text" id="shippedqty" name="FM_shippedqty" value="<%=shippedqty%>" />
                        </div>
                        <div class="form-group">
                            <label>Supplier invoice No. <span style="color:red;">(*)</span></label>
                            <input type="text" name="FM_supplier_invoiceno" value="<%=supplier_invoiceno%>"  />
                        </div>
                        <!--
                        <div class="form-group">
                            <label>Invoiced</label>
                            <select type="text">
                                <option>Invoiced</option>
                                <option>1</option>
                                <option>2</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Payed</label>
                            <select type="text">
                                <option>Payed</option>
                                <option>1</option>
                                <option>2</option>
                            </select>
                        </div>
                        -->
                        <div class="form-group">
                            <label>Invoice Note (visible on invoice)</label>
                            <textarea name="FM_beskrivelse" placeholder="Write your comment here"><%=beskrivelse %></textarea>
                        </div>
                   
                </section>
            </div>
            <!--ORDER INFO END -->
            <!--PRODUCTION INFO START -->
            <div id="price-terms" class="col-s-12 col-m-6">
                <section class="panel">
                    <header class="panel-heading">Production info</header>
                    <div class="panel-body">
                        <div class="form-group">
                            <label>1st. order comment</label>
                            <input type="text" name="FM_dt_firstorderc" value="<%=dt_firstorderc%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>LD app</label>
                            <input type="text" name="FM_dt_ldapp" value="<%=dt_ldapp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Exp Sizeset</label>
                            <input type="text" name="FM_dt_sizeexp" value="<%=dt_sizeexp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Sizeset app</label>
                            <input type="text" name="FM_dt_sizeapp" value="<%=dt_sizeapp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Exp PP</label>
                            <input type="text" name="FM_dt_ppexp" value="<%=dt_ppexp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>PP app</label>
                            <input type="text" name="FM_dt_ppapp" value="<%=dt_ppapp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>Exp SHS</label>
                            <input type="text" name="FM_dt_shsexp" value="<%=dt_shsexp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group">
                            <label>SHS app</label>
                            <input type="text" name="FM_dt_shsapp" value="<%=dt_shsapp%>" placeholder="dd-mm-yyyy"/>
                        </div>
                     
               
                </section>

                
                <section>  <div class="form-group">
                        <input type="submit" value="Submit >>">
                               </div></section>
                     </form>
         
                
                </div>
            <!--PRODUCTION INFO END -->
        </div>
    </div>



<%
 
    case else   

    

    if len(trim(request("post"))) <> 0 then
    post = 1
    else
    post = 0
    end if

    if len(trim(request("lastid"))) <> 0 then
    lastid = request("lastid")
    else
    lastid = 0
    end if

   if len(trim(request("rapporttype"))) <> 0 then
   rapporttype = request("rapporttype")
   response.cookies("orders")("rapporttype") = rapporttype 
   else
        if request.cookies("orders")("rapporttype") <> "" then
        rapporttype = request.cookies("orders")("rapporttype")
        else
        rapporttype = 0
        end if
    end if


   if len(trim(request("supplier"))) <> 0 then
   supplier = request("supplier")
   response.cookies("orders")("supplier") = supplier 
   else
        if request.cookies("orders")("supplier") <> "" then
        supplier = request.cookies("orders")("supplier")
        else
        supplier = 0
        end if
    end if

    

    
   if len(trim(request("salesrep"))) <> 0 then
   salesrep = request("salesrep")
   response.cookies("orders")("salesrep") = salesrep 
   else
        if request.cookies("orders")("salesrep") <> "" then
        salesrep = request.cookies("orders")("salesrep")
        else
        salesrep = 0
        end if
    end if

    if salesrep <> 0 then
    salesrepSQL = " AND j.jobans1 = "& salesrep
    else
    salesrepSQL = ""
    end if

    if supplier <> 0 then
    supplierSQL = " AND j.supplier = "& supplier
    else
    supplierSQL = ""
    end if

    'Rapporttyper
    '0: overview 
    '1: enquery  / case else
    '2: production
    '3: production no eco
    '4: Sales budget

    select case rapporttype
    case 1
    rapporttypeTxt = "Production / Enq. Overview"
    case 3
    rapporttypeTxt = "Production"
    case 4
    rapporttypeTxt = "Production (simple)"
    case 5
     rapporttypeTxt = "Sales budget"
    case else
    rapporttypeTxt = "Orders (overview)"
    end select

    if cint(post) = 1 then
    sogVal = request("FM_sog")
    response.cookies("orders")("sog") = sogVal
    else
        if request.cookies("orders")("sog") <> "" then
        sogVal = request.cookies("orders")("sog")
        else
        sogVal = ""
        end if
    end if

    if len(trim(request("FM_status1"))) <> 0 then
    strStatusSQL = " AND (jobstatus = 1"
    statusCHK1 = "CHECKED"
    response.cookies("orders")("status1") = "1" 
    else
        if cint(post) = 1 then
        strStatusSQL = " AND (jobstatus = -1"
        statusCHK1 = ""
        response.cookies("orders")("status1") = ""
        else
            if request.cookies("orders")("status1") <> "" then
            strStatusSQL = " AND (jobstatus = 1"
            statusCHK1 = "CHECKED"
            else
            
                if len(trim(request("rapporttype"))) <> 0 then 'altid sl�et til n�r der v�lges en rapport fra menu, uanset hvilken          
                strStatusSQL = " AND (jobstatus = 1"
                statusCHK1 = "CHECKED"
                else
                strStatusSQL = " AND (jobstatus = -1"
                statusCHK1 = ""
                end if
            end if
        end if
    end if


    if len(trim(request("FM_status2"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 2"
    statusCHK2 = "CHECKED"
    response.cookies("orders")("status2") = "1" 
    else
        if cint(post) = 1 then
        statusCHK2 = ""
        response.cookies("orders")("status2") = ""
        else
            if request.cookies("orders")("status2") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 2"
            statusCHK2 = "CHECKED"
            else
            statusCHK2 = ""
            end if
        end if
    end if

  if len(trim(request("FM_status3"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 3"
    statusCHK3 = "CHECKED"
    response.cookies("orders")("status3") = "1" 
    else
        if cint(post) = 1 then
        statusCHK3 = ""
        response.cookies("orders")("status3") = ""
        else
            if request.cookies("orders")("status3") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 3"
            statusCHK3 = "CHECKED"
            else
                
                  if request("rapporttype") = "1" then 'altid sl�et til n�r der v�lges en Enqueries fra menu  
                  strStatusSQL = strStatusSQL & " OR jobstatus = 3"
                  statusCHK3 = "CHECKED"  
                  else
                  statusCHK3 = ""
                  end if
            end if
        end if
    end if

   if len(trim(request("FM_status4"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 4"
    statusCHK4 = "CHECKED"
    response.cookies("orders")("status4") = "1" 
    else
        if cint(post) = 1 then
        statusCHK4 = ""
        response.cookies("orders")("status4") = ""
        else
            if request.cookies("orders")("status4") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 4"
            statusCHK4 = "CHECKED"
            else
            statusCHK4 = ""
            end if
        end if
    end if

       if len(trim(request("FM_status0"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 0"
    statusCHK0 = "CHECKED"
    response.cookies("orders")("status0") = "1" 
    else
        if cint(post) = 1 then
        statusCHK0 = ""
        response.cookies("orders")("status0") = ""
        else
            if request.cookies("orders")("status0") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 0"
            statusCHK0 = "CHECKED"
            else
            statusCHK0 = ""
            end if
        end if
    end if


    
    strStatusSQL = strStatusSQL & ")"


    if len(trim(request("sort"))) <> 0 then
    sort = request("sort")
    response.cookies("orders")("orderby") = sort
    else
         if request.cookies("orders")("orderby") <> "" then
         sort = request.cookies("orders")("orderby")
         else
         sort = 1
         end if 
    end if


     select case sort
        case "1"
        strSQLOdrBy = "kkundenavn, jobnavn"
        case "2"
        strSQLOdrBy = "jobstatus, kkundenavn, jobnavn"
        case "3"
        strSQLOdrBy = "dt_sour_dead"
        case "4"
        strSQLOdrBy = "dt_proto_dead"
        case "5"
        strSQLOdrBy = "dt_photo_dead"
        case "6"
        strSQLOdrBy = "dt_sms_dead"
        case "7"
        strSQLOdrBy = "jobnavn"
        case "8"
        strSQLOdrBy = "dt_sizeapp"
        case "9"
        strSQLOdrBy = "dt_ppapp"
        case "10"
        strSQLOdrBy = "dt_shsapp"
        case "11"
        strSQLOdrBy = "jobstartdato"
        case "14"
        strSQLOdrBy = "dt_confb_etd"
            case "15"
                strSQLOdrBy = "rekvnr"
            case "16"
                strSQLOdrBy = "supplier"
            case "17"
                strSQLOdrBy = "product_group"
            case "18"
                strSQLOdrBy = "jobnr"
        case "19"
                strSQLOdrBy = "dt_actual_etd"
        case "20"
                strSQLOdrBy = "dt_confs_etd"
        case else 
        strSQLOdrBy = "kkundenavn, jobnavn"
        end select

     
    strSQLOdrBy = strSQLOdrBy & ", kkundenavn, supplier, jobnavn"
    

    if len(trim(request("FM_dt_from"))) <> 0 then
    dt_from = request("FM_dt_from")
        
    dt_from = replace(dt_from, ".", "-")
    dt_from = replace(dt_from, "/", "-")
    dt_from = replace(dt_from, ":", "-")
    dt_from = replace(dt_from, " ", "")

     if isDate(dt_from) = false then
    dt_from = "2010-01-01"
    end if

    response.cookies("orders")("dt_from") = dt_from
    else

        if request.cookies("orders")("dt_from") <> "" then
        dt_from = request.cookies("orders")("dt_from")
        else
        dt_from = dateAdd("m", -3, now)
        end if

    end if

    dt_fromTxt = formatdatetime(dt_from, 2) 'day(dt_from) &"-"& month(dt_from) &"-"& year(dt_from)
    dt_fromSQL = year(dt_from) &"-"& month(dt_from) &"-"& day(dt_from)

    if len(trim(request("FM_dt_to"))) <> 0 then
    dt_to = request("FM_dt_to")

    dt_to = replace(dt_to, ".", "-")
    dt_to = replace(dt_to, "/", "-")
    dt_to = replace(dt_to, ":", "-")
    dt_to = replace(dt_to, " ", "")

    if isDate(dt_to) = false then
    dt_to = "2010-01-01"
    end if
     
    response.cookies("orders")("dt_to") = dt_to
    else

        if request.cookies("orders")("dt_to") <> "" then
        dt_to = request.cookies("orders")("dt_to")
        else
        dt_to = dateAdd("m", -3, now)
        end if

    end if

     dt_toTxt = formatdatetime(dt_to, 2) '&"-"& month(dt_to) &"-"& year(dt_to)
    dt_toSQL = year(dt_to) &"-"& month(dt_to) &"-"& day(dt_to)


    if len(trim(request("FM_append_to"))) <> 0 then
    append_to = request("FM_append_to")
    response.cookies("orders")("apend_to") = append_to 
    else
        if request.cookies("orders")("apend_to") <> "" then
        append_to = request.cookies("orders")("apend_to")
        else
        append_to = "-1"
        end if
    end if

  
    appto_1Sel = ""
    appto_2Sel = ""
    appto_3Sel = ""
    appto_4Sel = ""
    appto_5Sel = ""
    appto_6Sel = ""
    appto_7Sel = ""
    appto_8Sel = ""
     appto_9Sel = ""
     appto_10Sel = ""
     appto_11Sel = ""
     appto_12Sel = ""
     appto_13Sel = ""
     appto_14Sel = ""
    appto_15Sel = ""

    select case append_to
    case "-1"
    strSQLdtKri = ""
    case "1"
    strSQLdtKri = " AND (dt_sour_dead " 
    appto_1Sel = "SELECTED"
    case "2"
    strSQLdtKri = " AND (dt_proto_dead "
    appto_2Sel = "SELECTED"
    case "3"
    strSQLdtKri = " AND (dt_photo_dead "
    appto_3Sel = "SELECTED"
    case "4"
    strSQLdtKri = " AND (dt_sms_dead "
    appto_4Sel = "SELECTED"
     case "5"
    strSQLdtKri = " AND (dt_sizeapp "
    appto_5Sel = "SELECTED"
    case "6"
    strSQLdtKri = " AND (dt_ppapp "
    appto_6Sel = "SELECTED"
    case "7"
    strSQLdtKri = " AND (dt_shsapp "
    appto_7Sel = "SELECTED"
     case "8"
    strSQLdtKri = " AND (jobstartdato "
    appto_8Sel = "SELECTED"
    'case "9"
    'strSQLdtKri = " AND (dt_ppapp "
    'appto_9Sel = "SELECTED"
    'case "10"
    'strSQLdtKri = " AND (dt_shsapp "
    'appto_10Sel = "SELECTED"
     case "11"
    strSQLdtKri = " AND (jobstartdato "
    appto_11Sel = "SELECTED"
     case "12"
    strSQLdtKri = " AND (dt_sup_photo_dead "
    appto_12Sel = "SELECTED"
     case "13"
    strSQLdtKri = " AND (dt_sup_sms_dead "
    appto_13Sel = "SELECTED"
     case "14"
    strSQLdtKri = " AND (dt_confb_etd "
    appto_14Sel = "SELECTED"
    case "15"
    strSQLdtKri = " AND (dt_confs_etd "
    appto_15Sel = "SELECTED"
    end select

    if append_to <> "-1" then
    strSQLdtKri = strSQLdtKri & " BETWEEN '"& dt_fromSQL &"' AND '"& dt_toSQL &"' )"
    end if


    rapporttype0SEL = ""
    rapporttype1SEL = ""
    rapporttype3SEL = ""

    select case rapporttype
    case 0
    rapporttype0SEL = "SELECTED"
    case 1
    rapporttype1SEL = "SELECTED"
    case 3
    rapporttype3SEL = "SELECTED"
    case else
    rapporttype0SEL = "SELECTED"
    end select

    if media = "" then
    call menu_2014()
    end if




    %>
    



<%if media = "" then %>

        
        <div class="container">

     

        <div class="row">
            
<%end if %>

            <%if media = "" then %>
            <!--SEARCH START -->

            



            <div id="search" class="col-s-12 col-m-8">
                <section class="panel">
                    <header class="panel-heading">Search</header>
                    <form class="panel-body" method="post" action="job_nt.asp?post=1">
                        <input type="hidden" id="showfullscreen" value="0" />
                         
                       

                        <!--<div class="search-group">-->
                        <div class="form-group col-s-11 no-pad-left">
                         <label>Search:<br />Buyer, Style, Style, OR <b>Style NO.</b> or Order No ">" 1000 e.g. PO no. or Sup. Invoice NO. or<b> NT Invoice NO.</b></label>
                        <input type="search" name="FM_sog" value="<%=sogVal%>" placeholder="" />
                            </div>

                         

                         <div class="form-group col-s-1 no-pad-left" style="padding-top:33px;">
                        <span class="search-group-btn"><input type="submit" value="Search"></span>
                            
                        </div>

                        <br />
                         <div class="form-group col-s-4 no-pad-left">
                            <label>Date from:</label>
                            <input type="text" name="FM_dt_from" value="<%=dt_fromTxt%>" placeholder="dd-mm-yyyy" />
                        </div>

                         <div class="form-group col-s-4 no-pad-left">
                            <label>To:</label>
                            <input type="text" name="FM_dt_to" value="<%=dt_toTxt%>" placeholder="dd-mm-yyyy" />
                        </div>

                         <div class="form-group col-s-4 no-pad-left">
                            <label>Append to:</label>
                             <select name="FM_append_to" onchange="submit();">
                                 <option value="-1">Ignore dates</option>

                                 <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                           <option value="8" <%=appto_8Sel %>>Order date</option>
                            <option value="14" <%=appto_14Sel %>>ETD Buyer</option>

                                 <%end if %>

                                 <%if cint(rapporttype) = 1 then  %>
                        
                            <option value="2" <%=appto_2Sel %>>Proto DL</option>
                            <option value="3" <%=appto_3Sel %>>Photo Buyer DL</option>
                            <option value="12" <%=appto_12Sel %>>Photo Supp. DL</option>
                            <option value="4" <%=appto_4Sel %>>SMS Buyer DL</option>
                            <option value="13" <%=appto_13Sel %>>SMS Supp. DL</option>

                             <option value="6" <%=appto_6Sel %>>PP App</option>
                             <option value="7" <%=appto_7Sel %>>SHS App</option>
                               <option value="14" <%=appto_14Sel %>>ETD Buyer</option>

                                 <%end if %>

                                 <%if cint(rapporttype) = 3 then  %>
                            <option value="15" <%=appto_15Sel %>>ETD Suppl.</option>

                                 <%end if %>


                                 </select>
                        </div>

                       
                          <div class="form-group col-s-8 no-pad-left"><br />
                             <label>Status:</label>
                            <br />
                              <label><input type="checkbox" name="FM_status3" value="3" <%=statusCHK3 %>>Enquiry</label>
                            <label><input type="checkbox" name="FM_status1" value="1" <%=statusCHK1 %>>Active orders</label>
                            <label><input type="checkbox" name="FM_status2" value="2" <%=statusCHK2 %>>Shipped orders</label>
                            <label><input type="checkbox" name="FM_status0" value="0" <%=statusCHK0 %>>Invoiced/closed</label>
                            <label><input type="checkbox" name="FM_status4" value="4" <%=statusCHK4 %>>Cancelled</label>
                                  
                        </div>


                        <%  '**** Sales rep ***'
                       
                            call salesreplist(salesrep)
                        %>
                        
                          <div class="form-group col-s-3 no-pad-left"><br />
                            <label>Sales Rep.:</label> <select name="salesrep" style="width:240px;" onchange="submit();">
                           <%=strFil_Kpers_Txt %>
                            </select>
                             </div>


                           <div class="form-group col-s-8 no-pad-left"><br />
                            <label>Supplier:</label><select name="supplier" onchange="submit();">
                                  <option value="0">Choose..</option>
                               <%   strSQL = "SELECT kid, kkundenavn FROM kunder WHERE useasfak = 6 ORDER BY kkundenavn"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
        
                                    if cdbl(supplier) = oRec("kid") then
                                    ssel = "SELECTED"
                                    else
                                    ssel = ""
                                    end if

                             %>
                                  <option value="<%=oRec("kid")%>" <%=ssel %>><%=oRec("kkundenavn") %></option>
                                  <%
            
				
			                oRec.movenext
			                wend
			                oRec.close
%>
                                  </select>
                        </div>



                          <div class="form-group col-s-4 no-pad-left"><br />
                            <label>Order view:</label> <select name="rapporttype" style="width:240px;" onchange="submit();">
                            <option value="0" <%=rapporttype0SEL %>>Overview</option>
                            <option value="1" <%=rapporttype1SEL %>>Production / Enq. Overview</option>
                            <option value="3" <%=rapporttype3SEL %>>Overview (ext.)</option>
                            </select>
                              </div>


                       

                    
                       

                    </form>
                </section>
            </div><!--search-->

                     <div id="print" class="col-s-12 col-m-4">
                        <section class="panel">
                            <header class="panel-heading">Tools & Functions</header>

                            <div class="panel-body">

                                  <div class="col-s-12 no-pad-hoz">
                            
                             
                                     <div class="btn-group"><a href="job_nt.asp?media=exp" target="_blank">Export .CSV</a><br /><br />
                                    <a href="job_nt.asp?media=print" target="_blank">Print</a></div>

                                      <br /><br />
                                       <div class="btn-group">
                                       <%
                                        nWdt = 120
                                        nTxt = "New Order"
                                        nLnk = "job_nt.asp?func=opret"
                                        nTgt = ""
                                        call opretNy_2013(nWdt, nTxt, nLnk, nTgt) %>
                                            </div>
                                </div>

                            </div>



                        </section>
                         </div>

            <!--SEARCH END -->

            <%end if 'media %>

        



            <% 
                
             
            antal_orders = 0    
                
             if media <> "exp" then %>
            <!--TABLE START -->
            
                
               <div id="list" class="col-s-12">               
                <%
                
                call tableheader %>
              
                           
            <%end if %>



<%

'if len(trim(sogVal)) <> 0 then
'    sogValSQL = " AND (k.kkundenavn LIKE '"& sogVal &"%' OR jobnavn LIKE '"& sogVal &"%' OR jobnr LIKE '"& sogVal &"%' )"
'else
'    sogValSQL = ""
'end if


call basisValutaFN()
basisValKursUse = replace(basisValKurs, ".", ",")


'************************************************************************
'****************** S�ge funktion , Trim ********************************
'************************************************************************
if len(trim(sogVal)) <> 0 then 
                                    

                 


                    if instr(sogVal, ",") <> 0 then
                                    
                    sogValArr = split(sogVal, ",")
                                    

                                   

                    for j = 0 TO UBOUND(sogValArr)

                    sogValTxt = trim(sogValArr(j)) 

                    if j = 0 then
                        strsogValKri = " AND ((k.kkundenavn LIKE '%"& sogValTxt &"%' OR k.kkundenr = '"& sogValTxt &"') "
                    else
                        if j = 1 then
                        strsogValKri = strsogValKri & " AND ((jobnr LIKE '"& sogValTxt &"%' OR jobnavn LIKE '%"& sogValTxt &"%' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                        else
                        strsogValKri = strsogValKri & " OR (jobnr LIKE '"& sogValTxt &"%' OR jobnavn LIKE '%"& sogValTxt &"%' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                        end if
                    end if

                    next


                    strsogValKri = strsogValKri & "))"
                                    
                    else

                        if instr(sogVal, ">") <> 0 then
                        sogValTxt = trim(replace(sogVal, ">", "")) 
                            call erDetInt(SQLBless(trim(sogValTxt)))

                            if len(trim(sogValTxt)) > 0 AND isInt = 0 then
                            strsogValKri = " AND (jobnr > "& sogValTxt &")"
                            else
                                strsogValKri = " AND (jobnr < 0)"
                            end if 
                        else

                            if instr(sogVal, "<") <> 0 then
                                sogValTxt = trim(replace(sogVal, "<", ""))
                                            
                                call erDetInt(SQLBless(trim(sogValTxt)))
				                           
                                      
                                if len(trim(sogValTxt)) > 0 AND isInt = 0 then
                                strsogValKri = " AND (jobnr < "& sogValTxt &")"
                                else
                                strsogValKri = " AND (jobnr < 0)"
                                end if 
                            
                            else	



                                    '**** Finder jobnr p� fakturaer vhsi der er faktureret *****
                                    fakfundet = 0
                                    strSogFaknrJobids = " OR (jobnr = -1"
                                    strSQLfakjob = "SELECT fid, j.jobnr FROM fakturaer AS f LEFT JOIN job AS j ON (j.id = f.jobid) WHERE faknr LIKE '"& sogVal &"%' AND j.jobnr IS NOT NULL "
                    
                                    'response.Write strSQLfakjob
                                    'response.flush
    
    
                                    oRec4.open strSQLfakjob, oConn, 3
                                    while not oRec4.EOF 

                                    strSogFaknrJobids = strSogFaknrJobids & " OR jobnr = "& oRec4("jobnr")

                                    fakfundet = 1
                                    oRec4.movenext
                                    wend
                                    oRec4.close
                 

                                    if cint(fakfundet) = 1 then
                                    strSogFaknrJobids = strSogFaknrJobids & ")"
                                    else
                                    strSogFaknrJobids = ""
                                    end if
        
                                    'response.end


                                    	
		                    strsogValKri = " AND (jobnr LIKE '"& sogVal &"%' "& strSogFaknrJobids &" OR jobnavn LIKE '%"& sogVal &"%' OR k.kkundenavn LIKE '%"& sogVal &"%' OR k.kkundenr = '"& sogVal &"' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                                    
                            end if
                    end if
                    end if
		else
		strsogValKri = ""
		end if				


    sogValSQL = strsogValKri 

    '******************************************************


strSQLjobsel = "SELECT j.id, jobnavn, jobnr, k.kkundenavn, k.kid, k2.kkundenavn AS suppliername, jobstatus, supplier, mg.navn AS mgruppenavn, rekvnr, supplier_invoiceno, fastpris, "

if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
strSQLjobsel = strSQLjobsel &" jobstartdato, orderqty, shippedqty, product_group, kundekpers, m.mnavn AS salesrep, m.init, collection,"
strSQLjobsel = strSQLjobsel &" comm_pc, jo_dbproc, destination, dt_confb_etd, dt_confs_etd, dt_actual_etd, sales_price_pc, sales_price_pc_valuta, cost_price_pc, cost_price_pc_valuta, freight_pc, tax_pc,"
end if

if cint(rapporttype) = 1 then
strSQLjobsel = strSQLjobsel &" dt_proto_dead, dt_sms_dead, dt_photo_dead, dt_sour_dead, dt_proto_dead, dt_ppapp, dt_shsapp, dt_sup_photo_dead, dt_sup_sms_dead, dt_confb_etd, dt_confs_etd,"
end if
   
     
strSQLjobsel = strSQLjobsel &" jo_bruttooms, jo_udgifter_intern, alert"_
&" FROM job AS j "_
&" LEFT JOIN kunder AS k ON (k.kid = j.jobknr)"_
&" LEFT JOIN kunder AS k2 ON (k2.kid = j.supplier)"

strSQLjobsel = strSQLjobsel &" LEFT JOIN materiale_grp AS mg ON (mg.id = j.product_group)"

if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
strSQLjobsel = strSQLjobsel &" LEFT JOIN medarbejdere AS m ON (m.mid = j.jobans1)"
end if

strSQLjobsel = strSQLjobsel &" WHERE j.id <> 0 "& sogValSQL &" "& strStatusSQL &" "& strSQLdtKri &" "& salesrepSQL &" "& supplierSQL &" ORDER BY "& strSQLOdrBy &" LIMIT 2000" 


'if session("mid") = 1 then
'response.write strSQLjobsel
'response.Flush    
'end if

  jo_dbproc_bel_tot = 0
  jo_bruttooms_tot = 0
  jo_cost_tot = 0

  orderqtyTot = 0 
  shippedqtyTot = 0

  ordertypeTxt = ""
  strExpTxt = ""


oRec.open strSQLjobsel, oConn, 3
while not oRec.EOF




    if antal_orders > 0 AND media = "" then

    select case (antal_orders)
    case 15, 30, 45, 60, 75, 100, 115, 130, 145, 160, 175, 200, 215, 230, 245, 260, 275, 280, 305, 320
   



       
            %>
            <!--TABLE START -->
            <!-- 
            <div id="list" class="col-s-12">-->
            <%'call tableheader %>
            
               
                           
             

 <% end select
 end if 'antal_orders



    'if cint(antal_orders) = 0 then
    %>
                 <input type="hidden" id="fakhref_<%=oRec("id") %>" value="../timereg/erp_opr_faktura_fs.asp?visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" />
                                                        
    <%'end if



      if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
    
     if oRec("orderqty") <> 0 then
    orderqty = oRec("orderqty")
    orderqtyTot = orderqtyTot + oRec("orderqty")  
    else
    orderqty = ""
    end if

    

     if oRec("shippedqty") <> 0 then
    shippedqty = oRec("shippedqty")
    shippedqtyTot = shippedqtyTot + oRec("shippedqty")
    else
    shippedqty = ""
    end if

      if instr(oRec("jobstartdato"), "2010") <> 0 then
    dt_jobstartdato = ""
    else
    dt_jobstartdato = replace(formatdatetime(oRec("jobstartdato"), 2), "-", ".")
    end if

    

     end if

    if cint(rapporttype) = 1 then
    
    if instr(oRec("dt_sour_dead"), "2010") <> 0 then
    dt_sour_dead = ""
    else
    dt_sour_dead = replace(formatdatetime(oRec("dt_sour_dead"), 2), "-", ".") ' oRec("dt_sour_dead")
    end if

     if instr(oRec("dt_proto_dead"), "2010") <> 0 then
    dt_proto_dead = ""
    else
    dt_proto_dead = replace(formatdatetime(oRec("dt_proto_dead"), 2), "-", ".")
    end if
     

     if instr(oRec("dt_photo_dead"), "2010") <> 0 then
    dt_photo_dead = ""
    else
    dt_photo_dead = replace(formatdatetime(oRec("dt_photo_dead"), 2), "-", ".") 
    end if
    
    
    if instr(oRec("dt_sms_dead"), "2010") <> 0 then
    dt_sms_dead = ""
    else
    dt_sms_dead = replace(formatdatetime(oRec("dt_sms_dead"), 2), "-", ".")
    end if
   



    if instr(oRec("dt_proto_dead"), "2010") <> 0 then
    dt_proto_dead = ""
    else
    dt_proto_dead = replace(formatdatetime(oRec("dt_proto_dead"), 2), "-", ".") 
    end if
     

     if instr(oRec("dt_ppapp"), "2010") <> 0 then
    dt_ppapp = ""
    else
    dt_ppapp = replace(formatdatetime(oRec("dt_ppapp"), 2), "-", ".")
    end if
    
    
     if instr(oRec("dt_shsapp"), "2010") <> 0 then
    dt_shsapp = ""
     else
    dt_shsapp = replace(formatdatetime(oRec("dt_shsapp"), 2), "-", ".") 'oRec("dt_shsapp")
    end if
   

     if instr(oRec("dt_sup_photo_dead"), "2010") <> 0 then
    dt_sup_photo_dead = ""
     else
    dt_sup_photo_dead = replace(formatdatetime(oRec("dt_sup_photo_dead"), 2), "-", ".") 'oRec("dt_sup_photo_dead")
    end if

     if instr(oRec("dt_sup_sms_dead"), "2010") <> 0 then
    dt_sup_sms_dead = ""
     else
    dt_sup_sms_dead = replace(formatdatetime(oRec("dt_sup_sms_dead"), 2), "-", ".") 'oRec("dt_sup_sms_dead")
    end if

       

    



    end if



    
     if cint(rapporttype) = 1 OR cint(rapporttype) = 3 then

      if instr(oRec("dt_confs_etd"), "2010") <> 0 then
    dt_confs_etd = ""
     else
    dt_confs_etd = oRec("dt_confs_etd")
    end if

    end if

 


    if oRec("jo_udgifter_intern") <> 0 then
    jo_udgifter_intern = formatnumber(oRec("jo_udgifter_intern"), 2) 
    jo_udgifter_internTxt = formatnumber(oRec("jo_udgifter_intern"), 2) & " DKK"
    else
    jo_udgifter_intern = ""
    jo_udgifter_internTxt = ""
    end if

    if oRec("jo_bruttooms") <> 0 then
    jo_bruttoomsTxt = formatnumber(oRec("jo_bruttooms"), 2) & " DKK"
    jo_bruttooms = formatnumber(oRec("jo_bruttooms"), 2)
    else
    jo_bruttoomsTxt = ""
    jo_bruttooms = ""
    end if



    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then

        
        if oRec("sales_price_pc") <> 0 then

        call valutakode_fn(oRec("sales_price_pc_valuta"))
        sales_price_pc_val = valutaKode_CCC 
        sales_price_pc = formatnumber(oRec("sales_price_pc"), 2) 
        else
        sales_price_pc = ""
        sales_price_pc_val = ""
        end if
        
    
    
         if oRec("cost_price_pc") <> 0 then
        
        call valutakode_fn(oRec("cost_price_pc_valuta"))
        cost_price_pc_val = valutaKode_CCC
        cost_price_pc = formatnumber(oRec("cost_price_pc"), 2) 
        else
        cost_price_pc = ""
        cost_price_pc_val = ""
        end if
        


        if oRec("comm_pc") <> 0 then
        comm_pc = formatnumber(oRec("comm_pc"), 2) & " %" 
        else
        comm_pc = ""
        end if


        if oRec("cost_price_pc") <> 0 then

        valutaKursOR3(oRec("sales_price_pc_valuta"))
        salgsprisKurs = dblKursOR3

        valutaKursOR3(oRec("cost_price_pc_valuta"))
        kostprisKurs = dblKursOR3

        tax_pc_calc = 0
        tax_pc_calc = (oRec("cost_price_pc")/1 * (oRec("tax_pc")/100))
        profit_pc = formatnumber((oRec("sales_price_pc") * (salgsprisKurs/100)) - ((oRec("cost_price_pc") + tax_pc_calc + oRec("freight_pc")) * (kostprisKurs/100)) * basisValKursUse/100, 2) & " "& basisValISO '& " ("& basisValKursUse &")"
        else
        profit_pc = ""
        end if
       
        
        if oRec("jo_dbproc") <> 0 then
        jo_dbproc = formatnumber(oRec("jo_dbproc"), 2) & " %"
        else
        jo_dbproc = ""
        end if

        
        if oRec("destination") <> "0" then
        destination = oRec("destination")
        else
        destination = ""
        end if
        
       
        if instr(oRec("dt_actual_etd"), "2010") <> 0 then
        dt_actual_etd = ""
        else
        dt_actual_etd = oRec("dt_actual_etd")
        end if
        


    end if


     if instr(oRec("dt_confb_etd"), "2010") <> 0 then
    dt_confb_etd = ""
     else
    dt_confb_etd = oRec("dt_confb_etd")
    end if


     if jo_bruttooms <> "" AND jo_udgifter_intern <> "" then
     jo_dbproc_bel = formatnumber((oRec("jo_bruttooms") - oRec("jo_udgifter_intern")), 2)
     jo_dbproc_bel_tot = jo_dbproc_bel_tot + jo_dbproc_bel

     jo_dbproc_bel = jo_dbproc_bel 
     jo_dbproc_belTxt = jo_dbproc_bel & " DKK"
     else
     jo_dbproc_bel = ""
     jo_dbproc_belTxt = "" 
     end if
  
    jo_cost_tot = jo_cost_tot + oRec("jo_udgifter_intern")
    jo_bruttooms_tot = jo_bruttooms_tot + oRec("jo_bruttooms")
    

   
    select case oRec("jobstatus")
    case 0
    jobstatusTxt = "Closed"
    case 1
    jobstatusTxt = "Active"
    case 2
    jobstatusTxt = "Shipped"
    case 3
    jobstatusTxt = "Enquiry"
    case 4
    jobstatusTxt = "Review"
    end select

    select case oRec("fastpris")
    case 2
    ordertypeTxt = "Commission"
    case 3
    ordertypeTxt = "Salesorder"
    end select
                                  


    if media = "exp" then

    strExpTxt = strExpTxt & jobstatusTxt & ";" & oRec("kkundenavn") & ";"& oRec("rekvnr") &";"& oRec("suppliername") &";"& oRec("supplier_invoiceno") &";"& oRec("mgruppenavn") &";"& oRec("jobnavn") &";" & oRec("jobnr") & ";"& ordertypeTxt &";" 
    

     if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & dt_jobstartdato &";"& oRec("destination") &";"& oRec("collection") &";"& oRec("salesrep") &";"& dt_confb_etd &";"& dt_actual_etd &";" 
    end if


     if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt  & dt_confs_etd & ";"
    end if


      if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & orderqty & ";" & shippedqty & ";" 
    end if

    if cint(rapporttype) = 1 then 
    strExpTxt = strExpTxt  & dt_proto_dead & ";" & dt_photo_dead & ";"& dt_sup_photo_dead  &";" & dt_sms_dead & ";" & dt_sup_sms_dead & ";" & dt_ppapp & ";" & dt_shsapp &";" & dt_confb_etd & ";" & dt_confs_etd & ";"
    end if


    

     if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & cost_price_pc & ";" & cost_price_pc_val & ";"
    end if


    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & sales_price_pc & ";" & sales_price_pc_val & ";"
    end if

    if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & comm_pc & ";"& profit_pc & ";" & jo_udgifter_intern & ";DKK;"
    end if

    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & jo_bruttooms & ";DKK;"
    end if

    if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & jo_dbproc_bel & ";DKK;" & jo_dbproc & ";"
    end if




    strExpTxt = strExpTxt & ";;;xx99123sy#z"

    end if

                             if media <> "exp" then

                             if cdbl(lastid) = oRec("id") then
                             trbgcol = "#FFFFe1"
                             else
                             trbgcol = ""
                             end if
                             %>
                           
                             <tr style="background-color:<%=trbgcol%>;">
                                 <td style="width:<%=tdWdthTBD%>px;">
                                    <%=left(jobstatusTxt, 3) %>.
                                 </td>
                                 
                                 <td style="width:<%=tdWdthTBC%>px; text-align:left;">
                                    <%if media <> "print" then %>
                                    <a href="kunder.asp?func=red&id=<%=oRec("kid") %>"><%=left(oRec("kkundenavn"), 12) %></a>
                                    <%else %>
                                    <%=oRec("kkundenavn") %>
                                    <%end if %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(oRec("collection"), 15) %></span> 
                                </td>
                                 <td style="width:<%=tdWdthTBC%>px;"><%=oRec("rekvnr") %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(destination, 3)%></span> 
                                     
                                 </td>

                                  <td style="width:<%=tdWdthTBA%>px;"><%=left(oRec("suppliername"), 10) %>

                                      <%if oRec("jobstatus") = 2 AND len(trim(oRec("supplier_invoiceno"))) <> 0 then %>
                                      <br /><span style="font-size:9px; color:#999999;"><%=oRec("supplier_invoiceno") %></span>
                                      <%end if %>
                                  </td>

                                 
                                <!-- <td style="white-space:nowrap; width:<%=tdWdthTBC%>px; text-align:left;"><%=left(oRec("mgruppenavn"),10) %></td>-->

                                <td style="width:<%=tdWdthTBC%>px; text-align:left;">
                                    <%if cint(oRec("alert")) = 1 then %>
                                    <span style="color:red;"><b>!</b>&nbsp;</span>
                                    <%end if %>

                                    <%if media <> "print" then %>
                                    <a href="job_nt.asp?func=red&jobid=<%=oRec("id")%>"><%=left(oRec("jobnavn"), 12) %></a>
                                     <%else %>
                                    <%=oRec("jobnavn") %>
                                    <%end if %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(oRec("mgruppenavn"),15) %></span> 
                                    

                                </td>
                                
                                
                                <td style="width:<%=tdWdthTBA%>px;">
                                     <%if media <> "print" then %>
                                    <a href="job_nt.asp?func=red&jobid=<%=oRec("id")%>"><%=oRec("jobnr") %></a>
                                    <br /><span style="font-size:9px; color:#999999;"><%=ordertypeTxt %></span> 
                                    <%else %>
                                    <%=oRec("jobnr") %>
                                    <%end if %>
                                </td>
                                
                                

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_jobstartdato%></td>
                                 <!--<td style="white-space:nowrap; width:<%=tdWdthTBB%>px;"><%=left(destination, 3)%></td>-->
                                 
                                 <!--<td style="white-space:nowrap; width:<%=tdWdthTBA%>px;"><%=oRec("collection") %></td>-->
                                 
                                 <td style="white-space:nowrap; width:<%=tdWdthTBD%>px;"><%=oRec("init") %></td>
                                 
                                
                            
                                <%end if %>


                                   <%if cint(rapporttype) = 1 then %>
                                
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_proto_dead %></td>
                                <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_photo_dead %></td>
                                  <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_sup_photo_dead %></td>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_sms_dead%></td>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_sup_sms_dead%></td>
                                  <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_ppapp%></td>
                                   <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_shsapp%></td>
                                  <%end if 'rapporttype %>

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 1 OR cint(rapporttype) = 3 then%>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_confb_etd %></td>
                                  <%end if 'rapporttype %>

                                  <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then%>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_actual_etd %></td>
                                  <%end if 'rapporttype %>


                                  <%if cint(rapporttype) = 1 OR cint(rapporttype) = 3 then %>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBF%>px;"><%=dt_confs_etd %></td>
                                 <%end if 'rapporttype %>
                                 
                                  <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBA%>px;"><%=orderqty %></td>
                                 <td style="white-space:nowrap; width:<%=tdWdthTBA%>px;"><%=shippedqty %></td>
                                 <%end if %>


                                

                       

                                  <%if cint(rapporttype) = 3 then  %>
                                 <td style="width:<%=tdWdthTBE%>px;"><%=cost_price_pc &" "& cost_price_pc_val%></td>
                                 <%end if %>

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                                <td style="width:<%=tdWdthTBA%>px;"><%=sales_price_pc &" "& sales_price_pc_val%></td>
                                 <%end if %>

                                  <%if cint(rapporttype) = 3 then  %>
                                  <td style="width:<%=tdWdthTBE%>px;"><%=comm_pc%></td>
                                 <td style="width:<%=tdWdthTBE%>px;"><%=profit_pc%></td>
                                <td style="width:<%=tdWdthTBE%>px;"><%=jo_udgifter_internTxt%></td>
                                 <%end if %>

                                 <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                                <td style="width:<%=tdWdthTBA%>px;"><%=jo_bruttoomsTxt%></td>
                                 <%end if %>

                                  <%if cint(rapporttype) = 3 then  %>
                                
                                <td style="width:<%=tdWdthTBE%>px;"><%=jo_dbproc_belTxt%></td>
                                 <td style="width:<%=tdWdthTBE%>px;"><%=jo_dbproc%></td>
                                   <%end if %>


                                  


                                    <td style="white-space:nowrap; font-size:9px; width:<%=tdWdthTBA%>px; overflow:hidden;"><%if media <> "print" then %> 


                                        
                                                        <%'if ((level <= 2 OR level = 6)) then 'AND cint(oRec("jobstatus")) = 2) then 'shipped 
                                                            jobids = jobids & ","& oRec("id")
                                                            jobnrs = jobnrs & ","& oRec("jobnr")
                                                            %>
                                                      
                                                        <!--<a href="../timereg/erp_opr_faktura_fs.asp?visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" class=rmenu>Create Invoice >> </a>-->
                                                            
                                                            <%strSQLFakhist = "SELECT faknr, beloeb, fid AS fid, betalt, shadowcopy FROM fakturaer WHERE jobid = "& oRec("id") &" LIMIT 30"
                                                                
                                                                oRec6.open strSQLFakhist, oConn, 3
                                                                f = 0
                                                                while not oRec6.EOF

                                                                if f = 0 then
                                                                %>

                                                                <!--
                                                                <br />Invoices:
                                                                <span style="color:#999999; font-size:9px;">Not approved invoices<br /> will be deleted on create new invoice.</span>
                                                                -->
                                                                <%
                                                                end if

                                                                
                                                                     if oRec6("shadowcopy") <> 1 then 

                                                                         if oRec6("betalt") <> 1 then %> 
                                                                        <br /><a href="../timereg/erp_opr_faktura_fs.asp?func=red&id=<%=oRec6("fid")%>&visfaktura=2&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" style="color:#999999; font-size:9px;"><%=oRec6("faknr")%></a> 
                                                                                              
                                                                        <%else %>
                                                                        <br /><a href="../timereg/erp_opr_faktura_fs.asp?func=red&id=<%=oRec6("fid")%>&visfaktura=2&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" style="color:green; font-size:9px;"><i>V</i>&nbsp;&nbsp;<%=oRec6("faknr")%></a> 
                                                                   
                                                                        <%end if%>


                                                                     <%else 
                                                                         
                                                                        if oRec6("betalt") <> 1 then  %>
                                                                        <br /><span style="color:#999999; font-size:10px;">(<%=oRec6("faknr")%>)</span> 
                                                                        <%else %>
                                                                       
                                                                         <br /><span style="color:green; font-size:10px;"><i>V</i>&nbsp; <%=oRec6("faknr")%></span> 
                                                                        <%end if%>


                                                                    <%end if%>
                                                                    
                                                                     <%if oRec6("shadowcopy") <> 1 then %>
                                                                     <%'formatnumber(oRec6("beloeb"), 0) %>
                                                                     <%end if %>
                                                                
                                                                <%
                                                                f = f + 1
                                                                oRec6.movenext
                                                                wend 
                                                                oRec6.close%>            

                                        
                                                        <%'else %>
                                                        &nbsp;
                                                        <%'end if %>

                                                        

                                     <%else %>
                                     &nbsp;
                                     <%end if %>
                                 </td>

                             <%if media <> "print" then %>
                                 <td style="width:<%=tdWdthTBB%>px;"><a href="job_nt.asp?func=slet&id=<%=oRec("id") %>" style="color:red;">X</a></td>
                                 
                                 <td style="width:<%=tdWdthTBB%>px;"><input type="checkbox" value="<%=oRec("id") %>" id="bulk_jobid_<%=oRec("id") %>" name="FM_bulk_jobid" class="bulk_jobid" /></td>
                                <%end if %>

                            </tr>
                            
                            <%end if 'media %>
                       



<%
		
    antal_orders = antal_orders + 1

oRec.movenext
wend
oRec.close


    if media <> "exp" then
                
    %>
                              <td><b>Total:</b></td>
                            <td><%=antal_orders %></td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           <!-- prod grp  <td>&nbsp;</td>-->
                            <!-- DEST <td>&nbsp;</td>-->

                            <%if cint(rapporttype) = 0 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <%end if %>

                        
                            <!-- Collection <td>&nbsp;</td>-->
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            

                
                             <%if cint(rapporttype) = 0 then %>
                            <td>&nbsp;</td>
                            <%end if %>   
                         

                             <%if cint(rapporttype) = 1 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            
                            <%end if %>

                            
                               <%if cint(rapporttype) = 3 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           <td>&nbsp;</td>

                            
                            <%end if %>


                            <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                            <td>&nbsp;</td>
                            <td><%=formatnumber(orderqtyTot, 0) %></td>
                            <td><%=formatnumber(shippedqtyTot, 0) %></td>
                            <%end if %>

                            
                                   <%if cint(rapporttype) = 3 then %>
                             <td>&nbsp;</td>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           
                            
                            <%end if %>

                            
                             <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                          
                            <td><%=formatnumber(jo_cost_tot, 2) & " DKK" %></td>
                            <td><%=formatnumber(jo_bruttooms_tot, 2) & " DKK" %></td>
                            <%end if %>

                             <%if cint(rapporttype) = 0 then %>
                            <td>&nbsp;</td>
                            <%end if %>        


                             <%if cint(rapporttype) = 3 then %>
                            <td><%=formatnumber(jo_dbproc_bel_tot, 2) & " DKK" %></td>
                            <td>&nbsp;</td>
                             
                            
                            <%end if %>
                    
                          

                              <%if cint(rapporttype) = 3 then %>
                              <td>&nbsp;</td>
                            <%end if %>

                            <%if cint(rapporttype) = 1 then %>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            <%end if %>

                         <%if media <> "print" then %>
                         <td>&nbsp;</td> 
                        <td>&nbsp;</td>
                         <%end if %>

                       
                           
                           
                             </tbody>
                    </table><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />&nbsp;
               
              <!--  </section> -->
            
            </div><!--Table body div -->   
               
               
            </div><!-- list div -->
            <!--TABLE END -->

               <div id="dv_grandtotal" style="position:absolute; top:300px;left:880px;">
                 <section class="panel">
                    <header class="panel-heading">Totals for orders on list:</header>

                    <div class="panel-body">
                    <div id="dv_topGt">
                        
                        <table width="100%">
                            <tr>
                                <td>Order Qty.</td>
                                <td>Shipped Qty.</td>
                                <td>Total Costprice</td>
                                <td>Total Salesprice</td>
                                <%if cint(rapporttype) = 3 then %>
                                <td>Total Profit</td>
                                <%end if %>
                            </tr>
                            <tr>
                                <td><%=formatnumber(orderqtyTot, 0) %></td>
                                <td><%=formatnumber(shippedqtyTot, 0) %></td>
                                <td><%=formatnumber(jo_cost_tot, 2) & " DKK" %></td>
                                <td><%=formatnumber(jo_bruttooms_tot, 2) & " DKK" %></td>
                                <%if cint(rapporttype) = 3 then %>
                                <td><%=formatnumber(jo_dbproc_bel_tot, 2) & " DKK" %></td>
                                <%end if %>
                            </tr>
                           

                        </table>

                    </div>
                        </div>
                     </div>
                </section>
                </div>


        <%if media = "" then %>

                    

 </div>
</div>
                       <!--
                       </div class="container">
                       </div class="row">
                           -->
 
                          
        <%end if %>





                             <div id="dv_invoice" style="position:absolute; float:left; left:1200px; top:500px; z-index:2000; border:10px yellowgreen solid; padding:20px; visibility:hidden; display:none; background-color:#FFFFFF;">
                              
                               
                                     <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                     <tr><td style="border:0px;" align="left"><b>Create invoice on selected orders? </b>
                                 <br />
                                   <a id="ainvlink" href="#" target="_blank" class=vmenu>Yes, create Invoice on selected orders >> </a>

                                         <br /><br />
                                         <span style="color:#999999; font-size:9px;">
                                        
                                         Actual ETD, Buyer info, Bankaccount, currency, Invoice note etc.
                                         will always be obtained from the first selected order on the list.
                                             </span>

                                         </td><td valign="top" style="width:40px; border:0px;"><span id="sp_dv_invoice" style="color:#999999; font-size:14px;"><b>X</b></span></td></tr>
                                        </table>               


                                 </div>


                <div id="dv_bulk" style="position:absolute; width:800px; left:100px; top:200px; z-index:2000; border:10px #CCCCCC solid; padding:20px; visibility:hidden; display:none; background-color:#FFFFFF;">
                <form class="panel-body" id="joblist" method="post" action="job_nt.asp?func=bulk">

                     <section class="panel">
                     <input type="hidden" id="bulk_jobids" name="bulk_jobids" value="0" />

                    <header class="panel-heading">Bulk update production info</header>
                          
                    <div class="panel-body">
                      
                        <div class="form-group" style="float:right;">
                              <span style="color:red;" id="bulk_close">[X]</span>
                        </div>


                           <div>&nbsp;<br /><br /><b>Enquery info</b></div>

                         <div class="form-group col-s-4">
                            <label>SMS Buyer deadline</label>
                            <input type="text" name="FM_dt_sms_dead" value="" placeholder="dd-mm-yyyy" />
                        </div>
                         <div class="form-group col-s-4">
                            <label>SMS Supplier DL</label>
                            <input type="text" name="FM_dt_sup_sms_dead" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-4">
                            <label>SMS sent</label>
                            <input type="text" name="FM_dt_sms_sent" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-4">
                            <label>Photo Buyer Deadline</label>
                            <input type="text" name="FM_dt_photo_dead" value="" placeholder="dd-mm-yyyy" />
                        </div>
                          <div class="form-group col-s-4">
                            <label>Photo Supplier DL</label>
                            <input type="text" name="FM_dt_sup_photo_dead" value="" placeholder="dd-mm-yyyy" />
                        </div>

                        <div class="form-group col-s-4 no-pad-right">
                            <label>Photo sent</label>
                            <input type="text" name="FM_dt_photo_sent" value="" placeholder="dd-mm-yyyy" />
                        </div>


                        <div>&nbsp;<br /><br /><b>Order Dates</b></div>

                          <div class="form-group col-s-3">
                            <label>Confirmed buyer ETD</label>
                            <input type="text" name="FM_dt_confb_etd" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-3">
                            <label>Confirmed buyer ETA</label>
                            <input type="text" name="FM_dt_confb_eta" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-3">
                            <label>Confirmed supplier ETD</label>
                            <input type="text" name="FM_dt_confs_etd" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-3">
                            <label>Conf. suppl. ETA</label>
                            <input type="text" name="FM_dt_confs_eta" value="" placeholder="dd-mm-yyyy" />
                        </div>
                       
                        <div class="form-group col-s-6">
                            <label>Actual ETD</label>
                            <input type="text" name="FM_dt_actual_etd" value="" placeholder="dd-mm-yyyy" />
                        </div>
                        <div class="form-group col-s-6">
                            <label>Actual ETA</label>
                            <input type="text"  name="FM_dt_actual_eta" value="" placeholder="dd-mm-yyyy" />
                        </div>

                        
                        <div>&nbsp;<br /><br /><b>Production dates</b></div>

                        <div class="form-group col-s-3">
                            <label>1st. order comment</label>
                            <input type="text" name="FM_dt_firstorderc" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>LD app</label>
                            <input type="text" name="FM_dt_ldapp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>Exp Sizeset</label>
                            <input type="text" name="FM_dt_sizeexp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>Sizeset app</label>
                            <input type="text" name="FM_dt_sizeapp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>Exp PP</label>
                            <input type="text" name="FM_dt_ppexp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>PP app</label>
                            <input type="text" name="FM_dt_ppapp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>Exp SHS</label>
                            <input type="text" name="FM_dt_shsexp" value="" placeholder="dd-mm-yyyy"/>
                        </div>
                        <div class="form-group col-s-3">
                            <label>SHS app</label>
                            <input type="text" name="FM_dt_shsapp" value="" placeholder="dd-mm-yyyy"/>
                        </div>

                         <div class="form-group col-s-12 no-pad-right">
                            <input type="submit" value="Submit" />
                        </div>
                       
               
                </section>

                     </form> <!-- bulk -->


                </div>

    <%end if %>
   


            <%

               

                if media = "exp" then

                strExpTxt = replace(strExpTxt, "xx99123sy#z", vbcrlf)
	
	         
	            call TimeOutVersion()
	
	            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_nt.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
                
                strOskrifter = "Status;Buyer;Buyer PO no.;Supplier;Sup. Inv. no.;Product grp.;Style;Order No.;Order type;"
                

                if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "Order date;Destination;Collection;Sales Rep.;ETD Buyer;Actual ETD;"
                end if


                 if cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "ETD Suppl.;"
                end if

                  if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "Order Qty.;Shipped Qty.;"
                end if
                

                if cint(rapporttype) = 1 then 
                strOskrifter = strOskrifter & "Proto DL;Photo Buyer DL;Photo Suppl. DL;SMS Buyer DL;SMS Suppl. DL;PP App;SHS App;ETD Buyer;ETD Suppl.;"
                end if

                
               


                 
                if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Cost Price PC;Val;"
                end if

                if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Sales Price PC;Val;"
                end if

                 if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Commision PC;Profit PC;Total Cost Price;Val;"
                end if

                  if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Total Sales Price;Val;"
                end if

                 if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Total Profit;Val;Total Profit %;"
                end if

                


				objF.WriteLine(strOskrifter)
				objF.WriteLine(strExpTxt)
				objF.close


                response.redirect "../inc/log/data/"&file

                %>
                 
                 <!--
                <div style="position:absolute; left:90px; top:100px; width:400px; padding:20px; background-color:#FFFFFF; border:0px;">
                       <img src="../ill/outzource_logo_200.gif" /><br /><br /><br />
	              <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Your CSV. file is ready >></a>
                -->

                   

                   
                        </div>
                    
                <%

                end if'media

                if media = "print" then
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                end if


    
 end select %>
<!--#include file="../inc/regular/footer_inc.asp"-->