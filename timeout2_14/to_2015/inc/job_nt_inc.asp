
<%
    public fmatDate
    function fmatDate_fn(dt)

    fmatDate = replace(dt, ".", "-")
    fmatDate = replace(fmatDate, "/", "-")
    fmatDate = replace(fmatDate, ":", "-")
    fmatDate = replace(fmatDate, " ", "")
    

    if isDate(fmatDate) then
    fmatDate = year(fmatDate)&"-"&month(fmatDate)&"-"&day(fmatDate)
    else
    fmatDate = "2010-01-01"
    end if

    end function

    public tdWdthA, tdWdthB, tdWdthC, tdWdthD, tdWdthE, tdWdthF, tdWdthTBA, tdWdthTBB, tdWdthTBC, tdWdthTBD, tdWdthTBE, tdWdthTBF
 
    sub tableheader
    
    
        if media <> "print" then 
        tblBorder = 0
        else 
        tblBorder = 0
        end if


     


                if antal_orders = 0 then 

                   %>
                         <%if media <> "print" then %>
                        <div class="dataTables_wrapper" style="padding:20px;">
                        <table id="tb_jobliste" class="dataTable table table-hover table-bordered table-striped ui-datatable"> <!-- table display table-stripe dataTable table-bordered table-stripe table-hover table-stripe table-bordered ui-datatable -->
                            <!--<thead>
                                <tr>
                                    <th>A</th>
                                    <th>B</th>
                                    <th>C</th>
                                    <th>D</th>

                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td>xx</td>
                                    <td>cc</td>
                                    <td>vvv</td>
                                    <td>ee</td>
                                </tr>
                            </tbody>
                            </table>
                            -->
                            
                            
                          <%else %>
                             <table class="table">
                          <%end if %>

                    

                   
                

                <%end if
                    
                    
                     select case cint(rapporttype) 
                     case 0 'Overview
                     tdHeadMainSpan = 14
                     case 1 'Prod / Enquery
                     tdHeadMainSpan = 15 
                     case 3 'Overview ext
                     tdHeadMainSpan = 21
                     end select%>
                    
                       
                       
                        <thead>
                            
                            <tr>
                              
                                <th colspan="<%=tdHeadMainSpan %>" style="text-align:left;" ><%=rapporttypeTxt %></th>

                                 <%if media <> "print" then %>
                                <th colspan="2" style="text-align:right;">  <select id="bulk_jobid_action" class="form-control input-small" style="width:120px;">
                                        <option value="0">Action..</option>
                                        <option value="1">Bulk update</option>
                                            
                                        <option value="2">Create Invoice</option>
                                            

                                    </select> </th>

                                <%else %> <th colspan="2" style="text-align:right;" align="right">&nbsp;</th>

                                <%end if %>
                            </tr>

                      
                            
                            <tr>
                                <th style="white-space:nowrap; vertical-align:bottom;">Sta. <!--<%if media <> "print" then %><a href="job_nt.asp?sort=2">></a><%end if %>--><br /><span style="font-size:9px;">Status</span></th>
                                <th style="text-align:left; white-space:nowrap; vertical-align:bottom;">Buyer  <!--<%if media <> "print" then %><a href="job_nt.asp?sort=1">></a><%end if %>--><br /><span style="font-size:9px;">Collection</span></th></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Order/PO no <!--<%if media <> "print" then %><a href="job_nt.asp?sort=15">></a><%end if %>--><br /><span style="font-size:9px;">Destination</span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Supplier <!--<%if media <> "print" then %><a href="job_nt.asp?sort=16">></a><%end if %>--><br /><span style="font-size:9px;">Sup. inv. no.</span></th>
                                <!--<th >Product grp. <%if media <> "print" then %><a href="job_nt.asp?sort=17">></a><%end if %></th>-->
                                <th style="width:text-align:left;">Style <!--<%if media <> "print" then %><a href="job_nt.asp?sort=7">></a><%end if %> --><br />
                                    <span style="font-size:9px;">Prod.grp <!--<%if media <> "print" then %><a href="job_nt.asp?sort=17">></a><%end if %>--></span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Order No. <!--<%if media <> "print" then %><a href="job_nt.asp?sort=18">></a><%end if %>--><br /><span style="font-size:9px;">Order type</span>
                                </th>

                               

                                    <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                                <th style="white-space:nowrap; vertical-align:bottom;">Order<br /><span style="font-size:9px;">Date</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=11">></a><%end if %>--></th>
                                
                                <!--<th style="width:<%=tdWdthB%>px;">Dest.</th>-->
                                    
                                <!--<th >Coll.<br /><span style="font-size:9px;">Collection</span></th>-->
                             
                                <th style="white-space:nowrap; vertical-align:bottom;">Rep.<br /><span style="font-size:9px;">Sales</span></th>
                                 <%end if %>
                            
                                <%if cint(rapporttype) = 1 then 'Prod / Enquery %>
                            
                               <!--<th>Src DL <%if media <> "print" then %><a href="job_nt.asp?sort=3">></a><%end if %></th>-->
                                <th style="white-space:nowrap; vertical-align:bottom;">Proto<br /><span style="font-size:9px;">Deadline</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=4">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Photo<br /><span style="font-size:9px;">Buyer DL</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=5">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Photo<br /><span style="font-size:9px;">Suppl. DL</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=12">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">SMS<br /><span style="font-size:9px;">Buyer DL</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=6">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">SMS<br /><span style="font-size:9px;">Suppl. DL</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=13">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">PP<br /><span style="font-size:9px;"> App. Date</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=9">></a><%end if %>--></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">SHS<br /><span style="font-size:9px;">App. Date</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=10">></a><%end if %>--></th>
                                <%end if %>


                                <%if cint(rapporttype) = 0 OR cint(rapporttype) = 1 OR cint(rapporttype) = 3 then%>
                                <th style="white-space:nowrap; vertical-align:bottom;">ETD<br /><span style="font-size:9px;">Buyer</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=14">></a><%end if %>--></th>
                                <%end if %>

                                 <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then%>
                                <th style="white-space:nowrap; vertical-align:bottom;">ETD<br /><span style="font-size:9px;">Actual</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=19">></a><%end if %>--></th>
                                <%end if %>

                                 <%if cint(rapporttype) = 1 OR cint(rapporttype) = 3 then 'Prod / Enquery + ext. %>
                                 <th style="white-space:nowrap; vertical-align:bottom;">ETD<br /><span style="font-size:9px;">Suppl. date</span> <!--<%if media <> "print" then %><a href="job_nt.asp?sort=20">></a><%end if %>--></th>
                           
                                  <%end if %>


                                   <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                               
                                <th style="white-space:nowrap; vertical-align:bottom;">Order<br /><span style="font-size:9px;">Odr. Qty.</span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Ship.<br /><span style="font-size:9px;">Shp. Qty.</span></th>
                                <%end if %>


                                


                                
                                <%if cint(rapporttype) = 3 then %>
                                 <th style="white-space:nowrap; vertical-align:bottom;">Cost<br /><span style="font-size:9px;">Cost Price PC</span></th>
                                <%end if %>

                                <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                                <th style="white-space:nowrap; vertical-align:bottom;">Sales PC<br /> <span style="font-size:9px;">Sales Price PC</span></th>
                                <%end if %>

                                <%if cint(rapporttype) = 3 then %>
                                 <th style="white-space:nowrap; vertical-align:bottom;">Comm.<br /><span style="font-size:9px;"> PC %</span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Profit<br /><span style="font-size:9px;"> PC DKK</span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Total<br /><span style="font-size:9px;">Tot. Costprice DKK</span></th>
                                
                                <%end if %>

                                <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                                 <th style="white-space:nowrap; vertical-align:bottom;">Total<br /><span style="font-size:9px;">Salespri. DKK</span></th>
                                <%end if %>

                                <%if cint(rapporttype) = 3 then %>
                                <th style="white-space:nowrap; vertical-align:bottom;">Total<br /><span style="font-size:9px;">Profit DKK</span></th>
                                <th style="white-space:nowrap; vertical-align:bottom;">Total<br /><span style="font-size:9px;">Profit %</span></th>
                                <%end if %>
                            
                                   <th style="white-space:nowrap; vertical-align:bottom;">Invoice<br /><span style="font-size:9px;">Number</span></th>
                           

                                <%if media <> "print" then %>
                                
                                 <!--<th style="width:<%=tdWdthB%>px;">Del.</th>-->
                              
                                <th style="white-space:nowrap; vertical-align:bottom;">
                                   
                                    
                                        <%if antal_orders = 0 then %>

                                        <input type="checkbox" value="1" id="bulk_jobid" />

                                        <%else %>
                                        &nbsp;
                                        <%end if %>
                                  
                                 </th>
                                   <%end if %>
                            
                           
                            </tr>
                          

                        </thead>
                      
                      
                        <tbody> 
                    

    <%
    end sub



        public strFil_Kpers_Txt
    function salesreplist(sel)



         strFil_Kpers_Txt = "<option value=0>Choose..</option>" 

                       
                        
 
    
    
                        '**** Sales rep ***'
                      
                        strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
            
                        oRec.open strSQL, oConn, 3
                        while not oRec.EOF

                            if len(trim(oRec("email"))) <> 0 then
                            emlTxt = " ("& oRec("email") &")"
                            else
                            emlTxt = ""
                            end if    
    
                            if cdbl(sel) = cdbl(oRec("mid")) then       
                            kpersSel = "SELECTED"
                            else
                            kpersSel = ""
                            end if        

                          strFil_Kpers_Txt = strFil_Kpers_Txt & "<option value="& oRec("mid") &" "& kpersSel &">"& oRec("mnavn") &" "& emlTxt &"</option>" 
				
			            oRec.movenext
			            wend
			            oRec.close

    end function


    
    '*** Bearbejder data til input felter


                         '**** kunde ***'
                         strFil_Kon_Txt = "<option value=0>Choose.."& kid &"</option><option value=0></option>" 
                         lastLand = ""      
                         k = 0
                                
                        strSQLkunde = "SELECT k.kid, k.kkundenavn, kkundenr, k.land, betbet, levbet FROM kunder AS k WHERE kid <> 0 AND (useasfak = 0 OR useasfak = 1) AND ketype <> 'XX' ORDER BY k.land, k.kkundenavn"
                        oRec.open strSQLkunde, oConn, 3
                        while not oRec.EOF 
                        
                                if lastLand <> oRec("land") then

                                       if k <> 0 then
                                       strFil_Kon_Txt = strFil_Kon_Txt & "<option value='"& oRec("kid") &"'></option>"
                                       end if

                                       strFil_Kon_Txt = strFil_Kon_Txt & "<option value='"& oRec("kid") &"' disabled>"& oRec("land") &"</option>"

                                end if 


                                if len(trim(oRec("kkundenavn"))) <> 0 then
                                
                                if cdbl(jobknr) = cdbl(oRec("kid")) then
                                kSel = "SELECTED"
                                else
                                kSel = ""
                                end if 

                             

                                strFil_Kon_Txt = strFil_Kon_Txt & "<option value='"& oRec("kid") &"' "& kSel &">"& oRec("kkundenavn") & " ("& oRec("kkundenr") &") </option> "
                              
                               
                                end if

                        k = k + 1
                        lastLand = oRec("land")
                        oRec.movenext
                        wend 
                        oRec.close 

            

                          
            
                    


                         '**** Leverand / Supplier ***'
                         strFil_Sup_Txt = "<option value=0>Choose..</option>" 
                        
                          if func = "red" then
            
                            strSQL = "SELECT kid, kkundenavn FROM kunder WHERE land = '"& origin &"' AND useasfak = 6 ORDER BY kkundenavn"
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF
        
                                if cdbl(supplier) = oRec("kid") then
                                ksel = "SELECTED"
                                else
                                ksel = ""
                                end if

                               strFil_Sup_Txt = strFil_Sup_Txt & "<option value='"& oRec("kid") &"' "& kSel &">"& oRec("kkundenavn") & "</option> "
            
				
			                oRec.movenext
			                wend
			                oRec.close





                        end if' red



                        '*** materiale grupper / products groups 
                        strFil_PG_Txt = "<option value=0>Choose group..</option>" 
                        strSQL = "SELECT navn, id FROM materiale_grp WHERE id <> 0 ORDER BY navn" 
                        
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF
        
                                if cdbl(product_group) = oRec("id") then
                                pgsel = "SELECTED"
                                else
                                pgsel = ""
                                end if

                               strFil_PG_Txt = strFil_PG_Txt & "<option value='"& oRec("id") &"' "& pgSel &">"& oRec("navn") & "</option> "
            
				
			                oRec.movenext
			                wend
			                oRec.close

                        




'** Jobstatus ***'

    jobstatus0SEL = ""
    jobstatus1SEL = ""
    jobstatus2SEL = ""
    jobstatus3SEL = ""
    jobstatus4SEL = ""


select case jobstatus
case 1
    jobstatus1SEL = "SELECTED"
case 2
    jobstatus2SEL = "SELECTED"
case 3
    jobstatus3SEL = "SELECTED"
case 4
    jobstatus4SEL = "SELECTED"
case 0
    jobstatus0SEL = "SELECTED"
end select



    departmentMenSEL = ""
    departmentWomenSEL = ""
    departmentKidsSEL = ""
    departmentUnisexSEL = "" 


select case department
case "Men"
    departmentMenSEL = "SELECTED"
    case "Women"
    departmentWomenSEL = "SELECTED"
    case "Kids"
    departmentKidsSEL = "SELECTED"
    case "Unisex"
    departmentUnisexSEL = "SELECTED"
end select

    originTurSel = ""
    originChiSel = ""
    originBanSel = ""
    originIndSel = ""
    originVieSel = ""
    originHonSel = ""
    

select case origin
case "Turkey"
    originTurSel = "SELECTED"
case "China"
    originChiSel = "SELECTED"
case "Bangladesh"
    originBanSel = "SELECTED"
case "India"
    originIndSel = "SELECTED"
case "Vietnam"
    originVieSel = "SELECTED"
case "Hongkong"
    originHonSel = "SELECTED"
case "Hong kong"
    originHonSel = "SELECTED"
end select



    transportFlSel = "" 
    transportShSel = ""
    transportTrSel = ""
    transportRdSel = ""
    transportCurSel = ""


select case transport
case "By Air"
    transportFlSel = "SELECTED" 
case "By Sea"
    transportShSel = "SELECTED"
case "By Train"
    transportTrSel = "SELECTED"
case "By Truck"
    transportRdSel = "SELECTED"
case "By Currier"
    transportCurSel = "SELECTED"
end select


        select case kunde_levbetint
        case "1"
        kunde_levbetint1SEL = "SELECTED"
        case "2"
        kunde_levbetint2SEL = "SELECTED"
        case else
        kunde_levbetint0SEL = "SELECTED"
        end select


        select case lev_levbetint
        case "1"
        levlevbetint1SEL = "SELECTED"
        case "2"
        levlevbetint2SEL = "SELECTED"
        case else
        levlevbetint0SEL = "SELECTED"
        end select
 
  

fastpris2SEL = ""
fastpris3SEL = ""
select case fastpris
case 2
fastpris2SEL = "SELECTED"
case 3
fastpris3SEL = "SELECTED"
end select




    '**** enq dates

     
    if instr(dt_enq_st, "2010") <> 0 then
    dt_enq_st = "dd-mm-yyyy"
    end if

     if instr(dt_enq_end, "2010") <> 0 then
    dt_enq_end = "dd-mm-yyyy"
    end if
    

    
    if instr(dt_sour_dead, "2010") <> 0 then
    dt_sour_dead = "dd-mm-yyyy"
    end if

     if instr(dt_proto_dead, "2010") <> 0 then
    dt_proto_dead = "dd-mm-yyyy"
    end if
     

     if instr(dt_proto_sent, "2010") <> 0 then
    dt_proto_sent = "dd-mm-yyyy"
    end if
    
    
     if instr(dt_sms_dead, "2010") <> 0 then
    dt_sms_dead = "dd-mm-yyyy"
    end if
    
    
     if instr(dt_sms_sent, "2010") <> 0 then
    dt_sms_sent = "dd-mm-yyyy"
    end if
    
    
     if instr(dt_photo_dead, "2010") <> 0 then
    dt_photo_dead = "dd-mm-yyyy"
    end if
    
    
     if instr(dt_photo_sent, "2010") <> 0 then
    dt_photo_sent = "dd-mm-yyyy"
    end if
    
    
    if instr(dt_exp_order, "2010") <> 0 then
    dt_exp_order = "dd-mm-yyyy"
    end if
    

    '*** prod dates

    if instr(dt_confb_etd, "2010") <> 0 then
    dt_confb_etd = "dd-mm-yyyy"
    end if
    
    
    if instr(dt_confb_eta, "2010") <> 0 then
    dt_confb_eta = "dd-mm-yyyy"
    end if

    
    if instr(dt_confs_etd, "2010") <> 0 then
    dt_confs_etd = "dd-mm-yyyy"
    end if
    
  
    if instr(dt_confs_eta, "2010") <> 0 then
    dt_confs_eta = "dd-mm-yyyy"
    end if
    
    
    if instr(dt_actual_etd, "2010") OR instr(dt_actual_etd, "2013") <> 0 then 'for at nulstille
    dt_actual_etd = "dd-mm-yyyy"
    end if
    
    
    if instr(dt_actual_eta, "2010") <> 0 then
    dt_actual_eta = "dd-mm-yyyy"
    end if

      if instr(dt_sup_photo_dead, "2010") <> 0 then
    dt_sup_photo_dead = "dd-mm-yyyy"
    end if

      if instr(dt_sup_sms_dead, "2010") <> 0 then
    dt_sup_sms_dead = "dd-mm-yyyy"
    end if




    '**Orders Dates

      
     if instr(dt_firstorderc, "2010") <> 0 then
    dt_firstorderc = "dd-mm-yyyy"
    end if

    
     if instr(dt_ldapp, "2010") <> 0 then
    dt_ldapp = "dd-mm-yyyy"
    end if
     
    
     if instr(dt_sizeexp, "2010") <> 0 then
    dt_sizeexp = "dd-mm-yyyy"
    end if
     
    
     if instr(dt_sizeapp, "2010") <> 0 then
    dt_sizeapp = "dd-mm-yyyy"
    end if

     
     if instr(dt_ppexp, "2010") <> 0 then
    dt_ppexp = "dd-mm-yyyy"
    end if

    
    if instr(dt_ppapp, "2010") <> 0 then
    dt_ppapp = "dd-mm-yyyy"
    end if

    
     if instr(dt_shsexp, "2010") <> 0 then
    dt_shsexp = "dd-mm-yyyy"
    end if

   
     if instr(dt_shsapp, "2010") <> 0 then
    dt_shsapp = "dd-mm-yyyy"
    end if

'*********************


    function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function




public dt_sms_dead, dt_sup_sms_dea, dt_sms_sen, dt_photo_dea, dt_photo_sent, dt_sup_photo_dead
function enqDates()


    '**** Enq dates
    dt_sms_dead = request("FM_dt_sms_dead")
    call fmatDate_fn(dt_sms_dead)
    dt_sms_dead = fmatDate

    dt_sup_sms_dea = request("FM_dt_sup_sms_dea")
    call fmatDate_fn(dt_sup_sms_dea)
    dt_sup_sms_dea = fmatDate

    dt_sms_sent = request("FM_dt_sms_sent")
    call fmatDate_fn(dt_sms_sent)
    dt_sms_sent = fmatDate

     dt_photo_dead = request("FM_dt_photo_dead")
    call fmatDate_fn(dt_photo_dead)
    dt_photo_dead = fmatDate

     dt_sup_photo_dead = request("FM_dt_sup_photo_dead")
    call fmatDate_fn(dt_sup_photo_dead)
    dt_sup_photo_dead = fmatDate

    dt_photo_sent = request("FM_dt_photo_sent")
    call fmatDate_fn(dt_photo_sent)
    dt_photo_sent = fmatDate


end function



public dt_exp_order, dt_confb_eta, dt_confs_etd, dt_confs_eta, dt_actual_etd, dt_actual_eta, dt_firstorderc, dt_ldapp, dt_sizeexp
public dt_sizeapp, dt_ppexp, dt_ppapp, dt_shsexp, dt_shsapp, dt_sup_sms_dead , dt_confb_etd
function orderAndProdDates()


    '**** Order dates
    dt_exp_order = request("FM_dt_exp_order")
    call fmatDate_fn(dt_exp_order)
    dt_exp_order = fmatDate

  
    dt_confb_etd = request("FM_dt_confb_etd")
    call fmatDate_fn(dt_confb_etd)
    dt_confb_etd = fmatDate


    dt_confb_eta = request("FM_dt_confb_eta")
    call fmatDate_fn(dt_confb_eta)
    dt_confb_eta = fmatDate

 
    
    dt_confs_etd = request("FM_dt_confs_etd")
    call fmatDate_fn(dt_confs_etd)
    dt_confs_etd = fmatDate

     
    
    dt_confs_eta = request("FM_dt_confs_eta")
    call fmatDate_fn(dt_confs_eta)
    dt_confs_eta = fmatDate
  
    
    dt_actual_eta = request("FM_dt_actual_eta")
    call fmatDate_fn(dt_actual_eta)
    dt_actual_eta = fmatDate

  

    'dt_actual_etd = request("FM_dt_actual_etd")
    'call fmatDate_fn(dt_actual_etd)
    'dt_actual_etd = fmatDate


    if cint(kopier_ordre) = 1 then
    dt_actual_etd = "2010-01-01"
    else
    dt_actual_etd = request("FM_dt_actual_etd")
    call fmatDate_fn(dt_actual_etd)
    dt_actual_etd = fmatDate
    end if


    '*** production dates **

    dt_firstorderc = request("FM_dt_firstorderc")
    call fmatDate_fn(dt_firstorderc)
    dt_firstorderc = fmatDate

    dt_ldapp = request("FM_dt_ldapp")
    call fmatDate_fn(dt_ldapp)
    dt_ldapp = fmatDate


    dt_sizeexp = request("FM_dt_sizeexp")
    call fmatDate_fn(dt_sizeexp)
    dt_sizeexp = fmatDate

    dt_sizeapp = request("FM_dt_sizeapp")
    call fmatDate_fn(dt_sizeapp)
    dt_sizeapp = fmatDate
    
    dt_ppexp = request("FM_dt_ppexp")
    call fmatDate_fn(dt_ppexp)
    dt_ppexp = fmatDate
    
    dt_ppapp = request("FM_dt_ppapp")
    call fmatDate_fn(dt_ppapp)
    dt_ppapp = fmatDate

    dt_shsexp = request("FM_dt_shsexp")
    call fmatDate_fn(dt_shsexp)
    dt_shsexp = fmatDate

    dt_shsapp = request("FM_dt_shsapp")
    call fmatDate_fn(dt_shsapp)
    dt_shsapp = fmatDate
    
    

end function
%>
