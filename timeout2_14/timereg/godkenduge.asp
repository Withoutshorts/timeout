<%response.Buffer = true 
tloadA = now
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->




<%
    if len(session("user")) = 0 then

    %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else

    call smileyAfslutSettings()
    
    thisfile = "godkenduge.asp"
    rdir = "godkenduge.asp"

    func = request("func")

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = 0
    end if

    if len(trim(request("varTjDatoUS_man"))) <> 0 then

        varTjDatoUS_man = request("varTjDatoUS_man")
        
    else

        varTjDatoUS_man = day(now)&"-"&month(now)&"-"&year(now)

        dayofw = datePart("w", varTjDatoUS_man, 2,2)
        varTjDatoUS_man = dateAdd("d", dayofw - (dayofw-1), varTjDatoUS_man)
       

    end if


        varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
       
        prev_varTjDatoUS_man = dateAdd("d", -7, varTjDatoUS_man)
        next_varTjDatoUS_man = dateAdd("d", 7, varTjDatoUS_man)

        prevWeek = datepart("ww", dateAdd("d", -7, varTjDatoUS_man), 2,2) 
        nextWeek = datepart("ww", dateAdd("d", 7, varTjDatoUS_man), 2,2) 


      

        
        
        'response.write "<hr>SmiWeekOrMonth: "& SmiWeekOrMonth 
        
        if cint(SmiWeekOrMonth) = 0 then
        
        varTjDatoUS_tor = dateAdd("d", 3, varTjDatoUS_man) 'Altid torsdag, så vi er sikker på vi er i ugen

           if datepart("ww", varTjDatoUS_tor,2,2) = 53 then
           varTjDatoUS_tor = dateAdd("d", 3, varTjDatoUS_tor) 'ind i næste år
           end if

        sm_sidsteugedag = datePart("ww",varTjDatoUS_tor, 2,2)
        sm_aar = year(varTjDatoUS_tor)


        else
        varTjDatoUS_mdTest = dateAdd("d", 7, varTjDatoUS_man) 'Altid torsdag, så vi er sikker på vi er i MD

        if datePart("m", varTjDatoUS_mdTest, 2, 2) <> datePart("m", varTjDatoUS_man, 2, 2) then
        varTjDatoUS_tor = dateAdd("d", -7, varTjDatoUS_man) ' Sikrer vi er i MD
        else
        varTjDatoUS_tor = varTjDatoUS_man
        end if

        sm_sidsteugedag = datePart("m",varTjDatoUS_man, 2,2)
        sm_aar = year(varTjDatoUS_man)
        end if
        'sm_sidsteugedag = datePart("ww",varTjDatoUS_man, 2,2)
        
        sm_mid = usemrn


        

    call erugeAfslutte(sm_aar, sm_sidsteugedag, sm_mid)


    fmlink = "ugeseddel_2011.asp?usemrn="& usemrn &"&varTjDatoUS_man="& varTjDatoUS_man &"&varTjDatoUS_son= "& varTjDatoUS_son &"&nomenu=1&rdir=godkenduge"

    %>
    <div id="sindhold" style="position:absolute; left:40px; top:40px; z-index:0; width:400px; background-color:#FFFFFF; padding:20px;">
       

        <span style="float:right;">
        <table cellpadding=0 cellspacing=0 border=0 width=120>
	<tr>
	<td valign=top align=right style="padding:0px 10px 0px 0px;"><a href="godkenduge.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=prev_varTjDatoUS_man %>&varTjDatoUS_son=<%=prev_varTjDatoUS_son %>&nomenu=<%=nomenu %>">< uge <%=prevWeek %></a></td>
   <td style="padding-right:10px;" valign=top align=right><a href="godkenduge.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=next_varTjDatoUS_man %>&varTjDatoUS_son=<%=next_varTjDatoUS_son %>&nomenu=<%=nomenu %>">uge <%=nextWeek %> ></a></td>
	</tr>
	</table>
            </span>
        <br /><br />

        <%if cint(SmiWeekOrMonth) = 0 then 'uge %>
        <h4>Godkend ugeseddel uge <%=datepart("ww", varTjDatoUS_tor, 2,2)%>
            
            <%if datepart("ww", varTjDatoUS_tor, 2,2) = 53 then
            response.write " - "& datepart("yyyy", dateAdd("yyyy", - 1, varTjDatoUS_tor), 2,2) &" / "& datepart("yyyy", varTjDatoUS_tor, 2,2)
            else
            response.write " - "& datepart("yyyy", varTjDatoUS_tor, 2,2)
            end if
            %>
	    <%else
            
            select case lto
            case "tec", "esn"
            LukafslTxt = "Luk" 
            case else
            LukafslTxt = "Afslut"
            end select%>

        <h4><%=LukafslTxt &" "& monthname(datepart("m", varTjDatoUS_tor, 2,2)) & " - "& datepart("yyyy", varTjDatoUS_tor, 2,2) %>
        <%end if

        call meStamdata(usemrn)

        %>
        <br /><span style="font-size:11px; font-weight:lighter; color:#999999;"><i>For <%=meTxt %></i></span>
            </h4>
        
     
        <%
    
        if level <=2 OR level = 6 then %>
           
        <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>

        <%else %>
    
        Du har ikke adgang til at godkende ugeseddel for denne medarbejder.

        <%end if 'level

	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
