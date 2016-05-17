<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%
if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
	response.end
	end if

	func = request("func")
	id = request("id") 
	
	%>

<script language="javascript">
    function alterall(f) {
        for (i = 0; f.elements[i]; i++) {
            e = f.elements[i];
            if (e.type == "checkbox") {
                e.checked = (e.checked) ? false : true;
            }
        }
    }
</script>



<%call menu_2014 %>
<div class="wrapper">
 <div class="content">
  <div class="container">
   <div class="portlet">
    
<%
	select case func
	case "med"
	strSQL = "SELECT id, navn FROM brugergrupper WHERE id=" & id
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	gruppeNavn = oRec("Navn")
	end if
	oRec.close
	%>
       
       <!--Medarbliste-->    
       <h3 class="portlet-titlte"><u>Brugergrupper (medlemmer)</u></h3>

       <form name="brugergruppe" method="post" action="bgrupper.asp?func=update">

       <div class="portlet-body">
           <div>Medlemmer i <b><%=gruppeNavn%></b> gruppen</div> &nbsp
           <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
               <thead >
                   <tr>
                       <th style="width: 89%">Navn</th>
                       <th>Flyt medarbejder</th>
                   </tr>
               </thead>
               <tbody>
                   <%
	                strSQL = "SELECT Mnavn, Mid FROM medarbejdere WHERE brugergruppe = "&id&" ORDER BY Mnavn"
	                oRec.open strSQL, oConn, 3
	                'tæl brugere
	                strCountBrg = "SELECT count(*) AS AntBruger FROM medarbejdere WHERE brugergruppe = "&id&""
	                set BrgCount = oConn.execute(strCountBrg)
	                StrCountBruger = BrgCount("AntBruger")
	
	                strCHKIDS = "0"
	                while not oRec.EOF
	                strID = oRec("Mid")
	                strCHKIDS = strCHKIDS & "." & strID
	                %>
                   <tr>
                       <td><a href="medarb_red.asp?menu=medarb&func=red&id=<%=strID%>"><%=oRec("Mnavn")%> </a></td>
                       <td align=center><input type="checkbox" name="CHK<%=strID %>"" /></td>
                   </tr>
                   <%
	                oRec.movenext
	                wend
	                oRec.close
	                %>
               </tbody>
           </table>
                    <div>brugere i alt: <%=StrCountBruger%></div>
                    &nbsp

                    <div class="col-lg-6"><br />Flyt valgte medarbejdere til:
        <select name="brggrp" style="width:200px;">
        <%
        	strSQL2 = "SELECT id, navn FROM brugergrupper"
			oRec.open strSQL2, oConn, 3
			while not oRec.EOF 
			strID = oRec("id")
			strNavn = oRec("navn")
			if cint(id) = cint(strID) then
			strselected = "selected"
			end if
			%>

			<option value="<%=strID%>" <%=strselected%>><%=strNavn%></option>
			<%
			strselected = NULL
			oRec.movenext
	        wend
			oRec.close
			%>
        </select>
       <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
        <br />
        <% if StrCountBruger <> "0" then
        strDis = "disabled"
        end if %>
        <input type="hidden" name="curgrp" value="<%=id%>" />
        <input type="hidden" name="CheckID" value="<%=STRCHKIDS%>" /></div>
       </div>

<%
	case "update"
	
	'if request("grpnavn") <> "" then
	GlGruppe = request("curgrp")
	NyGruppe = request("brggrp")
	GrpNavn = request("grpnavn")
	
	'UpdateGruppeSQL = "Update brugergrupper SET Navn = '" & GrpNavn & "' WHERE id = " & GlGruppe
	'oConn.Execute(UpdateGruppeSQL)
	
'flyt medarbejdere til ny brugergruppe
if NOT GlGruppe = NyGruppe then	
    idsToCheck = request("CheckID")
	UpdateMedarbSQL = "UPDATE medarbejdere SET Brugergruppe = '" & NyGruppe & "' WHERE mid in ("	
	arr1 = split(idsToCheck,".")
    for i = 1 to ubound(arr1)
    strMedid = arr1(i)
        if request("CHK" & strMedid) = "on" then
            IDLIST = IDLIST & strMedid & ","
        end if
    next
   if len(IDLIST) > 0 then
    IDLIST = left(IDLIST,cint(len(IDLIST)-1))
    UpdateMedarbSQL = UpdateMedarbSQL & IDLIST & ")" 
    oConn.Execute(UpdateMedarbSQL)
   end if
end if
	'end if
	response.Redirect "bgrupper.asp"
	case "slet"%>
    <!--Slet side-->	
           <div class="container" style="width:500px;">
            <div class="porlet">       
            <h3 class="portlet-title">
               <u>brugergrupper</u>
            </h3>       
                <div class="portlet-body">
                    <div>Du er ved at <b>SLETTE</b> en brugergruppe. Er dette korrekt?
                    </div><br />
                    <div><b><a href="bgrupper.asp?func=sletok&id=<%=id%>">Ja</a></b>&nbsp&nbsp&nbsp&nbsp<a href="bgrupper.asp?"><b>Nej</b></a>
                    </div>
                    <br /><br />
                 </div>
            </div>
        </div>
	<%
	case "sletok"
	'tjek om gruppen er tom
	'tæl brugere
	strCountBrg = "SELECT count(*) AS AntBruger FROM medarbejdere WHERE brugergruppe = "&id&""
	set BrgCount = oConn.execute(strCountBrg)
	StrCountBruger = BrgCount("AntBruger")
	
	if strCountBruger = "0" then
	oConn.execute("DELETE FROM brugergrupper WHERE id = "& id)
	response.Redirect "bgrupper.asp"
	else
	response.Write "det er ikke muligt at slette denne gruppe da der stadig er medarbejdere tilnyttet"
	end if
	
	
	case "opret"
	strNavn = request("FM_grpnavn")
	if strNavn <> "" then
	strNavn = Replace(strNavn,"'","''")
	oConn.execute("INSERT INTO brugergrupper (rettigheder, navn) VALUES (1, '"& strNavn  &"')")
	response.Redirect "bgrupper.asp"
	else
	%>
     <!--Opret--> 
       <h3 class="portlet-title"><u>Brugergruppe (opret)</u></h3> 
           <div class="portlet-body">
               <div class="well well-white">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-2">gruppenavn:</div>
                <div class="col-lg-2"><input type="text" name="FM_grpnavn" value=""></div>
               </div>
           </div>

<% 
end if
    case else    
%>
                
       <h3 class="portlet-title"><u>Brugergrupper (rettigheder)</u></h3>
       <div class="portlet-body">
           <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
               <thead>
                   <tr>
                       <th>Navn</th>
                       <th>Rettigheds niveau</th>
                   </tr>
               </thead>
               <tbody>
                   <%
                    strSQL = "SELECT id, navn, rettigheder FROM brugergrupper ORDER BY navn"
                    x = 0
	
	                oRec.open strSQL, oConn, 3
	                while not oRec.EOF 
	
	                strSQL2 = "SELECT Mid FROM medarbejdere WHERE brugergruppe = "& oRec("id")
	                oRec2.open strSQL2, oConn, 3
	                while not oRec2.EOF 
	                x = x + 1
	                oRec2.movenext
	                wend
	                oRec2.close
	                Antal = x
                       %>
                   <tr>
                       <td><a href="bgrupper.asp?menu=medarb&func=med&id=<%=oRec("id")%>"><%=oRec("navn")%> </a>&nbsp;&nbsp;(<%=Antal%>)<br>
		                <%
		                SELECT CASE oRec("rettigheder")
		                case 1
		                varBeks = "Administrator gruppen har adgang til alle områder og alle funktioner."
		                case 2
		                varBeks = "Denne gruppe har adgang til alle områder (som ovenstående) og alle funktioner i Timeout på lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA/CRM [Indstillinger], og interne timepriser (løn)"
		                case 3
		                varBeks = "Denne gruppe har adgang til alle områder (som ovenstående) og alle funktioner i Timeout på lederniveau. <b><font color=darkred>Dog ikke adgang til</font></b> TSA [Statistik] og TSA [Fakturering], og begrænset adgang til TSA [Medarbejdere]."
		                case 4
		                varBeks = "Har adgang til TSA og CRM delen (TSA [timeregistrering] og TSA [medarbejdere], samt CRM [Kalender], CRM [Aktions historik] og CRM [Firma kontakter]).<br> Adgang <b>kun</b> på bruger niveau, det vil sige indtastning af timer, samt oversigter der vedrører medarbejderen selv."
		                case 6
		                varBeks = "Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 1)."
		                case 7
		                varBeks = "Har kun adgang til TSA delen (som gruppe TSA + CRM niveau 2)."
		                case 8
		                varBeks = "Har adgang til TSA [timeregistrering]."
		                end select 
		                Response.write varBeks
		                %></td>
		                <td align="center"><%=oRec("rettigheder")%></td>
                   </tr>
                   <%
                    x = 0
	                oRec.movenext
	                wend
	               %>
               </tbody>
           </table>
       </div>
    </form>




   </div>
  </div>
 </div>
</div>
<%end select %>
<!--#include file="../inc/regular/footer_inc.asp"-->