<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/header_hvd_inc.asp"-->

<%

faknr = request("faknr")
lto = request("lto")

jobid = request("jobid")
aftid = request("aftid")
id = request("id")
kid = request("kid")

    strFaktypeNavn = "faktura"  
    


    call TimeOutVersion()

         if cdbl(id) <> 0 then 
         
             strknavnSQL = "SELECT kkundenavn, fid, faktype FROM fakturaer LEFT JOIN kunder ON (kid = fakadr) WHERE fid = " & id
             oRec.open strknavnSQL, oConn, 3
             
              strknavn = "xx"
             if not oRec.EOF then
             strknavn = oRec("kkundenavn") 
             
             call kundenavnPDF(strknavn)
             strknavn = strKundenavnPDFtxt

             if oRec("faktype") = 1 then
             strFaktypeNavn = "kreditnota"
             else
             strFaktypeNavn = "faktura"
             end if

             end if
             oRec.close

             

         else
         strknavn = "oo"
         end if

    sprog = 2

    %><!--#include file="../inc/regular/global_func.asp"--><%

        '*** IVAL Datoer SKAL med for at få joblog med ud ****'

        Set Pdf = Server.CreateObject("Persits.Pdf")
        Set Doc = Pdf.CreateDocument
        'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&FM_usedatointerval="&request("FM_usedatointerval")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"&FM_fakint_ival="&request("FM_fakint_ival")&"&FM_aftnr="&sogaftnr&"&kid="&kid
        
         select case lto
         case "dencker"
         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"
          Doc.ImportFromUrl "https://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"

         case "synergi1", "xintranet - local"
         Doc.ImportFromUrl "https://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
         'Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
         case "essens"
         Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=55, RightMargin=0, TopMargin=80, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
         case else

         if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
         Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
         else
         '*** NY SERVER ****
         Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
         end if
          
         end select
        'Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
        
        
        Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"), True) 

        'Response.write filename
        'Response.end

Response.redirect "../inc/upload/"&lto&"/"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
%>


<!--
<br /><br /><br /><br />
<table cellspacing=0 cellpadding=0 border=0>
<tr><td style="padding:20px 20px 20px 20px;">
<h4>TimeOut PDF faktura</h4>

Din PDF fil er produceret og gemt på serveren. <br />
Du kan hente din DPF faktura <a href="../inc/upload/<=lto%>/faktura_<=lto%>_<=faknr%>.pdf">her</a>
</td></tr></table>
-->
<!--#include file="../inc/regular/footer_inc.asp"-->
