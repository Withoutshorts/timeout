

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%


lto = request("lto")

if len(request("id")) <> 0 then
id = request("id")
vis_indledning = request("FM_vis_indledning_pdf")
vis_footer = request("FM_vis_afslutning_pdf")

'vis_indledning = Server.HTMLEncode(vis_indledning)
vis_indledning = replace(vis_indledning, "'", "&#39;")

else
id = 0
end if



kid = request("kid")

if len(trim(request("gem_pdf_somfil"))) <> 0 then
gem_pdf_somfil = request("gem_pdf_somfil") 
else
gem_pdf_somfil = 0
end if

if len(trim(request("pdf_foid"))) <> 0 then
prFolderid = request("pdf_foid")
else
prFolderid = 1000
end if




if len(trim(request("pdf_skabelonnavn"))) <> 0 then
filnavnDato = request("pdf_skabelonnavn")
call kundenavnPDF(filnavnDato)
filnavnDato = strKundenavnPDFtxt
filnavnKlok = ""


filnavn = filnavnDato &".pdf"

else

filnavnDato = "pdf_"& year(now)&"_"&month(now)& "_"&day(now)
filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)

filnavn = filnavnDato &""& filnavnKlok &".pdf"

end if	

call TimeOutVersion()





if cint(gem_pdf_somfil) = 1 then

   
	prNavn = filnavn
	prTxt = vis_indledning '"" 'request("FM_vis_indledning")
	prEditor = session("user")
    pdf_doktype = request("pdf_doktype")

 
    '** Opdater eksisterende / eller gem ny ****'
    filfindes = 0
    strSQLfindesfil = "SELECT filnavn, id FROM filer WHERE filnavn = '"& prNavn &"' AND jobid = "& id 
    oRec3.open strSQLfindesfil, oConn, 3
    if not oRec3.EOF then
    filfindes = oRec3("id")
    end if
    oRec3.close


    filnavnDato = year(now)&"/"&month(now)& "/"&day(now)

    if filfindes <> 0 then

     call updFil(prNavn,prTxt,prEditor,filnavnDato,pdf_doktype,id,kid,prFolderid,filfindes)

    'strSQLinsfil = "UPDATE filer SET filnavn = '"& prNavn &"', "_
	'&" filertxt = '"&prTxt&"', editor = '"& prEditor &"', dato = '"& year(now)&"/"&month(now)& "/"&day(now) &"', type = 5, "_
	'&" jobid = "& id &", kundeid = "& kid &", folderid = "& prFolderid &""_
	'&" WHERE id = " & filfindes

    else

     call insFil(prNavn,prTxt,prEditor,filnavnDato,pdf_doktype,id,kid,prFolderid)
        
	'strSQLinsfil = "INSERT INTO filer (filnavn, filertxt, editor, dato, type, jobid, kundeid, folderid, "_
	'&" adg_kunde, adg_alle, incidentid) VALUES ('"& prNavn &"', '"&prTxt&"', '"& prEditor &"', "_
	'&" '"& year(now)&"/"&month(now)& "/"&day(now) &"', 5, "& id &", "& kid &", "& prFolderid &", 0, 1, 0)"
	end if
	
	'Response.Write strSQLinsfil & "<br><br>"
	'Response.end
	'oConn.execute(strSQLinsfil)
	

end if


if len(trim(vis_indledning)) <> 0 then

pdfTxt = vis_indledning

strSQLpdfval = "INSERT INTO pdf_values (pdf_txt, pdf_footer) VALUES ('"&vis_indledning&"', '"&vis_footer&"')"


'if lto = "synergi1" then
'Response.Write strSQLpdfval
'Response.flush
'end if

oConn.execute(strSQLpdfval)

strSQLsel = "SELECT id FROM pdf_values WHERE id <> 0 ORDER BY id DESC LIMIT 0, 1"
oRec.open strSQLsel, oConn, 3
if not oRec.EOF then

pdfvalid = oRec("id")

end if
oRec.close

'Response.Cookies("job")("pdfval") = vis_indledning
end if


'Response.end


if id <> 0 then



        Set Pdf = Server.CreateObject("Persits.Pdf")
        Set Doc = Pdf.CreateDocument

           
        
        if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_make_pdf.asp" then
        'Doc.ImportFromUrl "http://localhost/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=-20, RightMargin=0, TopMargin=0, BottomMargin=0, DrawBackground=True," 'PageWidth=590, PageHeight=792, 
        Doc.ImportFromUrl "http://localhost/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=20, RightMargin=20, TopMargin=20, BottomMargin=20, PageWidth=595, PageHeight=840, DrawBackground=True"  
        
        'str = "<HTML><TABLE><TR><TD><img src=""https://outzource.dk/timeout_xp/wwwroot/ver2_10/ill/synergi/synergi_logo_og_adresse_out_3.jpg""></TD><TD>Text2</TD></TR></TABLE></HTML>"
        'Doc.ImportFromUrl str 
        
        else
            select case lto 
            case "essens"
            Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=0, RightMargin=0, TopMargin=50, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
            'LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, Interpolate=True"
            case "synergi1"
            'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=20, RightMargin=0, TopMargin=0, BottomMargin=0, DrawBackground=True,"
            Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
            case "jttek"
            'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=20, RightMargin=0, TopMargin=0, BottomMargin=0, DrawBackground=True,"
            'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=40, RightMargin=0, TopMargin=-40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
            'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=-80, RightMargin=20, TopMargin=30, BottomMargin=20, PageWidth=595, PageHeight=840, DrawBackground=True" 'PageWidth=635, PageHeight=903,      
            Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=-80, RightMargin=20, TopMargin=30, BottomMargin=20, PageWidth=595, PageHeight=840, DrawBackground=True" 'PageWidth=635, PageHeight=903,      


            case "dencker"
            'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"
            Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"
         
            case else

                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=40, RightMargin=10, TopMargin=20, BottomMargin=5"
                else
                Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=40, RightMargin=10, TopMargin=20, BottomMargin=5"
                end if

            end select
        end if 


         'Response.Write filnavn
       'Response.end
         
        Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/"&filnavn&""), True) 
         
        
      


      

Response.redirect "../inc/upload/"&lto&"/"&filnavn


end if

%>


<!--
<br /><br /><br /><br />
<table cellspacing=0 cellpadding=0 border=0>
<tr><td style="padding:20px 20px 20px 20px;">
<h4>TimeOut PDF</h4>

Din PDF fil er produceret og gemt på serveren. <br />
Du kan hente din DPF faktura <a href="../inc/upload/"<%=lto%>"/jobtilbudprint_"<%=lto%>"_"<%= %>jobtilbudsnr&"_"&filnavnDato&"_"&filnavnKlok&".pdf">her</a>
</td></tr></table>
-->

<!--#include file="../inc/regular/footer_inc.asp"-->
