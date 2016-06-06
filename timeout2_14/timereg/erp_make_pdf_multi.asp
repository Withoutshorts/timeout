<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
 
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->

<%

'faknr = request("faknr")
lto = request("lto")

'jobid = request("jobid")
'aftid = request("aftid")
'id = request("id")
'kid = request("kid")

'Response.Write jobid
'Response.End 
if len(request("makepdf")) <> 0 then
makepdf = request("makepdf")
else
makepdf = 0
end if

if len(trim(request("doctype"))) <> 0 then
doctype = request("doctype")
else
doctype = ""
end if

if len(request("fakids")) <> 0 then
fakids = split(request("fakids"), ",")
else
fakids = 0
end if

'** emailmodtager / afsender
strSQL = "SELECT mnavn, email FROM medarbejdere WHERE mid = "& session("mid")
oRec.open strSQL, oConn, 3

if not oRec.EOF then

emAf = oRec("mnavn")
emAFmail = oRec("email")

end if
oRec.close

select case lto
case "execon", "immenso"
emMo = "Bogholderi"
emMomail = "bogholderi@execon.dk"
case else
emMo = emAf
emMomail = emAFmail
end select


    
    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\erp_make_pdf_multi.asp" then
                               
    'Opretter en instans af fil object **'
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    
                                 ' Sætter Charsettet til ISO-8859-1 
                                 Mailer.CharSet = 2
                                 ' Afsenderens navn 
                                 Mailer.FromName = "TimeOut Email Service"
                                 ' Afsenderens e-mail 
                                 Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                 Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
                                 Mailer.ContentType = "text/html"
                                 
                                'Mailens emne
                                 Mailer.Subject = "Faktura PDF/XML mail"
                                 'Modtagerens navn og e-mail
                                 Mailer.AddRecipient emMo, emMomail

    

    
 

    f = 9
    for f = 1 to UBOUND(fakids)
        
        if len(trim(fakids(f))) <> 0 then
        fakids(f) = fakids(f)
        else
        fakids(f) = 0
        end if
        
        strSQL = "SELECT faknr FROM fakturaer WHERE fid = " & fakids(f)
        oRec.open strSQL, oConn, 3
        faknr = 0
        if not oRec.EOF then
        
        faknr = oRec("faknr")
        
        End if
        oRec.close


         
         strFaktypeNavn = "faktura"
         if cdbl(fakids(f)) <> 0 then 
         
             strknavnSQL = "SELECT kkundenavn, fid, faktype FROM fakturaer LEFT JOIN kunder ON (kid = fakadr) WHERE fid = " & fakids(f)
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



        
            '*** Tilføjer PDF **'
             pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
	         
	 
        	If (fs.FileExists(pdfurl))=true Then
                    
                    Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                    bodyTXT = bodyTXT & vbcrlf &""&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                    
                    
            else
                        
                        '*** Opretter PDF og afsender **'
                        '*** Ved godkend faktura, Execon, Immenso.. ***'
                        if cint(makepdf) = 1 then
                        
                        
                        Set Pdf = Server.CreateObject("Persits.Pdf")
                        Set Doc = Pdf.CreateDocument


                         select case lto
                         case "dencker"
                         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?&visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"
                         Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"

                         case "synergi1", "intranet - local"
                         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=20, RightMargin=0, TopMargin=0, BottomMargin=0, DrawBackground=True,"
         
                         
                         Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?&visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
                         'Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"

       
                         case else

                             if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                              Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?&visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                             else 
                             '*** NY SERVER ****
                              Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                             end if

                         end select


                        ''Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&FM_usedatointerval="&request("FM_usedatointerval")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"&FM_fakint_ival="&request("FM_fakint_ival")&"&FM_aftnr="&sogaftnr&"&kid="&kid
                        ''Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                        ''Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id
                        
                        
                       'Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/faktura_"&lto&"_"&faknr&"_"&strknavn&".pdf"), True ) 
                       Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"), True)  
                       
                        
                        Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                        bodyTXT = bodyTXT & vbcrlf &""&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
               
                        
                        else
                        
                        Response.Write("")
                       
                        end if
                    
                    
                   
                    
            End If
            
            
             '*** Tilføjer XML **'
             xmlurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\"&lto&"\"&strFaktypeNavn&"_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
	    
	 
        	If (fs.FileExists(xmlurl))=true Then
                   Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\"&lto&"\"&strFaktypeNavn&"_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
                   bodyTXT = bodyTXT & vbcrlf &"faktura_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
                    
            End If
       
       
      
    next
    
    
                 
          
                 
                 ' Selve teksten
                  bodyTXT =  bodyTXT & vbcrlf & vbcrlf &"Denne mail er oprettet af:"& vbcrlf & emAf & ", "& emAFmail
                 
                  Mailer.BodyText = "Din mail med PDF dokumenter og XML filer er klar. Følgende filer er vedhæftet:" &vbcrlf & bodyTXT 
                  
       
        
                  If Mailer.SendMail Then
                        
                        
                        if makepdf = 1 then
                          Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                            Response.Write("<script language=""JavaScript"">window.opener.top.frames['erp2_2'].location.reload();</script>")
                            
                            Response.Write("<script language=""JavaScript"">window.close();</script>")
                           
                           else
                     
                       %>
                       
                       
                       
                  <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	           Mail med vedhæftede PDF'er og XML filer er afsendt.</a>
	            </td></tr>
	            </table>
                  <%    end if
                            
                  Else
                  Response.Write "Fejl...<br>" & Mailer.Response
                  End if
    
    
    set fs=nothing
    
    end if '** C:
    
%>



<!--#include file="../inc/regular/footer_inc.asp"-->
