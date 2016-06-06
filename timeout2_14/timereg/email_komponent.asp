<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	func = request("func")
	select case func
	case "dwldb"			
	
					
					strSQL = "SELECT Mid, email, mnavn FROM medarbejdere WHERE Mid=" & session("mid")
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					strEmail = oRec("email")
					strEditor = oRec("mnavn")
					end if
					oRec.close
					
	
					Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					' Sætter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "TimeOut DB Backup"
					' Afsenderens e-mail 
					Mailer.FromAddress = "support@outzource.dk"
					Mailer.RemoteHost = "smtp.tiscali.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient "Support OutZourCE", "support@outzource.dk"
					'Mailer.AddBCC "SK", "sk@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "DB backup er bestilt af: "& lto &" !"
					
					' Selve teksten
					Mailer.BodyText = "" & "Backup er bestilt af "& strNavn & vbCrLf _ 
					& "Sendes til: " &strEmail & "" & vbCrLf 
					
					Mailer.SendMail
					
					strDato = year(now)&"/"&month(now)&"/"&day(now)
					
						oConn.execute("INSERT INTO dbdownload (dato, editor, email) VALUES ("_
						&"'"& strDato &"',"_
						&"'"& strEditor &"',"_
						&"'"& strEmail &"')")
						
					Response.redirect "kontrolpanel.asp?menu=tok"
				
	case else
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<!--#include file="inc/dato.asp"-->
	<!--#include file="../inc/regular/topmenu_inc.asp"-->
	<span style="position:absolute; left:10; top:60; visibility:visible;">
	<table cellspacing="0" cellpadding="0" border="0" width="159">
	<tr>
		<td colspan="2" class="alt"><img src="../ill/personligemp_top_blaa.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
	<tr bgcolor="EFF3FF"><td style="padding-left:3;">
	Ny mail<br>
	<a href="email.asp?func=indbakke&menu=eml" class='vmenu'>Indbakke</a><br>
	Sendt mail<br>
	Slettet mail<br>
	</td></tr>
	</table><br><br>
	
	<table cellspacing="0" cellpadding="0" border="0" width="159">
	<tr>
		<td colspan="2" class="alt"><img src="../ill/kundemp_top_blaa.gif" alt="" border="0"></td>
	</tr>
	</table>
	<table cellspacing="0" cellpadding="2" border="0" bgcolor="#EFF3FF" width="159">
	<tr bgcolor="EFF3FF"><td>
	<table>
	<tr>
	<td><select name="kundereml" size="1" style="font-size : 11px; width:146 px;">
	<option value="0">Indbakke</option>
	<%
			strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder ORDER BY Kkundenavn"
			oRec.open strSQL, oConn, 3
			while not oRec.EOF
			
			if cint(request(name1)) = cint(oRec("Kid")) then
			isSelected = "SELECTED"
			else
			isSelected = ""
			end if
			%>
			<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=left(oRec("Kkundenavn"), 20)%></option>
			<%
			oRec.movenext
			wend
			oRec.close
			%>
	</select></td>
	</tr>
	<tr bgcolor="#EFF3FF"><td align="center"><input type="image" src="../ill/visdennemappe.gif" border="0">
	</td></form></tr>
	</table>
	</td></tr></table></span>
		
		
	<div id="sindhold" style="position:absolute; left:190; top:60; visibility:visible;">
	<table border=0 cellpadding=0 cellspacing=0 width="650">
	<tr>
	<td valign="top" width="163"><img src="../ill/logo_bg.gif" width="163" height="53" alt="" border="0"></td>
	<td valign="bottom"><b>Timeout Email.</b><br>
	Send og modtag email. Gem email historik på kunder og kontaktpersoner.</td>
	</tr>
	</table><br>
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#EFF3FF" width="671">
	<tr bgcolor="#5582D2">
		<td width="8" valign=top rowspan=2><img src="../ill/tabel_top_left.gif" width="8" height="32" alt="" border="0"></td>
		<td colspan=2 valign="top"><img src="../ill/tabel_top.gif" width="653" height="1" alt="" border="0"></td>
		<td align=right valign=top rowspan=2><img src="../ill/tabel_top_right.gif" width="8" height="32" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td class=alt colspan=2><b>Indbakke: (den valgte mappe)</b></td>
	</tr>
	<tr>
		<td colspan=4 style="padding-left:8; border-left:1px #003399 solid; border-right:1px #003399 solid;"><br>
		



<%

  'On Error resume next

  Server.ScriptTimeOut = 90000

  function Subst (strValue, strOldValue, strNewValue)
    intLoc = InStr(strValue, strOldValue)
    While intLoc > 0
      if intLoc > 1 then
        if intLoc = Len(strValue) then
          strValue = Left(strValue, intLoc-1) & strNewValue
        else
          strValue = Left(strValue, intLoc -1) & strNewValue & Right(strValue, Len(strValue)-(intLoc-Len(strOldValue)+1))
        end if
      else
        strValue = strNewValue & Right(strValue, Len(strValue)-1)
      end if
      intLoc = InStr(strValue, strOldValue)
    Wend
    Subst = strValue
  end function

  function FixUpItems (strItem)
    if strItem <> "" then
      strItem = Subst(strItem, "<", "&lt;")
      strItem = Subst(strItem, ">", "&gt;")
      FixUpItems = strItem
    else
      FixUpItems = "<br>"
    end if
  end function

  rem ******************************************
  rem * Get the list of all message headers and
  rem *  display the info to the client
  rem ******************************************
  sub ShowMessageList (strHost, strUid, strPwd)
    'Response.Write "<h2>Messages Currently On Server: " & strHost & "</h2>"
    Set Mailer = Server.CreateObject("POP3svg.Mailer")
    
	
	Mailer.RemoteHost = strHost
    Mailer.UserName = strUid
    Mailer.Password = strPwd
	
	
    rem ******************************************
    rem * GetPopHeaders will automatically open
    rem *  the connection, grab the POP header
    rem *  information and return a variant array
    rem *  consisting of strings:
    rem *  Msg No, Subject, Date, From, Sender, To,
    rem *  Reply-to, and Size
    rem *  the more mail you have the longer it
    rem *  takes to pull this information
    rem ******************************************
    
	
	if Mailer.GetPopHeaders then
    
	  'Response.Write "<br>Found " & Mailer.MessageCount & " messages on server.<p>" & VbCrLf
      Response.Write "<table border=0 cellpadding=0 cellspacing=0 width=""630"">" & VbCrLf
      Response.Write "<tr>" & VbCrLf
      'Response.Write "<td><b>" & "Msg #" & "<b></td>" & VbCrLf
      Response.Write "<td width=""150""><b>" & "Subject" & "<b></td>" & VbCrLf
      Response.Write "<td width=""150""><b>" & "Date" & "<b></td>" & VbCrLf
      Response.Write "<td><b>" & "From" & "<b></td>" & VbCrLf

      rem Response.Write "<td><b>" & "Sender" & "<b></td>" & VbCrLf
      rem Response.Write "<td><b>" & "To" & VbCrLf
      rem Response.Write "<td><b>" & "Reply-To" & "<b></td>" & VbCrLf

      rem Response.Write "<td><b>" & "Size" & "<b></td>" & VbCrLf
      Response.Write "<td><b>" & "Status" & "<b></td>" & VbCrLf
      Response.Write "<td><b>" & "Delete" & "<b></td>" & VbCrLf
	  Response.Write "<td><b>" & "Journaliser" & "<b></td>" & VbCrLf
      Response.Write "</tr>" & VbCrLf

      varArray = Mailer.MessageInfo
      if VarType(varArray) <> vbNull And IsEmpty(varArray) <>  True then
        ArrayLimit = UBound(varArray)
        For I = 0 to ArrayLimit
          Response.Write "<tr>"
          strMsgNo = Trim(varArray(I)(0))
          rem the random number is to prevent the browser from thinking that the
          rem   page is cached when you try and delete the same message number
          Randomize
          intRndNo = Int(500 * Rnd)
          'Response.Write "<td align=right>" & strMsgNo & "</td>"

          strSubject = varArray(I)(1)
          if strSubject = "" then strSubject = "(No Subject)"
          Response.Write "<td align=left style='padding-left:5;'>" & "<a href=email.asp?msgno=" & strMsgNo & "&rndno=" & intRndNo & " class='vmenu'>" & FixUpItems (strSubject) & "</a></td>" & VbCrLf

          Response.Write "<td align=left style='padding-left:5;'>" & varArray(I)(2) & "</td>" & VbCrLf

          Response.Write "<td align=left style='padding-left:5;'>" & FixUpItems (varArray(I)(3)) & "</td>" & VbCrLf

          rem skip the sender field for this demo
          'Response.Write "<td align=left>" & FixUpItems (varArray(I)(4)) & "</td>" & VbCrLf

          rem skip the to field for this demo
          rem Response.Write "<td align=left>" & FixUpItems (varArray(I)(5)) & "</td>" & VbCrLf

          rem skip the reply-to field for the demo
          rem Response.Write "<td align=left>" & FixUpItems (varArray(I)(6)) & "</td>"

          Rem Response.Write "<td align=left>" & varArray(I)(7) & "</td>"

          strStatus = varArray(I)(8)
          if (strStatus = "") then
            strStatus = "<b>Unread</b>"
          else
            strStatus = varArray(I)(8)
          end if
          Response.Write "<td align=left style='padding-left:5;'>" & strStatus & "</td>"
          Response.Write "<td align=left style='padding-left:5;'>" & "<a href=webmail2.asp?deletemsg=" & strMsgNo & "&rndno=" & intRndNo & ">Delete</a></td>"
		  Response.Write "<td align=left style='padding-left:5;'>Journaliser</td>"
          Response.Write "</tr>" & Chr(10) & Chr(13)
		  Response.write "<tr><td colspan=6 bgcolor=#003399 height=1><img src='../ill/blank.gif' width='1' height='1' alt='' border='0'></td></tr>"
        Next
      else
        Response.Write "<tr><td colspan=10 align=center><b>No messages on server</b></tr>"
      end if
      Response.Write "</table><center><p><br>"
      Response.Write "<a href=""webmail2.asp?rndno=" & intRndNo & """>Refresh</a></center>"
    else
      Response.Write "<p>Connection Failure. Check your mailhost, username and password."
    end if
    Response.Write "</blockquote>"
  end sub

  rem ******************************************
  rem * Shows the text for one specific message
  rem ******************************************
  sub ShowMessage(strHost, strUid, strPwd, strMsgNo)
    'Response.Write "<h2>Message #" & strMsgNo & " follows:</h2><pre>"
	Response.write "<pre>"
    Set Mailer = Server.CreateObject("POP3svg.Mailer")

    
    rem ******************************************
    rem * Make sure that you've got the following
    rem *   directory set up.  Also, note that 
    rem *   if you retrieve message number 1 it 
    rem *   will be saved as c:\temp\1.txt
    rem * ** DEMO WARNING **
    rem * For MULTIUSER use you'll need to point
    rem *   the base directories to a user specific
    rem *   directory to keep users from walking
    rem *   on top of each other.
    rem ******************************************

    strMailBaseDir = "c:\temp\"

    Mailer.MailDirectory = strMailBaseDir

    Mailer.RemoteHost = strHost
    Mailer.UserName = strUid
    Mailer.Password = strPwd
    Mailer.OpenPop3
    Mailer.Pop3Log = "c:\pop3log.txt"

    rem We could do multiple retrieves here but this demo only shows the
    rem  selected message.

    Mailer.Retrieve strMsgNo

    rem Mailer.RetrieveToFile strMsgNo, "TestMsg.txt"
    Mailer.ClosePop3

    Response.Write "<table border=0 bgcolor=#ffffff cellpadding=0 cellspacing=0 width=653>" 
	'Response.Write "<tr><td width=""20%"" align=right><b>" & "Msg ID:" & "<b></td><td>" & FixUpItems(Mailer.MessageID) & "</td></tr>" & VbCrLf
	Response.Write "<tr><td width=""120"" align=right style='padding-right:5px; border-left:1px #5582d2 solid; border-top:1px #5582d2 solid;'><b>" & "Date:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid; border-top:1px #5582d2 solid;'>" & Mailer.Date & "</td></tr>" & VbCrLf
    Response.Write "<tr><td width=""120"" align=right style='padding-right:5px; border-left:1px #5582d2 solid;'><b>" & "Subject:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid;'>" & FixUpItems(Mailer.Subject) & "</td></tr>" & VbCrLf
    Response.Write "<tr><td width=""120"" align=right style='padding-right:5px; border-left:1px #5582d2 solid;'><b>" & "From:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid;'>" & Mailer.FromName & " &lt;" & Mailer.FromAddress & "&gt;</td></tr>" & VbCrLf

    rem *******************************************
    rem You can also access any field in the header
    rem   using Mailer.GetHeaderField(FieldName)
    rem *******************************************
    strTemp = Mailer.GetHeaderField("Reply-To")
    if strTemp <> "" then
      Response.Write "<tr><td width=120 align=right style='padding-right:5px; border-left:1px #5582d2 solid;'><b>" & "Reply-To:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid;'>" & FixUpItems(strTemp) & "</td></tr>" & VbCrLf
    end if

    strTemp = Mailer.Recipients
    if strTemp <> "" then
      Response.Write "<tr><td width=120 align=right style='padding-right:5px; border-left:1px #5582d2 solid;'><b>" & "To:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid;'>" & FixUpItems(strTemp) & "</td></tr>" & VbCrLf
    end if

    strTemp = Mailer.CC
    if strTemp <> "" then
      Response.Write "<tr><td width=120 align=right style='padding-right:5px; border-left:1px #5582d2 solid;'><b>" & "CC:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid;'>" & FixUpItems(strTemp) & "</td></tr>" & VbCrLf
    end if

    Response.Write "<tr><td width=120 align=right style='padding-right:5px; border-left:1px #5582d2 solid; border-bottom:1px #5582d2 solid;'><b>" & "Attachments:" & "<b></td><td style='padding-left:5px; border-right:1px #5582d2 solid; border-bottom:1px #5582d2 solid;'>" & Mailer.AttachmentCount & "</td></tr>" & VbCrLf
    Response.Write "</table><br><br>"
	Response.Write "<table border=0 bgcolor=#ffffff cellpadding=0 cellspacing=0 width=653><tr><td style='border-left:1px #5582d2 solid; padding-left:5px; border-right:1px #5582d2 solid; border-top:1px #5582d2 solid; border-bottom:1px #5582d2 solid;'><pre><br><br>" 

    Response.Write Mailer.BodyText & VbCrLf

    if Mailer.AttachmentCount > 0 then
      Response.Write "<hr></pre>" & VbCrLf

      Response.Write "Attachments will be save to " & Mailer.MailDirectory
      Response.Write "<p>In a normal Web app you'll probably want to save these files "
      Response.Write "off to a directory accessible to the Web server, create the following "
      Response.Write "list of files as links, and allow the user to download specific file "
      Response.Write "attachments. You'll have to invent a scheme where the files are erased "
      Response.Write "(on session end would be a logical place)<p>"
      Response.Write "<pre><table border=1 width=""90%"">" & VbCrLf
      Response.Write "<tr><td>Attachment</td><td>ContentType</td><td>FileName</td><td>FileSize</td></tr>" & VbCrLf
      For intCount = 1 to Mailer.AttachmentCount
        if Mailer.GetAttachmentInfo (intCount) then
          Response.Write "<tr><td>" & intCount & "</td><td>" & Mailer.AttContentType & "</td><td>" & Mailer.AttFileName & "</td><td>" & Mailer.AttFileSize & "</td></tr>" & VbCrLf
          Mailer.SaveAttachment (intCount)
        end if
      Next
      Response.Write "</table>" & VbCrLf
    end if
	
	Response.write "</pre></td></tr></table>"
	
    rem *******
    rem NOTE that the physical file is our message # - 1 & ".txt"
    rem *******
    rem strFileName = strMailBaseDir & strMsgNo-1 & ".txt"
    rem Set FileObject = CreateObject("Scripting.FileSystemObject")
    rem Set MsgFile = FileObject.OpenTextFile(strFileName, 1, False, True)
    rem Do While MsgFile.AtEndOfStream <> True
    rem   strMsgLine = MsgFile.ReadLine
    rem   Response.Write strMsgLine & "<br>"
    rem Loop
    rem MsgFile.Close
    rem Response.Write "</pre>"
    rem Response.Write "<h2>" & strFileName & "</h2>"

    rem Erase the temporary file
    rem Mailer.EraseFile(strFileName)
  end sub

  rem ******************************************
  rem * DELETE the message PERMANENTLY from the
  rem *  server
  rem ******************************************
  sub DeleteMessage (strHost, strUid, strPwd, strMsgNo)
    Response.Write "<b>Deleting Message #" & strMsgNo & " From Server</b><p>"
    Set Mailer = Server.CreateObject("POP3svg.Mailer")

    Mailer.RemoteHost = strHost
    Mailer.UserName   = strUid
    Mailer.Password   = strPwd
    Mailer.OpenPop3

    rem We could do multiple deletes here but this demo only does 1. You must
    rem  close the server at this point or our message numbers in
    rem  ShowMessageList won't be correct.

    Mailer.Delete strMsgNo
    Mailer.ClosePop3

    ShowMessageList strHost, strUid, strPwd
  end sub

  
  rem ******************************************
  rem * Main code begins here
  rem ******************************************


  
  	strSQL = "SELECT emlremotehost, emlusername, emlpassword FROM medarbejdere WHERE Mid = "& session("mid")
	oRec.open strSQL, oConn, 3 
	
	if not oRec.EOF then
	strHost = "www.outzource.dk" 'oRec("emlremotehost")
  	strUid = "universe\logomotiv\support" 'oRec("emlusername") 
  	strPwd = "dolken2" 'oRec("emlpassword") 
	end if
	
	oRec.close
	
	Response.write strHost & "<br>" & strUid & "<br>" & strPwd
	
  strMsgNo = Request.QueryString("msgno")
  strDeleteNo  = Request.QueryString("deletemsg")

  rem We can log pop3 responses to a file by assigning Object.Pop3Log
  rem This is for DEBUG use only
  rem   eg. Mailer.Pop3Log = "c:\pop3log.txt"
  rem DO NOT use in a multi-user situation with the SAME file name

  'if (strUid = "") or (strPwd = "") or (strHost = "") then
    'ShowPopForm strHost, strUid, strPwd
  'else
    if (strMsgNo <> "") then
      ShowMessage strHost, strUid, strPwd, strMsgNo
    else
      if (strDeleteNo <> "") then
        DeleteMessage strHost, strUid, strPwd, strDeleteNo
      else
        ShowMessageList strHost, strUid, strPwd
      end if    
    end if
  'end if

%>



		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	</td>	
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
		<td colspan=2><br><br><br>&nbsp;</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="80" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="120" alt="" border="0"></td>
		<td colspan=2>
		<a href="javascript:NewWin_help('crmstatus.asp?menu=tok&ketype=e');" target="_self" class='vmenu'>#</a><br>
		<br><br>&nbsp;</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="120" alt="" border="0"></td>
	</tr>
	<tr>
		<td valign="top"><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
		<td colspan=2><br>&nbsp;</td>
		<td valign="top" align=right><img src="../ill/tabel_top.gif" width="1" height="60" alt="" border="0"></td>
	</tr>
	<tr bgcolor="#5582D2">
		<td valign="top"><img src="../ill/tabel_bund_left.gif" width="8" height="10" alt="" border="0"></td>
		<td colspan="2" valign="bottom"><img src="../ill/tabel_top.gif" width="653" height="1" alt="" border="0"></td>
		<td valign="top" align="right"><img src="../ill/tabel_bund_right.gif" width="8" height="10" alt="" border="0"></td>
	</tr>
	</table>
	<br><br>
	<br>
	<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
	<br>
	<br>
	</div>
	<%
	end select
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
