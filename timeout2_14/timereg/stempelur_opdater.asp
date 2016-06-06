<%response.buffer = true%>


<%
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")






'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;"
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
'strConnect = "driver={MySQL};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lysta;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_execon;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kringit;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_proveno;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oberg;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_norma;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_stejle;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_novo_qdc;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_external;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_perspektivait;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansoft;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bminds"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_workpartners"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gp"
'strConnect = "driver={MySQL};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_optimizer4u"


	

	
	if cbool(oConn.state) = false then
	oConn.open strConnect
	end if
	
	
			strSQL = "SELECT l.id AS lid, l.mid AS lmid, l.login, l.logud, "_
			&" l.stempelurindstilling, l.minutter FROM login_historik l"_
			&" WHERE (l.dato BETWEEN '2005/01/01' AND '2007/01/01') AND l.mid <> 0 "_
			&" ORDER BY lid" 
			
			'Response.write strSQL
			'Response.flush
			
			x = 0
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
			
			
				timerThis = 0
				timerThisDIFF = 0
				
				if len(oRec("login")) <> 0 AND len(oRec("logud")) <> 0 then
				
				loginTidAfr = left(formatdatetime(oRec("login"), 3), 5)
				logudTidAfr = left(formatdatetime(oRec("logud"), 3), 5)
				
				timerThisDIFF = datediff("s", loginTidAfr, logudTidAfr)/60
				
						'strSQL2 = "UPDATE login_historik SET minutter = "& timerThisDIFF &" WHERE id = "& oRec("lid")
						'oConn.execute strSQL2
				
				end if
			
			oRec.movenext
			wend
			oRec.close
	
	
	
	if cbool(oConn.state) = true then
	response.write strConnect & "&nbsp;&nbsp;ok!<hr>"
	response.flush
	oConn.close
	end if
	
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
