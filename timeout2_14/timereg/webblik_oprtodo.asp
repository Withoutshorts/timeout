
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<%


func = request("func")

'Response.Write "func" & func
'Response.end

if len(request("FM_parent")) <> 0 then
parent = request("FM_parent")
else
parent = 0
end if

id = request("id")

if len(request("oldone")) <> 0 then
oldones = request("oldone")
else
oldones = 0
end if

 if len(trim(request("nomenu"))) <> 0 then
    nomenu = request("nomenu")
    else
    nomenu = 0
    end if

select case func

case "slet"

strSQL = "SELECT id FROM todo_new WHERE parent = " & id
oRec.open strSQL, oConn, 3 
while not oRec.EOF 
		
		'**** Delete underliggende - underliggende LEVEL 3 *** 
		strSQL = "DELETE FROM todo_new WHERE parent = " & oRec("id")
		oConn.execute(strSQL)
		
oRec.movenext
wend
oRec.close 

'**** Delete underliggende *** 
strSQL = "DELETE FROM todo_new WHERE parent = " & id
oConn.execute(strSQL)

'*** Delete valgte *****
strSQL = "DELETE FROM todo_new WHERE id = " & id
oConn.execute(strSQL)

'**** Delete raltioner ****
strSQL = "DELETE FROM todo_rel_new WHERE todoid = " & id
oConn.execute(strSQL)

Response.redirect "webblik_todo.asp?FM_parent="&parent&"&oldone="&oldones&"&lastid="&id&"&nomenu=1"

case "xop"
            sortorder = request("FM_sortorder")

            %>
			<!--#include file="inc/isint_func.asp"-->
			<%
			call erDetInt(sortorder) 
			if isInt > 0 OR instr(sortorder, ".") <> 0 then
			
			else
			
            sqlToDoVsitDato = year(now) & "/" & month(now) & "/" & day(now)
            strSQL = "UPDATE todo_new SET sortorder = "& sortorder &", dato = '"& sqlToDoVsitDato &"' WHERE id = "& id
            
            'Response.Write strSQL
            'Response.Flush
            'Response.end
            
            oConn.execute(strSQL)

            end if

Response.redirect "webblik_todo.asp?FM_parent="&parent&"&oldone="&oldones&"&lastid="&id&"&nomenu="&nomenu

case "xned"

'*** Opdaterer dato på redigering på den viste todo ***
sqlToDoVsitDato = year(now) & "/" & month(now) & "/" & day(now)
strSQL = "UPDATE todo_new SET sortorder = (sortorder + 1), dato = '"& sqlToDoVsitDato &"' WHERE id = "& id
oConn.execute(strSQL)

Response.redirect "webblik_todo.asp?FM_parent="&parent&"&oldone="&oldones&"&lastid="&id&"&nomenu="&nomenu

case else

'**** Finder level ***

	strSQL = "SELECT level FROM todo_new WHERE id = " & parent
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	
	 level = oRec("level") + 1
	
	end if
	oRec.close 


if len(level) <> 0 then
level = level 
else
level = 1
end if 

navn = replace(request("FM_todo"), "'", "")
navn = replace(navn, "''", "")
navn = replace(navn, vbcrlf, " ") 


sqlDato = year(now) & "/" & month(now) & "/" & day(now)
if len(request("FM_afsluttet")) <> 0 then
afl = request("FM_afsluttet")
else
afl = 0
end if


if len(request("FM_opdater_u")) <> 0 then
opdaterU = request("FM_opdater_u")
else
opdaterU = 0
end if

tododato = request("FM_start_aar") &"/"& request("FM_start_mrd") &"/"& request("FM_start_dag")

'Response.Write "tododato " & tododato
'Response.end

                    
if IsDate(tododato) then
tododato = tododato
else
tododato = year(now) &"/"& month(now) &"/"& day(now)
end if
			        
			       
			      
if len(request("forvafsl")) <> 0 then
forvafsl = 1
else
forvafsl = 0
end if

if len(trim(request("FM_public"))) <> 0 then
todoPublic = request("FM_public")
else
todoPublic = 0
end if

if id <> 0 then '*** Rediger 
strSQL = "UPDATE todo_new SET navn = '"& navn &"', level = "& level &", "_
&" dato = '"& sqlDato &"', afsluttet = "& afl &", "_
&" tododato = '"& tododato &"', forvafsl = "& forvafsl &", public = "& todoPublic &" WHERE id = "& id
oConn.execute(strSQL)
	
	lastID = id
	
	'*** Tømmer tidligere realtions liste ***
	oConn.execute("DELETE FROM todo_rel_new WHERE todoid = "& lastID & "") ' AND medarbid <> " & session("mid"))
	oConn.execute("UPDATE todo_new SET delt = 0 WHERE id = "& lastID & "")
	
			if opdaterU = 1 then
				
				'*** Opdaterer underliggende relationer ***
				strSQL3 = "SELECT id FROM todo_new WHERE parent = " & lastID
				oRec3.open strSQL3, oConn, 3 
				while not oRec3.EOF 
						
						'*** Tømmer tidligere relations liste ***
						oConn.execute("DELETE FROM todo_rel_new WHERE todoid = "& oRec3("id") & "") 'AND medarbid <> " & session("mid"))
						oConn.execute("UPDATE todo_new SET delt = 0 WHERE id = "& oRec3("id") & "")
						
							'*** Opdaterer underliggende - underliggende ***
							strSQL2 = "SELECT id FROM todo_new WHERE parent = " & oRec3("id")
							oRec2.open strSQL2, oConn, 3 
							while not oRec2.EOF 
							
								oConn.execute("DELETE FROM todo_rel_new WHERE todoid = "& oRec2("id") & "") 'AND medarbid <> " & session("mid"))
								oConn.execute("UPDATE todo_new SET delt = 0 WHERE id = "& oRec2("id") & "")
								
					
							oRec2.movenext
							wend
							oRec2.close
				
				oRec3.movenext
				wend
				oRec3.close
			end if		
	
else '*** Opret

strSQL = "INSERT INTO todo_new (navn, parent, dato, level, afsluttet, tododato, forvafsl, public) VALUES "_
&" ('"& navn &"', "& parent &", '"& sqlDato &"', "& level &", "& afl &", '"& tododato &"', "& forvafsl &", "& todoPublic &")"
oConn.execute(strSQL)

	'*** Henter sidst oprettet id ************
	strSQL = "SELECT id FROM todo_new WHERE id <> 0 ORDER BY id DESC"
	oRec.open strSQL, oConn, 3 
	if not oRec.EOF then
	lastID = oRec("id")
	end if
	oRec.close 
	
	'*** Opretter relation til ejer **********
	'strSQL2 = "INSERT INTO todo_rel_new (medarbid, todoid) VALUES ("& session("mid") &", "& lastID &")"
	'oConn.execute(strSQL2)
			
	
end if



	'*** Relationer ****	
	dim relmedarb
	relmedarb = split(request("FM_todo_rel"), ",")
	m = 0
	for i = 0 to UBOUND(relmedarb)
			
			'*** Opdater valgte ***
			call opdaterrelationer(lastID, relmedarb(i))
			'Response.write "opdaterer valgte<br>"
			
			if opdaterU = 1 then
				'*** Opdaterer underliggende ***
				strSQL3 = "SELECT id FROM todo_new WHERE parent = " & lastID
				oRec3.open strSQL3, oConn, 3 
				while not oRec3.EOF 
						
						call opdaterrelationer(oRec3("id"), relmedarb(i))
						'Response.write "opdaterer underliggende<br>"
						
						'*** Opdaterer underliggende - underliggende ***
						strSQL2 = "SELECT id FROM todo_new WHERE parent = " & oRec3("id")
						oRec2.open strSQL2, oConn, 3 
						while not oRec2.EOF 
						
						call opdaterrelationer(oRec2("id"), relmedarb(i))
						'Response.write "opdaterer underliggende underliggende<br>"
				
						oRec2.movenext
						wend
						oRec2.close
				
				oRec3.movenext
				wend
				oRec3.close
			end if
			
    m = m + 1
	next
	
	if m = 0 then
	
	'*** Opretter relation til ejer **********
	strSQL2 = "INSERT INTO todo_rel_new (medarbid, todoid) VALUES ("& session("mid") &", "& lastID &")"
	oConn.execute(strSQL2)
	
	end if
	
	
	
	function opdaterrelationer(todoid, medarb)
	
	'*** Deler todo liste ***
			if i > 0 then
				oConn.execute("UPDATE todo_new SET delt = 1 WHERE id = "& todoid & "")
			end if
			
			'thisMid = 0
			'*** Opretter øvrige relationer **********
			'strSQL = "SELECT mid FROM medarbejdere WHERE email = '"& medarb &"'"
			'oRec.open strSQL, oConn, 3 
			'if not oRec.EOF then
			'thisMid = oRec("mid")
			'end if
			'oRec.close 
			
			thisMid = medarb
			
			if thisMid <> 0 then
			strSQL2 = "INSERT INTO todo_rel_new (medarbid, todoid) VALUES ("& thisMid &", "& todoid &")"
			oConn.execute(strSQL2)
			'Response.write strSQL2
			end if
	
	end function

Response.redirect "webblik_todo.asp?FM_parent="&parent&"&oldone="&oldones&"&lastid="&lastID&"&nomenu="&nomenu

end select
%>



