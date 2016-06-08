
<%
    Dim oConnAspX As New sqlConnection(strConnect)
    
 Sub antalMediPgrp(byRef pg)

    Dim strSQLam 'As String
    Dim antalMediPgrpX 'As Integer
    
    '** antal aktive og passive medarbejdere i grp. ***'
    strSQLam = "SELECT COUNT(pgrel2.MedarbejderId) AS antalmed FROM medarbejdere AS m "
    strSQLam = strSQLam & "LEFT JOIN progrupperelationer AS pgrel2 ON "
    strSQLam = strSQLam & "(pgrel2.MedarbejderId = m.mid AND pgrel2.projektgruppeid = " & pg & ") WHERE "
    strSQLam = strSQLam & "m.mansat <> 2 AND pgrel2.projektgruppeid = " & pg & " GROUP BY projektgruppeid"
            
    Response.Write(strSQLam)
    
    oRec4.open(strSQLam, oConnAspX, 3)
    If Not oRec4.EOF Then
        antalMediPgrpX = oRec4("antalmed")
    End If
    oRec4.close()

End Sub
    
%>


