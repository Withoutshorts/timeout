<%
function sdsktopmenu()
%><br>

<%if level = 1then%>
<a href="javascript:NewWin_help('sdsk_prioitet.asp?menu=tok');" target="_self" class='rmenu'>Incident Typer</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="javascript:NewWin_help('sdsk_prio_typ.asp?menu=tok');" target="_self" class='rmenu'>Prioiteter</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="javascript:NewWin_help('sdsk_prio_grp.asp?menu=tok');" target="_self" class='rmenu'>Aftalegrupper</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="javascript:NewWin_help('sdsk_status.asp?menu=tok');" target="_self" class='rmenu'>Incident Status</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<%
end if

if level <= 2 OR level = 6 then%>
<a href="javascript:NewWin_help('sdsk_typer.asp?menu=tok');" target="_self" class='rmenu'>Incident Kategorier</a>
<%
end if

end function
	

%>