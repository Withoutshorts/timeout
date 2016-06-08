<%
function budgettopmenu(vis)


select case vis
case 1
%><br>

<a href='budget_aar_dato.asp?menu=erp' class='rmenu'>Budget År/Dato</a>&nbsp;&nbsp;|&nbsp;&nbsp;

<%if level <= 2 OR level = 6 then  %>
<a href='budget_bruttonetto.asp?menu=erp' class='rmenu'>Brutto/Netto Dage & Timer</a>&nbsp;&nbsp;|&nbsp;&nbsp;
<a href='budget_medarb.asp?menu=erp' class='rmenu'>Budget medarbejdere</a>
<%end if %>



<!--<a href='budget_aar_dato.asp?menu=erp' class='rmenu'>Månedspuls</a>&nbsp;&nbsp;|&nbsp;&nbsp;-->
<!--<a href='budget_nogletal.asp?menu=erp' class='rmenu'>Nøgletal</a>-->
<br />
<%
end select
end function
%>