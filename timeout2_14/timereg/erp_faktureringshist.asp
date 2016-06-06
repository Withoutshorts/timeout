<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html><head><title>
</title></head>
<body>


TSA | CRM | SDSK | <a href="erp_faktureringshist.asp">ERP</a><br />

<font size=2><a href="erp_faktureringshist.asp">Fakturering</a> | Bogføring</font><br />

<font size=1><a href="../timereg/erp_tilfakturering.asp">Til Fakturering</a> | <a href="erp_faktureringshist.asp">Søg i Fakturaer</a> | Afstemning </font>


<table>
<tr><td><br />
<b>Funktioner:</b><br />
Kvikfakturering | 
Eksport | 
Print venlig | 
Ventetimer oversigt.  | Opret faktura skrivelse 
</td></tr>
</table>

<table>
<tr>
<td><br /><b>Søg:</b>

<br /> A)    <select id="Select1">
        <option>Kontakter</option>
    </select>
    &nbsp;
    <select id="Select1">
        <option>Konto</option>
    </select>
    &nbsp;
    <select id="Select1">
        <option>Kontakt ansvarlig</option>
    </select>
    &nbsp;
    <select id="Select1">
        <option>Job</option>
    </select>
    
    &nbsp;
    <select id="Select1">
        <option>Start Dato</option>
    </select>
    
     &nbsp;
    <select id="Select1">
        <option>Slut Dato</option>
    </select>
    
&nbsp;
    
      <select id="Select1">
        <option>Aftaler</option>
    </select>
    
    &nbsp;<br />
    B) 
    <input id="Text1" type="text" value="Faktura nr. / Jobnr, Aftale ID" /> % = Wildcard
&nbsp;<input id="Submit1" type="submit" value="submit" />
</td>

</tr>
<tr>
<td>
<br />
<b>Filter:</b> 
    <select id="Select1">
    <option>Faktura</option>
        <option>Kreditnota</option>
        <option>Rykker</option>
        <option>Alle</option>
        </select> Aktive Job, Aktive aftaler, Eksporterede aftaler, Fakturaer.
</td>
</tr>
</table>

<br />


<table width=80%>
<tr>
<td>Kontakt</td><td>Konto</td>
<td>Job</td><td>Aftale</td><td>Type</td><td>Faktura NR</td><td>Faktura dato</td><td>Fak. beløb<br />(excl. moms)</td>

</tr>
<tr>
<td>OutZourCE (5)</td>
<td>268420000</td>
<td>Hjemmeside (123)</td>
<td>Support Aft. (1234)</td>
<td>Faktura</td>
<td>386</td>
<td>20-04-2007</td>
<td>2965,65</td>
</tr>
<tr>
<td>OutZourCE (5)</td>
<td>268420000</td>
<td>Hjemmeside (123)</td>
<td></td>
<td>Faktura</td>
<td>387</td>
<td>20-04-2007</td>
<td>2965,65</td>
</tr>
<tr>
<td>OutZourCE (5)</td>
<td>268420000</td>
<td></td>
<td>Support Aft. (1234)</td>
<td>Rykker</td>
<td>388</td>
<td>20-04-2007</td>
<td>2965,65</td>
</tr>
<tr>
<td colspan=7>
    &nbsp;</td>
   <td><b>6965,65</b></td>
</tr>

</table>


</body>

</html>
