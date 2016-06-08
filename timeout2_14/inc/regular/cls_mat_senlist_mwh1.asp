<%
 
          


        public strSenesteMatHTML
        function senstematReg(usemrn, useSog, sogBilagOrJob, showonlypers, vasallemed, sogliste)
       

            strSenesteMatHTML = "<table width=""100%"" cellspacing=""0"" cellpadding=""2"" border=""0"">"

            strSenesteMatHTML = strSenesteMatHTML & "<tr bgcolor=""#5582d2"">"

            strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom class=alt>"

            strSenesteMatHTML = strSenesteMatHTML & tsa_txt_230 & "<br />"

            strSenesteMatHTML = strSenesteMatHTML & "<b>"& tsa_txt_066 &"</b><br>"

            strSenesteMatHTML = strSenesteMatHTML & tsa_txt_236 & "<br />"

            strSenesteMatHTML = strSenesteMatHTML & "~ " & tsa_txt_218
            
            if level <= 2 OR level = 6 then
            strSenesteMatHTML = strSenesteMatHTML & tsa_txt_213
            end if
            
            strSenesteMatHTML = strSenesteMatHTML & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom width=""80"" class= alt>"
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & tsa_txt_183 & "</b><br />"
                    strSenesteMatHTML = strSenesteMatHTML & tsa_txt_077
                strSenesteMatHTML = strSenesteMatHTML & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom align=right class=alt><b>" & tsa_txt_202 & "</b></td>"
                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom style=""padding:2px 5px 2px 5px;"" class= alt><b>" & tsa_txt_209 & "</b><br />Lokation</td>"
                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom align=center style=""padding:2px 5px 2px 5px;"" class=alt><b>" & tsa_txt_217 & "</b></td>"


                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom align=right style=""padding:2px 5px 2px 5px;"" class= alt>"
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & tsa_txt_219 & "</b><br />"

                    if level <= 2 OR level = 6 then
                    strSenesteMatHTML = strSenesteMatHTML & tsa_txt_220
                    end if

                    strSenesteMatHTML = strSenesteMatHTML & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom align=right style=""padding:2px 5px 2px 5px;"" class= alt>"
                strSenesteMatHTML = strSenesteMatHTML & "<b>" & tsa_txt_248 & "</b>"
                strSenesteMatHTML = strSenesteMatHTML & "</td>"


                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom class= alt>"
                    strSenesteMatHTML = strSenesteMatHTML & left(tsa_txt_251, 3) & "."
                strSenesteMatHTML = strSenesteMatHTML & "</td>"
                strSenesteMatHTML = strSenesteMatHTML & "<td valign=bottom class= alt>" & tsa_txt_221 & "</td>"
            strSenesteMatHTML = strSenesteMatHTML & "</tr>"
 
           
            'if cint(useSog) = 1 then
            'if cint(sogBilagOrJob) = 0 then
            sqlWh = "(mf.bilagsnr LIKE '"& sogliste &"' OR j.jobnr LIKE '"& sogliste &"' OR j.jobnavn LIKE '"& sogliste &"%' OR matnavn LIKE '"& sogliste &"%')"
            'else
            'sqlWh = "j.jobnr LIKE '"& sogliste &"'"
            'end if
            'else
            'sqlWh = "j.jobnr <> 0"
            'end if

            if cint(showonlypers) = 1 then
            strSQLper = " AND personlig = 1"
            else
            strSQLper = ""
            end if

            if cint(vasallemed) = 1 then
            sqlWh = sqlWh &  " AND (usrid <> 0) "
            else
            sqlWh = sqlWh & " AND (usrid = "& usemrn & ")"
            end if

            strSQLmat = "SELECT m.mnavn AS medarbejdernavn, mnr, init, mf.id AS mfid, mf.matvarenr AS varenr, mg.navn AS gnavn, mf.matenhed AS enhed, "_
            &" mf.matnavn AS navn, mf.matantal AS antal, mf.dato, mf.editor, "_
            & "mf.matkobspris, mf.matsalgspris, mf.jobid, mf.matgrp, "_
            &" mf.usrid, mf.forbrugsdato, j.id, j.jobnr, j.jobnavn, "_
            &" mg.av, f.fakdato, k.kkundenavn, "_
            &" k.kkundenr, mf.valuta, mf.intkode, mf.bilagsnr, v.valutakode, mf.personlig, j.serviceaft, "_
            &" s.navn AS aftnavn, s.aftalenr, mf.kurs, ma.lokation, j.risiko, a.navn AS aktnavn, a.id AS aktid "_
            &" FROM materiale_forbrug mf"_
            &" LEFT JOIN materialer ma ON (ma.id = matid) "_
            &" LEFT JOIN materiale_grp mg ON (mg.id = mf.matgrp) "_
            &" LEFT JOIN medarbejdere m ON (mid = usrid) "_
            &" LEFT JOIN job j ON (j.id = mf.jobid) "_
            &" LEFT JOIN aktiviteter a ON (a.id = mf.aktid) "_
            &" LEFT JOIN serviceaft s ON (s.id = j.serviceaft) "_
            &" LEFT JOIN fakturaer f ON (f.jobid = mf.jobid AND f.faktype = 0) "_
            &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
            &" LEFT JOIN valutaer v ON (v.id = mf.valuta) "_
            &" WHERE "& sqlWh &" "& strSQLper &" GROUP BY mf.id ORDER BY mf.id DESC, mf.forbrugsdato DESC, f.fakdato DESC LIMIT 100"

            'mf.jobid = "& id &"
            'if session("mid") = 21 then
            'response.write strSQLmat
            'Response.flush
            'end if
            
            matkobsprisialt = 0
            s = 0
            oRec.open strSQLmat, oConn, 3
            while not oRec.EOF


            if len(oRec("fakdato")) <> 0 then
            fakdato = oRec("fakdato")
            else
            fakdato = "01/01/2002"
            end if

            if oRec("mfid") = cdbl(lastId) then
            bgthis = "#FFFF99"
            else
            select case right(s, 1)
            case 0,2,4,6,8
            bgthis = "#EFf3FF"
            case else
            bgthis = "#FFFFff"
            end select
            end if
           

            strSenesteMatHTML = strSenesteMatHTML & "<tr bgcolor=" & bgthis & " class=lille>"

                useBr = ""
               
                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top style=""padding:5px 2px 2px 2px; width:250px; white-space:nowrap; border-bottom:1px d6dff5 solid;"" class=lille>"
                    if len(oRec("bilagsnr")) <> 0 then
                    strSenesteMatHTML = strSenesteMatHTML & "<span style=""color:#999999;"">"& oRec("bilagsnr") & "</span><br>"
                    'else
                    'useBr = "<br />"
                    end if

                    strSenesteMatHTML = strSenesteMatHTML & "<b>"& left(oRec("kkundenavn"), 15) & " </b><br />"
                    strSenesteMatHTML = strSenesteMatHTML & left(oRec("jobnavn"), 20) & " ("& oRec("jobnr") &")"

                    if oRec("aktid") <> 0 then
                    strSenesteMatHTML = strSenesteMatHTML & "<br /><span style=""color:#5582d2;"">" & left(oRec("aktnavn"), 20) & "</span>"
                    end if

                    if oRec("serviceaft") <> "0" then
                    strSenesteMatHTML = strSenesteMatHTML & "<br /><span style=""color:#999999;"">Aft:" & oRec("aftnavn") & "(" & oRec("aftalenr") & ")</span>"
                    end if

                    strSenesteMatHTML = strSenesteMatHTML & "<font class=""megetlillesort"">"
                        if len(oRec("gnavn")) <> 0 then
                        strSenesteMatHTML = strSenesteMatHTML & "<br /> ~ " & oRec("gnavn")
                            if level <= 2 OR level = 6 then
                            strSenesteMatHTML = strSenesteMatHTML & "(" & oRec("av") & "%)"
                            end if
                        end if
                    strSenesteMatHTML = strSenesteMatHTML & "</font>&nbsp;"

                strSenesteMatHTML = strSenesteMatHTML & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top class=lille style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; white-space:nowrap;"">"
                    
    
                    if len(oRec("forbrugsdato")) <> 0 then
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & formatdatetime(oRec("forbrugsdato"), 2) & "</b><br />"
                    end if

                    strSenesteMatHTML = strSenesteMatHTML & "<span style=""color:#5C75AA; font-size:9px;"">" & left(oRec("medarbejdernavn"), 15) & " ("& oRec("mnr") &")" & "</span>"

                    strSenesteMatHTML = strSenesteMatHTML & "<span style=""color:#999999; font-size:9px;"">"

                        select case oRec("intkode")
                        case 0
                        intKode = "-"
                        case 1
                        intKode = tsa_txt_239 'intern
                        case 2
                        intKode = tsa_txt_240 'ekstern
                        'case 3
                        'intKode = tsa_txt_241
                        end select

                        strSenesteMatHTML = strSenesteMatHTML & "<br />"

                        if intKode <> "-" then
                        strSenesteMatHTML = strSenesteMatHTML & intKode 
                        end if

                        if oRec("personlig") <> 0 then
                        strSenesteMatHTML = strSenesteMatHTML &" - "& tsa_txt_234
                        end if

                    strSenesteMatHTML = strSenesteMatHTML & "</span></td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top align=right style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; white-space:nowrap;"" class=lille>"
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & oRec("antal") & "</b>&nbsp;"

                    if len(oRec("enhed")) <> 0 then
                    enh = oRec("enhed")
                    else
                    enh = tsa_txt_222 '"Stk."
                    end if

                 
                strSenesteMatHTML = strSenesteMatHTML & enh & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;"" class=lille>"
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & oRec("navn") & "</b>&nbsp;"
                    if len(trim(oRec("lokation"))) <> 0 then
                    strSenesteMatHTML = strSenesteMatHTML & "<span style=""font-size:9px; color:#999999;"">" & oRec("lokation") & "</span>"
                    end if
                strSenesteMatHTML = strSenesteMatHTML & "</td>"
                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top align=right style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;"" class=lille>" & oRec("varenr") & "</td>"


                    strSenesteMatHTML = strSenesteMatHTML & "<td valign=top align=right style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;"" class=lille>"
                    strSenesteMatHTML = strSenesteMatHTML & "<b>" & formatnumber(oRec("matkobspris"), 2) & "</b>"

                    
                    matkobsprisIalt = matkobsprisialt + (oRec("matkobspris")/1) * (oRec("kurs")/100)


                    if level <= 2 OR level = 6 then
                    strSenesteMatHTML = strSenesteMatHTML & "<br />"
                    strSenesteMatHTML = strSenesteMatHTML & formatnumber(oRec("matsalgspris"), 2)
                    end if

                strSenesteMatHTML = strSenesteMatHTML & "</td>"

                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top align=right style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid; white-space:nowrap;"" class=lille>"
                    
                kobsprisialt = formatnumber(oRec("antal") * oRec("matkobspris"), 2)
                
    
                strSenesteMatHTML = strSenesteMatHTML & "<b>" & kobsprisialt & "&nbsp;" & oRec("valutakode") & "</b>"

                    if level <= 2 OR level = 6 then
                    strSenesteMatHTML = strSenesteMatHTML & "<br />"
                    salgsprisialt = formatnumber(oRec("antal") * oRec("matsalgspris"), 2)
                    strSenesteMatHTML = strSenesteMatHTML & salgsprisialt & "&nbsp;" & oRec("valutakode")

                    end if

                strSenesteMatHTML = strSenesteMatHTML & "</td>"


                strSenesteMatHTML = strSenesteMatHTML & "<td valign=top style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;"">"

                    
                    '*** Er uge alfsuttet af medarb, er smiley og autogk slået til
                    erugeafsluttet = instr(afslUgerMedab(usemrn), "#"&datepart("ww", oRec("forbrugsdato"),2,2)&"_"& datepart("yyyy", oRec("forbrugsdato")) &"#")

                    'Response.Write "erugeafsluttet --" & erugeafsluttet  &"<br>"
                    'Response.flush


                    strMrd_sm = datepart("m", oRec("forbrugsdato"), 2, 2)
                    strAar_sm = datepart("yyyy", oRec("forbrugsdato"), 2, 2)
                    strWeek = datepart("ww", oRec("forbrugsdato"), 2, 2)
                    strAar = datepart("yyyy", oRec("forbrugsdato"), 2, 2)

                    if cint(SmiWeekOrMonth) = 0 then
                    usePeriod = strWeek
                    useYear = strAar
                    else
                    usePeriod = strMrd_sm
                    useYear = strAar_sm
                    end if


                    call erugeAfslutte(useYear, usePeriod, usemrn)

                    'Response.Write "smilaktiv: "& smilaktiv & "<br>"
                    'Response.Write "SmiWeekOrMonth: "& SmiWeekOrMonth &" ugeNrAfsluttet: "& ugeNrAfsluttet & " tjkDag: "& tjkDag &"<br>"
                    'Response.Write "autolukvdatodato: "& autolukvdatodato & "<br>"
                    'Response.Write "tjkDag: "& tjkDag & "<br>"
                    'Response.Write "autolukvdato: "& autolukvdato & "<br>"
                    'Response.Write "erugeafsluttet:" & erugeafsluttet & "<br>"

                    call lonKorsel_lukketPer(oRec("forbrugsdato"), oRec("risiko"))

                    'if (cint(erugeafsluttet) <> 0 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                    if ( (( datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0) OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) = year(now) AND DatePart("m", oRec("forbrugsdato")) < month(now)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) = 12)) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", oRec("forbrugsdato")) < year(now) AND DatePart("m", oRec("forbrugsdato")) <> 12) OR _
                    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", oRec("forbrugsdato")) > 1))) OR cint(lonKorsel_lukketIO) = 1 then

                    ugeerAfsl_og_autogk_smil = 1
                    else
                    ugeerAfsl_og_autogk_smil = 0
                    end if



                    if (ugeerAfsl_og_autogk_smil = 0 _
                    OR ugeerAfsl_og_autogk_smil = 1 AND level = 1) _
                    AND cdate(fakdato) < cdate(oRec("forbrugsdato")) OR (oRec("intkode") <> 2)  then 'intkode <> 2 ekstern


                    '*** Kun materialer der ikke er oprettet på laver skal kunne redigeres ***'
                    '*** Ændre denne så man vælger ved flueben fra matreg. om det skal oprettes på lager **'

                    if oRec("varenr") = "0" then	
                    strSenesteMatHTML = strSenesteMatHTML & "<a href=""materialer_indtast.asp?id=" & oRec("id") & "&func=red&matregid=" & oRec("mfid") & "&lastid=" & oRec("mfid") & "&fromsdsk=" & fromsdsk & "&aftid=" & aftid &  "&vis=otf&mid="&usemrn&"""><img src=""../ill/ac0059-16.gif"" alt=""& tsa_txt_251 &"" border=0 /></a>&nbsp;"
                    else
                    strSenesteMatHTML = strSenesteMatHTML & "&nbsp;"
                    end if

                    strSenesteMatHTML = strSenesteMatHTML & "</td>"
                    strSenesteMatHTML = strSenesteMatHTML & "<td valign=top style=""padding:5px 2px 2px 2px; border-bottom:1px d6dff5 solid;"">"
                    strSenesteMatHTML = strSenesteMatHTML & "&nbsp;<a href=""materialer_indtast.asp?id="& oRec("id") &"&func=slet&matregid="& oRec("mfid") & "&fromsdsk=" & fromsdsk & "&aftid=" & aftid & "&vis=" & vis & "&mid="& usemrn &"""><img src=""../ill/slet_16.gif"" alt="& tsa_txt_221 &" border=0 /></a>"
                    else
                    strSenesteMatHTML = strSenesteMatHTML & "&nbsp;"
                    strSenesteMatHTML = strSenesteMatHTML & "</td>"
                    strSenesteMatHTML = strSenesteMatHTML & "<td>"
                    strSenesteMatHTML = strSenesteMatHTML & "&nbsp;"
                    end if

                strSenesteMatHTML = strSenesteMatHTML & "</td>"

            strSenesteMatHTML = strSenesteMatHTML & "</tr>"

            
            s = s + 1
            oRec.movenext
            wend

            oRec.close
            
            strSenesteMatHTML = strSenesteMatHTML & "<tr>"
            strSenesteMatHTML = strSenesteMatHTML & "<td colspan=9 style=""padding:5px 2px 2px 2px;"">Ialt: " & s & " registreringer, Købspris ialt: " & formatnumber(matkobsprisialt, 2) &" "& basisValISO &"</td>"
            strSenesteMatHTML = strSenesteMatHTML & "</tr>"
            strSenesteMatHTML = strSenesteMatHTML & "</table>"


            
            '*** ÆØÅ **'
            call jq_format(strSenesteMatHTML)
            strSenesteMatHTML = jq_formatTxt
            response.write strSenesteMatHTML

        end function

%>