
<%
    
     select case lto
     case "epi", "epi_sta", "epi_ab", "epi_no", "epi_cati", "tec", "esn"
     arrayHigh = 1400
     case "esn", "tec", "intranet - local"
     arrayHigh = 600
     case else
     arrayHigh = 150
     end select
 
 
public  fTimer, ifTimer, km, sTimer, flexTimer,  sundTimer, pausTimer 
public	fpTimer, feriePLTimer, fefriTimer, fefriTimerBr, ferieAFTimer, ferieAFPerTimer, ferieAFPerTimertimer 
public	ferieOptjtimer, ferieUdbTimer, fefriTimerUdb, sygTimer, BarnsygTimer
public	afspTimer, afspTimerTim, afspTimerBr, afspTimerUdb, afspTimerBrPer
public	afspTimerOUdb, natTimer, weekendTimer, weekendnattimer 
public	aftenTimer, aftenweekendTimer, adhocTimer, stkAntal
public 	lageTimer, e1Timer, e2Timer
public dagTimer, fefriplTimer, ferieAFulonTimer, fefriTimerPerBr
public ferieaarSlut, ferieaarStart, afspTimerTimPer, afspTimerPer, sygaarSt
public sygDage, sygDagePer, sygTimerPer, barnSygTimerPer, barnSygDagePer, barnSygDage
public startdatoFerieLabel, slutdatoFerieLabel, barsel, ferieAFPerstDato, feriefriAFPerstDato, barselPerstDato, barnsygPerstDato, sygPerstDato, omsorgPerstDato, afspadPerstDato  
public  omsorg, senior, divfritimer, fefriTimerPerBrTimer, ferieFriaarStart, ferieFriaarSlut, startdatoFerieFriLabel, slutdatoFerieFriLabel
public fefriTimerUdbPer, fefriTimerUdbPerTimer
public ferieAFulonPer, ferieAFulonPerTimer, ferieUdbPer, ferieUdbPerTimer, ferieOptjUlontimer, ferieOptjOverforttimer, korrektionKomG, korrektionReal, tjenestefri 
public aldersreduktionOpj, aldersreduktionBr, aldersreduktionUdb
public aldersreduktionPl, omsorg2pl, omsorg10pl, omsorgKpl, ulempe1706udb, ulempeWudb, rejseDage, e3timer 
	    

        redim fTimer(arrayHigh), ifTimer(arrayHigh), km(arrayHigh), sTimer(arrayHigh), flexTimer(arrayHigh),  sundTimer(arrayHigh), pausTimer(arrayHigh) 
        redim fpTimer(arrayHigh), feriePLTimer(arrayHigh), fefriTimer(arrayHigh), fefriTimerBr(arrayHigh), afspTimerBrPer(arrayHigh), ferieAFTimer(arrayHigh) 
        redim ferieOptjtimer(arrayHigh), ferieUdbTimer(arrayHigh), fefriTimerUdb(arrayHigh), sygTimer(arrayHigh), BarnsygTimer(arrayHigh), fefriTimerUdbPer(arrayHigh), fefriTimerUdbPerTimer(arrayHigh)
        redim afspTimer(arrayHigh), afspTimerTim(arrayHigh), afspTimerBr(arrayHigh), afspTimerUdb(arrayHigh), ferieAFPerTimer(arrayHigh), ferieAFPerTimertimer(arrayHigh) 
        redim afspTimerOUdb(arrayHigh), natTimer(arrayHigh), weekendTimer(arrayHigh), weekendnattimer(arrayHigh)
        redim aftenTimer(arrayHigh), aftenweekendTimer(arrayHigh), adhocTimer(arrayHigh), stkAntal(arrayHigh)
        redim lageTimer(arrayHigh), e1Timer(arrayHigh), e2Timer(arrayHigh)
        redim dagTimer(arrayHigh), fefriplTimer(arrayHigh), ferieAFulonTimer(arrayHigh), fefriTimerPerBr(arrayHigh)
        redim afspTimerTimPer(arrayHigh), afspTimerPer(arrayHigh), sygDage(arrayHigh), sygTimerPer(arrayHigh), sygDagePer(arrayHigh)
        redim barnSygTimerPer(arrayHigh), barnSygDagePer(arrayHigh), barnSygDage(arrayHigh), barsel(arrayHigh)
        redim ferieAFPerstDato(arrayHigh), feriefriAFPerstDato(arrayHigh), barselPerstDato(arrayHigh), barnsygPerstDato(arrayHigh), sygPerstDato(arrayHigh), omsorgPerstDato(arrayHigh), afspadPerstDato(arrayHigh) 
        redim omsorg(arrayHigh), senior(arrayHigh), divfritimer(arrayHigh), fefriTimerPerBrTimer(arrayHigh)
        redim ferieAFulonPer(arrayHigh), ferieAFulonPerTimer(arrayHigh), ferieUdbPer(arrayHigh), ferieUdbPerTimer(arrayHigh)
        redim ferieOptjUlontimer(arrayHigh), ferieOptjOverforttimer(arrayHigh), korrektionKomG(arrayHigh), korrektionReal(arrayHigh), tjenestefri(arrayHigh)
        redim aldersreduktionOpj(arrayHigh), aldersreduktionBr(arrayHigh), aldersreduktionUdb(arrayHigh)
        redim omsorg2pl(arrayHigh), omsorg10pl(arrayHigh), omsorgKpl(arrayHigh), aldersreduktionPl(arrayHigh), ulempe1706udb(arrayHigh), ulempeWudb(arrayHigh), rejseDage(arrayHigh), e3timer(arrayHigh) 

	     
function fordelpaaaktType(intMid, startDato, slutDato, visning, akttype_sel, x)
         
         'Response.Write  "startDato, slutDato, visning: " & startDato &","& slutDato &","& visning 
         
	     fTimer(x) = 0
	     ifTimer(x) = 0
	     
	     km(x) = 0
	     
	     sTimer(x) = 0
	     
	     flexTimer(x) = 0
	     sundTimer(x) = 0
	     pausTimer(x) = 0
	     fpTimer(x) = 0
	     
	               
	     fefriTimer(x) = 0
	     fefriplTimer(x) = 0
	     
	     fefriTimerBr(x) = 0
	     fefriTimerPerBr(x) = 0 'dage
	     fefriTimerPerBrTimer(x) = 0

	     ferieAFTimer(x) = 0
	     ferieAFPerTimer(x) = 0 'dage
	     ferieAFPerTimertimer(x) = 0

	     ferieAFulonTimer(x) = 0
	     ferieOptjtimer(x) = 0
	     ferieOptjUlontimer(x) = 0
         ferieOptjOverforttimer(x) = 0


	     ferieUdbTimer(x) = 0

         ferieAFulonPer(x) = 0
         ferieAFulonPerTimer(x) = 0
         ferieUdbPer(x) = 0 
         ferieUdbPerTimer(x) = 0
	     
	     fefriTimerUdb(x) = 0
         fefriTimerUdbPer(x) = 0
         fefriTimerUdbPerTimer(x) = 0
	     
	     sygTimer(x) = 0
	     sygDage(x) = 0
	     sygTimerPer(x) = 0
	     sygDagePer(x) = 0
	     
	     BarnsygTimer(x) = 0
	     barnSygTimerPer(x) = 0
         barnSygDage(x) = 0
         barnSygDagePer(x) = 0
	     
	     afspTimer(x) = 0
	     afspTimerTim(x) = 0
	     
	     afspTimerTimPer(x) = 0
	     afspTimerPer(x) = 0
	     
	     afspTimerBr(x) = 0
	     afspTimerBrPer(x) = 0
	     afspTimerUdb(x) = 0
	     afspTimerOUdb(x) = 0
	     
	     dagTimer(x) = 0
	     
	     natTimer(x) = 0
	     
	     weekendTimer(x) = 0
	     
	     weekendnattimer(x) = 0
	     
	     aftenTimer(x) = 0
	     
	     aftenweekendTimer(x) = 0
	     
	     adhocTimer(x) = 0
	     
	     stkAntal(x) = 0
	     
	     lageTimer(x) = 0
	    
	     e1Timer(x) = 0
	     e2Timer(x) = 0
         e3Timer(x) = 0
    
         barsel(x) = 0


         omsorg(x) = 0
         senior(x) = 0
         divfritimer(x) = 0
         
         korrektionKomG(x) = 0
         korrektionReal(x) = 0

        tjenestefri(x) = 0

        aldersreduktionOpj(x) = 0
        aldersreduktionBr(x) = 0
        aldersreduktionUdb(x) = 0 

        omsorg2pl(x) = 0
        omsorg10pl(x) = 0
        omsorgKpl(x) = 0
        aldersreduktionPl(x) = 0 


        ulempe1706udb(x) = 0
        ulempeWudb(x) = 0
        
        rejseDage(x) = 0


        per13wrt = 0
        per14wrt = 0
        per20wrt = 0
        per21wrt = 0
        per30wrt = 0
        per31wrt = 0
        
        afstemnul(x) = 0 
        
        aktiveTyper = akttype_sel

        ferieAFPerstDato(x) = "2001/12/31"
        feriefriAFPerstDato(x) = "2001/12/31"
        barselPerstDato(x) = "2001/12/31"
        barnsygPerstDato(x) = "2001/12/31"
        sygPerstDato(x) = "2001/12/31"
        omsorgPerstDato(x) = "2001/12/31"
        afspadPerstDato(x) = "2001/12/31" 
            
            
         '*** Ferie pl. kun fra DagsDato ***'
	    
	    startdatoSQLNow = year(now) &"/"& month(now) & "/"& day(now)
	    startdatoSQL = startdato
	    
            select case lto

            case "akelius", "xintranet - local"

                    slutdatoFerie = slutdato
                    startdatoFerie = datepart("yyyy", slutdato, 2, 2) & "-01-01"
                    slutdatoFeriePl = datepart("yyyy", slutdato, 2, 2) & "-12-31"
        
            case else
	  
	                '**Indstiller ferieperiode til ferie år (planlagt felter > dd) '***
	                if cdate(slutdato) >= cdate("1-5-"& datepart("yyyy", slutdato, 2, 2)) AND cdate(slutdato) <= cdate("31-12-"& datepart("yyyy", slutdato, 2, 2)) then
	                'slutdatoFerie = datepart("yyyy", (dateadd("yyyy", "1",slutdato)), 2, 2) & "-04-30"
	                startdatoFerie = datepart("yyyy", slutdato, 2, 2) & "-05-01"
                    slutdatoFerie = slutdato
                    slutdatoFeriePl = datepart("yyyy", (dateadd("yyyy", "1",slutdato)), 2, 2) & "-04-30"

	                else
	                'slutdatoFerie = datepart("yyyy", slutdato, 2, 2) & "-04-30"
	                slutdatoFerie = slutdato
                    startdatoFerie = datepart("yyyy", (dateadd("yyyy", "-1",slutdato)), 2, 2) & "-05-01"
                    slutdatoFeriePl = datepart("yyyy", slutdato, 2, 2) & "-04-30"
	                end if

            end select

	        
	        ferieaarStart = datepart("yyyy", startdatoFerie, 2, 2)
	        ferieaarSlut = datepart("yyyy", slutdatoFerie, 2, 2)
	        
	        slutdatoFerieLabel = datepart("d",slutdatoFerie,2,2) &"/"& datepart("m",slutdatoFerie,2,2) &"/"& datepart("yyyy",slutdatoFerie,2,2)
            startdatoFerieLabel = datepart("d",startdatoFerie,2,2) &"/"& datepart("m",startdatoFerie,2,2) &"/"& datepart("yyyy",startdatoFerie,2,2)

            
            '****************************************************
            '**** Følger feriefridag kalender år eller ferie år REGEL: Kalender år
            '****************************************************
            select case lcase(lto)
            case "mi", "intranet - local" '** kalender år
            
            'if session("mid") = 1 then
            'Response.write " Startdato: "& startdato &" slutdato:" & slutdato & "::"& year(now) & " visning;" & visning & "<br>"
            'end if       
            

            if cDate(now) >= cDate("1-5-"& year(now)) then
            useFeDtKri = startdato
            else
            useFeDtKri = slutdato
            end if    
              

            if visning <> 4 then '' saldo på timereg.
            startdatoFerieFri = year(useFeDtKri)&"/1/1"
            slutdatoFerieFri = year(useFeDtKri)&"/12/31" '"&month(useFeDtKri)&"/"&day(useFeDtKri)
            else
            startdatoFerieFri = year(useFeDtKri)&"/1/1"
            slutdatoFerieFri = year(useFeDtKri)&"/12/31" '&month(startdato)&"/"&day(startdato)
            end if

            ferieFriaarStart = datepart("yyyy", startdatoFerieFri, 2, 2)
	        ferieFriaarSlut = datepart("yyyy", slutdatoFerieFri, 2, 2)
	        
	        slutdatoFerieFriLabel = datepart("d",slutdatoFerieFri,2,2) &"/"& datepart("m",slutdatoFerieFri,2,2) &"/"& datepart("yyyy",slutdatoFerieFri,2,2)
            startdatoFerieFriLabel = datepart("d",startdatoFerieFri,2,2) &"/"& datepart("m",startdatoFerieFri,2,2) &"/"& datepart("yyyy",startdatoFerieFri,2,2)


            case else ''** ferieår
            startdatoFerieFri = startdatoFerie
            slutdatoFerieFri = slutdatoFerie


            ferieFriaarStart = datepart("yyyy", startdatoFerieFri, 2, 2)
	        ferieFriaarSlut = datepart("yyyy", slutdatoFerieFri, 2, 2)
	        
	        slutdatoFerieFriLabel = datepart("d",slutdatoFerieFri,2,2) &"/"& datepart("m",slutdatoFerieFri,2,2) &"/"& datepart("yyyy",slutdatoFerieFri,2,2)
            startdatoFerieFriLabel = datepart("d",startdatoFerieFri,2,2) &"/"& datepart("m",startdatoFerieFri,2,2) &"/"& datepart("yyyy",startdatoFerieFri,2,2)

            end select

            'Response.write "Her:" & ferieFriaarStart & "<br>"

            
	        
          

	        
	        '*** hvis ansat dato er senere end licens st. dato **'
	        if cDate(meAnsatDato) > cDate(lisStDato) then
	        lisStDatoSQL = year(meAnsatDato) &"/"& month(meAnsatDato) &"/"& day(meAnsatDato) 'meAnsatDato
	        else
	        lisStDatoSQL = year(lisStDato) &"/"& month(lisStDato) &"/"& day(lisStDato)
	        end if
	        
	        if visning = 1 OR visning = 5 then '** MD --> Dato
	        lisStDatoSQL = lisStDatoSQL 'year(lisStDato) &"/"& month(lisStDato) &"/"& day(lisStDato)
	        else
	        lisStDatoSQL = year(startdato) &"/"& month(startdato) &"/"& day(startdato) 'startdato
	        slutdatoFerie = startdato
	        slutdatoFerie = slutdato
	        end if

            

              '*** Syg ***'
            sygaarSt = year(startdato) &"/1/1" 'year(slutdato) &"/1/1"

            if cint(visning) = 6 OR cint(visning) = 7 OR visning = 77 OR visning = 14  then
            sygaarSt = startdato
            end if


            '*** Nulstil ikke saldo på disse typer, slevom år afsluttes
             select case lto
            case "tec", "esn"
                if cint(visning) = 5 then 'GrandTotal
                strDatoKriViderefor = lisStDatoSQL
                else
                 strDatoKriViderefor = startdato
                end if
            case else
            strDatoKriViderefor = startdato
            end select
	        
	        
	        'Response.Write "<hr>" & aktiveTyper & "<hr>"
	        
	        alleaktivetyperSQL = aktiveTyper
	       
	       '** fjerner sumtyper **'
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-1#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-2#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-3#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-4#", "")
	       alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-5#", "")
	       
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-10#", "")
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #-20#", "")
	       
	       '** Trimer og sætter dato interval på hver enekelt type. Ferie skal altid være i ferie år **' 
	        alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #11#", " OR (af.tfaktim = 11 AND af.tdato BETWEEN '"& startdatoSQLNow &"' AND '"& slutdatoFeriePl &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #12#", " OR (af.tfaktim = 12 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #13#", " OR (af.tfaktim = 13 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #14#", " OR (af.tfaktim = 14 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #15#", " OR (af.tfaktim = 15 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #16#", " OR (af.tfaktim = 16 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #17#", " OR (af.tfaktim = 17 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #18#", " OR (af.tfaktim = 18 AND af.tdato BETWEEN '"& startdatoFerieFri &"' AND '"& slutdatoFerieFri &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #19#", " OR (af.tfaktim = 19 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
    	   
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #111#", " OR (af.tfaktim = 111 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #112#", " OR (af.tfaktim = 112 AND af.tdato BETWEEN '"& startdatoFerie &"' AND '"& slutdatoFerie &"')")

    	    '** Sætter periode på andre felter til valgt periode **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #1#", " OR (af.tfaktim = 1 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #2#", " OR (af.tfaktim = 2 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #5#", " OR (af.tfaktim = 5 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #6#", " OR (af.tfaktim = 6 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #7#", " OR (af.tfaktim = 7 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #8#", " OR (af.tfaktim = 8 AND af.tdato BETWEEN '"& strDatoKriViderefor &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #9#", " OR (af.tfaktim = 9 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #10#", " OR (af.tfaktim = 10 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    '*** Sygdom / Barn Syg / Barsel / Omsorg mm **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #20#", " OR (af.tfaktim = 20 AND af.tdato BETWEEN '"& sygaarSt &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #21#", " OR (af.tfaktim = 21 AND af.tdato BETWEEN '"& sygaarSt &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #22#", " OR (af.tfaktim = 22 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #23#", " OR (af.tfaktim = 23 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #24#", " OR (af.tfaktim = 24 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #25#", " OR (af.tfaktim = 25 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            
            '*/ Omsrog Pl **'
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #120#", " OR (af.tfaktim = 120 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #121#", " OR (af.tfaktim = 121 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #122#", " OR (af.tfaktim = 122 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")


            '** Aldersreduktion ***'
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #26#", " OR (af.tfaktim = 26 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #27#", " OR (af.tfaktim = 27 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #28#", " OR (af.tfaktim = 28 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #29#", " OR (af.tfaktim = 29 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
        
    	    
    	    '** Overarb / Afspad **'
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #30#", " OR (af.tfaktim = 30 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #31#", " OR (af.tfaktim = 31 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #32#", " OR (af.tfaktim = 32 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #33#", " OR (af.tfaktim = 33 AND af.tdato BETWEEN '"& lisStDatoSQL &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #50#", " OR (af.tfaktim = 50 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")

           
            '** / Nat Weekend / Omsorg 2 / 10
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #51#", " OR (af.tfaktim = 51 AND af.tdato BETWEEN '"& strDatoKriViderefor &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #52#", " OR (af.tfaktim = 52 AND af.tdato BETWEEN '"& strDatoKriViderefor &"' AND '"& slutdato &"')")
            '*** 
    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #53#", " OR (af.tfaktim = 53 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #54#", " OR (af.tfaktim = 54 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #55#", " OR (af.tfaktim = 55 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")

            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #113#", " OR (af.tfaktim = 113 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #114#", " OR (af.tfaktim = 114 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")

    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #60#", " OR (af.tfaktim = 60 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #61#", " OR (af.tfaktim = 61 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #81#", " OR (af.tfaktim = 81 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #90#", " OR (af.tfaktim = 90 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #91#", " OR (af.tfaktim = 91 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #92#", " OR (af.tfaktim = 92 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
           
            '*** /Tjensetefri / TEC: Omsorgsdage konverteret afspadsering afholdt
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #115#", " OR (af.tfaktim = 115 AND af.tdato BETWEEN '"& strDatoKriViderefor &"' AND '"& slutdato &"')")


            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #123#", " OR (af.tfaktim = 123 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
            alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #124#", " OR (af.tfaktim = 124 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
    	   
           alleaktivetyperSQL = replace(alleaktivetyperSQL, ", #125#", " OR (af.tfaktim = 125 AND af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"')")
      

            '****'
    	    
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99# OR ", "")
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99#,", "af.tfaktim = -99")

            
    	    alleaktivetyperSQL = replace(alleaktivetyperSQL, ",", "")
    	    
    	    'alleaktivetyperSQL = replace(alleaktivetyperSQL, "#-99#", "")
	        'alleaktivetyperSQL_len = len(alleaktivetyperSQL)
	        'alleaktivetyperSQL = left(alleaktivetyperSQL, alleaktivetyperSQL_len - 13) 
    	   
    	    
	        aktiveTyper = ""
	        'end if
	    
	    'end if
	    
	    
	    if trim(alleaktivetyperSQL) <> "#-99#" AND trim(alleaktivetyperSQL) <> "#-99#," then
	    alleaktivetyperSQL = alleaktivetyperSQL
	    else
	    alleaktivetyperSQL = "af.tfaktim = -99"
	    end if
	    
	    'Response.Write "<hr>" & alleaktivetyperSQL & "<hr>"
	    'Response.flush
	    
	    '**** Timefordeling på akt.typer *****'
	    strSQLtim = "SELECT tfaktim, sum(af.timer) AS sumtimer, "_
	    &" sum(af.timer * a.faktor) AS sumenheder, af.tdato,  "_
	    &" a.faktor FROM timer af "_
	    &" LEFT JOIN aktiviteter a ON (a.id = af.taktivitetid) "_
	    &" WHERE af.tmnr = "& intMid &" AND ("& alleaktivetyperSQL &") "_
	    &" GROUP BY af.tfaktim, af.tmnr, af.tdato ORDER BY tdato DESC"
	    

        'Response.write strSQLtim & "<br><br>"
        'Response.flush
	    
	    'AND af.tdato BETWEEN '"& startdatoSQL &"' AND '"& slutdato &"'
	    'if visning <> 1 then
	    'alleaktivetyperSQL = ""
	    'akttype_sel = ""
	    'end if
	    
	    'af.tfaktim = "& akttype_sel_arr(t) &"
	    'if session("mid") = 1 then
        'Response.Write "<br>Visning: "& visning &"<br>"& strSQLtim &"<hr>"
	    'Response.flush
	    'end if

        'if session("mid") = 1 then
	    'Response.Write "x:"& meAnsatDato &" "& x &" visning:" & visning &" sql: <br>"& strSQLtim & "<br><br>"
	    'Response.flush
        'end if
	   
	    oRec6.open strSQLtim, oConn, 3
	    while not oRec6.EOF
	        
	        
	     select case oRec6("tfaktim") 'akttype_sel_arr(t)
	     case 1
	     fTimer(x) = fTimer(x) + oRec6("sumtimer")
	     case 2
	     ifTimer(x) = ifTimer(x) + oRec6("sumtimer")
	     case 5
	     km(x) = km(x) + oRec6("sumtimer")
	     case 6
	    

                    select case lto 
                    case "esn"
                     sTimer(x) = sTimer(x) + oRec6("sumenheder")
	                case else
                     sTimer(x) = sTimer(x) + oRec6("sumtimer")
                    end select

	     case 7
	     flexTimer(x) = flexTimer(x) + oRec6("sumtimer")
	     case 8
	     sundTimer(x) = sundTimer(x) + oRec6("sumtimer")
	     case 9
	     pausTimer(x) = pausTimer(x) + oRec6("sumtimer")
	     case 10
	     fpTimer(x) = fpTimer(x) + oRec6("sumtimer")
	     
         case 11    '**** Ferie pl ***'
	     
	     call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

          feriePLTimer(x) = feriePLTimer(x) + oRec6("sumtimer") / ntimPerUse         
	               
	                
	     case 12
	     
         '*** nomrtid skal være staddanrd uge uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6, 0)
	     
         if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

          fefriTimer(x) =  fefriTimer(x) + oRec6("sumtimer") / normTimerGns5

	     case 13
	    
         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

          fefriTimerBr(x) =  fefriTimerBr(x) + oRec6("sumtimer") / ntimPerUse
	                
                    'Response.write "per13wrt: "& per13wrt

                    if cint(per13wrt) = 0 then 

	                      ''*** FerieFridage afholdt i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 13 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                    
                        'if session("mid") = 1 then
	                    'Response.Write "####"& strSQLper
	                    'Response.flush
                        'end if
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        fefriTimerPerBr(x) = fefriTimerPerBr(x) + oRec5("sumtimerPer") / ntimPerUse
                        fefriTimerPerBrTimer(x) = fefriTimerPerBrTimer(x) + (oRec5("sumtimerPer")/1)

                        '*** Første dage i peridode ***'
                        feriefriAFPerstDato(x) = oRec5("tdato")

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per13wrt = 1

                    end if
	     
	     case 14

         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieAFTimer(x) = ferieAFTimer(x) + (oRec6("sumtimer") / ntimPerUse)
       
         
         'Response.Write oRec6("sumtimer") & " " & oRec6("tdato") & " " & ntimPerUse & "<br>"
        


                    if cint(per14wrt) = 0 then 

	                     ''*** Ferie afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 14 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieAFPerTimer(x) = ferieAFPerTimer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieAFPerTimertimer(x) = ferieAFPerTimertimer(x) + (oRec5("sumtimerPer") / 1)

                        '*** Første dage i peridode ***'
                        ferieAFPerstDato(x) = oRec5("tdato") 
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per14wrt = 1

                    end if
	                
	     
	     
	     case 15
	     
         '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6, 0)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjtimer(x) = ferieOptjtimer(x) + oRec6("sumtimer") / normTimerGns5


         'Response.write "<br>"
         'Response.write " ferieOptjtimer(x)" & ferieOptjtimer(x)

         case 111

         '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6, 0)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjOverforttimer(x) = ferieOptjOverforttimer(x) + oRec6("sumtimer") / normTimerGns5
         
         
         case 112

          '*** nomrtid skal være staddanrd uge, baseret på 5 arb. dage uanset helleidag mm. **'
         call normtimerPer(intMid, oRec6("tdato"), 6, 0)
	     if ntimPer <> 0 then
         'ntimPerUse = ntimPer/antalDageMtimer
         normTimerGns5 = (ntimManIgnHellig + ntimTirIgnHellig + ntimOnsIgnHellig + ntimTorIgnHellig + ntimFreIgnHellig + ntimLorIgnHellig + ntimSonIgnHellig)  / 5
         else
         normTimerGns5 = 1
         end if 

         ferieOptjUlontimer(x) = ferieOptjOverforttimer(x) + oRec6("sumtimer") / normTimerGns5

         
         
         case 16
	   
        
         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieUdbTimer(x) = ferieUdbTimer(x) + oRec6("sumtimer") / ntimPerUse

                    
                       if cint(per16wrt) = 0 then 

	                     ''*** Ferie udbetalt afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 16 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieUdbPer(x) = ferieUdbPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieUdbPerTimer(x) = ferieUdbPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per16wrt = 1

                    end if
                     

	     case 17
	    
         
         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         fefriTimerUdb(x) = fefriTimerUdb(x) + oRec6("sumtimer") / ntimPerUse


                    
                       if cint(per17wrt) = 0 then 

	                     ''*** FerieFridage udbetalt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 17 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        fefriTimerUdbPer(x) = fefriTimerUdbPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        fefriTimerUdbPerTimer(x) = fefriTimerUdbPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per17wrt = 1

                    end if



	     
	     case 18 '**** Ferie frdage pl ****'
	     

         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         fefriplTimer(x) = fefriplTimer(x) + oRec6("sumtimer") / ntimPerUse

	     case 19           
	     
         
          call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         ferieAFulonTimer(x) = ferieAFulonTimer(x) + (oRec6("sumtimer") / ntimPerUse)

            
            
                        
                        if cint(per19wrt) = 0 then 

	                     ''*** Ferie U løn afholdt i per **'
	                     strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato FROM timer af WHERE af.tmnr = "& intMid &" AND af.tfaktim = 19 AND "_
	                     &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    'Response.Write strSQLper
	                    'Response.flush
	                   
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        ferieAFulonPer(x) = ferieAFulonPer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        ferieAFulonPerTimer(x) = ferieAFulonPerTimer(x) + (oRec5("sumtimerPer") / 1)

                      
                        
                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per19wrt = 1

                        end if
              
	     case 20
	    

         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         sygTimer(x) = sygTimer(x) + oRec6("sumtimer")
	     sygDage(x) = sygDage(x) + (oRec6("sumtimer") / ntimPerUse)           
	                
	                
	                

                     if cint(per20wrt) = 0 then 

	                     ''*** Syg. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                    &" FROM timer af "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 20 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                    

                        'if session("mid") = 1 then
                        'Response.write strSQLper & "<br><br>"
                        'end if

	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        sygTimerPer(x) = sygTimerPer(x) + oRec5("sumtimerPer")
                        sygDagePer(x) = sygDagePer(x) + (oRec5("sumtimerPer") / ntimPerUse)
                        sygPerstDato(x) = oRec5("tdato")
                        
                      

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per20wrt = 1

                    end if
	     
	     case 21
	     
         
         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         BarnsygTimer(x) = BarnsygTimer(x) + oRec6("sumtimer")
	     barnSygDage(x) = barnSygDage(x) + (oRec6("sumtimer") / ntimPerUse)
	                
	                
	               

                     if cint(per21wrt) = 0 then 

	                     ''*** Barnsyg. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                    &" FROM timer af "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 21 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                        barnSygTimerPer(x) = barnSygTimerPer(x) + oRec5("sumtimerPer") 
                        barnSygDagePer(x) = barnSygDagePer(x) + (oRec5("sumtimerPer") / ntimPerUse)

                         
                        barnsygPerstDato(x) = oRec5("tdato")

                        oRec5.movenext
	                    wend
	                    oRec5.close

                         per21wrt = 1

                    end if

	     
         case 22
         
	     
         
         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         barsel(x) = barsel(x) + (oRec6("sumtimer") / ntimPerUse) 
         barselPerstDato(x) = oRec6("tdato")
     


         case 23 'omsorg(x)
         
          call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

         omsorg(x) = omsorg(x) + (oRec6("sumtimer") / ntimPerUse) 
         omsorgPerstDato(x) = oRec6("tdato")
       
         
         case 24 'senior(x)

     
         senior(x) = senior(x) + oRec6("sumtimer")

         case 25 'divfritimer(x) 

         
         
	     call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 
     
         divfritimer(x) = divfritimer(x) + oRec6("sumtimer")

        case 26
        aldersreduktionPl(x) = aldersreduktionPl(x) + oRec6("sumtimer") 

        'response.write "aldersreduktionPl(x):" & aldersreduktionPl(x) & "<br>"

        case 27
        aldersreduktionOpj(x) = aldersreduktionOpj(x) + oRec6("sumtimer")
        case 28
        aldersreduktionBr(x) = aldersreduktionBr(x) + oRec6("sumtimer")
        case 29
        aldersreduktionUdb(x) = aldersreduktionUdb(x) + oRec6("sumtimer")


	     case 30
	     afspTimer(x) = afspTimer(x) + oRec6("sumenheder")
	     afspTimerTim(x) = afspTimerTim(x) + oRec6("sumtimer")
	                
	                
	               
	                

                    if cint(per30wrt) = 0 then 

	                     ''*** Overarb. i per **'
	                    strSQLper = "SELECT SUM(timer) AS sumTimerPer, sum(af.timer * a.faktor) AS sumenhederPer, tdato "_
	                    &" FROM timer af "_
	                    &" LEFT JOIN aktiviteter a ON (a.id = af.taktivitetid) "_
	                    &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 30 AND "_
	                    &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato"
	                
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                       
                         afspTimerPer(x) = afspTimerPer(x) + oRec5("sumenhederPer")
	                    afspTimerTimPer(x) = afspTimerTimPer(x) + oRec5("sumTimerPer")

	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per30wrt = 1

                    end if

	     
	     case 31
	     afspTimerBr(x) = afspTimerBr(x) + oRec6("sumtimer")
	     
	                

                    if cint(per31wrt) = 0 then 

	                     ''*** Afspadsering i per **'
	                strSQLper = "SELECT SUM(timer) AS sumTimerPer, tdato "_
	                &" FROM timer af "_
	                &" WHERE af.tmnr = "& intMid &" AND af.tfaktim = 31 AND "_
	                &" af.tdato BETWEEN '"& startdato &"' AND '"& slutdato &"' GROUP BY af.tfaktim, af.tdato ORDER BY tdato DESC"
	                
	                
	                    oRec5.open strSQLper, oConn, 3
	                    while not oRec5.EOF 

                         call normtimerPer(intMid, oRec5("tdato"), 0, 0)
	                     if ntimPer <> 0 then
                         ntimPerUse = ntimPer
                         else
                         ntimPerUse = 1
                         end if 

                       
                        afspTimerBrPer(x) = afspTimerBrPer(x) + oRec5("sumTimerPer")
                        afspadPerstDato(x) = oRec5("tdato")

	                    
	                    oRec5.movenext
	                    wend
	                    oRec5.close

                         per31wrt = 1

                    end if

	     
	     case 32
	     afspTimerUdb(x) = afspTimerUdb(x) + oRec6("sumtimer")
	     case 33
	     afspTimerOUdb(x) =  afspTimerOUdb(x) + oRec6("sumtimer")
	     case 50
	     dagTimer(x) = dagTimer(x) + oRec6("sumtimer")
	     case 51
	     natTimer(x) = natTimer(x) + oRec6("sumtimer")
	     case 52
	     weekendTimer(x) = weekendTimer(x) + oRec6("sumtimer")
	     case 53
	     weekendnattimer(x) = weekendnattimer(x) + oRec6("sumtimer")
	     case 54
	     aftenTimer(x) = aftenTimer(x) + oRec6("sumtimer")
	     case 55
	     aftenweekendTimer(x) = aftenweekendTimer(x) + oRec6("sumtimer")
	     case 60
                    select case lto 
                    case "esn"
                    adhocTimer(x) = adhocTimer(x) + oRec6("sumenheder")
	                case else
                    adhocTimer(x) = adhocTimer(x) + oRec6("sumtimer")
                    end select
	     case 61
	     stkAntal(x) = stkAntal(x) + oRec6("sumtimer")
	     case 81
	     lageTimer(x) = lageTimer(x) + oRec6("sumtimer")
	     case 90
	     e1Timer(x) = e1Timer(x) + oRec6("sumtimer")
	     case 91
	     e2Timer(x) = e2Timer(x) + oRec6("sumtimer")
         case 92
	     e3Timer(x) = e3Timer(x) + oRec6("sumtimer")

         case 113
         korrektionKomG(x) = korrektionKomG(x) + oRec6("sumtimer")
         case 114
         korrektionReal(x) = korrektionReal(x) + oRec6("sumtimer")
         case 115
         tjenestefri(x) =  tjenestefri(x) + oRec6("sumtimer")
        
        case 120
        omsorg2pl(x) = omsorg2pl(x) + oRec6("sumtimer")
        case 121
        omsorg10pl(x) = omsorg10pl(x) + oRec6("sumtimer")
        case 122
        omsorgKpl(x) = omsorgKpl(x) + oRec6("sumtimer")

        case 123
        ulempe1706udb(x) = ulempe1706udb(x) + oRec6("sumenheder")
        case 124
        ulempeWudb(x) = ulempeWudb(x) +  oRec6("sumenheder")
        case 125

         call normtimerPer(intMid, oRec6("tdato"), 0, 0)
	     if ntimPer <> 0 then
         ntimPerUse = ntimPer
         else
         ntimPerUse = 1
         end if 

        

        rejseDage(x) = rejseDage(x) +  (oRec6("sumtimer") / ntimPerUse)

        end select
	    
	    
	    
	    oRec6.movenext
	    wend
	    oRec6.close
	    
	    if media <> "export" then
	    Response.flush
        end if
	    

	
	         '* Afspadsering ***'
	         
	        if afspTimer(x) <> 0 then
	         afspTimer(x) = afspTimer(x)
	         afstemnul(x) = 235
	         else
	         afspTimer(x) = 0
	         end if
	         
	         if afspTimerTim(x) <> 0 then
	         afspTimerTim(x) = afspTimerTim(x)
	         afstemnul(x) = 234
	         else
	         afspTimerTim(x) = 0
	         end if
	        
	        
	        
	    
	    
	    '** Afspad. Brugt **'
	   
	    
	         if afspTimerBr(x) <> 0 then
	         afspTimerBr(x) = afspTimerBr(x)
	         afstemnul(x) = 233
	         else
	         afspTimerBr(x) = 0
	         end if
	    
	    
	    '*** Afspad Udbetalt **'
	        if afspTimerUdb(x) <> 0 then
	         afspTimerUdb(x) = afspTimerUdb(x)
	         afstemnul(x) = 232
	         else
	         afspTimerUdb(x) = 0
	         end if
	         
	         '*** Afspad Ønskes Udbetalt **'
	         if afspTimerOUdb(x) <> 0 then
	         afspTimerOUdb(x) = afspTimerOUdb(x)
	         afstemnul(x) = 231
	         else
	         afspTimerOUdb(x) = 0
	         end if
	         
	         
	       
	  
	    
	    '** Ferie fridage ***'
	    
	    '*** Optj **'
	     if fefriTimer(x) <> 0 then
	     fefriTimer(x) = fefriTimer(x)
	     afstemnul(x) = 230
	     else
	     fefriTimer(x) = 0
	     end if
	     



	     '*** PL ***'
	     if fefriplTimer(x) <> 0 then
	      fefriplTimer(x) =  fefriplTimer(x)
	      afstemnul(x) = 229
	     else
	      fefriplTimer(x) = 0
	     end if
	     
	    
	    
	    
	    '** Brugt **'
	    
	    if fefriTimerBr(x) <> 0 then
	     fefriTimerBr(x) = fefriTimerBr(x)
	     afstemnul(x) = 228
	     else
	    fefriTimerBr(x) = 0
	     end if
	    
	    
	    '** Udbetalt **'
	     if fefriTimerUdb(x) <> 0 then
	     fefriTimerUdb(x) = fefriTimerUdb(x)
	     afstemnul(x) = 227
	     else
	    fefriTimerUdb(x) = 0
	     end if
	    

         '** Udbetalt Periode **'
	     if fefriTimerUdbPer(x) <> 0 then
	     fefriTimerUdbPer(x) = fefriTimerUdbPer(x)
	     afstemnul(x) = 227
	     else
	    fefriTimerUdbPer(x) = 0
	     end if
	    
	   
	    
	    '** Ferie Planlagt **'
	    if feriePLTimer(x) <> 0 then
	    feriePLTimer(x) = feriePLTimer(x)
	    afstemnul(x) = 226
	    else
	    feriePLTimer(x) = 0
	    end if
	    
	    
	    
	    '** Ferie Afholdt **'
	    if ferieAFTimer(x) <> 0 then
	    ferieAFTimer(x) = ferieAFTimer(x)
	    afstemnul(x) = 225
	    else
	    ferieAFTimer(x) = 0
	    end if
	   
	     '** Ferie Afholdt ulon **'
	    if ferieAFulonTimer(x) <> 0 then
	    ferieAFulonTimer(x) = ferieAFulonTimer(x)
	    afstemnul(x) = 224
	    else
	    ferieAFulonTimer(x) = 0
	    end if
	  
	   
	    
	    '** Ferie Optjent **'
	    if  ferieOptjtimer(x) <> 0 then
	    ferieOptjtimer(x) =  ferieOptjtimer(x)
	    afstemnul(x) = 223
	    else
	    ferieOptjtimer(x) = 0
	    end if
	    
        '*** Ferie optjent overført **'
        if ferieOptjOverforttimer(x) <> 0 then
	    ferieOptjOverforttimer(x) = ferieOptjOverforttimer(x)
	    afstemnul(x) = 240
	    else
	    ferieOptjOverforttimer(x) = 0
	    end if

        '** SENESTE afstemnul(x) 241 2012.12.21 
        if ferieOptjUlontimer(x) <> 0 then
	    ferieOptjUlontimer(x) = ferieOptjUlontimer(x)
	    afstemnul(x) = 241
	    else
	    ferieOptjUlontimer(x) = 0
	    end if
         


	    
	    '** Ferie  Udb **'
	    if ferieUdbTimer(x) <> 0 then
	    ferieUdbTimer(x) = ferieUdbTimer(x)
	    afstemnul(x) = 222
	    else
	    ferieUdbTimer(x) = 0
	    end if
	    
	    
	      '*** Fakturerbaretimer #1 **'
	    if fTimer(x) <> 0 then
	    fTimer(x) = fTimer(x)
	    afstemnul(x) = 221
	    else
	    fTimer(x) = 0
	    end if
	      
	     
	     
	    '*** Ikke Fakturerbare timer #2 ***'
	    if ifTimer(x) <> 0 then
        ifTimer(x) = ifTimer(x)
        afstemnul(x) = 220
        else
        ifTimer(x) = 0
        end if
	    
	    '*** Salg & Newbizz timer***'
	    if sTimer(x) <> 0 then
        sTimer(x) = sTimer(x)
        afstemnul(x) = 219
        else
        sTimer(x) = 0
        end if
	    
	    
	    
	        
	    '*** Frokost ***'
	    if fpTimer(x) <> 0 then
	    fpTimer(x) = fpTimer(x)
	    afstemnul(x) = 218
	    else
	    fpTimer(x) = 0
	    end if
	       
	    
	    
	         
	    '*** Syg ****'
	    if sygTimer(x) <> 0 then
         sygTimer(x) = sygTimer(x)
         afstemnul(x) = 217
         else
         sygTimer(x) = 0
         end if
	   
         
          if sygDage(x) <> 0 then
         sygDage(x) = sygDage(x)
        else
         sygDage(x) = 0
         end if

          if sygDagePer(x) <> 0 then
         sygDagePer(x) = sygDagePer(x)
        else
         sygDagePer(x) = 0
         end if


         if barsel(x) <> 0 then
         barsel(x) = barsel(x)
         afstemnul(x) = 236
         else
         barsel(x) = 0
         end if

          if omsorg(x) <> 0 then
         omsorg(x) = omsorg(x)
         afstemnul(x) = 237
         else
         omsorg(x) = 0
         end if

          if senior(x) <> 0 then
         senior(x) = senior(x)
         afstemnul(x) = 238
         else
         senior(x) = 0
         end if

          if divfritimer(x) <> 0 then
         divfritimer(x) = divfritimer(x)
         afstemnul(x) = 239
         else
         divfritimer(x) = 0
         end if


	   
	   '*** Barn syg ***'
	    if BarnsygTimer(x) <> 0 then
         BarnsygTimer(x) = BarnsygTimer(x)
         afstemnul(x) = 216
         else
         BarnsygTimer(x) = 0
         end if
	    
	    if barnSygDage(x) <> 0 then
         barnSygDage(x) = barnSygDage(x)
        else
         barnSygDage(x) = 0
         end if

          if barnSygDagePer(x) <> 0 then
         barnSygDagePer(x) = barnSygDagePer(x)
        else
         barnSygDagePer(x) = 0
         end if
	    
	   
	    
	    
	    
	    '*** KM ***'
	    if km(x) <> 0 then
         km(x) = km(x)
         afstemnul(x) = 215
         else
         km(x) = 0
         end if
	         
	         
	      '*** sundheds timer ***'
	    if sundTimer(x) <> 0 then
	    sundTimer(x) = sundTimer(x)
	    afstemnul(x) = 214
	    else
	    sundTimer(x) = 0
	    end if
	    
	   
	    
	   
	    '*** Flex timer ***'
	    if flexTimer(x) <> 0 then
        flexTimer(x) = flexTimer(x)
        afstemnul(x) = 213
        else
        flexTimer(x) = 0
        end if

	  
	    
	   
	    
	    
	    '** Pause timer **'
	    if pausTimer(x) <> 0 then
	    pausTimer(x) = pausTimer(x)
	    afstemnul(x) = 212
	    else
	    pausTimer(x) = 0
	    end if
	    
	    
	     '*** Dag ***'
	   if dagTimer(x) <> 0 then
         dagTimer(x) = dagTimer(x)
         afstemnul(x) = 211
         else
         dagTimer(x) = 0
         end if
	    
	   '*** Nat ***'
	   if natTimer(x) <> 0 then
         natTimer(x) = natTimer(x)
         afstemnul(x) = 210
         else
         natTimer(x) = 0
         end if
	    
	    
	  
	    '*** Weekend ***'
	   if weekendTimer(x) <> 0 then
	         weekendTimer(x) = weekendTimer(x)
	         afstemnul(x) = 209
	         else
	         weekendTimer(x) = 0
	         end if
	    
	    
	    '*** Weekend Nat ***'
	   if weekendnattimer(x) <> 0 then
	         weekendnattimer(x) = weekendnattimer(x)
	         afstemnul(x) = 208
	         else
	         weekendnattimer(x) = 0
	         end if
	    
	    
	   
	 '*** Aften ***'
	 if aftenTimer(x) <> 0 then
     aftenTimer(x) = aftenTimer(x)
     afstemnul(x) = 207
     else
     aftenTimer(x) = 0
     end if
	    
	    
	    
	   '*** Aften Weekend ***'
	    if aftenweekendTimer(x) <> 0 then
         aftenweekendTimer(x) = aftenweekendTimer(x)
         afstemnul(x) = 206
         else
         aftenweekendTimer(x) = 0
         end if
	    
	    
	   
	    
	    
	    
	    '*** Adhoc ***'
	   
	         if adhocTimer(x) <> 0 then
	         adhocTimer(x) = adhocTimer(x)
	         afstemnul(x) = 205
	         else
	         adhocTimer(x) = 0
	         end if
	    
	  
	    
	    
	    '*** StkAntal ***'
	   if stkAntal(x) <> 0 then
	         stkAntal(x) = stkAntal(x)
	         afstemnul(x) = 204
	         else
	         stkAntal(x) = 0
	         end if
	    
	  
	     
	    
	    
	     '*** Læge Timer ***'
	    if lageTimer(x) <> 0 then
	     lageTimer(x) = lageTimer(x)
	     afstemnul(x) = 203
         else
         lageTimer(x) = 0
         end if
         
         
         '*** E1 Timer ***'
	     if e1Timer(x) <> 0 then
	     e1Timer(x) = e1Timer(x)
	     afstemnul(x) = 202
         else
         e1Timer(x) = 0
         end if
         
         '*** E2 Timer ***'
	     if e2Timer(x) <> 0 then
	     e2Timer(x) = e2Timer(x)
	     afstemnul(x) = 201
         else
         e2Timer(x) = 0
         end if

        
	     

         if korrektionKomG(x) <> 0 then
	     korrektionKomG(x) = korrektionKomG(x)
	     afstemnul(x) = 242
         else
         korrektionKomG(x) = 0
         end if
	     

         
         if korrektionReal(x) <> 0 then
	     korrektionReal(x) = korrektionReal(x)
	     afstemnul(x) = 243
         else
         korrektionReal(x) = 0
         end if

         if tjenestefri(x) <> 0 then
	     tjenestefri(x) = tjenestefri(x)
	     afstemnul(x) = 244
         else
         tjenestefri(x) = 0
         end if


    
         if aldersreduktionOpj(x) <> 0 then
	     aldersreduktionOpj(x) = aldersreduktionOpj(x)
	     afstemnul(x) = 245
         else
         aldersreduktionOpj(x) = 0
         end if


    
         if  aldersreduktionBr(x) <> 0 then
	     aldersreduktionBr(x) = aldersreduktionBr(x)
	     afstemnul(x) = 246
         else
         aldersreduktionBr(x) = 0
         end if


    
         if aldersreduktionUdb(x) <> 0 then
	     aldersreduktionUdb(x) = aldersreduktionUdb(x)
	     afstemnul(x) = 247
         else
         aldersreduktionUdb(x) = 0
         end if

        
    
         if omsorg2pl(x) <> 0 then
	     omsorg2pl(x) = omsorg2pl(x)
	     afstemnul(x) = 248
         else
         omsorg2pl(x) = 0
         end if

         if omsorg10pl(x) <> 0 then
	     omsorg10pl(x) = omsorg10pl(x)
	     afstemnul(x) = 249
         else
         omsorg10pl(x) = 0
         end if

         if omsorgKpl(x) <> 0 then
	     omsorgKpl(x) = omsorgKpl(x)
	     afstemnul(x) = 250
         else
         omsorgKpl(x) = 0
         end if

         if aldersreduktionPl(x) <> 0 then
	     aldersreduktionPl(x) = aldersreduktionPl(x)
	     afstemnul(x) = 251
         else
         aldersreduktionPl(x) = 0
         end if


         if ulempe1706udb(x) <> 0 then
	     ulempe1706udb(x) = ulempe1706udb(x)
	     afstemnul(x) = 252
         else
         ulempe1706udb(x) = 0
         end if

          if ulempeWudb(x) <> 0 then
	     ulempeWudb(x) = ulempeWudb(x)
	     afstemnul(x) = 253
         else
         ulempeWudb(x) = 0
         end if

       
         '*** E3 Timer ***'
	     if e3Timer(x) <> 0 then
	     e3Timer(x) = e3Timer(x)
	     afstemnul(x) = 254
         else
         e3Timer(x) = 0
         end if
        

    
         if rejseDage(x) <> 0 then
	     rejseDage(x) = rejseDage(x)
	     afstemnul(x) = 255
         else
         rejseDage(x) = 0
         end if



end function   
%>