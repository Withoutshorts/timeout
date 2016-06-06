<html>

<body>

<%



'Definer variabler

dim fs, fileIndput, fileOutput, IndputFileName, OutputFileName, datoFieldNo



'Definer indput og output filer

IndputFileName = "../inc/log/data/"& file  '"medarbafexp_2011_10_26__13_4_16_fk.csv"

OutputFileName = Replace(IndputFileName,".csv","_silkeborg.csv")



'To cifres kode for kommune

strInstkode = "FB" 'FB_MFKA



'Tjenstekode værdier 0, 1, 2, 3

strTjenstekode = "0"



'Åben .csv fil til indlæsning

set fs=Server.CreateObject("Scripting.FileSystemObject")

set fileIndput=fs.OpenTextFile(Server.MapPath(IndputFileName),1,true)

set fileOutput=fs.OpenTextFile(Server.MapPath(OutputFileName),8,true)




'Læs første linie da det er header

fileIndput.ReadLine



'Gennemløb hver linje i .csv filen

Do Until fileIndput.AtEndOfStream

    sSeg = Split( fileIndput.ReadLine , ";" )



    'CreateSDLine(csvLine, Tjenstekode, Afdeling, Konto, loenArt1, loenArt2, enheder)

    'csvLine(7) Ferie, afholdelse  :Ferie 730+130 (skal være timer istedet for dage)

    if len(trim(sSeg(11))) <> 0 AND sSeg(11) > 0 then
        
        datoFieldNo = 12
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "730", "", sSeg(11), "") '7

        fileOutput.WriteLine(sdLine)


    end if



    'csvLine(8) 6. ferieuge, afholdelse : Ferie afholdt 752+130 (skal være timer istedet for dage)

     if len(trim(sSeg(6))) <> 0 AND sSeg(6) > 0 then
        
        datoFieldNo = 7
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "752", "", sSeg(6), "") '8 

        fileOutput.WriteLine(sdLine)

    end if



    'csvLine(4) Seniordage, afholdelse : 774+130 !!!!SKAL VÆRE TIMER ISTEDET FOR DAGE!!!!
    '28
    if len(trim(sSeg(29))) <> 0 AND sSeg(29) > 0 then

        datoFieldNo = 31 '30
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "774", "", sSeg(29), "") '5
        
        fileOutput.WriteLine(sdLine)

    end if

    

    'csvLine(3) Omsorgsdage, afholdelse : Afveksling af flextimer  650   !!!!HVILKEN KOLONNE HAR ANTAL TIMER!!!!
    '26
    if len(trim(sSeg(27))) <> 0 AND sSeg(27) > 0 then
        
        datoFieldNo = 28 '27
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "520", "", sSeg(27), "") '130

        fileOutput.WriteLine(sdLine)

    end if


     'Afspadsering   

    'if len(trim(sSeg(16))) <> 0 AND sSeg(16) > 0 then
        
    '    datoFieldNo = 17
    '    sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "650", "", sSeg(16)) '3

    '    fileOutput.WriteLine(sdLine)

    'end if
    

    

    'csvLine(13) 1. maj timer, afholdelse : Feriefridage 730+130    

    if len(trim(sSeg(13))) <> 0 AND sSeg(13) > 0 then

        datoFieldNo = 31
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "650", "", sSeg(13), "") '13
    
        fileOutput.WriteLine(sdLine)

    end if

    

    

    'csvLine(14) Sygdom    :  Sygdom 790+130 !!!!SKAL VÆRE TIMER ISTEDET FOR DAGE!!!!
    '15
    if len(trim(sSeg(16))) <> 0 AND sSeg(16) > 0 then

        datoFieldNo = 18 '17
        
        'afrunder til dage
        'sSeg(17) = formatnumber(""&sSeg(17)&"", 0) 
        sygdomHeleDage = sSeg(17) 
    
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode , "790", "130", sygdomHeleDage, sSeg(16))  

        fileOutput.WriteLine(sdLine)

    end if

    

    

    'csvLine(18) Barn syg  : Barn syg 910+710
    '20
    if len(trim(sSeg(21))) <> 0 AND sSeg(21) > 0 then
        
        datoFieldNo = 23 '22
        
        barnsygHeleDage = sSeg(22)
       
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode, "910", "710", barnsygHeleDage, sSeg(21)) 

        fileOutput.WriteLine(sdLine)

    end if

    

    

    'csvLine(23) Barsel : Barsel lønart 770+130 !!!!SKAL VÆRE TIMER ISTEDET FOR DAGE!!!!
    '** ER SAT til IKKE FAKTURERBAR i TO så den ikke kommer med ud. PGA VISMA ikke kan håndtere barsel for mænd.
    '** Er med i gang pr. 5 nov. 2013 da Visma got ka nhåndete barsel
    '24
    if len(trim(sSeg(25))) <> 0 AND sSeg(25) > 0 then

        datoFieldNo = 26 '25
        sdLine = CreateSDLine(sSeg, strInstkode, strTjenstekode  , "770", "", sSeg(25), "") 

        fileOutput.WriteLine(sdLine)

    end if

Loop



fileIndput.Close

fileOutput.Close



set fileIndput=Nothing

set fileOutput=Nothing



set fs=Nothing



Response.Write "Script er udført"



Function CreateSDLine(csvLine, Instkode, Tjenstekode, loenArt1, loenArt2, enheder, enheder2)



    if InStr(enheder,",") > 0 then

         enheder = Replace(enheder,",","")
        
    
    end if

    if (loenArt1 = "790") then


        ''** Syg altid ,00 i dage ***'
        
        enhederDecimaler_len = len(enheder)
        enhederDecimaler_left = left(enheder, enhederDecimaler_len-2)
        enheder = enhederDecimaler_left & "00"
       
    end if


     if (loenArt1 = "910") then 


        ''** Barnsyg altid min 1,00 i dage ***'
        
        enhederDecimaler_len = len(enheder)
        enhederDecimaler_left = left(enheder, enhederDecimaler_len-2)

        if cint(enhederDecimaler_left) < 1 then
        enheder = "100"
        else
        enheder = enhederDecimaler_left & "00"
        end if

        'Response.write "LA:" & loenArt1 & " enh: "& enheder & " ..."
    'Response.flush 
    'strEnheder = "" & Lpad(enheder,"0",5) '4

    'Response.write "strEnheder:" & strEnheder &"<br><br>"
    'Response.flush 
       
    end if


    strEnheder = "" & Lpad(enheder,"0",5) '4




    if InStr(enheder,"-") > 0  then

        strChar = ""



        'Formattering af enheder til håndtering af minus jft sd løn format

        if StrComp(Right(strEnheder,1),"0") = 0 then strEnheder = left(strEnheder,4) & "å" end if

        if StrComp(right(strEnheder,1),"1") = 0 then strEnheder = left(strEnheder,4) & "J" end if

        if StrComp(right(strEnheder,1),"2") = 0 then strEnheder = left(strEnheder,4) & "K" end if

        if StrComp(right(strEnheder,1),"3") = 0 then strEnheder = left(strEnheder,4) & "L" end if

        if StrComp(right(strEnheder,1),"4") = 0 then strEnheder = left(strEnheder,4) & "M" end if

        if StrComp(right(strEnheder,1),"5") = 0 then strEnheder = left(strEnheder,4) & "N" end if

        if StrComp(right(strEnheder,1),"6") = 0 then strEnheder = left(strEnheder,4) & "O" end if

        if StrComp(right(strEnheder,1),"7") = 0 then strEnheder = left(strEnheder,4) & "P" end if

        if StrComp(right(strEnheder,1),"8") = 0 then strEnheder = left(strEnheder,4) & "Q" end if

        if StrComp(right(strEnheder,1),"9") = 0 then strEnheder = left(strEnheder,4) & "R" end if

 

    end if


    
    
    '**** Enheder 2 ******
    if InStr(enheder2,",") > 0 then

        enheder2 = Replace(enheder2,",","")

    end if

    
    'Response.write "LA:" & loenArt2 & " enh: "& enheder2 & "<br>"
    'Response.flush 
    strEnheder2 = "" & Lpad(enheder2,"0",5) '4



   



    sdLine = ""



   '001-002 Inst.kode

    sdLine =  Instkode 'Lpad(Instkode," ",2)



    '003-004 Filler

    sdLine = sdLine & "  "



    '005-005  KortArt altid på 7

    sdLine = sdLine & "7"

    

    '006 Tjenstekode værdier 0, 1, 2, 3

    sdLine = sdLine & Tjenstekode 



    '007-011 Tjenstenummer

    sdLine = sdLine & Lpad(trim(csvLine(1)),"0",5) '"0"



    '012-015 Afdeling

    sdLine = sdLine & "    "



    '016-017 År

    'Response.write csvLine(datoFieldNo) & "<br>"
    'Response.flush

    if len(trim(csvLine(datoFieldNo))) <> 0 AND datoFieldNo <> 32 then
    csvLine(datoFieldNo) = csvLine(datoFieldNo)
    else
    csvLine(datoFieldNo) = day(now) &"-"& month(now) &"-"& year(now)
    end if 

    yVal = datepart("yyyy", csvLine(datoFieldNo), 2,2)
    yVal = right(yVal, 2)
    sdLine = sdLine & yVal 
    'sdLine = sdLine & csvLine(27) &"##"


    '018-019 Måned
    
    mdVal = datepart("m", csvLine(datoFieldNo), 2,2)
    'mdVal = mid(csvLine(27),6,2)
    if len(trim(mdVal)) = 1 then ' md < 10
    mdVal = "0"&left(mdVal, 1)
    else
    mdVal = mdVal
    end if
    sdLine = sdLine & mdVal

   'sdLine = sdLine & left(Right(csvLine(25),7),2)
   
    '020-021 Uge / Blank

    sdLine = sdLine & "  "



    '022-023 dag
    dVal = datepart("d", csvLine(datoFieldNo), 2,2)
    if len(trim((dVal))) = 1 then 'dag < 10
    sdLine = sdLine & "0"&dVal
    else
    sdLine = sdLine & dVal
    end if
    
    'sdLine = sdLine & "dval:#" & dVal & "#"

    '024-026 lønart

    sdLine = sdLine & loenArt1



    '027-031 Enheder

    sdLine = sdLine & strEnheder



    '032 - 034 
    if Len(loenArt2) > 0 then

        sdLine = sdLine & loenArt2 & Lpad("","0",3)
        '035 - 038 'Sekunder lønart i timer 
        sdLine = sdLine & strEnheder2

    else
        
        sdLine = sdLine & "00000000"

    end if


    


    '096-099 Filler

    sdLine = Rpad(sdLine," ",99)



    '100-109 Konto

    sdLine =  Rpad(sdLine," ",109)



    '110-115 Filler

    sdLine =  Rpad(sdLine," ",115)



    '116 Intern behandlingskode Værdien skal være ’K’ for KMD eller ’T’ for Tabulex

    sdLine = sdLine & "T"



    '117 Særlig kode 6 = sygdom, 8 = ferie

    if StrComp(loenArt2,"790") = 0 then

        sdLine = sdLine & "6"

    end if



    if StrComp(loenArt2,"752") = 0 then

        sdLine = sdLine & "8"

    end if





    if StrComp(loenArt2,"752") <> 0 and StrComp(loenArt2,"790") <> 0  then

        sdLine = sdLine & " "

    end if



    '118-128 Filler

    sdLine = Rpad(sdLine," ",128)

   

    CreateSDLine = sdLine

End Function


'*** Bagvedstillede nuller ****
Function Rpad (sValue, sPadchar, iLength)
  
  if len(trim(sValue)) <> 0 then
  Rpad = sValue & string(iLength - Len(sValue), sPadchar)
  end if

End Function


'*** foranstillede nuller ****
Function Lpad (sValue, sPadchar, iLength)

  if len(trim(sPadchar)) <> 0 then
  sPadchar = sPadchar
  else
  sPadchar = 0
  end if

  

  if len(trim(sValue)) <> 0 then
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
  end if

 'Response.Write "iLength: " & iLength & " sPadchar: "& sPadchar & " sValue: " & sValue & " Lpad: "&  Lpad &"<br>"
 'Response.flush


End Function



%>

</body>

</html>