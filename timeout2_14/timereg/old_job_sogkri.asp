<%if sort = "9999999" then%>
		<td><a href="jobs.asp?menu=job&sort=navn&filt=<%=filt%>&shokselector=1&l=<%=usedletter%>" class='alt'>Jobnavn</a>&nbsp;(<a href="jobs.asp?menu=job&sort=nr&shokselector=1&filt=<%=filt%>&l=<%=usedletter%>" class='alt'>Jobnr.</a>)&nbsp;/ &nbsp;<a href="jobs.asp?menu=job&shokselector=1&filt=<%=filt%>&l=<%=usedletter%>" class='alt'>Kunde</a></td>
		<%end if%>

Job der begynder med:
<%if sort = "kunkunde" then%>
	&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Alle</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=a&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">A</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=b&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">B</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=c&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">C</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=d&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">D</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=e&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">E</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=f&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">F</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=g&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">G</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=h&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">H</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=i&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">I</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=j&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">J</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=k&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">K</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=l&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">L</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=m&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">M</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=n&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">N</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=o&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">O</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=p&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">P</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=q&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Q</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=r&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">R</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=s&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">S</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=t&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">T</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=u&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">U</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=v&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">V</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=w&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">W</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=x&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">X</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=y&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Y</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=z&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Z</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=æ&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Æ</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=ø&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Ø</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=å&fm_kunde=<%=request("FM_kunde")%>&shokselector=1">Å</a>&nbsp;|
	
	
	<%else%>
	&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&shokselector=1">Alle</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=a&shokselector=1">A</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=b&shokselector=1">B</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=c&shokselector=1">C</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=d&shokselector=1">D</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=e&shokselector=1">E</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=f&shokselector=1">F</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=g&shokselector=1">G</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=h&shokselector=1">H</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=i&shokselector=1">I</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=j&shokselector=1">J</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=k&shokselector=1">K</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=l&shokselector=1">L</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=m&shokselector=1">M</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=n&shokselector=1">N</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=o&shokselector=1">O</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=p&shokselector=1">P</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=q&shokselector=1">Q</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=r&shokselector=1">R</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=s&shokselector=1">S</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=t&shokselector=1">T</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=u&shokselector=1">U</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=v&shokselector=1">V</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=w&shokselector=1">W</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=x&shokselector=1">X</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=y&shokselector=1">Y</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=z&shokselector=1">Z</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=æ&shokselector=1">Æ</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=ø&shokselector=1">Ø</a>&nbsp;|
	&nbsp;&nbsp;<a href="jobs.asp?menu=job&sort=<%=sort%>&filt=<%=filt%>&l=å&shokselector=1">Å</a>&nbsp;|
	<%end if%>
