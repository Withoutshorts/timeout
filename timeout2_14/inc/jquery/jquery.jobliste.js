var pxMulti = 2;
var AddJobListDB = new Array();
var BeginDateOffset;
var EndDateOffset;
var BeginDate;
var EndDate;
var midMarkerDateArr = [];


function prekon_0(thisid) {

    
    var idtrim = thisid;


    $.post("?", "func=ajaxupdate&datatype=prekon_0&jid=" + idtrim);
    preconditions_met_Txt = "<br><span class='prekon_1' id='prekon_" + idtrim + "' style='color:#6CAE1C; font-size:10px; background:#DCF5BD;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'>Pre-konditioner opfyldt</span>"

    $("#prekon_container_" + idtrim).html(preconditions_met_Txt)

    $(".prekon_1").bind('mouseover', function () {
        $(this).css('cursor', 'pointer');
    });


    $(".prekon_1").bind('click', function () {
        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)


        prekon_1(idtrim);

    });


}


function prekon_1(thisid) {


    var idtrim = thisid;


    $.post("?", "func=ajaxupdate&datatype=prekon_1&jid=" + idtrim);
    preconditions_met_Txt = "<br><span class='prekon_2' id='prekon_" + idtrim + "' style='color:red; font-size:10px; background:pink;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'>Pre-konditioner ikke opfyldt!</span>"

    $("#prekon_container_" + idtrim).html(preconditions_met_Txt)

    $(".prekon_2").bind('mouseover', function () {
        $(this).css('cursor', 'pointer');
    });


    $(".prekon_2").bind('click', function () {
        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)


        prekon_2(idtrim);

    });


}



function prekon_2(thisid) {


    var idtrim = thisid;


    $.post("?", "func=ajaxupdate&datatype=prekon_2&jid=" + idtrim);

    preconditions_met_Txt = "<br><span class='prekon_0' id='prekon_" + idtrim + "' style='color:#000000; font-size:10px;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'><u>Pre-kondition +/-</u></span>";

    $("#prekon_container_" + idtrim).html(preconditions_met_Txt)

    $(".prekon_0").bind('mouseover', function () {
        $(this).css('cursor', 'pointer');
    });


    $(".prekon_0").bind('click', function () {
        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)


        prekon_0(idtrim);

    });


}



function opdater_prioitet(thisid) {

  var idtrim = thisid;

  prioitetVal = $("#opd_prio_t_" + idtrim).val()

  //alert(prioitetVal)

  $.post("?", "func=ajaxupdate&datatype=prioitet&jid=" + idtrim + "&prioval=" + prioitetVal);

    $("#filterform").submit();



}

function parseDate(str) {
	try {
		var mdy = str.split("-");
		return new Date(mdy[2], mdy[1] - 1, mdy[0]);
	}
	catch (err) {
		return str;
	}
}
function daydiff(first, second) {

	return (second.getTime() - first.getTime()) / (1000 * 60 * 60 * 24)
}
function CheckIntVal(Int) {
	if (Int == undefined) { Int = 0; }
	return Int;
}

function IsValidDate(input) {
	var dayfield = input.split("/")[0];
	var monthfield = input.split("/")[1];
	var yearfield = input.split("/")[2];
	var returnval;
	var dayobj = new Date(yearfield, monthfield - 1, dayfield)
	if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
		alert("Forkert datoformat, skriv venligst dd/mm/yyyy eller d/m/yyyy");
		returnval = false;
	}
	else {
		returnval = true;
	}
	return returnval;
}


  

function positionMarker(obj, DataElement) {
	var hasDateLength; var start; var length;
	(DataElement.startdate != null && DataElement.enddate != null && DataElement.enddate != undefined && DataElement.enddate != "") ? hasDateLength = true : hasDateLength = false;

	

	//alert(DataElement.startdate + "::" +pxMulti)

	if (DataElement.alternatestartdate != null) {
		var AltBeginDate = parseDate(DataElement.alternatestartdate);
		start = (daydiff(AltBeginDate, parseDate(DataElement.startdate)) * pxMulti)
	    //alert("gher 1")
    }
else {

    if ($("#interval1").is(':checked') == true) {
        //alert(DataElement.ElmType)
        start = (daydiff(BeginDate, parseDate(DataElement.startdate)) * pxMulti * 5) + BeginDateOffset.left - 22;
    } else {
        //alert(BeginDate)
        start = (daydiff(BeginDate, parseDate(DataElement.startdate)) * pxMulti) + BeginDateOffset.left - 9;
    }

		//alert("DayDiff: " + daydiff(BeginDate, parseDate(DataElement.startdate)) +" - Invt. BeginDate: " + BeginDate + " - Job st. dato: " + DataElement.startdate + " - pxMulti: " + pxMulti + " - BeginDateOffset.left: " + BeginDateOffset.left)
    }



    if (DataElement.absmarkcorr == "mid") {
        if ($.browser.msie) { start += 3};

    } else {

        //start -= 0
    }

    //alert(start)

    if (DataElement.absmarkcorr == "today") {
        if ($.browser.msie) { start -= 0 };

    } else {

        //alert(DataElement.ElmType)
        //start -= 0
    }

    if (DataElement.absmarkcorr == "mil") {
        if ($.browser.msie) { start -= 0 };

    } else {

        //start -= 100
    }





   
	

    // år ell. kvartal / uge
	if ($("#interval1").is(':checked') == true) {
	    //define length and start position
	    //alert(DataElement.startdate)
	    if (hasDateLength) { length = daydiff(parseDate(DataElement.startdate), parseDate(DataElement.enddate)) * pxMulti * 5 + 10; }
	} else {
	    //define length and start position
	    if (hasDateLength) { length = daydiff(parseDate(DataElement.startdate), parseDate(DataElement.enddate)) * pxMulti; }
	}

	//alert("DataElement.startdate: " + DataElement.startdate + " start:" + start + "length:" + length + " pxMulti:" + pxMulti)

	obj.show();

	if (obj.hasClass("ui-icon-seek-next") || obj.hasClass("ui-icon-seek-prev")) obj.removeClass("ui-icon-orange").removeClass("ui-icon-seek-next").removeClass("ui-icon-seek-prev");
	//perform correctional actions if there is a datelength defined.

	if (hasDateLength) {
		//if nothing of the percentbar should be visible inside the frame hide it and the containing job
		if (start + length < BeginDateOffset.left) {
			(obj.parent(".job").length > 0) ? obj.hide().parent(".job").remove() : start = BeginDateOffset.left, length = 16, obj.css("overflow", "hidden").progressbar("destroy").addClass("ui-icon-orange").addClass("ui-icon-seek-next");
		}
		else {
			//else if only part of it should be visible, adapt.
			if (start < BeginDateOffset.left) {
				length = length - (BeginDateOffset.left - start);
				start = BeginDateOffset.left;
			}
		}

		//if nothing of the percentbar should be visible inside the frame hide it and the containing job
		if (start > EndDateOffset.left) {
			(obj.parent(".job").length > 0) ? obj.hide().parent(".job").remove() : start = EndDateOffset.left, length = 10, obj.css("overflow", "hidden").progressbar("destroy").addClass("ui-icon-orange").addClass("ui-icon-seek-prev");
		}
		else {
			//else if only part of it should be visible, adapt.
			if (start + length > EndDateOffset.left) {
				length = EndDateOffset.left - start;
			}
		}

        //alert("Her" + length)
		obj.width(length);
	}

    //alert(start)
	obj.css("left", start);

















}

$.fn.makeMilePael = function(DataElement) {
	var milePael = $(this);
	DataElement.absmarkcorr = "mil";
	DataElement.datatype = "milepole";
	milePael.attr("title", "(" + DataElement.startdate + ") " + DataElement.name + " - " + DataElement.type);
	positionMarker(milePael, DataElement);
}


$.fn.makePercentBar = function (DataElement) {


    var percentbar = $(this);

    //make sure that calculations can be made, and that they haven't been made before

    if (DataElement.precalc != 0) {

        //if (Math.round(DataElement.precalc) >= Math.round(DataElement.realised)) {
        DataElement.realised = DataElement.realised.replace(".", "")
        DataElement.precalc = DataElement.precalc.replace(".", "")
        //alert((DataElement.realised / DataElement.precalc))
        var percentval = Math.round((DataElement.realised / DataElement.precalc) * 100);
        //} else {
        //    var percentval = Math.round((DataElement.precalc / DataElement.realised) * 100);
        //}

    } else {

        if (DataElement.realised != 0) {
            var percentval = 100
        } else {
            var percentval = 0
        }

    }

    //if precalc is equal to 0, meaning that there is no estimate, make the percentage 100
    isFinite(percentval) ? percentval = percentval : (DataElement.realised > DataElement.precalc) ? percentval = 100 : percentval = 0;

    percentbar.attr("title", DataElement.startdate + " - " + DataElement.enddate).progressbar({ value: percentval })
		.children(".ui-progressbar-value").text(percentval + "%").css("text-align", "center");


    if (percentval > 100) {
        percentbar.css({ borderColor: "#FF8F8F" });
        percentbar.children(".ui-progressbar-value").css({ borderColor: "#CF0000", background: "#FF1F1F url(../inc/jquery/images/ui-bg_gloss-wave_45_e14f1c_500x100.png) repeat-x scroll 50% 50%" });

    }
    if (percentval == 100) {

        percentbar.css({ borderColor: "#B2FF5F" });
        percentbar.children(".ui-progressbar-value").css({ borderColor: "#4B8F00", background: "#FF1F1F url(../inc/jquery/images/ui-bg_gloss-wave_50_6eac2c_500x100.png) repeat-x scroll 50% 50%" });

    }
    if (DataElement.jobresp1 != null) {
        //define jobresponsible variables
        var jobresp1split = DataElement.jobresp1.split('-');
        var jobresp1 = jobresp1split[1];
        var jobresp1mail = jobresp1split[2];
        var jobresp2split = DataElement.jobresp2.split('-');
        var jobresp2 = jobresp2split[1];
        var jobresp2mail = jobresp2split[2];

        percentbar.siblings(".percentInfo").html("<span style='font-size:10px; color:#000000;'>Fork.: " + DanDeci(DataElement.precalc) + " t. Real.: " + DanDeci(DataElement.realised) + " t. Res.: " + DanDeci(DataElement.resource) + " t. <br>Prioitet: <input type='text' id='opd_prio_t_" + DataElement.jid + "' style='width:40px; font-size:8px;' value='" + DanDeci(DataElement.prioritet) + "'><span class='opd_prio' id='opd_prio_" + DataElement.jid + "' style='color:#6CAE1C; font-size:10px; font-weight:bolder;'> >></span> <br> Jobansv: <a class='rmenu' href='mailto:" + jobresp1mail + "'>" + jobresp1 + "</a> | <a class='rmenu' href='mailto:" + jobresp2mail + "'>" + jobresp2 + "</a></span>");
    }
    else {
        percentbar.siblings(".percentInfo").html("<span style='font-size:10px; color:#000000;'>Fork.: " + DanDeci(DataElement.precalc) + " t. Real.: " + DanDeci(DataElement.realised) + " t. Res.: " + DanDeci(DataElement.resource) + " t. </span>");
    }


    //<span style='font-size:9px; color:#999999;'>Forkalk. timer: " + f.precalc + " Realiseret: " + f.realised + " Ressource: " + f.resource + " Prioitet: " + f.prioritet + "</span>











    positionMarker(percentbar, DataElement, true);
    //add style to jobtable
    $(".job:odd").css({ "background-color": "#F7F7F7", "border-top": "0px solid #ccc", "border-bottom": "0px solid #ccc" });
    $(".job:even").css({ "background-color": "#FFFFFF", "border": "0" });




    //Today marker
    
    //alert($("#JobGraph").height())
    //var hgt = $("#JobGraph").height() - 300
    //alert(hgt)
    //$("#todayMarker").height(hgt);
    $("#todayMarker").height($("#JobGraph").height());

    if ($("#interval1").is(':checked') == true) {
        $("#todayMarker").css({ "top": "35px" });
    } else {
        $("#todayMarker").css({ "top": "35px" });
    }


    //Start grid
    $("#startMarker").height($("#JobGraph").height());

    if ($("#interval1").is(':checked') == true) {
        $("#startMarker").css({ "top": "70px" });
    } else {
        $("#startMarker").css({ "top": "75px" });
    }


    //end grid
    $("#endMarker").height($("#JobGraph").height());

    if ($("#interval1").is(':checked') == true) {
        $("#endMarker").css({ "top": "70px" });
    } else {
        $("#endMarker").css({ "top": "75px" });
    }

    $(".midMarker").height($("#JobGraph").height());

    if ($("#interval1").is(':checked') == true) {
        $(".midMarker").css({ "top": "70px" });
    } else {
        $(".midMarker").css({ "top": "75px" });
    }




}

function DanDeci(EngDeci) {
	return EngDeci.toString().replace(',', '.');
}







function AddJobs(options) {
	defaults = { print: false, searchId: "-1" };
	opts = $.extend(defaults, options);
	var print = opts.print;
	var SearchId = opts.searchId;
	var AddJobList = AddJobListDB;
	var cnt = 0;

	//alert(AddJobList.length + SearchId + options)

	var ij = AddJobList.length; while (ij--) {

    cnt = cnt + 1

	    //alert(AddJobList[ij].jid)

        var ArrActivities = AddJobList[ij].ArrActivities;
		var ArrMilePael = AddJobList[ij].ArrMilePael;
		var customer = AddJobList[ij].customer;
		var name = AddJobList[ij].name;
		var jid = AddJobList[ij].jid;
		var precalc = AddJobList[ij].precalc;
		var realised = AddJobList[ij].realised;
		var resource = AddJobList[ij].resource;
		
		var datediff = Math.round(daydiff(parseDate(AddJobList[ij].startdate), parseDate(AddJobList[ij].enddate)))
	    var startdate = AddJobList[ij].startdate;
		//alert(startdate)
		var enddate = AddJobList[ij].enddate;
		var jnr = AddJobList[ij].jnr;
		var cnr = AddJobList[ij].cnr;
		var prioritet = AddJobList[ij].prioritet;

		//define jobresponsible variables
		var jobresp1split = AddJobList[ij].jobresp1.split('-');
		var jobresp1 = jobresp1split[1];
		var jobresp1mail = jobresp1split[2];
		var jobresp2split = AddJobList[ij].jobresp2.split('-');
		var jobresp2 = jobresp2split[1];
		var jobresp2mail = jobresp2split[2];

		var preconditions_met = AddJobList[ij].preconditions_met



		var preconditions_met_Txt = "<br><span class='prekon_0' id='prekon_" + jid + "' style='color:#000000; font-size:10px;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'><u>Pre-kondition +/-</u></span>";
		//alert(preconditions_met)
        if (preconditions_met == '1') {
            preconditions_met_Txt = "<br><span class='prekon_1' id='prekon_" + jid + "' style='color:#6CAE1C; font-size:10px; background:#DCF5BD;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'>Pre-konditioner opfyldt</span>"
		}
        
        if (preconditions_met == '2') {
            preconditions_met_Txt = "<br><span class='prekon_2' id='prekon_" + jid + "' style='color:red; font-size:10px; background:pink;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'>Pre-konditioner ikke opfyldt!</span>"
		}

        //preconditions_met_Txt_lnk = "<span style='color:#000000; font-size:10px;' title=''>+</span>&nbsp;<span style='color:#000000; font-size:10px;' title=''>-</span>"

		//var ElmType = 1;

		//create milepael elements on job
		var MilePaelStr = "";
		for (var i = 0, il = ArrMilePael.length; i < il; i++) {
			MilePaelStr += "<div class\=\'ui-icon-orange ui-icon-flag milepael\' style='position\:absolute\' ><\/div>";
		}




$("#JobGraph").append("<div class='job' id=" + ij + "><span class='jobInfo' style='padding-left:3px;'>" + (customer + " (" + cnr + ")") + "<br><a href='jobs.asp?menu=job&func=red&id=" + jid + "&int=1&rdir=webblik_joblisten2&nomenu=1' class='vmenu' title='" + name + " (" + jnr + ")" + "' target='_blank'>" + (name + " (" + jnr + ")").substring(0, 31) + "</a><br><span style='font-size:9px; color:#999999;'>" + startdate + " - " + enddate + " <b>" + datediff + " dage</b></span><span class='prekon_container' id='prekon_container_" + jid + "'>" + preconditions_met_Txt + "</span></span>" + MilePaelStr + "<br><div class='percentBar' title='" + startdate + " - " + enddate + "'></div><span class='percentInfo'></span><br class='clear' /><div class='activities'></div><center class='clear'><div class='closeOpenIcon ui-icon ui-icon-carat-1-s' /></center></div>");
		var newJob = $("#JobGraph div.job:last");
		//fade if priority is hidden
		if (prioritet == -1 || prioritet == -2) { newJob.fadeTo(1, 0.8); }
		newJob.data("dataElm", AddJobList[ij]);
		var ActivityDiv = newJob.children("div.activities");
		newJob.children("div.percentBar").makePercentBar(AddJobList[ij]);

		//if (newJob.length > 0) {
			//check whether a dateline should be inserted
			//if ((ij / 7).toString().indexOf('.') == -1 && ij + 5 < AddJobList.length) {
			//	try {
			//		$("#JobGraphHeader").clone().css("margin", "5px 0px 5px 0px").insertAfter(newJob);
			//	}
			//	catch (err) { }
			//}
		//}

		newJob.children("div.milepael").each(function(index) {
	    $(this).data("dataElm", ArrMilePael[index]).makeMilePael(ArrMilePael[index]);
		});
		//show job if it's searched for
		if (SearchId == jid | SearchId.toLowerCase() == name.toLowerCase())
		{ ActivityDiv.addClass("searchResult"); }
	}

	
	//perform final wrap up





    // Activities HTML
	function actString() {



        markerStopFn();

	    job.children(".activities").stop().slideDown(200, function () {
	        $("#todayMarker").height($("#JobGraph").height() - markerStop);
	        $("#startMarker").height($("#JobGraph").height() - markerStop);
	        $("#endMarker").height($("#JobGraph").height() - markerStop);
	        $(".midMarker").height($("#JobGraph").height() - markerStop);
	    });


        if (!(job.hasClass("hasActivities"))) {

	    var DataElm = job.data("dataElm");
	    var ArrActivities = DataElm.ArrActivities;



	    //create subbars with activities
	    //Activity elements must look like ArrActivities[i] = { "name" : (), "precalc" : (), "realised" : (), "startdate" : (), "enddate" : (), "id" : () })
	    var ActivitiesStr = "";
	    var datediffA = 0


	    for (var i = 0, il = ArrActivities.length; i < il; i++) {
	        datediffA = Math.round(daydiff(parseDate(ArrActivities[i].startdate), parseDate(ArrActivities[i].enddate)))
	        ActivitiesStr += "<br class='clear' /><div class='activity' style='border-top:0px #8cAAe6 dashed;'><span class='jobInfo' style='font-size:10px; padding-left:20px; line-height:12px; color:#5582d2;' title='" + (ArrActivities[i].name + " (#" + ArrActivities[i].id + ")") + "'><b>" + (ArrActivities[i].name + "").substring(0, 300) + "</b><br><span style='font-size:9px; color:#999999;'>" + ArrActivities[i].startdate + " - " + ArrActivities[i].enddate + " <b>" + datediffA + " dage</b></span></span><div class='percentBar' title='" + ArrActivities[i].startdate + " - " + ArrActivities[i].enddate + "'></div><span class='percentInfo'></span></div>";
	    }
	    //ActivitiesStr += '<br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir="><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	    //ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'>Opret ny Milep?l</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	    if (!(print)) {
	        ActivitiesStr += '<br class="clear" /><br class="clear" /><br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank"><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	        ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milepæl" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","600","200","200")\'>Opret ny Milepæl</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milepæl" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","600","200","200")\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	        //ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'>Opret ny Milep?l</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';

	    }

	    job.children(".activities").html(ActivitiesStr)

            .children(".activity").each(function (index) {
                $(this).data("dataElm", ArrActivities[index]);
                //get activity percentbar and pass on for processing on position
                $(this).children("div.percentBar").makePercentBar(ArrActivities[index]);
            });

	    Draggable();
	    job.addClass("hasActivities");

        }

	}



    /////////////////////////////////////////
    // Åbner default aktiviteter ved 1 job åben kun 
    ///
    //alert(cnt)
    if (cnt == 1) {

	   var job = $("#JobGraph div.job:last");

	   actString();



	   $("#JobGraph div.job:last").removeClass("ui-icon-carat-1-s").addClass("ui-icon-carat-1-n");
	    //////////////////////////

	} // ij == 1


	if (cnt != 1) {
        //hide activities
	    if (!(print)) $("#JobGraph .activities").hide();
	}


	$("#JobGraph .loadingIcon").remove();



    //create dialog box for dragging
	if (!(print)) $("#JobGraph").append("<div id='jobDialog' title='ændring af dato'><p>Datoen på <span class='jobName'>0</span> vil blive ændret <span class='jobNotice'>0</span></p><br /><span id='ajaxFromDate'>Fra:<br /><input type='text' id='ajaxStartDate' /></span><br /><span id='ajaxToDate'>Til:<br /><input type='text' id='ajaxEndDate' /></span></div>");
	
    
    
    
    
    //create toggle methods for activity divs

	$("#JobGraph .closeOpenIcon").toggle(function () {


	    var job = $(this).parent().parent(".job");

	    
        //actString();

	    markerStopFn();

	    job.children(".activities").stop().slideDown(200, function () {
	        $("#todayMarker").height($("#JobGraph").height() - markerStop);
	        $("#startMarker").height($("#JobGraph").height() - markerStop);
	        $("#endMarker").height($("#JobGraph").height() - markerStop);
	        $(".midMarker").height($("#JobGraph").height() - markerStop);
	    });


	    if (!(job.hasClass("hasActivities"))) {

	        var DataElm = job.data("dataElm");
	        var ArrActivities = DataElm.ArrActivities;



	        //create subbars with activities
	        //Activity elements must look like ArrActivities[i] = { "name" : (), "precalc" : (), "realised" : (), "startdate" : (), "enddate" : (), "id" : () })
	        var ActivitiesStr = "";
	        var datediffA = 0


	        for (var i = 0, il = ArrActivities.length; i < il; i++) {
	            datediffA = Math.round(daydiff(parseDate(ArrActivities[i].startdate), parseDate(ArrActivities[i].enddate)))
	            ActivitiesStr += "<br class='clear' /><div class='activity' style='border-top:0px #8cAAe6 dashed;'><span class='jobInfo' style='font-size:10px; padding-left:20px; line-height:12px; color:#5582d2;' title='" + (ArrActivities[i].name + " (#" + ArrActivities[i].id + ")") + "'><b>" + (ArrActivities[i].name + "").substring(0, 300) + "</b><br><span style='font-size:9px; color:#999999;'>" + ArrActivities[i].startdate + " - " + ArrActivities[i].enddate + " <b>" + datediffA + " dage</b></span></span><div class='percentBar' title='" + ArrActivities[i].startdate + " - " + ArrActivities[i].enddate + "'></div><span class='percentInfo'></span></div>";
	        }
	        //ActivitiesStr += '<br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir="><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	        //ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'>Opret ny Milep?l</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	        if (!(print)) {
	            ActivitiesStr += '<br class="clear" /><br class="clear" /><br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank"><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	            ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milepæl" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","600","200","200")\'>Opret ny Milepæl</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milepæl" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","600","200","200")\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	            //ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'>Opret ny Milep?l</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';

	        }

	        job.children(".activities").html(ActivitiesStr)

            .children(".activity").each(function (index) {
                $(this).data("dataElm", ArrActivities[index]);
                //get activity percentbar and pass on for processing on position
                $(this).children("div.percentBar").makePercentBar(ArrActivities[index]);
            });

	        Draggable();
	        job.addClass("hasActivities");

	    }



	    $(this).removeClass("ui-icon-carat-1-s").addClass("ui-icon-carat-1-n");



	},




	function () {
	    var job = $(this).parent().parent(".job");

	    markerStopFn();

	    job.children(".activities").stop().slideUp(200, function () {

	        $("#todayMarker").height($("#JobGraph").height() - markerStop);
	        $("#startMarker").height($("#JobGraph").height() - markerStop);
	        $("#endMarker").height($("#JobGraph").height() - markerStop);
	        $(".midMarker").height($("#JobGraph").height() - markerStop);
	    });

	    $(this).removeClass("ui-icon-carat-1-n").addClass("ui-icon-carat-1-s");
	});



    var markerStop = 100
    function markerStopFn() {

        if ($("#interval1").is(':checked') == true) {
            markerStop = 135
        } else {
            markerStop = 95
        }

    }


  


    /* /Bind 
    $(".prekon_1").bind('mouseover', function () {
        $(this).css('cursor', 'pointer');
    });

    $(".prekon_1").bind('click', function () {
        var thisid = this.id

        var idlngt = thisid.length
        var idtrim = thisid.slice(7, idlngt)

        //alert(idtrim)
        $.post("?", "func=ajaxupdate&datatype=prekon_1&jid=" + idtrim);
        preconditions_met_Txt = "<br><span class='prekon_2' id='prekon_" + idtrim + "' style='color:red; font-size:10px; background:pink;' title='Angiv og jobbet er klar til produktion, er der bestilt materialer mv.'>Pre-konditioner ikke opfyldt!</span>"

        $("#prekon_container_" + idtrim).html(preconditions_met_Txt)

    }); */




                    function Draggable() {


                              //create drag methods
		                        var ObjOffset;
		                        var SelPercentBar;
		                        var SelPercentActivities;
		                        var SelMilePael;
		                        var SelPercentParent;
		                        $(".percentBar, .milepael").draggable({ axis: 'x',
		        
                                        start: function () {
		                                SelPercentBar = $(this).filter(".percentBar");
		                                SelMilePael = $(this).filter(".milepael");
		                                SelPercentActivities = SelPercentBar.parent(".job").children(".activities").children(".activity");
		                                SelPercentParent = SelPercentBar.parent(".activity").parent(".activities").parent(".job");
		                                ObjOffset = $(this).offset();
		                                //scroll down activities if they are not currently displayed

		                                
                                        
		        
                                        if ($(this).parent(".job").children(".activities").css("display") == "none") { $(this).parent(".job").click(); }}, stop: function () {
		                                var ObjBeginDate = new Date(BeginDate);
		                                var ObjNewOffset = $(this).offset();
		                                var DataElm;
		                                ($(this).data("dataElm") != null) ? DataElm = $(this).data("dataElm") : DataElm = $(this).parent().data("dataElm");
		                                var hasDateLength;
		                                (DataElm.startdate != null && DataElm.enddate != null) ? hasDateLength = true : hasDateLength = false;

		                                if (hasDateLength) { var datediff = Math.round(daydiff(parseDate(DataElm.startdate), parseDate(DataElm.enddate))); }

		                                // år ell. kvartal / uge
		                                if ($("#interval1").is(':checked') == true) {
		                                    //alert("berenger dato" + ObjBeginDate.getDate() + " ObjNewOffset.left: " + ObjNewOffset.left + " BeginDateOffset.left: " + BeginDateOffset.left + " pxMulti " + pxMulti)
		                                    ObjBeginDate.setDate(ObjBeginDate.getDate() + ((ObjNewOffset.left - BeginDateOffset.left - 0) / (pxMulti * 5)));
		                                } else {
		                                    ObjBeginDate.setDate(ObjBeginDate.getDate() + ((ObjNewOffset.left - BeginDateOffset.left - 0) / pxMulti));
		                                }
		                                //$("#jobDialog").dialog('open');

		                                $("#jobDialog .jobName").html(DataElm.name);
		                                DataElm.jid != undefined ? jobNotice = " Eventuelle under-aktiviteter vil blive rykket med" : jobNotice = "";

		                                //format dateinput boxes default, and make sure that enddate follows startdate

		                                $("#jobDialog #ajaxEndDate").val($.datepicker.formatDate('d/m/yy', ObjBeginDate)).change(function () {
		                                    try {
		                                        var ajaxEndDate = $.datepicker.parseDate('dd/mm/yy', $("#ajaxEndDate").val());
		                                        $(".ui-dialog-buttonpane button:first").disabled = false;
		                                    }
		                                    catch (err) {
		                                        alert("forkert datoformat skriv venligst d/m/åå eller dd/mm/yy");
		                                        $(".ui-dialog-buttonpane button:first").disabled = true;
		                                    }
		                                });
		                                $("#jobDialog #ajaxEndDate").change()
		                                // if there is an end date, activate the enddate change to take affect, else hide the enddate
		                                hasDateLength ? $("#jobDialog #ajaxToDate").show() : $("#jobDialog #ajaxToDate").hide();

		                                $("#jobDialog #ajaxStartDate").val($.datepicker.formatDate('d/m/yy', ObjBeginDate)).change(function () {
		                                    try {
		                                        var ajaxEndDate = $.datepicker.parseDate('dd/mm/yy', $("#ajaxStartDate").val());
		                                        ajaxEndDate.setDate(ajaxEndDate.getDate() + datediff);
		                                        $("#jobDialog #ajaxEndDate").val($.datepicker.formatDate('d/m/yy', ajaxEndDate));
		                                        $(".ui-dialog-buttonpane button:first").disabled = false;
		                                    }
		                                    catch (err) {
		                                        alert("forkert datoformat skriv venligst d/m/åå eller dd/mm/yy");
		                                        $(".ui-dialog-buttonpane button:first").disabled = true;
		                                    }
		                                });
		                                $("#jobDialog #ajaxStartDate").change();


		                                $("#jobDialog .jobNotice").html(jobNotice);

		                                $("#jobDialog").dialog('option', 'buttons', { 'Accept': function () {

		                                    var asdsDate = $.datepicker.parseDate('dd/mm/yy', $("#jobDialog #ajaxStartDate").val());
		                                    if (hasDateLength) { var aedsDate = $.datepicker.parseDate('dd/mm/yy', $("#jobDialog #ajaxEndDate").val()); }

		                                    var ajaxDatediff = Math.round(daydiff(parseDate(DataElm.startdate), asdsDate));
		                                    $.post("?", "jid=" + DataElm.jid + "&cnr=" + DataElm.cnr + "&dateDiff=" + ajaxDatediff + "&id=" + DataElm.id + "&startdate = " + $.datepicker.formatDate('yy/mm/d', asdsDate) + "&enddate = " + $.datepicker.formatDate('yy/mm/d', aedsDate) + "&func=ajaxupdate&datatype=" + DataElm.datatype + "&ajid=" + DataElm.ajid);

		                                    $(this).dialog('close');
		                                    //reposition and rewrite activities
		                                    if (hasDateLength) {
		                                        SelPercentActivities.each(function () {
		                                            var activityDataElement = $(this).data("dataElm");

		                                            //write new start and enddates to dataelement
		                                            activityDataElement.startdate = $.datepicker.parseDate('dd-mm-yy', activityDataElement.startdate); (activityDataElement.startdate).setDate(((activityDataElement.startdate).getDate() + ajaxDatediff)); activityDataElement.startdate = $.datepicker.formatDate('dd-mm-yy', activityDataElement.startdate);
		                                            activityDataElement.enddate = $.datepicker.parseDate('dd-mm-yy', activityDataElement.enddate); (activityDataElement.enddate).setDate((activityDataElement.enddate).getDate() + ajaxDatediff); activityDataElement.enddate = $.datepicker.formatDate('dd-mm-yy', activityDataElement.enddate);

		                                            $(this).children(".percentBar").makePercentBar(activityDataElement);

		                                            //return activityDataElement rewritten
		                                            $(this).data("dataElm", activityDataElement);
		                                        });

		                                        //reposition and rewrite milepael
		                                        SelPercentBar.parent().children(".milepael").each(function () {
		                                            var milePaelDataElement = $(this).data("dataElm");
		                                            //write new data to milepael
		                                            milePaelDataElement.startdate = $.datepicker.parseDate('dd-mm-yy', milePaelDataElement.startdate); (milePaelDataElement.startdate).setDate(((milePaelDataElement.startdate).getDate() + ajaxDatediff)); milePaelDataElement.startdate = $.datepicker.formatDate('dd-mm-yy', milePaelDataElement.startdate);
		                                            $(this).data("dataElm", milePaelDataElement).makeMilePael(milePaelDataElement);
		                                        });

		                                        //Rewrite DataElement and return to job/activity
		                                        DataElm.startdate = $.datepicker.formatDate('dd-mm-yy', asdsDate);

		                                        if (SelPercentBar.length > 0) {
		                                            DataElm.enddate = $("#ajaxEndDate").val().replace('/', '-').replace('/', '-');
		                                            SelPercentBar.parent(".job").data("dataElm", DataElm);
		                                            SelPercentBar.makePercentBar(DataElm);

		                                            var parentDataElm = SelPercentParent.data("dataElm");

		                                            if (parentDataElm != null) {
		                                                if (($.datepicker.parseDate('dd-mm-yy', DataElm.startdate)).getTime() < ($.datepicker.parseDate('dd-mm-yy', parentDataElm.startdate)).getTime()) { parentDataElm.startdate = DataElm.startdate; }
		                                                if (($.datepicker.parseDate('dd-mm-yy', DataElm.enddate)).getTime() > ($.datepicker.parseDate('dd-mm-yy', parentDataElm.enddate)).getTime()) { parentDataElm.enddate = DataElm.enddate; }
		                                                SelPercentParent.data("dataElm", parentDataElm);

		                                                SelPercentParent.children(".percentBar").makePercentBar(parentDataElm);
		                                            }
		                                        }
		                                        if (SelMilePael.length > 0) {
		                                            SelMilePael.data("dataElm", DataElm).makeMilePael(DataElm);
		                                        }
		                                    }
		                                },
		                                    Cancel: function () {
		                                        $(this).dialog('close');
		                                        if (SelPercentBar.length > 0) {
		                                            SelPercentBar.makePercentBar(DataElm);
		                                        }
		                                        if (SelMilePael.length > 0) {
		                                            SelMilePael.makeMilePael(DataElm);
		                                        }
		                                    }
		                                });

		                                $("#jobDialog").dialog('open');

		                            }
		                        });
                        } //Dragable



                   
                

                      



                       

                       




	if (!(print)) {
		Draggable();
		//make dialogbox accesible
		$("#jobDialog").dialog({
			bgiframe: true,
			resizable: true,
			autoOpen: false,
			height: 420,
			modal: true,
			overlay: {
				backgroundColor: '#000',
				opacity: 0.5
			},
			buttons: { 'Accept': function() { $(this).dialog('close'); },
				Cancel: function() {
					$(this).dialog('close');
					$(this).css({ left: ObjOffset.left, top: ObjOffset.top });
				}
			}
		});

		$.datepicker.setDefaults($.extend({ showAnim: 'slide', showOptions: { direction: 'drop' }, duration: -50, showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd/m/yy', changeMonth: true, changeYear: true }));
		
        $("#ajaxStartDate").datepicker();
		$("#ajaxEndDate").datepicker();

		$("#JobGraph .ui-icon").hover(function () {
            $(this).removeClass("ui-icon").addClass("ui-icon-orange");
		}, function() {
			$(this).removeClass("ui-icon-orange").addClass("ui-icon");
		});
	}


          









	if (print) {
	    $(".job").each(function () {
	        var job = $(this);

	        actString();

	        /*
            markerStopFn();

	        job.children(".activities").stop().slideDown(200, function () {
	            $("#todayMarker").height($("#JobGraph").height() - markerStop);
	            $("#startMarker").height($("#JobGraph").height() - markerStop);
	            $("#endMarker").height($("#JobGraph").height() - markerStop);
	            $(".midMarker").height($("#JobGraph").height() - markerStop);
	        });

	            if (!(job.hasClass("hasActivities"))) {
	            var DataElm = job.data("dataElm");
	            var ArrActivities = DataElm.ArrActivities;
	            //create subbars with activities
	            //Activity elements must look like ArrActivities[i] = { "name" : (), "precalc" : (), "realised" : (), "startdate" : (), "enddate" : (), "id" : () })
	            var ActivitiesStr = "";
	            for (var i = 0, il = ArrActivities.length; i < il; i++) {
	                ActivitiesStr += "<br class='clear' /><div class='activity'><span class='jobInfo' title='" + (ArrActivities[i].name + " (#" + ArrActivities[i].id + ")") + "'>" + (ArrActivities[i].name + " (" + ArrActivities[i].id + ")").substring(0, 31) + "</span><div class='percentBar' title='" + ArrActivities[i].startdate + " - " + ArrActivities[i].enddate + "'></div><span class='percentInfo'></span></div>";
	            }

	            //ActivitiesStr += '<br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir="><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	            //ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'>Opret ny Milep?l</a></td><td style="padding: 3px 0px 0px;"><a target="_top" alt="Opret ny Milep?l" class="vmenu" href=\'javascript:popUp("milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik","650","500","250","120");\'><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	            if (!(print)) {
	                ActivitiesStr += '<br class="clear" /><div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 130px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank">Opret ny aktivitet</a></td><td style="padding: 3px 0px 0px;"><a alt="Opret ny aktivitet" class="vmenu" href="aktiv.asp?menu=job&amp;func=opret&amp;jobid=' + DataElm.jid + '&amp;id=0&amp;jobnavn=' + DataElm.name + '&amp;fb=1&amp;rdir=webbl21&nomenu=1" target="_blank"><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>'
	                ActivitiesStr += '<div style="border: 1px solid rgb(140, 170, 230); padding: 3px; margin:3px; float:left; width: 140px; background-color: rgb(255, 255, 255);"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tbody><tr><td><a target="_blank" alt="Opret ny Milepæl" class="vmenu" href="milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik">Opret ny Milepæl</a></td><td style="padding: 3px 0px 0px;"><a target="_blank" alt="Opret ny Milepæl" class="vmenu" href="milepale.asp?menu=job&amp;func=opr&amp;jid=' + DataElm.jid + '&amp;rdir=webblik"><img border="0" src="../ill/add2.png"/></a></td></tr></tbody></table></div>';
	            }

	            job.children(".activities").html(ActivitiesStr)
		        .children(".activity").each(function (index) {
		        $(this).data("dataElm", ArrActivities[index]);
		        //get activity percentbar and pass on for processing on position
		        $(this).children("div.percentBar").makePercentBar(ArrActivities[index]);
		  });
	            job.addClass("hasActivities");
	        }
            */

	        $(this).removeClass("ui-icon-carat-1-s").addClass("ui-icon-carat-1-n");
	    });



	    } //print

  
	
	//toggle search results
	$(".searchResult").parent().children("center").children(".closeOpenIcon").click();
}














function PopulateGraphFrame() {
	//create monthnavigation
	var FM_startdate = $("#FM_startdate");
	var naviDate = $.datepicker.parseDate('d/m/yy', FM_startdate.val());
	//alert(naviDate)

	if ($("#interval1").is(':checked') == true) {
    var varAddM = 1
	} else {
    var varAddM = 3
    }

    $("#navPrevMonth").click(function () {
        naviDate.setMonth(naviDate.getMonth() - varAddM);
        FM_startdate.val($.datepicker.formatDate('d/m/yy', naviDate));
        var datestring = FM_startdate.val(); var datesplit = datestring.split('/'); $("#FM_start_dag").val(datesplit[0]); $("#FM_start_mrd").val(datesplit[1]); $("#FM_start_aar").val(datesplit[2]);
        $("#filterform").submit();
    });
	$("#navNextMonth").click(function() {
	    naviDate.setMonth(naviDate.getMonth() + varAddM);
		$("#FM_startdate").val($.datepicker.formatDate('d/m/yy', naviDate));
		var datestring = FM_startdate.val(); var datesplit = datestring.split('/'); $("#FM_start_dag").val(datesplit[0]); $("#FM_start_mrd").val(datesplit[1]); $("#FM_start_aar").val(datesplit[2]);
		$("#filterform").submit();
	});


 

    
	//alert("her")

	//Create multidimensional array and ordinary array with month names
	// år ell. kvartal / uge
    if ($("#interval0").is(':checked') == true) {

    //Create list items with correct width and name them
    var CalBeginDate = $.datepicker.parseDate('d/m/yy', FM_startdate.val()); CalBeginDate.setMonth(CalBeginDate.getMonth() - 6); CalBeginDate.setDate(1);

	
	    var m_names = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
	    var Ival = 12

	    MonthArr = new Array("", "");

	    for (var i = 0; i < Ival; i++) {
	        CalBeginDate.setMonth(CalBeginDate.getMonth() + i);
	        MonthArr[i] = { "days": daysInMonth(CalBeginDate.getMonth() + 1, CalBeginDate.getFullYear()), "name": m_names[CalBeginDate.getMonth()] + " " + CalBeginDate.getFullYear(), "date": CalBeginDate };
	       


	        CalBeginDate.setMonth(CalBeginDate.getMonth() - i);

	        $("#JobGraphHeader").append("<li class='month'>&nbsp;</li>");
	        var newLI = $("#JobGraphHeader li.month:last");
	        //-5 because this is border (1px) + 2*margin(2px)
	        newLI.width(((MonthArr[i].days) * pxMulti) - 1)
		    .text(MonthArr[i].name)
		    .data("date", MonthArr[i].date);

	        if (i == 0) {
	            BeginDate = new Date(MonthArr[i].date);
	            BeginDateOffset = newLI.offset(); 
                BeginDateOffset.left = BeginDateOffset.left - 20;
	            //alert(BeginDate)
            }
	        if (i == 11) { EndDate = new Date(MonthArr[i].date); EndDate.setMonth(EndDate.getMonth() + 12); EndDateOffset = newLI.offset(); EndDateOffset.left += newLI.width(); EndDateOffset.left -= 20; }
	    }

	} else {

        

	    Date.prototype.getWeek = function () {
	        var onejan = new Date(this.getFullYear(), 0, 1);
	        var today = new Date(this.getFullYear(), this.getMonth(), this.getDate());
	        var dayOfYear = ((today - onejan + 1) / 86400000);
	        return Math.ceil(dayOfYear / 7)
	    };

	   
       //Create list items with correct width and name them

	    var CalBeginDate = $.datepicker.parseDate('d/m/yy', FM_startdate.val());
	    CalBeginDate.setMonth(CalBeginDate.getMonth() - 1);
	    CalBeginDate.setDate(1);

	   
	    //alert(CalBeginDate)

	    var BeginDateUse = $.datepicker.parseDate('d/m/yy', FM_startdate.val());
	    BeginDateUse.setYear(BeginDateUse.getFullYear())
	    BeginDateUse.setMonth(BeginDateUse.getMonth() - 1);
	    BeginDateUse.setDate(1);

	    //alert(BeginDateUse.getMonth() + " // " + BeginDateUse)

	 
	 
	    var weekday = new Array(7);
	    weekday[0] = "m";
	    weekday[1] = "t";
	    weekday[2] = "o";
	    weekday[3] = "t";
	    weekday[4] = "f";
	    weekday[5] = "l";
	    weekday[6] = "s";

	    var dt = new Date();
	    dt = BeginDateUse
        var dtnW = new Date();
	    
        //var m_names = new Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53");
        var m_names = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
        var Ival = 11

	    MonthArr = new Array("", "");

	    var eUsed = 0;
	    var daysInFirstWeek = 0;
	    var dtnWDno = 0;

	    // antal dage i første uge
	    dtnW.setDate(1);
	    dtnW.setMonth(BeginDateUse.getMonth())
	    dtnW.setYear(BeginDateUse.getFullYear())

	    

	    //alert(BeginDateUse)

	    dtnWDno = BeginDateUse.getUTCDay()

	    //alert(dtnWDno)

	    if (dtnWDno == 0) {
	        daysInFirstWeek = 7
	    } else {
	        daysInFirstWeek = (7 - dtnWDno)
        }


	    

	    for (var i = 0; i < Ival; i++) {



	        //daysInFirstWeek = 0
	        if (i == 0) {
	            CalBeginDate.setDate(CalBeginDate.getDate() + (((i * 7) - 7) + daysInFirstWeek));
	            CalBeginDate.setMonth(CalBeginDate.getMonth());
	            CalBeginDate.setFullYear(CalBeginDate.getFullYear());
	            
                //ThisWeek.setMonth(CalBeginDate.getMonth() - 1)
	        } else {


	            CalBeginDate.setDate(dt.getDate() + 1);
                CalBeginDate.setMonth(dt.getMonth());
	            CalBeginDate.setFullYear(dt.getFullYear());

	            //ThisWeek.setDate(dt.getDate())
	            //ThisWeek.setMonth(dt.getMonth() - 1)
	            //ThisWeek.setFullYear(dt.getFullYear())
	           
	                //alert(CalBeginDate)
            }


            if (CalBeginDate.getDate() <= 2) {
                MonthArr[i] = { "days": daysInMonth(CalBeginDate.getMonth(), CalBeginDate.getFullYear()), "name": ".", "date": BeginDateUse };
                } else {
                MonthArr[i] = { "days": daysInMonth(CalBeginDate.getMonth(), CalBeginDate.getFullYear()), "name": " [ " + CalBeginDate.getDate() + "." + m_names[CalBeginDate.getMonth()] + "." + CalBeginDate.getFullYear() + " ]  Uge: " + CalBeginDate.getWeek(), "date": BeginDateUse };

            }


            //+ "........................ Uge: " + CalBeginDate.getWeek()
	       
	            $("#JobGraphHeader").append("<li class='month'>&nbsp;</li>");
	        //alert(CalBeginDate.getDate() + "(((i:" + i + " * 7 ) - 7) daysInFirstWeek:" + daysInFirstWeek + "dato: " + CalBeginDate)

	        d = 0;

	        if (i == 0) {
	            tpArr = daysInFirstWeek
	        } else {
	            tpArr = 7
	        }


	        for (var d = 0; d < tpArr; d++) {


	            
	            if (i == 0 && d == 0) {
	                dt.setDate(1) // + (i * 7) + d + daysInFirstWeek
                    dt.setMonth(BeginDateUse.getMonth())
	                dt.setFullYear(BeginDateUse.getFullYear())
	            } else {
	                dt.setDate(dt.getDate() + 1) // + (i * 7) + d + daysInFirstWeek
                    dt.setMonth(dt.getMonth())
	                dt.setFullYear(dt.getFullYear())
	            }
                

	            //if (d == 0) {
	            //    alert(dt)
	            //}

                dtn = dt.getUTCDay()
                var dtname = weekday[dtn]
	            
                //alert(dt)

	            //var dname = dt.getUTCDay()

                    if (dtn == 5 || dtn == 6) {
                    bgc = "#CCCCCC" 
                    } else {
                    bgc = "#EEEFF3"
                    }



                        $("#JobGraphHeader_weekday").append("<li class='day' style='background-color:" + bgc + "'>" + dtname + "<br><span style='font-size:7px;'> " + dt.getDate() + "</span></li>");
             
            }



             //CalBeginDate.setDate(CalBeginDate.getDate() - (((i * 7) - 7) + daysInFirstWeek));


                   
                    


	        var newLI = $("#JobGraphHeader li.month:last");
	        var expxl = 0;

	        if ($.browser.msie) {
	            expxl = 0;
	        } else {
	            expxl = 6;
            };

	        if (i != 0) {
	            newLI.width(70 + expxl)
		        .text(MonthArr[i].name)
		        .data("date", MonthArr[i].date);
	        } else {
	            //alert(dtn)
	            newLI.width(((daysInFirstWeek) * 10))
            }

            
            if (i == 0) {
	            BeginDate = new Date(MonthArr[i].date);
	            BeginDateOffset = newLI.offset(); 
                BeginDateOffset.left = BeginDateOffset.left;
	            //alert(BeginDate)
            }

	        if (i == 10) {
	            EndDate = new Date(MonthArr[i].date); EndDate.setMonth(EndDate.getMonth()); EndDateOffset = newLI.offset(); EndDateOffset.left += newLI.width(); EndDateOffset.left -= 20;
	        }


	        


	    }




    }
    
   

 








	function daysInMonth(month, year) {
		var dd = new Date(year, month, 0);
		return dd.getDate();
    }



 

        //var nArr = midMarkerDateArr
        var dtnW = new Date();
        var midMarkerDate = new Date()

        //Strat And End Marker
        //$("#JobGraph").prepend("<div title='d.d.' id='startMarker' style='width:1px;'></div>");
        //positionMarker($("#startMarker"), { "startdate": BeginDate, "absmarkcorr": true });

        //$("#JobGraph").prepend("<div title='d.d.' id='endMarker' style='width:1px;'></div>");
        //positionMarker($("#endMarker"), { "startdate": EndDate, "absmarkcorr": true });


        var CalBeginDate = $.datepicker.parseDate('d/m/yy', FM_startdate.val());

        if ($("#interval1").is(':checked') == true) {
            CalBeginDate.setMonth(CalBeginDate.getMonth() - 1);
            CalBeginDate.setDate(1);
        } else {
            CalBeginDate.setMonth(CalBeginDate.getMonth() - 6);
            CalBeginDate.setDate(1);
        }

        //alert(CalBeginDate)

        // antal dage i første uge
        dtnW.setDate(1);
        dtnW.setMonth(CalBeginDate.getMonth())
        dtnW.setYear(CalBeginDate.getFullYear())

        var dtnWDno = CalBeginDate.getUTCDay()

        //alert(dtnWDno)

        if (dtnWDno == 0) {
            daysInFirstWeek = 7
        } else {
            daysInFirstWeek = (7 - dtnWDno)
        }


        //alert(CalBeginDate + "//" + daysInFirstWeek)

        //if ($("#interval1").is(':checked') == true) {
            BeginDate = new Date(CalBeginDate);
        //} else {
        //    BeginDate = BeginDate
        //}

        for (var i = 0; i < 11; i++) {


            if (i == 0) {
                if ($("#interval1").is(':checked') == true) {
                    CalBeginDate.setDate(CalBeginDate.getDate() + daysInFirstWeek - 2);
                } else {
                    CalBeginDate.setDate(CalBeginDate.getDate() + 30);
                }
                CalBeginDate.setMonth(CalBeginDate.getMonth());
                CalBeginDate.setFullYear(CalBeginDate.getFullYear());
            } else {

                if ($("#interval1").is(':checked') == true) {
                    CalBeginDate.setDate(CalBeginDate.getDate() + (7));
                } else {
                    CalBeginDate.setDate(CalBeginDate.getDate() + (30));
                }

                CalBeginDate.setMonth(CalBeginDate.getMonth());
                CalBeginDate.setFullYear(CalBeginDate.getFullYear());
            }




            midMarkerDate.setDate(CalBeginDate.getDate() + 1);
            midMarkerDate.setMonth(CalBeginDate.getMonth());
            midMarkerDate.setFullYear(CalBeginDate.getFullYear());




            midMarkerDateArr[i] = midMarkerDate




            //if (i != 0) {
            //if (i == 0) {
            //alert(midMarkerDateArr[i] + "//" + i + "//" + daysInFirstWeek)
            //}

            if (CalBeginDate.getDate() <= 1) {

            } else {

                $("#JobGraph").prepend("<div title='d.d.' id='midMarker_" + i + "' class='midMarker' style='width:1px;'></div>");
                positionMarker($("#midMarker_" + i), { "startdate": midMarkerDateArr[i], "absmarkcorr": "mid" });

            }
            //}



        }

    



	             var today = new Date();
	             //today.setDate(today.getDate() - 1);
	         
	//check if todays date is in the dateinterval
	if (BeginDate.getTime() < today.getTime() && today.getTime() < EndDate.getTime()) {

	    $("#JobGraph").prepend("<div title='d.d.' id='todayMarker' style='width:1px;'></div>");
	   
		//position todaymarker
		//alert(today)
        positionMarker($("#todayMarker"), { "startdate": today, "absmarkcorr": "today" });
        $("#JobGraphHeader").append("<li class='infoField' style='border-right:0;'>&nbsp;Jobinfo:</li>");
        }





  
     
            

} // populate graphframe





































    /*

    //PopulateGraphFrame();
    //AddJobs();
    //$("#todayMarker").height($("#JobGraph").height());

    //$("#tilLabel").hide();
    //$("#FM_enddate").css("display", "none");

    //$(".activities").css('visibility', 'visible', 'display', '');

    
    $("#JobGraph .activity").css("height", "500px");
    $(".activities").css("display", "");
    $(".activities").css("visibility", "visible");
    $(".activities").show("fast");


        $("#JobGraph .activity").css("display", "");
    $("#JobGraph .activity").css("visibility", "visible");
    $("#JobGraph .activity").css("height", "500px");
    $("#JobGraph .activity").show("fast");

   

      */






  
    //$(".activities").css("display", "");
    //$(".activities").css("visibility", "visible");
    //$(".activities").show("fast");

    //$(".activities").slideDown(200)

      //$(".job").each(function () {
	  //      var job = $(this);

            

	        //alert(this)

	        //job.children(".activities").show('slow', function () {
	           // //job.children(".activities").slideDown(200)
	        //});






    //});

  


   



