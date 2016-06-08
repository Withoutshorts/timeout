$(document).ready(function() {
    
    

	$.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));
	$("#FM_forfaldsdato").datepicker({ buttonImage: '../ill/popupcal.gif' })

   .change(function() {
   	$("#FM_betbetint").val("-1");
    });

   

   $("#FM_betbetint, #FM_istdato, #FM_istdato2").change(function() { GetBetIntVal(); });
   $("[name=FM_brugfakdatolabel]").click(function() {
	    GetBetIntVal();
	});

	function daydiff(first, second) {
		return (second.getTime() - first.getTime()) / (1000 * 60 * 60 * 24);
	}
	function daysInMonth(month, year) {
		var dd = new Date(year, month, 0);
		return dd.getDate();
	}



	function GetBetIntVal() {
		var numdays;
		var intval = $("#FM_betbetint").val();
		(intval == "0") ? $("#FM_forfaldsdato").attr("disabled",true) : $("#FM_forfaldsdato").removeAttr("disabled");

		
		if (intval != "-1") {
			var concatdate = $("#FM_start_dag").val() + "." + $("#FM_start_mrd").val() + "." + $("#FM_start_aar").val();
			var orgdate;
			$("[name=FM_brugfakdatolabel]").attr("checked") ? orgdate = $.datepicker.parseDate('d.m.yy', $("#FM_istdato2").val()) : orgdate = $.datepicker.parseDate('d.m.yy', concatdate)
			var newdate;
			if (intval == "1") numdays = 8;
			if (intval == "2") numdays = 14;
			if (intval == "5") numdays = 21;
			if (intval == "3") numdays = 30;
			if (intval == "6") numdays = 45;
			if (intval == "4") numdays = (daysInMonth(orgdate.getMonth(), orgdate.getYear()) - orgdate.getDate()) + 15;
			if (intval == "7") numdays = (daysInMonth(orgdate.getMonth(), orgdate.getYear()) - orgdate.getDate()) + 45;
			if (intval == "8") numdays = (daysInMonth(orgdate.getMonth(), orgdate.getYear()) - orgdate.getDate()) + 30;
			
			orgdate.setDate(orgdate.getDate() + numdays);
			while (orgdate.getDay() == 6 || orgdate.getDay() == 7) {
				orgdate.setDate(orgdate.getDate() + 1);
			}

			alert("ger:" + intval + "#" + numdays)
			
			$("#FM_forfaldsdato").datepicker('setDate', orgdate);
		}

	}

	GetBetIntVal();

});