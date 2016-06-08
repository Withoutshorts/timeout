//Dato func jq //
function IsValidDate(input) {
    var dayfield = input.split(".")[0];
    var monthfield = input.split(".")[1];
    var yearfield = input.split(".")[2];
    var returnval;
    var dayobj = new Date(yearfield, monthfield - 1, dayfield)
    if ((dayobj.getMonth() + 1 != monthfield) || (dayobj.getDate() != dayfield) || (dayobj.getFullYear() != yearfield)) {
        alert("Forkert datoformat, skriv venligst dd.mm.yyyy eller d.m.yyyy");
        returnval = false;
    }
    else {
        returnval = true;
    }
    return returnval;
}



//Make datepickers
//$.datepicker.setDefaults($.extend({ showOn: 'button', constrainInput: true, showOtherMonths: true, showWeeks: true, minDate: new Date(2002, 1, 1), firstDay: 1, changeFirstDay: false, buttonImage: '../ill/popupcal.gif', start: 6, buttonImageOnly: true, dateFormat: 'd.m.yy', changeMonth: true, changeYear: true }));


alert("her")

//$("#FM_stdato").datepicker();
//$("#FM_stdato").change(function () { if (IsValidDate($("#FM_stdato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[id=FM_start_dag]").val(datesplit[0]); $("[id=FM_start_mrd]").val(datesplit[1]); $("[id=FM_start_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
$("#FM_stdato").css("display", "inline");

$("#FM_sldato").datepicker();
$("#FM_sldato").change(function () { if (IsValidDate($("#FM_sldato").val()) == true) { var datestring = $(this).val(); var datesplit = datestring.split('.'); $("[name=FM_slut_dag]").val(datesplit[0]); $("[name=FM_slut_mrd]").val(datesplit[1]); $("[name=FM_slut_aar]").val(datesplit[2]); } else { alert("dato ikke gemt"); } });
$("#FM_sldato").css("display", "inline");