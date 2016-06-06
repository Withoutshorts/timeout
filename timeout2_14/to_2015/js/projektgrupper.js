$(document).ready(function () {


    // Setup - add a text input to each header cell
    $('#XXtb_progrp_ikkemed thead th').not(":eq(1),:eq(2),:eq(3),:eq(4)").each(function () {

        var title = $('#XXtb_progrp_ikkemed thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#XXtb_progrp_ikkemed').DataTable({
    });




    // Setup - add a text input to each header cell
    $('#progrp_list thead th').not(":eq(0),:eq(2),:eq(3),:eq(4),:eq(5)").each(function () {

        var title = $('#progrp_list thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    lto = $('#lto').val();

    //alert(lto)

    if (lto == 'tec' || lto == 'esn') {
        var table = $('#progrp_list').DataTable({
            "iDisplayLength": 50,
            "order": [1, "asc"]
        });

    } else {
        var table = $('#progrp_list').DataTable({
            "order": [1, "asc"]
        });
    }

        // Apply the search
        table.columns().every(function () {

            var coulmn = this;

            $('input', this.header()).on('keyup change', function () {
                if (coulmn.search() !== this.value) {
                    coulmn
                        .search(this.value)
                        .draw();
                }
            });
    });


    
    

});