$(document).ready(function () {

    $("#FM_chk_all_add").click(function () {

        //alert("her" + $("#FM_chk_all_add").is(':checked'))

        if ($("#FM_chk_all_add").is(':checked') == true) {
            $(".chk_medlem_clas").prop('checked', true);

        } else {
            $(".chk_medlem_clas").prop('checked', false);
        }
       

    });

    
  
    /*

    // Setup - add a text input to each header cell
    $('#xxtb_progrp_ikkemed thead th').not(":eq(1),:eq(2),:eq(3),:eq(4)").each(function () {

        var title = $('#xxtb_progrp_ikkemed thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('xx#tb_progrp_ikkemed').DataTable({
    });



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

    */
    
    

});