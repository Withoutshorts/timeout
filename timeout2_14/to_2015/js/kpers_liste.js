$(document).ready(function () {


    


    
    // Setup - add a text input to each header cell
    $('#tb_kontaktsliste thead th').not(":eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#tb_kontaktsliste thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#tb_kontaktsliste').DataTable({


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

});
