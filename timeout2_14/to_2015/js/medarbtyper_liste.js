$(document).ready(function () {


    // Setup - add a text input to each header cell
    $('#example thead th').not(":eq(0), :eq(1),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#example thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });


    // DataTable
    lto = $('#lto').val();

    //alert(lto)

    if (lto == 'tec' || lto == 'esn') {
        var table = $('#example').DataTable({
            "iDisplayLength": 50,
            "order": [2, "asc"]
        });

    } else {
        var table = $('#example').DataTable({
            "order": [2, "asc"]
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