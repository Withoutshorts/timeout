$(document).ready(function () {


    // Setup - add a text input to each header cell
    $('#kundetyper thead th').not(":eq(0),:eq(2),:eq(3),:eq(4),:eq(5),:eq(6),:eq(7),:eq(8),:eq(9)").each(function () {

        var title = $('#kundetyper thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Search ' + title + '" />');
    });


    // DataTable
    var table = $('#kundetyper').DataTable({

        "aLengthMenu": [10, 25, 50],
        "iDisplayLength": 10,
        "order": [1, "asc"]

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