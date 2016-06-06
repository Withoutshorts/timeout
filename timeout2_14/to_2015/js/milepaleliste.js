$(document).ready(function () {


    // Setup - add a text input to each header cell
    $('#milepaleliste thead th').not(":eq(0),:eq(2),:eq(3)").each(function () {

        var title = $('#milepaleliste thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Search ' + title + '" />');
    });

    // DataTable
    var table = $('#milepaleliste').DataTable({


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