$(document).ready(function () {

  
   

    // Setup - add a text input to each header cell
    $('#medarb_list thead th').not(":eq(1),:eq(2),:eq(3),:eq(5),:eq(6),:eq(7),:eq(8)").each(function () {

        var title = $('#medarb_list thead th').eq($(this).index()).text();
        $(this).html('<input type="text" placeholder="Søg ' + title + '" />');
    });

    // DataTable
    var table = $('#medarb_list').DataTable({
        
        "aLengthMenu": [10, 25, 50, 75, 100, 500, 1000],
        "iDisplayLength": 100,
        "order": [0, "asc"]
    
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


