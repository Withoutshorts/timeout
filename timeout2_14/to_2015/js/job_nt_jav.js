// JScript File
$(window).load(function () {
    // run code
    //$("#loadbar").hide(1000);
});







$(document).ready(function () {

    //$(".portlet-body").css("left", "2");


    
    

    $('.date').datepicker({

    });


    $("#td_listtotals").html($("#listtotals").val());
 

    rapporttype = $("#rapporttype").val()

    //alert(rapporttype)
    if (rapporttype == 0) { // Orverview
        $("#tb_jobliste").DataTable(
            {
                "iDisplayLength": 25,
                "columnDefs": [
               { "type": "date", "targets": 6 },
               { "type": "date", "targets": 8 },
               { "type": "num-fmt", "targets": 10 },
               { "type": "num-fmt", "targets": 11 }
                ]
            }
        );
        
    }


    if (rapporttype == 1) { // Enquery 
        $("#tb_jobliste").DataTable(
            {
                "iDisplayLength": 25,
                "columnDefs":[
                { "type": "date", "targets": 6 },
                { "type": "date", "targets": 7 },
                { "type": "date", "targets": 8 },
                { "type": "date", "targets": 9 },
                { "type": "date", "targets": 10 },
                { "type": "date", "targets": 11 },
                { "type": "date", "targets": 12 },
                { "type": "date", "targets": 13 },
                { "type": "date", "targets": 14 }
                ]
            }
        );
    }


    if (rapporttype == 3) { // OverView Extended 
        $("#tb_jobliste").DataTable(
            {
                "iDisplayLength": 25,
                "columnDefs": [
                    { "type": "date", "targets": 6 },
                    { "type": "date", "targets": 8 },
                    { "type": "date", "targets": 9 },
                    { "type": "date", "targets": 10 },
                { "type": "num-fmt", "targets": 12 },
                { "type": "num-fmt", "targets": 13 },
                { "type": "num-fmt", "targets": 14 },
                { "type": "num-fmt", "targets": 15 },
                { "type": "num-fmt", "targets": 16 },
                { "type": "num-fmt", "targets": 17 },
                { "type": "num-fmt", "targets": 18 },
                { "type": "num-fmt", "targets": 19 },
                { "type": "num-fmt", "targets": 20 }
                ]
            }
        );
    }


    //  "fixedHeader": true,
    //"columnDefs": [{ "orderable": false, "targets": 0 }]

// "aLengthMenu": [10, 25, 50, 75, 100],
//"iDisplayLength": 25,
  

    //"columnDefs": [
    //       { "type": "Date", "targets": 8 }
    //]

    //"sPaginationType": "full_numbers"


//{ type: 'date-de', targets: 7 }
    //{

    //    "aLengthMenu": [10, 25, 50, 75, 100],
    //    "iDisplayLength": 25,
    //    "sPaginationType": "full_numbers",
          
    //    }

    //"sDom": '<"row view-filter"<"col-sm-12"<"pull-left"l><"pull-right"f><"clearfix">>>t<"row view-pager"<"col-sm-12"<"text-center"ip>>>'
  
    
        
    //$('.wrapper').css("left", "20px")



 



    if ($("#fastpris").val() == "2") { //commision
        //$("#cc").attr("disabled", "")
        $("#cc").removeAttr("disabled")
    };

    

    $("#fastpris").change(function () {

     
       
        if ($("#fastpris").val() == "2") { //commision
            $("#cc").removeAttr("disabled")
        } else {
            $("#cc").attr("disabled", "disabled")
        }

        salestypefn();

    });


    $("#cc").click(function () {
        
        if ($("#cc").is(':checked') == true) {

            $("#sales_price_pc_label").css("visibility", "hidden")
            $("#sales_price_pc_label").css("display", "none")

            $("#sales_price_pc").css("visibility", "visible")
            $("#sales_price_pc").css("display", "")

            $("#comm_pc_label").css("visibility", "visible")
            $("#comm_pc_label").css("display", "")

            $("#comm_pc").css("visibility", "hidden")
            $("#comm_pc").css("display", "none")

        } else {


            $("#sales_price_pc_label").css("visibility", "visible")
            $("#sales_price_pc_label").css("display", "")

            $("#sales_price_pc").css("visibility", "hidden")
            $("#sales_price_pc").css("display", "none")


            $("#comm_pc_label").css("visibility", "hidden")
            $("#comm_pc_label").css("display", "none")

            $("#comm_pc").css("visibility", "visible")
            $("#comm_pc").css("display", "")


        }
       

    });

   




    function salestypefn() {

        //alert($("#fastpris").val())

        if ($("#fastpris").val() == "2") { //commision

            $("#sales_price_pc").css("visibility", "hidden")
            $("#sales_price_pc").css("display", "none")

            $("#sales_price_pc_label").css("visibility", "visible")
            $("#sales_price_pc_label").css("display", "")


            $("#comm_pc_label").css("visibility", "hidden")
            $("#comm_pc_label").css("display", "none")

            $("#comm_pc").css("visibility", "visible")
            $("#comm_pc").css("display", "")



            $("#freight_pc").css("visibility", "hidden")
            $("#freight_pc").css("display", "none")
            
            $("#freight_pc_label").css("visibility", "visible")
            $("#freight_pc_label").css("display", "")

            $("#tax_pc").css("visibility", "hidden")
            $("#tax_pc").css("display", "none")

            $("#tax_pc_label").css("visibility", "visible")
            $("#tax_pc_label").css("display", "")

            $("#sales_price_pc").val(("#cost_price_pc").val())
            $("#sales_price_pc_label").val(("#cost_price_pc").val())

            $("#freight_pc").val(0)
            $("#tax_pc").val(0)

        } else {

            $("#sales_price_pc_label").css("visibility", "hidden")
            $("#sales_price_pc_label").css("display", "none")

            $("#sales_price_pc").css("visibility", "visible")
            $("#sales_price_pc").css("display", "")


            $("#comm_pc_label").css("visibility", "hidden")
            $("#comm_pc_label").css("display", "none")

            $("#comm_pc").css("visibility", "visible")
            $("#comm_pc").css("display", "")



            $("#freight_pc").css("visibility", "visible")
            $("#freight_pc").css("display", "")

            $("#freight_pc_label").css("visibility", "hidden")
            $("#freight_pc_label").css("display", "none")

            $("#tax_pc").css("visibility", "visible")
            $("#tax_pc").css("display", "")

            $("#tax_pc_label").css("visibility", "hidden")
            $("#tax_pc_label").css("display", "none")

        }
    };



    $("#FM_valuta_cost_price_pc_valuta").change(function () {
  
        newVal = $("#FM_valuta_cost_price_pc_valuta").val()
        $("#FM_valuta_freight_price_pc_valuta").val(newVal)
    });



     
    $("#FM_valuta_sales_price_pc_valuta, #FM_valuta_cost_price_pc_valuta").change(function () {

        beregn_order()

    });


    $("#FM_kunde").change(function () {


        
        getBetLev();

    });




    $("#shippedqty").keyup(function () {

        if (window.event.keyCode != '9') {

            illchartjk(this.id);

            if ($("#shippedqty").val() > 0) {

                $("#status").val(2)

            }
        }

    });
    

   
    $("#cost_price_pc, #sales_price_pc").keyup(function () {

        //alert("Hej NT - dette er en test, vent et øjeblik")

        if (window.event.keyCode != '9') {

            illchartjk(this.id);

            fastpris = $("#fastpris").val() //2: commision 3:salesorder

            //alert(fastpris)

            if (fastpris == "2") {
          
                if ($("#cc").is(':checked') == true) {

                    cost_price_pc = $("#cost_price_pc").val().replace(",", ".")
                    sales_price_pc = $("#sales_price_pc").val().replace(",", ".")

                } else {

                    cost_price_pc = $("#cost_price_pc").val().replace(",", ".")
                    sales_price_pc = $("#cost_price_pc").val().replace(",", ".")

                }

                
            } else {
                cost_price_pc = $("#cost_price_pc").val().replace(",", ".")
                sales_price_pc = $("#sales_price_pc").val().replace(",", ".")
            }

            //alert("fastpris: " + fastpris + " kost: " + cost_price_pc)

            if (fastpris == "2") { //comm

                if ($("#cc").is(':checked') == true) {

                } else {

                $("#sales_price_pc").val(sales_price_pc.replace(".", ","))
                $("#sales_price_pc_label").val(sales_price_pc.replace(".", ","))

                }

                comm_pc = 100 - ((cost_price_pc / sales_price_pc) * 100)
                comm_pc = Math.round(comm_pc * 100) / 100
                comm_pc = String(comm_pc).replace(".", ",")

                $("#comm_pc").val(comm_pc)
                $("#comm_pc_label").val(comm_pc)

                if ($("#comm_pc").val() == "NaN" || ($("#comm_pc").val() == "-Infinity")) {
                    $("#comm_pc").val(0)
                    $("#comm_pc_label").val(0)
                }


            }

            if (this.id == "cost_price_pc") {
                $("#cost_price_pc_base").val($("#cost_price_pc").val().replace(".", ","))
                $("#cost_price_pc_label").text($("#cost_price_pc").val().replace(".", ","))

            }



            if ($("#cc").is(':checked') == true) {

            } else {
            cost_price_pc = cost_price_pc.replace(".", ",")
            $("#cost_price_pc").val(cost_price_pc)

            sales_price_pc = sales_price_pc.replace(".", ",")
            $("#sales_price_pc").val(sales_price_pc)
            }


            beregn_order()

        } 

    });




    $("#comm_pc").keyup(function () {


        if (window.event.keyCode != '9') {

        illchartjk(this.id);

        fastpris = $("#fastpris").val() //2: commision 3:salesorder

        
        if (fastpris == 2) {


        sales_price_pc = $("#sales_price_pc").val().replace(",", ".")
        comm_pc = $("#comm_pc").val().replace(",", ".")

        if (comm_pc < 100) {
            cost_price_pc = sales_price_pc - ((comm_pc * sales_price_pc) / 100)
        } else {
            cost_price_pc = 0
        }
        cost_price_pc = Math.round(cost_price_pc * 10000) / 10000
        cost_price_pc = String(cost_price_pc).replace(".", ",")

        $("#cost_price_pc").val(cost_price_pc)

     
        }


        if (fastpris == 3) {


            cost_price_pc_opr = $("#cost_price_pc_base").val().replace(",", ".")
            comm_pc = $("#comm_pc").val().replace(",", ".")

            if (comm_pc < 100) {
                cost_price_pc = cost_price_pc_opr - ((comm_pc * cost_price_pc_opr) / 100)
            } else {
                cost_price_pc = 0
            }
            cost_price_pc = Math.round(cost_price_pc * 100) / 100
            cost_price_pc = String(cost_price_pc).replace(".", ",")

            $("#cost_price_pc").val(cost_price_pc)


        }


       
        beregn_order()

        } //keycode

    });

    
   // $("#orderqty, #sales_price_pc, #freight_pc, #tax_pc, #cost_price_pc").keyup(function () {


        function illchartjk(id){
            var str = ""
            str = $("#" + id).val()
            //alert(str)

            passedVal = str

            if (id == "orderqty") {
                invalidChars = "/+:;<>abcdefghijklmnopqrstuvwxyzæøå,."
            } else {
                invalidChars = "/+:;<>abcdefghijklmnopqrstuvwxyzæøå"
            }

            if (passedVal == "") {
                return false
            }

            for (i = 0; i < invalidChars.length; i++) {
                badChar = invalidChars.charAt(i)
                if (passedVal.indexOf(badChar, 0) != -1) {
                    alert("You have used an illegal character!")

                    str = str.substring(0, str.length - 1);
                    $("#" + id).val(str)

                    return false
                }
            }

            /*public String method(String str) {
            if (str.length > 0 && (str.charAt(str.length - 1) == ',' || str.charAt(str.length - 1) == '.')) {
                str = str.substring(0, str.length - 1);
                $("#" + this.id).val(str)
                
            } */

            //beregn_order()

        };



    $("#orderqty, #freight_pc, #tax_pc").keyup(function () {

        if (window.event.keyCode != '9') {

            illchartjk(this.id);

            beregn_order()

        }

    });



    $("#bulk_jobid_action").change(function () {

        //alert("her")

        action = $("#bulk_jobid_action").val()

      
        if (action == '1') {


            bulk_jobids = "0"

            var values = $('input:checkbox:checked.bulk_jobid').map(function () {
                bulk_jobids = bulk_jobids + "," + this.value
            }).get();

            $("#bulk_jobids").val(bulk_jobids)

            $("#dv_bulk").css("visibility", "visible");
            $("#dv_bulk").css("display", "");
            $("#dv_bulk").show(1000)

        } 


        if (action == '2') {

           
            //alert($("#fakhref").val())

            fakhref = ""// $("#fakhref").val()
            jobids = "0" 

            lp = 0;

            var values = $('input:checkbox:checked.bulk_jobid').map(function () {
                //alert(this.value)
                if (lp == 0) {
                    fakhref = $("#fakhref_" + this.value).val();
                    //alert($("#fakhref_" + this.value).val())
                }
                jobids = jobids + "," + this.value;
                lp = lp + 1;
            }).get();

            fakhref = fakhref + "&jobids=" + jobids

            //alert(fakhref)
         
            $("#ainvlink").attr('href', fakhref);

            $("#dv_invoice").css("visibility", "visible");
            $("#dv_invoice").css("display", "");
            $("#dv_invoice").show(1000)
           

        } 

        if (action == '0') {

            $("#dv_bulk").hide(300)
            $("#dv_invoice").hide(300)

        } 
                

    });
    

    $("#ainvlink").click(function () {
    
        $("#dv_invoice").hide(1000)

    });


    
    $("#to_top").click(function () {


        //alert($("#showfullscreen").val())

        if ($("#showfullscreen").val() == 0) {

        $("#table_header").css("top", "100px")
        $("#table_body").css("top", "186px")
        $("#table_body").css("height", "600px")

        $("#search").hide("fast");
        $("#dv_grandtotal").hide("fast");
        $("#print").hide("fast");
        
        

        $("#showfullscreen").val('1')

        } else {
        
            $("#table_header").css("top", "458px")
            $("#table_body").css("top", "544px")
            $("#table_body").css("height", "300px")

            $("#showfullscreen").val('0')

            $("#search").show("fast");
            $("#dv_grandtotal").show("fast");
            $("#print").show("fast");
            
            
        
        }

        //$("#search").hide("fast");
        //$.scrollTo($('#table_header'), 2000, { offset: -300 });
        //$.scrollTo('-=100px', 1500);

    });


    $("#to_top").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#bulk_jobid").click(function () {

    

        if ($("#bulk_jobid").is(':checked') == true) {
            $(".bulk_jobid").attr("checked", "checked");

           

        } else {

            $(".bulk_jobid").removeAttr("checked");
            
        }
       

    });


    
    $("#sp_dv_invoice").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#sp_dv_invoice").click(function () {
        $("#bulk_jobid").removeAttr("checked");
        $(".bulk_jobid").removeAttr("checked");
        
        $("#dv_invoice").hide(300)
        $("#bulk_jobid_action").val(0)
    });

    
    $("#bulk_close").mouseover(function () {
        $(this).css('cursor', 'pointer');
    });

    $("#bulk_close").click(function () {
        $("#bulk_jobid").removeAttr("checked");
        $(".bulk_jobid").removeAttr("checked");
        $("#dv_bulk").hide(300)

        $("#bulk_jobid_action").val(0)
    });




    $("#FM_kunde").change(function () {

       

        getAfd();
        //getKpers();

    });


    $("#FM_origin").change(function () {

        getSup();
     
    });





    // BEt og LEv betingelser //
    function getBetLev() {


        var kid_val = $("#FM_kunde").val()

       
        $.post("?jq_kid=" + kid_val, { control: "FN_betlev", AjaxUpdateField: "true", cust: 0 }, function (data) {
            //$("#FM_modtageradr").val(data);

           
            $("#FM_t5").val(data);

            $("#FM_betbetint").val($("#FM_t5").val());
            //alert(data)

        });


    }



    // Afdeling //
    function getAfd() {
     

        

        var kid_val = $("#FM_kunde").val()

       

        $.post("?jq_kid=" + kid_val, { control: "FN_afd", AjaxUpdateField: "true", cust: 0 }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#FM_afd").html(data);
            //$("#jobid").html(data);
            //alert(data)

        });


    }


    // Supplier //
    function getSup() {


        var origin_val = $("#FM_origin").val()

        //alert("her")

        //alert(func)

        $.post("?jq_origin=" + origin_val, { control: "FN_sup", AjaxUpdateField: "true", cust: 0 }, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#FM_supplier").html(data);
            //$("#jobid").html(data);
            //alert(data)

        });


    }




    // kontakpers //
    function xgetKpers() {
     

        var kid_val = $("#FM_kunde").val()

    

        $.post("?jq_kid=" + kid_val, { control: "FN_kpers", AjaxUpdateField: "true", cust: 0}, function (data) {
            //$("#FM_modtageradr").val(data);
            $("#FM_kpers").html(data);
            //$("#jobid").html(data);
            //alert(data)

        });


    }









function beregn_order() {


    fastpris = $("#fastpris").val() //2: commision 3:salesorder

    //alert(fastpris)

    //if (fastpris == 2) {

        orderqty = $("#orderqty").val().replace(",", ".")

        sales_price_pc = $("#sales_price_pc").val().replace(",", ".")
        cost_price_pc = $("#cost_price_pc").val().replace(",", ".")

        cost_price_pc_id = $("#FM_valuta_cost_price_pc_valuta").val()
        valuta_kurs_c = $("#valuta_kurs_" + cost_price_pc_id).val().replace(",", ".")
        cost_price_pc = (cost_price_pc / 1 * valuta_kurs_c / 1) / 100
        cost_price_pc = (orderqty * cost_price_pc)

        //alert(valuta_kurs_c)

        sales_price_pc_id = $("#FM_valuta_sales_price_pc_valuta").val()
        valuta_kurs_s = $("#valuta_kurs_"+ sales_price_pc_id).val().replace(",", ".")
        sales_price_pc = (sales_price_pc / 1 * valuta_kurs_s / 1) / 100

        //alert(sales_price_pc)

        tax = $("#tax_pc").val().replace(",", ".") / 1
        costprice_tax = 0
        if (tax != 0) {
            costprice_tax = ((cost_price_pc * (tax / 1 + 100 / 1)) / 100)
        } else {
            costprice_tax = (cost_price_pc * 1)
        }



        freight = $("#freight_pc").val().replace(",", ".") / 1
        freight = Math.round((freight) * 100) / 100
        freight = ((orderqty * freight) * (valuta_kurs_c / 100))

       
        
        
        
        //alert(tax)
        //tax = Math.round(((orderqty * tax / 100)) * 100) / 100

        totalcostprice = 0
        totalcostprice = ((costprice_tax + freight / 1))
        //alert(cpvatfre)


        totalsalesprice = Math.round(((orderqty * sales_price_pc)) * 100) / 100
        totalcostprice = Math.round((totalcostprice) * 100) / 100
        //totalsalesprice = toString(totalsalesprice)

        jo_dbproc = Math.round((totalcostprice / totalsalesprice)* 100) / 100
        jo_dbproc = 100 - ((jo_dbproc) * 100)
        jo_dbproc = Math.round((jo_dbproc) * 100) / 100

        jo_dbproc_bel = (totalsalesprice - totalcostprice)
        jo_dbproc_bel = Math.round((jo_dbproc_bel) * 100) / 100

        totalcostprice = String(totalcostprice).replace(".", ",")
        totalsalesprice = String(totalsalesprice).replace(".", ",")
        jo_dbproc = String(jo_dbproc).replace(".", ",")
        //alert(totalsalesprice)

        $("#bruttooms").val(totalsalesprice)
        $("#bruttooms_label").val(totalsalesprice)

        $("#udgifter_intern").val(totalcostprice)
        $("#udgifter_intern_label").val(totalcostprice)

        $("#jo_dbproc").val(jo_dbproc)
        $("#jo_dbproc_label").val(jo_dbproc)

        $("#jo_dbproc_bel").val(jo_dbproc_bel)

        if ($("#bruttooms").val() == "NaN" || ($("#bruttooms").val() == "-Infinity")) {
            $("#bruttooms").val(0)
            $("#bruttooms_label").val(0)
        }

        if ($("#udgifter_intern").val() == "NaN" || ($("#udgifter_intern").val() == "-Infinity")) {
            $("#udgifter_intern").val(0)
            $("#udgifter_intern_label").val(0)
        }

     
        if ($("#jo_dbproc").val() == "NaN" || ($("#jo_dbproc").val() == "-Infinity")) {
            $("#jo_dbproc").val(0)
            $("#jo_dbproc_label").val(0)
        }

        if ($("#jo_dbproc_bel").val() == "NaN" || ($("#jo_dbproc_bel").val() == "-Infinity")) {
            $("#jo_dbproc_bel").val(0)
        }

        


    //}

}

    
    salestypefn();



    $('.panel-collapse:not(".in")')
       .collapse('show');


});