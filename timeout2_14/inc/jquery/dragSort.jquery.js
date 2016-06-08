jq1101(function () {
    var defaults = {
        sortControl: ".SortOrder" //SortOrder
    }
    //Lei edited from td:first > :input[name=rowId] to input:[name=rowId]


    $(".drag").mouseover(function () {
        //alert("her")
        $(this).css('cursor', 'pointer');


        //startPos = $(this).position();
       
        //var x = startPos.left;
        //var y = startPos.top;

        //alert(x +" "+ y)
      
        gtThan = 0

        if ($("#thisfile").val() == "erp_fak.asp") {

            var thisidx = this.id
            var idlngt = thisidx.length
            var thisidx = thisidx.slice(5, idlngt)

            //alert($("#tblname_" + thisidx).val())

            thisTableID = "#" + $("#tblname_" + thisidx).val() + " tbody"
        } else {
            thisTableID = "#incidentlist tbody"
        }

        //alert(thisTableID)

        var opts = $.extend(defaults, { items: "tr:gt("+gtThan+"):has(:input[name=rowId])", IdControlNode: "input:[name=rowId]" });
        //var opts = $.extend(defaults, { items: "tr:gt(1):has(:input.rowId)", IdControlNode: "input.rowId" });
        var sortitems = opts.items;

      
        jq1101(thisTableID).sortable({
        items: sortitems,
        sort: function (event, ui) {
            var sortValueStart = $(this).children(sortitems + ":first").find(opts.sortControl).val();

            
            TONotifie("Du kan nu sortere i listen, for at gemme din sortering skal du klikke på knappen nedenunder<br /><input type='button' value='Gem sortering' id='saveSortOrderBtn'/>", false);
            var saveBtn = $("#saveSortOrderBtn");
            saveBtn.click(function () {
                //Lei edited from this to #incidentlist tbody
                $(thisTableID).children(sortitems).each(function (index) {
                    //Lei edited from this to #incidentlist tbody, .val() to [index].value
                    var IdControlVal = $(thisTableID).find(opts.IdControlNode)[index].value;


                    
                    if ($("#altsort_" + IdControlVal).val() == "1") {
                        useControl = "FM_sortOrder_alt"
                        lt = "F"
                    } else {
                        useControl = "FM_sortOrder"
                        lt = "A"
                    }
                    
                    sortVal = parseInt(sortValueStart) + index
                  
                  

                    if ($("#thisfile").val() == "erp_fak.asp") {
                       
                        //alert(IdControlVal + "#" + sortVal + "#" + index)

                        var thisid = IdControlVal
                        var idlngt = thisid.length
                        var IdControlVal = thisid.slice(1, idlngt)

                        $("#aktsort_" + IdControlVal + "").val(sortVal)

                        if ($("#altsort_" + lt + "" + IdControlVal).val() == "1") {
                            $("#sort_F" + IdControlVal + "").html(sortVal)
                            //alert(IdControlVal +"#"+ sortVal + "#" + index)
                        } else {
                            $("#sort_A" + IdControlVal + "").html(sortVal)
                        }






                    } else {

                        //alert(IdControlVal + " sortVal:" + sortVal)







                        $.ajax({
                            url: "?",
                            dataType: "text",
                            type: "POST",
                            data: { AjaxUpdateField: "true", control: useControl, value: sortVal, id: IdControlVal },
                            error: function (err) {
                                //alert("sort table error: " + err.toString());



                            },
                            success: function (data) {
                                //alert("success!"+ data);
                                //$(".drag").disableSelection();
                                //if ($(this).hasClass("cancel")) {
                                    //return false;
                                //}


                            }



                        });

                    }

                    $("#TONotifyWindow").dialog("close");
                });
            });
       
            

         
   
            
            
            
        }

      }); //.disableSelection()




















    //DO NOT COPY THE FOLLOWING TEST CODE TO PRODUCTION
//    $.fn.AjaxUpdateField = function (options) {
//        var defaults = {
//            file: "?",
//            parameter: "",
//            parent: "",
//            subselector: ""
//        };
//        var opts = $.extend(defaults, options);
//        var file = opts.file;
//        var parent = opts.parent;
//        var subselector = opts.subselector;
//        var IdControlNode;
//        $(this).change(function () {
//            IdControlNode = (parent != "") ? $(this).closest(parent).find(subselector) : $(this).find(subselector);
//            var senddata = { AjaxUpdateField: "true", control: $(this).attr("name"), value: $(this).val(), id: IdControlNode.val() };
//            $.post(file, senddata, function (data) { TONotifie(data, true); });
//        });
//    };
//    function TONotifie(msg, autoclose) {
//        var NotifyWindow;
//        if ($("#TONotifyWindow").length > 0) {
//            NotifyWindow = $("#TONotifyWindow");
//        }
//        else {
//            $("body").append("<div id='TONotifyWindow'></div>");
//            NotifyWindow = $("#TONotifyWindow");
//            NotifyWindow.dialog({ title: 'Timeout Notifier', resizable: false, height: 200, draggable: false, width: 200, position: ['right', 'bottom'], autoOpen: false })

//            var Dialog = NotifyWindow.parent();
//            $(window).scroll(function () {
//                Dialog.css({
//                    top: $(document).scrollTop() + $(window).height() - Dialog.height() - 8
//                });
//            });
//        }
//        //change contents of Notify window
//        NotifyWindow.html(msg);
//        if (!NotifyWindow.dialog("isOpen")) {
//            NotifyWindow.dialog("open");
//        }

//        //Make dialog close in 4 seconds, and then remove it
//        if (autoclose == true) {
//            setTimeout(function () { NotifyWindow.dialog("close"); }, 4000);
//        }
//    };
    //END OF TEST CODE

    });

    
  
});

