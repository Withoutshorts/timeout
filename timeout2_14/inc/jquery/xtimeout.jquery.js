﻿//Jquery starts here
$(window).load(function(){
//global jquery edits to documents goes here
$("input[type=text]").focus(function(){$(this).addClass("focusField");});
$("input[type=text]").blur(function(){$(this).removeClass("focusField");});

});

$.fn.table_checkall = function() {
	var table = $(this);
	$("#" + table.attr("id") + " .checkAll").click(function() {

		var ChkBtn = $(this);
		var td = ChkBtn.parent("td");
		var tr = td.parent("tr");

		var tdIndex = tr.children("td").index(td);

		//find the tablecell in all the other rows
		table.children("tbody").children("tr").not(":first").each(function() {
			var selCell = $(this).children("td").eq(tdIndex);
			//if the cell contains checkboxes or radios check them
			selCell.children(":radio").click();
			selCell.children(":checkbox").click();
			//if there are no radios or checkboxes copy values from the input in the first td to the rest of the row
			if (selCell.children(":radio").attr("checked") == null && selCell.children(":checkbox").attr("checked") == null) {
				var ValueColumn = tr.next().children("td").eq(tdIndex);
				ValueColumn.children(":input").each(function(index) {
					selCellControl = selCell.children(":input").eq(index);
					selCellControl.attr("value", $(this).val());
				});
			}
		});
	});
};

$.fn.table_sort = function(options, callback) {
	var defaults = {
sortControl : ":input[name=SortOrder]"
}
var opts = $.extend(defaults, options);
var sortitems = opts.items;
$(this).sortable({
	items: sortitems,
	start: StartSort
});
};

function StartSort() {
	var sortValueStart = $(this).children(sortitems + ":first").find(opts.sortControl).val();
	
	/*function StopItem(e, ui) {
	console.log($(this).children(sortitems).index(ui.item) - ui.item.data("orgIndex"));
	console.log
	}*/

	TONotifie("Du kan nu sortere i listen, for at gemme din sortering skal du klikke på knappen nedenunder<br /><input type='button' value='Gem sortering' id='saveSortOrderBtn'/>", false);
	var saveBtn = $("#saveSortOrderBtn");
	saveBtn.click(function() {
	$(this).children(sortitems).each(function(index) {
	var IdControlVal = $(this).find(opts.IdControlNode).val();
				$.post("?", { AjaxUpdateField: "true", control: "FM_sortOrder", value: (parseInt(sortValueStart) + index), id: IdControlVal });
			});
		$("#TONotifyWindow").dialog("close");
	});
}




	var opts = $.extend(defaults, options);
	var file = opts.file;
	var parent = opts.parent;
	var subselector = opts.subselector;
	var IdControlNode;
	$(this).change(function() {
		IdControlNode = (parent != "") ? $(this).closest(parent).find(subselector) : $(this).find(subselector);
		var senddata = { AjaxUpdateField: "true", control: $(this).attr("name"), value: $(this).val(), id: IdControlNode.val() };
		$.post(file, senddata, function(data) { TONotifie(data, true); });
	});


	else {
		$("body").append("<div id='TONotifyWindow'></div>");
		NotifyWindow = $("#TONotifyWindow");
		NotifyWindow.dialog({ title: 'Timeout Notifier', resizable: false, height: 200, draggable: false, width: 200, position: ['right', 'bottom'], autoOpen: false })
	
		var Dialog = NotifyWindow.parent();











