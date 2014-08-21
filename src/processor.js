function ProcessingDialog($el, options) {
	function init() {
		if ($el.find("button").length === 0) {
			var top = options.top,
				left = options.left;
			if (top == "center") {
				top = ($(window).height() - options.height) / 2;
			}
			if (left == "center") {
				left = ($(window).width() - options.width) / 2;
			}
			$el.empty().css({
				"display" : "none",
				"position" : "fixed",
				"top" : top,
				"left" : left,
				"width" : options.width,
				"height" : options.height,
				"border" : "solid 1px #000000",
				"background-color": "#F0F0F0",
				"text-align" : "center",
				"z-index" : 1000
			}).append(
				"<div style='margin:10px 0'><img src='" + options.baseUrl + "/assets/images/loading.gif'/></div>" +
				"<div style='margin:10px 0'><span>" + Status.Prepared.text + "...</span></div>" +
				"<div style='margin-right:10px;text-align:right'><button class='btn btn-warning'>" + MSG.Cancel + "</button></div>" + 
				"<form action='#' method='post'></form>"
			);
		}
		$msgSpan = $el.find("span");
		$form = $el.find("form");
		$cancelBtn = $el.find("button");
		$cancelBtn.unbind("click");
	}
	function show() {
		$el.show();
	}
	function close() {
		$el.hide();
	}
	function text(str) {
		$msgSpan.text(str);
	}
	function download(ticket) {
		$form.attr("action", makeUrl(options, "/finish/" + ticket)).submit();
	}
	var $msgSpan, $form, $cancelBtn;
	init();

	$.extend(this, {
		"show" : show,
		"close" : close,
		"text" : text,
		"download" : download,
		"cancelButton" : $cancelBtn
	});
}
function Processor(user, template, options) {
	function init() {
		var id = "excelReport-download-dialog",
			$el = $("#" + id);
		if ($el.length === 0) {
			$el = $("<div/>");
			$el.attr("id", id);
			$("body").append($el);
		}
		dialog = new ProcessingDialog($el, options);
		dialog.cancelButton.click(cancel);
	}
	function prepare(data, func) {
		ticket = null;
		canceled = false;
		callback = func;
		var params = {
			"url" : makeUrl(options, "/prepare/" + user + "/" + template),
			"type" : "POST",
			"data" : data,
			"success" : function(data) {
				var status = Status.fromName(data.status);
				if (status == Status.Error) {
					options.error(data.msg);
				} else if (status == Status.Prepared) {
					ticket = data.ticket;
					if (options.showLoading) {
						dialog.show();
					}
					setTimeout(checkStatus, 5000);
				} else {
					//not occur
					console.log("Unknown data", data);
					options.error(JSON.stringify(data));
				}
			},
			"error" : function(xhr, status, e) {
				options.error(status + ", " + e);
			}
		};
		if (options.apikey) {
			addAuthorization(params, options.apikey);
		}
		if (isFormData(data)) {
			params.processData = false;
			params.contentType = false;
		}
		$.ajax(params);
	}
	function checkStatus() {
		if (!ticket || canceled) {
			return;
		}
		$.ajax({
			"url" : makeUrl(options, "/status/" + ticket),
			"type" : "GET",
			"success" : function(data) {
				var status = Status.fromName(data.status);
				if (status == Status.Prepared || status == Status.Processing) {
					dialog.text(status.text + "...");
					setTimeout(checkStatus, 1000);
				} else if (status == Status.Finished) {
					dialog.close();
					if (isUpload()) {
						finishUpload();
					} else {
						dialog.download(ticket);
					}
				} else if (status == Status.Error) {
					dialog.close();
					options.error(data.msg);
				} else if (status == Status.Canceled) {
					dialog.close();
				} else {
					//not occur
					console.log("Unknown data", data);
					options.error(JSON.stringify(data));
				}
			},
			"error" : function(xhr, status, e) {
				options.error(status + ", " + e);
			}
		});
	}
	function cancel() {
		if (!ticket || canceled) {
			return;
		}
		canceled = true;
		$.ajax({
			"url" : makeUrl(options, "/cancel/" + ticket),
			"type" : "GET",
			"success" : function(data) {
				console.log("Cancel: " + data);
			},
			"error" : function(xhr, status, e) {
				options.error(status + ", " + e);
			}
		});
		dialog.close();
	}
	function finishUpload() {
		if (!ticket || canceled) {
			return;
		}
		$.ajax({
			"url" : makeUrl(options, "/finish/" + ticket),
			"type" : "GET",
			"success" : function(data) {
				if (callback) {
					callback(data);
				}
			},
			"error" : function(xhr, status, e) {
				options.error(status + ", " + e);
			}
		});
	}
	options = $.extend(defaults, options || {});
	var self = this,
		canceled = false,
		ticket = null,
		dialog = null,
		callback = null;
	init();
	$.extend(this, {
		"prepare" : prepare
	});
}
