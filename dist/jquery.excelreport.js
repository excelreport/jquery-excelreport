if (typeof(flect) == "undefined") flect = {};

(function ($) {
"use strict";

var 
	VT_ANY = 0,
	VT_INT = 1,
	VT_DECIMAL = 2,
	VT_LIST = 3,
	VT_DATE = 4,
	VT_TIME = 5,
	VT_TEXT_LENGTH = 6,
	VT_FORMULA = 7,
	MSG = {
		"Prepared" : "Preparing",
		"Processing" : "Processing",
		"Finished" : "Finished",
		"Error" : "Error",
		"Canceled" : "Canceled",
		"Cancel" : "Cancel"
	}, defaults = {
		"baseUrl" : "",
		"contextUrl" : "download",
		"showLoading" : true,
		"width" : 320,
		"height" : 120,
		"top" : "center",
		"left" : "center",
		"error" : function(msg) {
			alert(msg);
		},
		"onRule" : defaultOnRule
	};
if (flect.excelreport && flect.excelreport.MSG) {
	MSG = $.extend(MSG, flect.excelreport.MSG);
}

function Enum(values) {
	for (var i=0; i<values.length; i++) {
		var v = values[i];
		var text = MSG[v.name];
		if (!text) {
			text = v.name;
		}
		v.text = text;
		this[v.name] = v;
	}
	$.extend(this, {
		"fromName" : function(v) {
			for (var i=0; i<values.length; i++) {
				if (values[i].name == v) return values[i];
			}
			return null;
		}
	});
}
var Status = new Enum([
	{ "name" : "Prepared"},
	{ "name" : "Processing"},
	{ "name" : "Finished"},
	{ "name" : "Error"},
	{ "name" : "Canceled"},
]);

function isFormData(data) {
	if (!data) {
		return false;
	}
	var name = data.constructor.name;
	if (name) {
		return name === "FormData";
	}
	return !$.isPlainObject(data);
}
function isNumeric(str) {
	return $.isNumeric(str.replace(/,/g, ""));
}
function makeUrl(options, path) {
	return options.baseUrl + "/" + options.contextUrl + path;
}
function addAuthorization(params, apikey) {
	params.headers = {
		"X-Auth-Token": apikey
	};
}

function defaultOnRule(eventName, params) {
	switch (eventName) {
		case "prompt":
			break;
		case "error":
			break;
	}
}
function RuleManager(rules) {
	function getColStr(str) {
		return str.match(/[A-Z]+/)[0];
	}
	function isTargetCell(range, id) {
		var idx = range.indexOf(":");
		if (idx == -1) {
			return id == range;
		} else {
			var topLeft = range.substring(0, idx),
				bottomRight = range.substring(idx + 1),
				left = getColStr(topLeft),
				top = parseInt(topLeft.substring(left.length)),
				right = getColStr(bottomRight),
				bottom = parseInt(bottomRight.substring(right.length)),
				col = getColStr(id),
				row = parseInt(id.substring(col.length));
			return col >= left && col <= right && row >= top && row <= bottom;
		}
	}
	function getRule(name, id) {
		var target = null,
			ret = null;
		$.each(rules, function(key, value) {
			if (key == name) {
				target = value;
				return false;
			}
			return true;
		});
		if (target) {
			$.each(target, function(idx, rule) {
				for (var i=0; i<rule.regions.length; i++) {
					if (isTargetCell(rule.regions[i], id)) {
						ret = rule;
						return false;
					}
				}
				return true;
			});
		}
		return ret;
	}
	function buildInput(name, id) {
		var rule = getRule(name, id);
		if (rule && rule.vt === VT_LIST) {
			var $sel = $("<select><option value=''></option></select>");
			$.each(rule.list, function(idx, value) {
				var $op = $("<option/>");
				$op.attr("value", value);
				$op.text(value);
				$sel.append($op);
			});
			return $sel;
		}
		if (rule && isNumberRule(rule)) {
			var $input = $("<input type='text'/>");
			$input.css("text-align", "right");
			return $input;
		}
		return null;
	}
	function isNumberRule(rule) {
		return rule.vt == VT_INT || rule.vt == VT_DECIMAL;
	}
	$.extend(this, {
		"buildInput" : buildInput
	});
}

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
	function isUpload() {
		return options.contextUrl == "upload";
	}
	function prepare(data, func) {
		ticket = null;
		canceled = false;
		callback = func;
		var params = {
			"url" : makeUrl(options, "/prepare"),
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
		if (!isUpload()) {
			 params.url += "/" + user + "/" + template;
		}
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

flect.ExcelReport = function(baseUrl, user) {
	function isMultiNamedCell($div) {
		var name = $div.attr("data-name"),
			$prev = $div.prev("div:eq(0)"),
			$next = $div.next("div:eq(0)");
		return $prev.attr("data-name") == name || $next.attr("data-name") == name;
	}
	function createOptions(options, context) {
		var ret = options || {};
		ret.baseUrl = baseUrl;
		ret.contextUrl = context;
		return ret;
	}
	function buildForm($el, ruleMan, buildInput) {
		function getRealWidth($span) {
			var $temp = $("<span style='display:none;'/>"),
				ret = 0;
			$temp.text($span.text());
			$span.parent().append($temp);
			ret = $temp.width();
			$temp.remove();
			return ret;
		}
		$el.find(".namedCell").each(function() {
			var $div = $(this),
				$span = $div.find("span"),
				id = $div.attr("id"),
				name = $div.attr("data-name"),
				text = $span.text(),
				isFormula = !!$div.attr("data-formula");
			if (!isFormula) {
				var $input = buildInput ? buildInput($div) : null,
					h = $div.innerHeight(),
					w = $div.innerWidth();
				if ($input === null && ruleMan) {
					$input = ruleMan.buildInput(name, id);
				}
				if ($input === null) {
					if (h > 40) {
						$input = $("<textarea></textarea>");
					} else {
						$input = $("<input type='text'/>");
						if ($span.hasClass("cell-ar")) {
							$input.css("text-align", "right");
						} else if ($span.hasClass("cell-ac")) {
							$input.css("text-align", "center");
						} 
					}
				}
				if (typeof($input) != "string") {
					if (text) {
						if (isMultiNamedCell($div)) {
							$input.val(text);
							$span.remove();
						} else {
							var labelWidth = getRealWidth($span);
							w -= labelWidth + 4;
							$span.css("width", (labelWidth + 8) + "px");
							$input.css("margin-left", (labelWidth + 8) + "px");
						}
					} else {
						$span.remove();
					}
					$input.css({
						"width" : (w - 8) + "px",
						"height" : (h - 2) + "px"
					});
					$input.addClass("exrep-input");
					$input.attr("name", name);
				}
				$div.append($input);
			}
		});
	}
	function buildLiveForm($el, template, form, buildInput) {
		function insertLogo() {
			var $img = $("<img/>");
			$img.attr("src", baseUrl + "/assets/images/PoweredByExcelReport.png");
			$img.css({
				"position" : "absolute",
				"right" : 0,
				"bottom" : 0,
				"z-index" : 1000
			});
			$el.append($img);
		}
		function buildUrl() {
			var url = "";
			if (!baseUrl) {
				url = (location.protocol == "https:" ? "wss" : "ws") + "://" + location.host;
			} else if (baseUrl.indexOf("https:") === 0) {
				url = "wss" + baseUrl.substring(5);
			} else if (baseUrl.indexOf("http:") === 0) {
				url = "ws" + baseUrl.substring(4);
			} else if (baseUrl.indexOf("//") === 0) {
				url = (location.protocol == "https:" ? "wss" : "ws") + ":" + baseUrl;
			} else {
				defaults.error("Invalid baseUrl: " + baseUrl);
			}
			return url + "/liveform/" + user + "/" + template;
		}
		var url = buildUrl(),
			con = new room.Connection({
				"url" : url
			});
		con.on("calced", function(data) {
			$.each(data, function(key, value) {
				var $span = $("#" + key + " span");
				$span.empty().append(value);
				if (isNumeric(value) && !$span.hasClass("cell-ac")) {
					$span.removeClass("cell-al").addClass("cell-ar");
				}
			});
		});
		con.on("chart", function(data) {
			if ($.fn.excelToChart) {
				var chart = data.chart,
					$div = $el.find(".excel-chart").eq(data.index);
				$div.excelToChart(chart);
				$div.css("position", "absolute");
			}
		});
		con.onOpen(function(event) {
			con.request({
				"command" : "available",
				"data" : {
					"rules" : form
				},
				"success" : function(data) {
					if (data.error) {
						alert(data.error);
						con.close();
						$.removeData($el.get(0), "connection");
					} else {
						if (data.license == "Free") {
							insertLogo();
						}
						if (form) {
							var ruleMan = new RuleManager(data.rules);
							buildForm($el, ruleMan, buildInput);
						}
						$el.find(":input").change(function() {
							var $input = $(this),
								$div = $input.parent("div"),
								id = $div.attr("id"),
								value = $input.val();
							con.request({
								"command" : "modified",
								"data" : {
									"id" : id,
									"value" : value
								}
							});
						});
					}
				}
			});
		});
		con.sendNoop(30, !room.utils.isMobile());
		$.data($el.get(0), "connection", con);
	}
	function downloadExcel(template, data, options) {
		var processor = new Processor(user, template, createOptions(options, "download"));
		processor.prepare(data);
	}
	function downloadPdf(template, data, options) {
		options = createOptions(options, "pdf");
		var processor = new Processor(user, template, options);
		processor.prepare(data);
	}
	function uploadExcel(formData, options, callback) {
		switch (arguments.length) {
			case 0: 
				throw new Exception("upload file is required.");
			case 1:
				break;
			case 2:
				if ($.isFunction(arguments[1])) {
					callback = options;
					options = null;
				}
				break;
		}
		options = createOptions(options, "upload");
		var processor = new Processor("", "", options);
		processor.prepare(formData, callback);
	}
	function report($el, template, data, options) {
		var params = {
			"url" : baseUrl + "/report/json/" + user + "/" + template,
			"type" : "POST",
			"data" : data,
			"success" : function(data, status, xhr) {
				var ct = xhr.getResponseHeader("content-type") || "json",
					form = options.form,
					live = options.live && WebSocket && room && room.Connection;
				if (ct.indexOf("json") != -1) {
					$el.excelToCanvas(data);
					if (form) {
						if (live) {
							buildLiveForm($el, template, true, options.buildInput);
						} else {
							buildForm($el, null, options.buildInput);
						}
					} else if (live) {
						buildLiveForm($el, template, false);
					}
					if (options.callback) {
						options.callback($el, data);
					}
				} else if (ct.indexOf("html") != -1) {
					$el.html(data);
				}
			} 
		}, con = $.data($el.get(0), "connection");
		if (con) {
			con.close();
			$.removeData($el.get(0), "connection");
		}
		if (isFormData(data)) {
			params.processData = false;
			params.contentType = false;
		}
		if (options.apikey) {
			addAuthorization(params, options.apikey);
		}
		$.ajax(params);
	}
	switch (arguments.length) {
		case 0: 
			alert("user is required.");
			break;
		case 1:
			user = baseUrl;
			baseUrl = defaults.baseUrl;
			break;
		case 2:
			break;
	}
	$.extend(this, {
		"downloadExcel" : downloadExcel,
		"downloadPdf" : downloadPdf,
		"uploadExcel" : uploadExcel,
		"report" : report
	});
};

$.fn.uploadExcel = function(options, callback) {
	if (arguments.length == 1) {
		callback = options;
		options = null;
	}
	options = options || {};
	options.contextUrl = "upload";
	var processor = new Processor("", "", options),
		formData = new FormData($(this)[0]);
	processor.prepare(formData, callback);
};
$.fn.excelReport = function(method, params, param2) {
	//Private
	function normalizeData($el, data) {
		var ret = {};
		$.each(data, function(key, value) {
			var $div = $el.find("#" + key);
			if ($div.length) {
				ret[key] = value;
			} else {
				$div = $el.find("[data-name=" + key + "]");
				if ($div.length) {
					if ($.isArray(value)) {
						var idx = 0;
						$div.each(function() {
							var $cell = $(this),
								id = $cell.attr("id"),
								isFormula = $cell.attr("data-formula");
							if (!isFormula) {
								ret[id] = value[idx++];
							}
							return idx < value.length;
						});
					} else {
						var id = $($div.get(0)).attr("id");
						ret[id] = value;
					}
				}
			}
		});
		return ret;
	}
	//Public
	function showExcel($el, params) {
		var baseUrl = params.baseUrl || defaults.baseUrl,
			user = params.user,
			template = params.template,
			sheet = sheet,
			data = params.data,
			callback = params.callback,
			position = $el.css("position");
		if (sheet) {
			template = template + "/" + sheet;
		}
		if (position === "static") {
			$el.css("position", "relative");
		}
		new flect.ExcelReport(baseUrl, user).report($el, template, data, params);
	}
	function data($el, includeEmptyValue) {
		var ret = {};
		$el.find(":input").each(function() {
			var $input = $(this),
				name = $input.attr("name"),
				value = $input.val(),
				prev = ret[name],
				type = typeof(prev);
			if (type === "undefined") {
				ret[name] = value;
			} else if (type === "string") {
				var array = [];
				array.push(prev);
				array.push(value);
				ret[name] = array;
			} else {
				prev.push(value);
			}
		});
		if (!includeEmptyValue) {
			var remove = [];
			for (var key in ret) {
				var value = ret[key];
				if ($.isArray(value)) {
					while (value.length > 0) {
						if (value[value.length-1]) {
							break;
						} else {
							value.pop();
						}
					}
					if (value.length === 0) {
						remove.push(key);
					}
				} else if (!value) {
					remove.push(key);
				}
			}
			for (var i=0; i<remove.length; i++) {
				delete ret[remove[i]];
			}
		}
		return ret;
	}
	function updateCells($el, p1, p2) {
		var params = p1,
			con = $.data($el.get(0), "connection");
		if (typeof(p1) === "string" && typeof(p2) !== "undefined") {
			params = {};
			params[p1] = p2;
		}
		params = normalizeData($el, params);
		$.each(params, function(key, value) {
			var $div = $el.find("#" + key);
			if ($div.length) {
				var $input = $div.find(":input");
				if ($input.length) {
					$input.val(value);
				} else {
					var $span = $div.find("span");
					$span.text(value);
					if (isNumeric("" + value) && !$span.hasClass("cell-ac")) {
						$span.removeClass("cell-al").addClass("cell-ar");
					}
				}
			}
		});
		if (con) {
			con.request({
				"command" : "updateCells",
				"data" : params
			});
		}
	}
	function download($el, method, params) {
		var bForm = $el[0].tagName.toLowerCase() == "form",
			postData = bForm ? new FormData($el[0]) : data($el),
			newParams = $.extend({}, params);
		newParams.data = postData;
		$.excelReport(method, newParams);
	}
	//jQuery
	if (typeof method === "object") {
		params = method;
		method = "show";
	}
	switch (method) {
		case "show":
		case "showExcel": 
			showExcel(this, params);
			return this;
		case "data":
			return data(this, !!params);
		case "update":
		case "updateCells":
			updateCells(this, params, param2);
			return this;
		case "pdf":
		case "downloadPdf":
			download(this, method, params);
			break;
		case "excel":
		case "downloadExcel":
			download(this, method, params);
			break;
		default:
			throw "Unknown method: " + method;
	}
};
$.excelReport = function(method, params) {
	function configure(params) {
		if (typeof(params) === "string") {
			params = {
				"baseUrl" : params
			};
		}
		if (params.msg) {
			$.extend(MSG, params.msg);
			delete params.msg;
		}
		$.extend(defaults, params);
	}
	function download(params, pdf) {
		function appendData(key, value) {
			if (isFormData(data)) {
				data.append(key, value);
			} else {
				data[key] = value;
			}
		}
		var user = params.user,
			template = params.template,
			data = params.data,
			options = $.extend({}, params);
		delete options.user;
		delete options.template;
		delete options.data;
		if (!options.baseUrl) {
			options.baseUrl = defaults.baseUrl;
		}
		if (pdf && options.pdf) {
			$.each(options.pdf, function(key, value) {
				appendData("pdf." + key, value);
			});
			delete options.pdf;
		}
		if (options.filename) {
			appendData("filename", options.filename);
			delete options.filename;
		}
		options.contextUrl = pdf ? "pdf" : "download";
		new Processor(user, template, options).prepare(data);
	}
	if (typeof method === "object") {
		params = method;
		method = "configure";
	}
	switch (method) {
		case "configure":
			configure(params);
			break;
		case "pdf":
		case "downloadPdf":
			download(params, true);
			break;
		case "excel":
		case "downloadExcel":
			download(params, false);
			break;
		default:
			throw "Unknown method: " + method;
	}
};


})(jQuery);
