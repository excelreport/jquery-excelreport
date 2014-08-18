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

