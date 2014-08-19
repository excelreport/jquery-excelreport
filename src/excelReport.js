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
console.log("test1: ", id, $input);
				if ($input === null && ruleMan) {
					$input = ruleMan.buildInput(name, id);
				}
console.log("test2: ", id, $input);
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
console.log("test3: ", id, $span.attr("class"));
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
					if (ruleMan) {
						$input.focus(ruleMan.onFocus);
					}
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
								name = $input.attr("name"),
								value = $input.val();
							con.request({
								"command" : "modified",
								"data" : {
									"id" : id,
									"name" : name,
									"value" : value
								},
								"success" : function(data) {
									defaults.onRule.call($input[0], "error", data);
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
