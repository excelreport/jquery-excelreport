flect.ExcelReport = function($el, baseUrl, user, template, sheet, options) {
	function getCacheKey() {
		var ret = "excelreport-" + user + "-" + template;
		if (sheet) {
			ret += "-" + sheet;
		}
		return ret;
	}
	function isMultiNamedCell($div) {
		var name = $div.attr("data-name"),
			$prev = $div.prev("div:eq(0)"),
			$next = $div.next("div:eq(0)");
		return $prev.attr("data-name") == name || $next.attr("data-name") == name;
	}
	function buildForm() {
		function getRealWidth($span) {
			var $temp = $("<span style='display:none;'/>"),
				ret = 0;
			$temp.text($span.text());
			$span.parent().append($temp);
			ret = $temp.width();
			$temp.remove();
			return ret;
		}
		debug("ExcelReport.buildForm");
		$el.find(".namedCell").each(function() {
			var $div = $(this),
				$span = $div.find("span"),
				id = $div.attr("id"),
				name = $div.attr("data-name"),
				text = $span.text(),
				isFormula = !!$div.attr("data-formula");
			if (!isFormula) {
				var $input = options.buildInput ? options.buildInput(name, id) : null,
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
					if (ruleMan) {
						$input.focus(ruleMan.onFocus);
					}
				}
				$div.append($input);
			}
		});
		if (con) {
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
						if (data == "OK") {
							defaults.onRule.call($input[0], "modified", value);
						} else {
							defaults.onRule.call($input[0], "error", data);
						}
					}
				});
			});
		}
		finish();
	}
	function openConnection() {
		function buildUrl() {
			var url = "",
				lang = defaults.lang || $("html").attr("lang") || "";
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
			url += "/liveform/" + user + "/" + template;
			if (lang) {
				url += "?language=" + lang;
			}
			return url;
		}
		debug("ExcelReport.openConnection");
		var params = {
				"url": buildUrl()
			},
			first = con === null;
		if (defaults.debug) {
			params.logger = console;
		}
		con = new room.Connection(params);
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
			if (first) {
				con.request({
					"command" : "available",
					"data" : {
						"rules" : ruleMan === null
					},
					"success" : function(data) {
						if (data.error) {
							alert(data.error);
							con.close();
							con = null;
						} else {
							license = data.license;
							if (!ruleMan) {
								ruleMan = new RuleManager(data.rules);
							}
							if (options.form) {
								buildForm();
							}
						}
					}
				});
			} else {
				$el.excelReport("update", $el.excelReport("data"));
			}
		});
		con.sendNoop(30, !room.utils.isMobile());
	}
	function getInfo(callback) {
		debug("ExcelReport.getInfo");
		$.ajax({
			"url" : baseUrl + "/report/info/" + user + "/" + template,
			"success" : callback
		});
	}
	function show(data) {
		debug("ExcelReport.show");
		var hasData = data && !$.isEmptyObject(data);
		if (options.cache && !hasData) {
			var cachedObj = storage.getAsJson(getCacheKey());
			if (cachedObj) {
				getInfo(function(info) {
					lastModified = info.lastModified;
					if (info.lastModified == cachedObj.lastModified) {
						if (cachedObj.rules) {
							ruleMan = new RuleManager(cachedObj.rules);
						}
						if (cachedObj.json) {
							json = cachedObj.json;
							showJson(json);
							return;
						}
					}
					loadJson(data, hasData);
				});
				return;
			}
		}
		loadJson(data, hasData);
	}
	function loadJson(data, hasData) {
		debug("ExcelReport.loadJson");
		var params = {
			"url" : baseUrl + "/report/json/" + user + "/" + template,
			"type" : "POST",
			"data" : data,
			"success" : function(data, status, xhr) {
				var ct = xhr.getResponseHeader("content-type") || "json";
				if (ct.indexOf("json") != -1) {
					if (!hasData) {
						json = data;
					}
					showJson(data);
				} else if (ct.indexOf("html") != -1) {
					$el.html(data);
				}
			}
		};
		if (sheet) {
			params.url += "/" + sheet;
		}
		if (isFormData(data)) {
			params.processData = false;
			params.contentType = false;
		}
		if (apikey) {
			addAuthorization(params, options.apikey);
		}
		$.ajax(params);
	}
	function showJson(data) {
		debug("ExcelReport.showJson");
		$el.excelToCanvas(data);
		if (options.live) {
			openConnection();
		} else if (options.form) {
			buildForm();
		} else {
			finish();
		}
	}
	function finish() {
		debug("ExcelReport.finish");
		if (license == "Free") {
			var $img = $("<img/>");
			$img.attr("src", baseUrl + "/assets/images/PoweredByExcelReport.png");
			$img.css({
				"position" : "absolute",
				"right" : 0,
				"bottom" : 0,
				"z-index" : 1000
			});
			$img.addClass("exrep-logo");
			$el.append($img);
		}
		if (options.callback) {
			options.callback($el, data);
		}
		if (options.cache && (json || ruleMan)) {
			var obj = {};
			if (json) {
				obj.json = json;
			}
			if (ruleMan) {
				obj.rules = ruleMan.rules;
			}
			if (lastModified) {
				obj.lastModified = lastModified;
				storage.put(getCacheKey(), obj);
			} else {
				getInfo(function(info) {
					obj.lastModified = info.lastModified;
					storage.put(getCacheKey(), obj);
				});
			}
		}
	}
	function release() {
		debug("ExcelReport.release");
		if (con) {
			con.close();
		}
		storage = null;
		lastModified = null;
		con = null;
		license = null;
		json = null;
		ruleMan = null;
	}
	debug("ExcelReport.init", user, template, sheet);
	var storage = new room.Cache(localStorage),
		lastModified = null,
		con = null,
		license = null,
		json = null,
		ruleMan = null;
	$.extend(this, {
		"show" : show,
		"release" : release
	});
};
