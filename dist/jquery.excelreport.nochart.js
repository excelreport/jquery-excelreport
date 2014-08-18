if (typeof(room) === "undefined") room = {};
if (typeof(room.utils) === "undefined") room.utils = {};

$(function() {
	"use strict";
	var visibilityPrefix = (function() {
		var key = "hidden",
			prefix = ["webkit", "moz", "ms"];
		if (key in document) {
			return "";
		}
		key = "Hidden";
		for (var i=0; i<prefix.length; i++) {
			if (prefix + key in document) {
				return prefix;
			}
		}
		return "";
	})(),
	visibilityProp = visibilityPrefix ? visibilityPrefix + "Hidden" : "hidden",
	visibilityChangeProp = visibilityPrefix + "visibilitychange",
	nullLogger = {
		"log" : function() {}
	}, 
	defaultSettings = {
		"maxRetry" : 5,
		"authCommand" : "room.auth",
		"authError" : null,
		"retryInterval" : 1000,
		"noopCommand" : "noop",
		"logger" : nullLogger
	};
	if (isIOS()) {
		visibilityChangeProp = "pageshow";
	}
	function isDocumentVisible() {
		return !document[visibilityProp];
	}
	function isMobile() {
		return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
	}
	function isIOS() {
		return /iPhone|iPad|iPod/i.test(navigator.userAgent);
	}

	/**
	 * settings
	 * - url
	 * - maxRetry
	 * - retryInterval
	 * - logger
	 * - authToken
	 * - onOpen(event)
	 * - onClose(event)
	 * - onMessage(event)
	 * - onSocketError(event)
	 * - onRequest(command, data)
	 * - onServerError(data)
	 */
	room.Connection = function(settings) {
		function request(params) {
			logger.log("request", params);
			if (!isConnected()) {
				if (openning || retryCount < settings.maxRetry) {
					ready(function() {
						request(params);
					});
					if (!openning) {
						socket = createWebSocket();
					}
				}
				return;
			}
			if (settings.onRequest) {
				settings.onRequest(params.command, params.data);
			}
			var startTime = new Date().getTime(),
				id = ++requestId;
			times[id] = startTime;
			if (params.success) {
				listeners[id] = params.success;
			}
			if (params.error) {
				errors[id] = params.error;
			}
			var msg = {
				"id" : id,
				"command" : params.command,
				"data" : params.data
			};
			if (params.log) {
				msg.log = params.log;
			}
			socket.send(JSON.stringify(msg));
			return self;
		}
		function on(name, sl, el) {
			if (sl) {
				listeners[name] = sl;
			}
			if (el) {
				errors[name] = el;
			}
			return self;
		}
		function off(name) {
			delete listeners[name];
			delete errors[name];
			return self;
		}
		function onOpen(event) {
			function authError(data) {
				logger.log("authError", settings.url);
				retryCount = settings.maxRetry;
				if (settings.authError) {
					settings.authError(data);
				}
			}
			openning = false;
			logger.log("onOpen", settings.url);
			if (settings.onOpen) {
				settings.onOpen(event);
			}
			retryCount = 0;
			if (settings.authToken) {
				request({
					"command" : settings.authCommand,
					"data" : settings.authToken,
					"success" : function(data) {
						settings.authToken = data;
					},
					"error" : authError
				});
			}
			for (var i=0; i<readyFuncs.length; i++) {
				readyFuncs[i]();
			}
			readyFuncs = [];
		}
		function onMessage(event) {
			logger.log("receive", event.data);
			if (settings.onMessage) {
				settings.onMessage(event);
			}
			var data = JSON.parse(event.data),
				startTime = times[data.id],
				func = null;
			if (startTime) {
				delete times[data.id];
			}
			if (data.type == "error") {
				if (data.id && errors[data.id]) {
					func = errors[data.id];
				} else if (data.command && errors[data.command]) {
					func = errors[data.command];
				} else if (settings.onServerError) {
					func = settings.onServerError;
				}
			} else {
				if (data.id && listeners[data.id]) {
					func = listeners[data.id];
				} else if (data.command && listeners[data.command]) {
					func = listeners[data.command];
				}
			}
			if (data.id) {
				delete listeners[data.id];
				delete errors[data.id];
			}
			if (func) {
				func(data.data);
			} else if (data.type != "error") {
				logger.log("UnknownMessage", event.data);
			}
		}
		function onClose(event) {
			function isRetry() {
				if (isMobile() && !isDocumentVisible()) {
					return false;
				}
				return retryCount < settings.maxRetry;
			}
			logger.log("close", settings.url);
			if (settings.onClose) {
				settings.onClose(event);
			}
			if (isRetry()) {
				setTimeout(function() {
					if (!isConnected()) {
						socket = createWebSocket();
					}
				}, retryCount * settings.retryInterval);
				retryCount++;
			}
		}
		function onError(event) {
			if (settings.onSocketError) {
				settings.onSocketError(event);
			}
		}
		function polling(interval, params) {
			return setInterval(function() {
				if (isConnected()) {
					request($.extend(true, {}, params));
				}
			}, interval);
		}
		function ready(func) {
			if (isConnected()) {
				func();
			} else {
				readyFuncs.push(func);
			}
		}
		function close() {
			if (isConnected()) {
				retryCount = settings.maxRetry;
				socket.close();
			}
		}
		function isConnected() {
			return socket.readyState == 1;//OPEN
		}
		function createWebSocket() {
			openning = true;
			var socket = new WebSocket(settings.url);
			socket.onopen = onOpen;
			socket.onmessage = onMessage;
			socket.onerror = onError;
			socket.onclose = onClose;
			return socket;
		}
		function sendNoop(interval, sendIfHidden, commandName) {
			return setInterval(function() {
				if (isConnected() && (sendIfHidden || isDocumentVisible())) {
					request({
						"command" : commandName || "noop"
					});
				}
			}, interval * 1000);
		}
		if (typeof(settings) === "string") {
			settings = {
				"url" : settings
			};
		}
		settings = $.extend({}, defaultSettings, settings);
		var self = this,
			logger = settings.logger,
			requestId = 0,
			times = {},
			listeners = {},
			errors = {},
			readyFuncs = [],
			openning = false,
			retryCount = 0,
			socket = createWebSocket();
		$(window).on("beforeunload", close);
		$(document).on(visibilityChangeProp, function() {
			var bVisible = isDocumentVisible();
			logger.log(visibilityChangeProp, "visible=" + bVisible);
			if (bVisible && !isConnected()) {
				socket = createWebSocket();
			}
		});
		if (settings.noopCommand) {
			on(settings.noopCommand, function() {});
		}
		$.extend(this, {
			"request" : request,
			/** deprecated use on method.*/
			"addEventListener" : on,
			/** deprecated use off method.*/
			"removeEventListener" : off,
			"on" : on,
			"off" : off,
			"polling" : polling,
			"ready" : ready,
			"close" : close,
			"isConnected" : isConnected,
			"sendNoop" : sendNoop,
			"onOpen" : function(func) { settings.onOpen = func; return this;},
			"onClose" : function(func) { settings.onClose = func; return this;},
			"onRequest" : function(func) { settings.onRequest = func; return this;},
			"onMessage" : function(func) { settings.onMessage = func; return this;},
			"onSocketError" : function(func) { settings.onSocketError = func; return this;},
			"onServerError" : function(func) { settings.onServerError = func; return this;}
		});
	};
	$.extend(room.utils, {
		"isIOS" : isIOS,
		"isMobile" : isMobile,
		"isDocumentVisible" : isDocumentVisible
	});
});
if (typeof(room) === "undefined") room = {};

$(function() {
	room.Cache = function(storage, logErrors) {
		function get(key) {
			if (storage) {
				try {
					return storage.getItem(key);
				} catch (e) {
					//Ignore. Happened when cookie is disabled. 
					if (logErrors && console) {
						console.log(e);
					}
				}
			}
			return null;
		}
		function getAsJson(key) {
			var ret = get(key);
			if (ret) {
				ret = JSON.parse(ret);
			}
			return ret;
		}
		function put(key, value) {
			if (storage) {
				if (typeof(value) === "object") {
					value = JSON.stringify(value);
				}
				try {
					storage.setItem(key, value);
				} catch (e) {
					//Ignore. Happened when cookie is disabled, or quota exceeded.
					if (logErrors && console) {
						console.log(e);
					}
				}
			}
		}
		function remove(key) {
			if (storage) {
				try {
					storage.removeItem(key);
				} catch (e) {
					//Ignore. Happened when cookie is disabled, or quota exceeded.
					if (logErrors && console) {
						console.log(e);
					}
				}
			}
		}
		function keys() {
			var ret = [];
			if (storage) {
				try {
					for (var i=0; i<storage.length; i++) {
						ret.push(storage.key(i));
					}
				} catch (e) {
					//Ignore. Happened when cookie is disabled. 
					if (logErrors && console) {
						console.log(e);
					}
				}
			}
			return ret;
		}
		function size() {
			try {
				return sotrage ? storage.length : 0;
			} catch (e) {
				if (logErrors && console) {
					console.log(e);
				}
			}
		}
		function clear() {
			if (storage) {
				try {
					storage.clear();
				} catch (e) {
					//Ignore. Happened when cookie is disabled. 
					if (logErrors && console) {
						console.log(e);
					}
				}
			}
		}
		$.extend(this, {
			"get" : get,
			"getAsJson" : getAsJson,
			"put" : put,
			"remove" : remove,
			"size" : size,
			"clear" : clear,
			"keys" : keys
		});
	};
});
if (typeof(room) === "undefined") room = {};
if (typeof(room.logger) === "undefined") room.logger = {};
if (typeof(room.utils) === "undefined") room.utils = {};

$(function() {
	function stripFunc(obj) {
		var type = typeof(obj);
		if (type !== "object") {
			return obj;
		} else if ($.isArray(obj)) {
			var newArray = [];
			for (var i=0; i<obj.length; i++) {
				type = typeof(obj[i]);
				if (type === "function") {
					newArray.push("(function)");
				} else if (type === "object") {
					newArray.push(stripFunc(obj[i]));
				} else {
					newArray.push(obj[i]);
				}
			}
			return newArray;
		} else {
			var newObj = {};
			for (var key in obj) {
				type = typeof(key);
				if (type === "function") {
					newObj[key] = "(function)";
				} else if (type === "object") {
					newObj[key] = stripFunc(obj[key]);
				} else {
					newObj[key] = obj[key];
				}
			}
			return newObj;
		}
	}
	room.logger.WsLogger = function(wsUrl, commandName) {
		var ws = new room.Connection(wsUrl);
		commandName = commandName || "log";
		this.log = function() {
			ws.request({
				"command": commandName,
				"data": stripFunc(arguments)
			});
		};
	};
	room.logger.DivLogger = function($div) {
		this.log = function() {
			var msgs = [];
			for (var i=0; i<arguments.length; i++) {
				if (typeof(arguments[i]) == "object") {
					msgs.push(stripFunc(arguments[i]));
				} else {
					msgs.push(arguments[i]);
				}
			}
			$("<p/>").text(JSON.stringify(msgs)).prependTo($div);
		};
	};
	room.logger.nullLogger = {
		"log" : function() {}
	};
	$.extend(room.utils, {
		"stripFunc" : stripFunc
	});
});
(function ($) {
	var
		BORDER_THIN                = 1,
		BORDER_MEDIUM              = 2,
		BORDER_DASHED              = 3,
		BORDER_HAIR                = 4,
		BORDER_THICK               = 5,
		BORDER_DOUBLE              = 6,
		BORDER_DOTTED              = 7,
		BORDER_MEDIUM_DASHED       = 8,
		BORDER_DASH_DOT            = 9,
		BORDER_MEDIUM_DASH_DOT     = 10,
		BORDER_DASH_DOT_DOT        = 11,
		BORDER_MEDIUM_DASH_DOT_DOT = 12,
		BORDER_SLANTED_DASH_DOT    = 13,
		context;
	
	function hasTooltip() {
		return !!$.fn.tooltip
	}
	function fillStyle(data, fill) {
		var back = fill.back,
			fore = fill.fore,
			pattern = fill.pattern;
		if (fill.styleRef) {
			var styles = data.styles[fill.styleRef].split("|");
			back = styles[0];
			fore = styles[1];
			pattern = styles[2];
		}
		//ToDo back, pattern
		if (fore) {
			context.fillStyle = fore;
			context.fillRect(fill.p[0], fill.p[1], fill.p[2], fill.p[3]);
		}
	}
	function drawLine(line) {
		var kind = line.kind ? line.kind : 1,
			w = 1,
			x1 = line.p[0],
			y1 = line.p[1],
			x2 = line.p[2],
			y2 = line.p[3],
			horizontal = y1 == y2;
		
		if (kind == BORDER_MEDIUM || 
		    kind == BORDER_MEDIUM_DASHED || 
		    kind == BORDER_MEDIUM_DASH_DOT ||
		    kind == BORDER_MEDIUM_DASH_DOT_DOT) 
		{
			w = 2;
		} else if (kind == BORDER_THICK) {
			w = 3;
		}
		context.lineWidth = w;
		if (w == 1 || w == 3) {
			if (horizontal) {
				y1 = y1 == 0 ? 0.5 : y1 - 0.5;
				y2 = y2 == 0 ? 0.5 : y2 - 0.5;
			} else {
				x1 = x1 == 0 ? 0.5 : x1 - 0.5;
				x2 = x2 == 0 ? 0.5 : x2 - 0.5;
			}
		} else {//w == 2
			if (horizontal && y1 == 0) {
				y1 = 1;
				y2 = 1;
			} else if (!horizontal && x1 == 0) {
				x1 = 1;
				x2 = 1;
			}
		}
		if (line.color) {
			context.strokeStyle = line.color;
		} else {
			context.strokeStyle = "#000000";
		}
		if (kind == BORDER_DOUBLE) {
			context.beginPath();
			if (horizontal) {
				context.moveTo(x1, y1 - 1);
				context.lineTo(x2, y2 - 1);
				context.moveTo(x1, y1 + 1);
				context.lineTo(x2, y2 + 1);
			} else {
				context.moveTo(x1 - 1, y1);
				context.lineTo(x2 - 1, y2);
				context.moveTo(x1 + 1, y1);
				context.lineTo(x2 + 1, y2);
			}
			context.stroke();
			context.closePath();
		} else if (kind == BORDER_DASHED ||
		           kind == BORDER_HAIR ||
		           kind == BORDER_DOTTED ||
		           kind == BORDER_MEDIUM_DASHED)
		{
			var bw, sw;
			switch (kind) {
				case BORDER_DASHED:
					bw = 4; sw = 2;
					break;
				case BORDER_HAIR:
					bw = 2; sw = 2;
					break;
				case BORDER_DOTTED:
					bw = 1; sw = 1;
					break;
				case BORDER_MEDIUM_DASHED:
					bw = 8; sw = 3;
					break;
			}
			
			var bar = true;
			context.beginPath();
			context.moveTo(x1, y1);
			if (horizontal) {
				var y = y1,
					cx = x1,
					ex = x2;
				while (cx < ex) {
					var nx = bar ? bw : sw;
					cx += nx;
					if (cx > ex) {
						cx = ex;
					}
					if (bar) {
						context.lineTo(cx, y);
					} else {
						context.moveTo(cx, y);
					}
					bar = !bar;
				}
			} else {
				var x = x1,
					cy = y1,
					ey = y2;
				while (cy < ey) {
					var ny = bar ? bw : sw;
					cy += ny;
					if (cy > ey) {
						cy = ey;
					}
					if (bar) {
						context.lineTo(x, cy);
					} else {
						context.moveTo(x, cy);
					}
					bar = !bar;
				}
			}
			context.stroke();
			context.closePath();
		} else {
			context.beginPath();
			context.moveTo(x1, y1);
			context.lineTo(x2, y2);
			context.stroke();
			context.closePath();
		}
	}
	$.fn.excelToCanvas = function(data, mergeData) {
		var holder, canvas;
		if (this[0].tagName == "canvas") {
			canvas = this;
			holder = canvas.parent();
		} else {
			holder = this;
			canvas = holder.find("canvas");
			if (canvas.length == 0) {
				canvas = $("<canvas style='position:absolute;left:0;top:0;z-index:0'></canavas>").appendTo(holder);
			}
		}
		if (typeof FlashCanvas !== "undefined") {
			FlashCanvas.initElement(canvas[0]);
		}
		
		canvas.attr("width", data.width).attr("height", data.height);
		holder.css({
			"width" : data.width,
			"height" : data.height
		});
		context = canvas[0].getContext("2d");
		context.fillStyle = "white";
		context.fillRect(0, 0, data.width, data.height);
		if (data.font) {
			context.font = data.font;
		}
		
		if (data.fills) {
			for (var i=0; i<data.fills.length; i++) {
				var fill = data.fills[i];
				fillStyle(data, fill);
			}
		}
		if (data.lines) {
			for (var i=0; i<data.lines.length; i++) {
				var line = data.lines[i];
				drawLine(line);
			}
		}
		
		if (data.strs) {
			for (var i=0; i<data.strs.length; i++) {
				var str = data.strs[i],
					style = str.style ? str.style : data.styles[str.styleRef],
					span = $("<span style='" + style + "'></span>");
				if (str.link) {
					var link = $("<a target='_blank'></a>");
					link.append(str.text);
					link.attr("href", str.link);
					span.append(link);
				} else {
					span.append(str.text);
				}
				var align = str.align;
				if (align) {
					span.addClass("cell-a" + align[0]);
					span.addClass("cell-v" + align[1]);
					if (align[1] == "c" || align[1] == "j") {
						var n = str.text.split("<br>").length;
						if (n > 1) {
							span.css("margin-top", "-" + (n / 2) + "em");
						}
					}
				}
				if (str.rawdata) {
					span.attr("data-raw", str.rawdata);
				}
				var div = $("<div id='" + str.id + "' class='cell'></div>").append(span);
				div.css({
					"left" : str.p[0],
					"top" : str.p[1],
					"width" : str.p[2],
					"height" : str.p[3]
				});
				if (str.clazz) {
					div.addClass(str.clazz);
				}
				if (str.dataAttrs) {
					$.each(str.dataAttrs, function(key, value) {
						div.attr("data-" + key, value);
					});
				}
				if (str.comment && hasTooltip()) {
					context.strokeStyle = "red";
					context.fillStyle = "red";
					
					context.beginPath();
					context.moveTo(str.p[0] + str.commentWidth - 4,  str.p[1]);
					context.lineTo(str.p[0] + str.commentWidth, str.p[1]);
					context.lineTo(str.p[0] + str.commentWidth, str.p[1] + 4);
					context.lineTo(str.p[0] + str.commentWidth - 4, str.p[1]);
					context.stroke();
					context.fill();
					context.closePath();
					div.tooltip({
						"title" : str.comment,
						"html" : true
					});
				}
				holder.append(div);
			}
		}
		
		if (data.pictures) {
			for (var i=0; i<data.pictures.length; i++) {
				var pic = data.pictures[i],
					img = $("<img class='excel-img'/>");
				img.attr("src", pic.uri);
				img.css({
					"left" : pic.p[0],
					"top" : pic.p[1],
					"width" : pic.p[2],
					"height" : pic.p[3]
				});
				holder.append(img);
			}
		}
		if (data.charts && $.fn.excelToChart) {
			for (var i=0; i<data.charts.length; i++) {
				var chart = data.charts[i],
					options = {},
					chartDiv = $("<div class='excel-chart'></div>");
				chartDiv.css({
					"left" : chart.p[0],
					"top" : chart.p[1],
					"width" : chart.p[2],
					"height" : chart.p[3]
				});
				holder.append(chartDiv);
				if (data.font) {
					options.HtmlText = true;
				}
				chartDiv.excelToChart(chart.chart, options);
				chartDiv.css("position", "absolute");
			}
		}
		if (mergeData) {
			for (var key in mergeData) {
				var div = $("#" + key);
				if (div.length > 0) {
					var userData = mergeData[key];
					if (typeof (userData) === "string") {
						div.find("span").text(userData);
					} else if (userData.element) {
						var el = $(document.createElement(userData.element));
						for (attr in userData) {
							if (attr != "element") {
								el.attr(attr, userData[attr]);
							}
						}
						div.css("overflow", "visible").empty().append(el);
					}
				}
			}
		}
		return this;
	}
})(jQuery);

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
	function closePrompt() {
		popup.hide();
		$input.unbind("blur change", closePrompt);
	}
	var popup = null,
		$input = $(this);
	switch (eventName) {
		case "prompt":
			popup = new Popup(params.title, params.text, false);
			$input.blur(closePrompt);
			break;
		case "error":
			popup = new Popup(params.title, params.text, true);
			$input.change(closePrompt);
			break;
	}
	if (popup) {
		popup.show($input);
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
	function onFocus() {
		var $input = $(this),
			name = $input.attr("name"),
			id = $input.parent("div").attr("id"),
			rule = getRule(name, id);
		if (rule && rule.prompt) {
			defaults.onRule.call(this, "prompt", rule.prompt);
		}
	}
	$.extend(this, {
		"buildInput" : buildInput,
		"onFocus" : onFocus
	});
}

var $promptDiv = $("<div class='exrep-popup' style='display:none;'/>"),
	$errorDiv = $("<div class='exrep-popup exrep-error' style='display:none;'/>");
function Popup(title, text, error) {
	function show($input) {
		var $parent = $input.parent("div"),
			offset = $parent.position(),
			top = error ? offset.top + $parent.height() + 20 : offset.top - 60,
			left = offset.left + 20;

		$div.appendTo($parent.parent()).css({
			"top" : top,
			"left" : left
		}).show();
	}
	function hide() {
		$div.hide();
	}
	var $div = error ? $errorDiv : $promptDiv,
		$title = null,
		$content = $("<div class='exrep-popup-content'/>");
	$div.empty();
	if (title) {
		$title = $("<div class='exrep-popup-title'/>");
		$title.text(title);
		$div.append($title);
		if (error) {
			$("<div class='exrep-popup-close'>×</div>")
				.appendTo($title)
				.click(hide);
		}
	}

	$content.text(text);
	$div.append($content);

	$.extend(this, {
		"show" : show,
		"hide" : hide
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
