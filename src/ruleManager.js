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
