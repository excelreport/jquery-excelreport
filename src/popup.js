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