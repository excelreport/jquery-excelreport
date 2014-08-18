var $promptDiv = $("<div class='exrep-popup' style='display:none;'/>"),
	$errorDiv = $("<div class='exrep-popup exrep-error' style='display:none;'/>");
function Popup(title, text, error) {
	function show($input) {
		var $parent = $input.parent("div"),
			offset = $parent.position();
		$div.appendTo($parent.parent()).css({
			"top" : offset.top - 60,
			"left" : offset.left + 20
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
	}
	$content.text(text);
	$div.append($content);

	$.extend(this, {
		"show" : show,
		"hide" : hide
	});
}