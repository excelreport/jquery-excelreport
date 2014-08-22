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
function debug() {
	if (defaults.debug) {
		console.log.apply(console, arguments);
	}
}
function makeUrl(options, path) {
	return options.baseUrl + "/" + options.contextUrl + path;
}
function addAuthorization(params, apikey) {
	params.headers = {
		"X-Auth-Token": apikey
	};
}
