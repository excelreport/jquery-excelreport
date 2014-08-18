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
