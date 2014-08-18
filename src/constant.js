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
