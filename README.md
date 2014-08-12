# jquery.excelreport
Client library for [ExcelReport](http://excel-report2.herokuapp.com)

## Dependencies
This library is a client library for using ExcelReport service.
If you don't have ExcelReport account, please [sign up](http://excel-report2.herokuapp.com) first.

Some features are depend on other js libraries.  
If you want to use these features, you must include followings. Or use all-in-one version.

### [excel2canvas](https://github.com/shunjikonishi/excel2canvas)
If you want to use showExcel method, this library is required.

### [flotr2](https://github.com/HumbleSoftware/Flotr2)
If you want to use showExcel method with drawing chart, this library is required.

### [roomframework](https://github.com/shunjikonishi/roomframework)
If you want to use liveform, this library is required.


## Files
All distributed files are in dist directory.

- jquery.excelreport[.min].js - core library
- i18n/excelreport.msg_ja.[.min].js - language option
- jquery.excelreport.full[.min].js - Concatenate all dependencies.
- jquery.excelreport.nochart[.min].js- Concatenete dependencies expect the chart.
- jquery.excel2canvas[.min].css - CSS for showExcel method


### All-In-One
All in one files include follwings.

- jquery.excelreport.full[.min].js
  - flotr2.js
  - roomframework.js
  - jquery.excel2canvas.js
  - jquery.excel2chart.flotr2.js
  - jquery.excelreport.js

- jquery.excelreport.nochart[.min].js
  - roomframework.js
  - jquery.excel2canvas.js
  - jquery.excelreport.js

### License
MIT