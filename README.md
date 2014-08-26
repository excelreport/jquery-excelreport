# jquery.excelreport
Client library for [ExcelReport](https://excelreport.net)

If you don't have ExcelReport account, please [sign up](https://excelreport.net) first.

## Files
All distributed files are in dist directory.

- jquery.excelreport.full[.min].js - Concatenate all dependencies.
- jquery.excelreport.nochart[.min].js- Concatenete dependencies expect the chart.
- jquery.excelreport[.min].css - CSS for show method
- i18n/excelreport.msg_ja.[.min].js - language option


## Dependencies
All in one files include follwings.

- jquery.excelreport.full[.min].js
  - [flotr2.js](https://github.com/HumbleSoftware/Flotr2)
  - [roomframework.js](https://github.com/shunjikonishi/roomframework)
  - [jquery.excel2canvas.js](https://github.com/shunjikonishi/excel2canvas)
  - [jquery.excel2chart.flotr2.js](https://github.com/shunjikonishi/excel2canvas)
  - jquery.excelreport.js

- jquery.excelreport.nochart[.min].js
  - roomframework.js
  - jquery.excel2canvas.js
  - jquery.excelreport.js

## License
MIT