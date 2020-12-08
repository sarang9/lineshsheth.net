<!DOCTYPE html>
<html xmlns=“http://www.w3.org/1999/xhtml”>
<head>
    <title>Export Hierarchical HTML Using JQuery – Sibeesh|Passion</title>
    <script src=“Contents/jquery-1.9.1.js”></script>
    <style>
    td,th,thead,caption,tr{
        text-align:center;border:1px solid #ccc;
    }
        caption {
            background-color:#ccc;
        }
</style>
    <script type=“text/javascript”>
        var exportThis = (function () {
            var uri = ‘data:application/vnd.ms-excel;base64,’,
                template = ‘<html xmlns:o=”urn:schemas-microsoft-com:office:office” xmlns:x=”urn:schemas-microsoft-com:office:excel”  xmlns=”http://www.w3.org/TR/REC-html40″><head> <!–[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets> <x:ExcelWorksheet><x:Name>{worksheet}</x:Name> <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions> </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook> </xml><![endif]–></head><body> <table>{table}</table></body></html>’,
                base64 = function (s) {
                    return window.btoa(unescape(encodeURIComponent(s)))
                },
                format = function (s, c) {
                    return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; })
                }
            return function () {
                var ctx = { worksheet: ‘Multi Level Export Table Example’ || ‘Worksheet’, table: document.getElementById(“multiLevelTable”).innerHTML }
                window.location.href = uri + base64(format(template, ctx))
            }
        })()
        var exportThisWithParameter = (function () {
            var uri = ‘data:application/vnd.ms-excel;base64,’,
                template = ‘<html xmlns:o=”urn:schemas-microsoft-com:office:office” xmlns:x=”urn:schemas-microsoft-com:office:excel”  xmlns=”http://www.w3.org/TR/REC-html40″><head> <!–[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets> <x:ExcelWorksheet><x:Name>{worksheet}</x:Name> <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions> </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook> </xml><![endif]–></head><body> <table>{table}</table></body></html>’,
                base64 = function (s) {
                    return window.btoa(unescape(encodeURIComponent(s)))
                },
                format = function (s, c) {
                    return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; })
                }
            return function (tableID, excelName) {
                tableID = document.getElementById(tableID)
                var ctx = { worksheet: excelName || ‘Worksheet’, table: tableID.innerHTML }
                window.location.href = uri + base64(format(template, ctx))
            }
        })()
    </script>
</head>
<body>
    <div onclick=“exportThis()” style=“cursor: pointer; border: 1px solid #ccc; text-align: center;width:19%;”>Export Multi Level Table to Excel</div>
    <br />
    <div onclick=“exportThisWithParameter(‘multiLevelTable’, ‘Multi Level Export Table Example’)” style=“cursor: pointer; border: 1px solid #ccc; text-align: center;width:19%;”>Export Multi Level Table to Excel With Parameter</div>
    <br />
    <br />
    <table id=“multiLevelTable”>
        <caption>My Multi Level Table</caption>
        <thead>
            <tr>
                <th colspan=“2” rowspan=“2” style=“background-color:aqua;”>Column 0</th>
                <th rowspan=“2”>Column 1</th>
                <th colspan=“2”>Column 2</th>
            </tr>
            <tr>
                <th>Column 2a</th>
                <th>Column 2b</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <th rowspan=“2”>Row 1</th>
                <th>Row 1a</th>
                <td>123</td>
                <td>456</td>
                <td>789</td>
            </tr>
            <tr>
                <th>Row 1b</th>
                <td>123</td>
                <td>456</td>
                <td>789</td>
            </tr>
            <tr>
                <th colspan=“2”>Row 2</th>
                <td>123</td>
                <td>456</td>
                <td>789</td>
            </tr>
        </tbody>
    </table>
</body>
</html>