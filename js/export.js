// exportStuff --> generate --> getWorkbook

var exportStuff = function () {
    return generate()
        .then(function (blob) {
            if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
            } else {
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement("a");
                document.body.appendChild(a);
                a.href = url;
                a.download = "out.xlsx";
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            }
        })
        .catch(function (err) {
            alert(err.message || err);
            throw err;
        });
}

function generate(options) {
    options = options || {};

    return getWorkbook()
        .then(function (workbook) {


            //prepare data
            var headers = [],
                dataRow = [],
                data = [];
            //prepare data - header - 10 columns
            for (var i = 0; i < 10; i++) {
                headers.push("header" + i);
            }

            //prepare data - body - 100k rows
            for (var m = 0; m < 100000; m++) {
                dataRow = [];

                for (var n = 0; n < 10; n++) {
                    dataRow.push("data_" + m + "_" + n);
                }
                data.push(dataRow);
            }

            //data - headers
            workbook.sheet(0).range("A1:T1").value([headers]);

            //data - body
            workbook.sheet(0).range("A2:T100002").value(data);

            return workbook.outputAsync();
        })
}

function getWorkbook() {
    return new Promise(function (resolve, reject) {
        var req = new XMLHttpRequest();
        var url = "/templates/templateEn.xlsx";
        req.open("GET", url, true);
        req.responseType = "arraybuffer";
        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 200) {
                    resolve(XlsxPopulate.fromDataAsync(req.response));
                } else {
                    reject("Received a " + req.status + " HTTP code.");
                }
            }
        };

        req.send();
    });
}