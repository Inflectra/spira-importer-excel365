function ajaxImport(artifact, objTemplate) {
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/${artifact}${userInfo.auth}`,
        success: function (data) {
            let valueArray = [];
            for (let obj in data){
                valueArray.push(jsonToArray(data[obj], objTemplate));
            }
            toExcel(artifact, valueArray);
        },
        error: function () {
            console.log("failed to import");
        }
    });
}

function jsonToArray(oldObj, objTemplate) {
    let valArray = [];
    for (let prop in objTemplate) {
        objTemplate[prop] = oldObj[prop];
        valArray.push(objTemplate[prop]);
    }
    return valArray;
}

function toExcel(artifact, newValues) {
    return Excel.run(function (context) {
        let sheetName = toSheetName(artifact);
        let valRange = columnRanges[artifact] + (newValues.length + 2);
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let sheetValues = sheet.getRange(valRange);
        sheetValues.values = newValues;
        enableButtons();
        return context.sync();
    });
}