function ajaxImport(artifact, objTemplate) {
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/${artifact}${userInfo.auth}`,
        success: function (data) {
            let valueArray = [];
            for (let obj in data) {
                valueArray.push(jsonToArray(data[obj], objTemplate));
            }
            toExcel(artifact, valueArray);
        },
        error: function () {

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

function loadCustomFields(artifact, project) {
    let artifactNum = undefined;
    if (artifact == "requirements") {
        artifactNum = 1;
    }
    disableButtons();
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects/${project}/custom-properties/${artifactNum}${userInfo.auth}`,
        success: function (data) {
            if (data.length < 1) {
                populateCustomFieldNames([{ "Name": "" }], artifact);
            }
            else {
                for (let customProp of data) {
                    customFieldInfo = {};
                    customFieldInfo.Name = customProp.Name;
                    customFieldInfo.Type = customProp.CustomPropertyTypeName;
                    customFieldNames.push(customFieldInfo);
                }
                populateCustomFieldNames(customFieldNames, artifact);
                enableButtons();
            }
        },
        error: function () {
            enableButtons();
            populateCustomFieldNames([{ "Name": "" }], artifact);
        }
    });
}

function populateCustomFieldNames(cusObj, artifact) {
    let newNames = [];
    for (let info of cusObj) {
        newNames.push(info.Name);
    }
    if (newNames.length < 30) {
        for (let i = newNames.length; i < 30; i++) {
            newNames.push("");
        }
    }
    return Excel.run(function (context) {
        let sheetName = toSheetName(artifact);
        let customFieldNameRange = customFieldRanges[artifact][0] + "2:" + customFieldRanges[artifact][1] + "2";
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let names = sheet.getRange(customFieldNameRange);
        names.values = [newNames];
        enableButtons();
        return context.sync();
    });
    enableButtons();
}

function returnId(newId, artifact, row) {
    let currentRow = (row + 2);
    return Excel.run(function (context) {
        let sheetName = toSheetName(artifact);
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let IdCell = sheet.getCell(currentRow, 0);
        IdCell.values = [[newId]];
        return context.sync();
    });
}