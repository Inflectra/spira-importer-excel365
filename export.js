function grabExcelValues(rows, artifact, objTemplate, customFieldRange) {
    if (rows === null) {
        getRows(artifact, objTemplate, customFieldRange);
    }
    else {
        return Excel.run(function (context) {
            let sheet = context.workbook.worksheets.getItem(toSheetName(artifact));
            let inputRange = columnRanges[artifact] + rows;
            let customRange = (customFieldRange[0] + 3) + ":" + (customFieldRange[1] + rows);
            let inputValues = sheet.getRange(inputRange);
            let customFields = sheet.getRange(customRange);
            inputValues.load();
            customFields.load();
            return context.sync()
                .then(function () {
                    customFieldObjArr = customFields.values.map(customFieldObjCreate);
                    ajaxExport(inputValues.values, artifact, objTemplate, customFieldObjArr);
                })
        });
    }
}

function getRows(artifact, objTemplate, customFieldRange) {
    return Excel.run(function (context) {
        let sheet = context.workbook.worksheets.getItem(toSheetName(artifact));
        let sheetRange = sheet.getUsedRange();
        sheetRange.load();
        return context.sync()
            .then(function () {
                grabExcelValues(sheetRange.values.length, artifact, objTemplate, customFieldRange);
            })
            .catch(function (error) {
                console.log(error);
            })
    });
}

function ajaxExport(valueArray, artifact, objTemplate, customFieldObjArr) {
    let objArray = buildobjects(valueArray, artifact, objTemplate);
    for (let i in customFieldObjArr){
        objArray[i].CustomProperties = customFieldObjArr[i];
    }
    postNew(objArray, artifact, 0);
    console.log(objArray);
}

function buildobjects(valueArray, artifact, objTemplate) {
    let objArray = [];
    for (let i = 0; i < valueArray.length; i++) {
        let j = 0;
        for (let prop in objTemplate) {
            //grabs the digit from Importance Name field for the id
            if (prop == "ImportanceId") {
                objTemplate[prop] = valueArray[i][j].charAt(0);
            }
            else {
                objTemplate[prop] = valueArray[i][j];
            }
            j++
        }
        objArray.push(cleanObject(objTemplate));
    }
    return objArray;
}

function postNew(toSend, artifact, reqNum) {
    let id = $('#projects').val();
    if (toSend[reqNum] && toSend[reqNum].hasOwnProperty(toIdString(artifact))) {
        $("<p>RequirementId " + toSend[reqNum].RequirementId + " was not updated<p>").appendTo('#error-box');
        postNew(toSend, artifact, (reqNum + 1));
    }
    else if (reqNum < toSend.length) {
        $.ajax({
            async: true,
            method: "POST",
            crossDomain: true,
            contentType: "application/json",
            dataType: "json",
            url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects/${id}/${artifact}${userInfo.auth}`,
            data: JSON.stringify(toSend[reqNum]),
            success: function (data, textStatus, response) {
                returnId(data.RequirementId, artifact, reqNum);
                $("<p>" + toSend[reqNum].Name + " sent successfully<p>").appendTo('#error-box');
            },
            error: function () {
                $("<p>" + toSend[reqNum].Name + " failed to send<p>").appendTo('#error-box');
                enableButtons();
            }
        }).done(function (data, textStatus, response) {
            if (toSend[reqNum + 1] && toSend[reqNum + 1].hasOwnProperty(toIdString(artifact))) {
                $("<p>RequirementId " + toSend[reqNum + 1].RequirementId + " was not updated<p>").appendTo('#error-box');
                postNew(toSend, artifact, (reqNum + 2));
            }
            else {
                postNew(toSend, artifact, (reqNum + 1));
            }
        })
    } else {
        enableButtons();
    }
}

function customFieldObjCreate(valueArray){
    let newArray = valueArray.filter((val) => val != "");
    for (let i in newArray){
        let cusObj = {};
        let valType = "";
        switch (customFieldNames[i].Type){
            case "Text": valType = "StringValue";
            break;
            case "Decimal": valType = "DecimalValue";
            break;
            case "Date": valType = "DateTimeValue";
            break;
            case "MultiList": valType = "IntegerListValue";
            break;
            default: valType = "IntegerValue";
        }
        if (valType == "DateTimeValue"){
            if (!Number.isInteger(newArray[i])){
                newArray[i] = null;//if they enter an invalid date or NaN, date will default to null
            }
            else{
                newArray[i] = daysToMseconds(newArray[i]);
            }
        }
        else if (valType == "IntegerValue"){
            if (!Number.isInteger(newArray[i])){
                newArray[i] = null;
            }
        }
        else if (valType == "DecimalValue"){
            if (!$.isNumeric(newArray[i])){
                newArray[i] = null; //avoids trying to send invalid data
            }
        }
        else if (valType == "IntegerListValue"){
            newArray[i] = multilistConvert(newArray[i]);
        }
        cusObj[valType] = newArray[i];
        cusObj.PropertyNumber = parseInt(i) + 1;
        newArray[i] = cusObj;
    }
    return newArray;
}