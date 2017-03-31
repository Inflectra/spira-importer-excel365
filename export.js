function grabExcelValues(rows, artifact, objTemplate, customFieldRange) {
    if (rows === null) {
        getRows(artifact, objTemplate, customFieldRange);
    }
    else {
        return Excel.run(function (context) {
            let sheet = context.workbook.worksheets.getItem(convertToSheetName(artifact));
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
        let sheet = context.workbook.worksheets.getItem(convertToSheetName(artifact));
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
    for (let i = 0; i < customFieldObjArr.length; i++){
        objArray[i].CustomProperties = customFieldObjArr[i];
    }
    postNew(objArray, artifact, 0);
    console.log(objArray);
}

function buildobjects(valueArray, artifact, objTemplate) {
    /*valueArray is the values pulled from the Excel row,
    artifact is the artifact name and objTemplate
    is a pre-made default object with the keys in the correct
    order for iterating over along with valueArray*/
    let objArray = [];
    for (let i = 0; i < valueArray.length; i++) {
        let j = 0;
        for (let k = 0; k < Object.keys(objTemplate).length; k++) {
            newObject = objTemplate;
            //grabs the digit from Importance Name field for the id
            if (Object.keys(newObject)[k] == "ImportanceId") {
                console.log(valueArray[i][j]);
                newObject[Object.keys(newObject)[k]] = valueArray[i][j].toString().charAt(0);
            }
            else if (Object.keys(newObject)[k] == "ReleaseId"){
                newObject[Object.keys(newObject)[k]] = currentReleases[valueArray[i][j]];
            }
            else if ((Object.keys(newObject)[k] == "AuthorId") || (Object.keys(newObject)[k] == "OwnerId")){
                newObject[Object.keys(newObject)[k]] = currentUsers[valueArray[i][j]];
            }
            else if (Object.keys(newObject)[k] == "RequirementTypeId"){
                newObject[Object.keys(newObject)[k]] = requirementType[valueArray[i][j]];
            }
            else if (Object.keys(newObject)[k] == "StatusId"){
                newObject[Object.keys(newObject)[k]] = reqStatus[valueArray[i][j]];
            }
            else if (Object.keys(newObject)[k] == "ComponentId"){
                newObject[Object.keys(newObject)[k]] = currentComponents[valueArray[i][j]];
            }
            else {
                newObject[Object.keys(newObject)[k]] = valueArray[i][j];
            }
            j++
        }
        objArray.push(cleanObject(newObject));
    }
    return objArray;
}

function postNew(toSend, artifact, rowNum) {
    let id = $('#projects').val();
    //Check to make sure object doesn't already have RequirementId and move on to 
    //the next row if it does
    if (toSend[rowNum] && toSend[rowNum].hasOwnProperty(convertToIdKey(artifact))) {
        $("<p>RequirementId " + toSend[rowNum].RequirementId + " was not updated<p>").appendTo('#error-box');
        postNew(toSend, artifact, (rowNum + 1));
    }
    else if (rowNum < toSend.length) {
        $.ajax({
            async: true,
            method: "POST",
            crossDomain: true,
            contentType: "application/json",
            dataType: "json",
            url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects/' + id + '/' + artifact + userInfo.auth,
            data: JSON.stringify(toSend[rowNum]),
            success: function (data, textStatus, response) {
                returnId(data.RequirementId, artifact, rowNum);
                $("<p>" + toSend[rowNum].Name + " sent successfully<p>").appendTo('#error-box');
            },
            error: function () {
                $("<p>" + toSend[rowNum].Name + " failed to send<p>").appendTo('#error-box');
                enableButtons();
            }
        }).done(function (data, textStatus, response) {
            if (toSend[rowNum + 1] && toSend[rowNum + 1].hasOwnProperty(convertToIdKey(artifact))) {
                $("<p>RequirementId " + toSend[rowNum + 1].RequirementId + " was not updated<p>").appendTo('#error-box');
                postNew(toSend, artifact, (rowNum + 2));
            }
            else {
                console.log(toSend[rowNum + 1]);
                postNew(toSend, artifact, (rowNum + 1));
            }
        })
    } else {
        enableButtons();
    }
}

function customFieldObjCreate(valueArray){
    let newArray = valueArray.filter(function(val) {return val != ""});
    for (let i = 0; i < newArray.length; i++){
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