function grabExcelValues(rows, artifact, objTemplate) {
    if (rows === null) {
        getRows(artifact, objTemplate);
    }
    else {
        return Excel.run(function (context) {
            let sheet = context.workbook.worksheets.getItem(toSheetName(artifact));
            let inputRange = columnRanges[artifact] + rows;
            let inputValues = sheet.getRange(inputRange);
            inputValues.load();
            return context.sync()
                .then(function () {
                    ajaxExport(inputValues.values, artifact, objTemplate);
                })
        });
    }
}

function getRows(artifact, objTemplate) {
    return Excel.run(function (context) {
        let sheet = context.workbook.worksheets.getItem(toSheetName(artifact));
        let sheetRange = sheet.getUsedRange();
        sheetRange.load();
        return context.sync()
            .then(function () {
                grabExcelValues(sheetRange.values.length, artifact, objTemplate);
            })
            .catch(function (error) {
                console.log(error);
            })
    });
}

function ajaxExport(valueArray, artifact, objTemplate) {
    let objArray = buildobjects(valueArray, artifact, objTemplate);
    postNew(objArray, artifact, 0);
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
                $("<p>" + toSend[reqNum].Name + " sent successfully<p>").appendTo('#error-box');
            },
            error: function () {
                $("<p>" + toSend[reqNum].Name + " failed to send<p>").appendTo('#error-box');
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

function addCustomFields(toSend, artifact, reqNum, customFieldRange){
    let customPropObj ={};
    let range = (customFieldRange[0] + (reqNum + 3)) + ":" + (customFieldRange[1] + (reqNum + 3));
    return Excel.run(function (context) {
        let sheet = context.workbook.worksheets.getItem(toSheetName(artifact));
        let customVals = sheet.getRange(range);
        customVals.load();
        return context.sync()
            .then(function () {
                customFieldObjCreate(customVals.values);
            })
            .catch(function (error) {
                console.log(error);
            })
    });
}

function customFieldObjCreate(valueArray){
    let newArray = valueArray[0].filter((val) => val != "");
    for (let i in newArray){
        newArray[i] = {
            "PropertyNumber": parseInt(i) + 1,
            "StringValue": newArray[i],
        }
    }
    console.log(newArray);
}

/*"CustomProperties": [
    {
      "BooleanValue": null,
      "DateTimeValue": null,
      "DecimalValue": null,
      "Definition": {
        "ArtifactTypeId": 1,
        "CustomList": null,
        "CustomPropertyFieldName": "Custom_01",
        "CustomPropertyId": 1,
        "CustomPropertyTypeId": 1,
        "CustomPropertyTypeName": "Text",
        "IsDeleted": false,
        "Name": "URL",
        "Options": null,
        "ProjectId": 1,
        "PropertyNumber": 1,
        "SystemDataType": "System.String"
      },
      "IntegerListValue": null,
      "IntegerValue": null,
      "PropertyNumber": 1,
      "StringValue": null
    },

let testObj = {
    "AuthorName": "Fred Bloggs",
    "Description": "haha",
    "EstimatePoints": 18.5,
    "ImportanceId": "3",
    "Name": "poo",
    "OwnerName": "Rodrigo Pereira",
    "ReleaseVersionNumber": 1,
    "RequirementTypeName": "Package",
    "StatusName": "In Progress"
}
*/