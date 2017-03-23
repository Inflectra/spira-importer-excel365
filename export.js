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
            objTemplate[prop] = valueArray[i][j];
            j++
        }
        objArray.push(cleanObject(objTemplate));
    }
    return objArray;
}

function postNew(toSend, artifact, reqNum) {
    let id = $('#projects').val();
    if (toSend[reqNum].hasOwnProperty(toIdString(artifact))) {
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
            if (toSend[reqNum + 1].hasOwnProperty(toIdString(artifact))) {
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