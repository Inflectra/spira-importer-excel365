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

function getComponents(project){
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects/${project}/`
        + `components?active_only=true&include_deleted=false&`
        +`username=${userInfo.username}&api-key=${userInfo.apikey}`,
        success: function (data) {
            populateComponents(data);
            enableButtons();
        },
        error: function () {
            console.log("error retrieving components");
            populateComponents([]);
            enableButtons();
        }
    });
}

function getReleases(project){
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects/${project}/releases${userInfo.auth}`,
        success: function (data) {
            populateReleases(data);
            getComponents(project);
        },
        error: function () {
            console.log("error retrieving releases");
            populateReleases([]);
            getComponents(project);
        }
    });
}

function getUsers(project){
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects/${project}/users${userInfo.auth}`,
        success: function (data) {
            populateUsers(data);
            getReleases(project);
        },
        error: function () {
            console.log("error retrieving users");
            populateUsers([]);
            getReleases(project);
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
    if (project == -1){
        return null;
    }
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
                getUsers(project);
            }
            else {
                for (let customProp of data) {
                    customFieldInfo = {};
                    customFieldInfo.Name = customProp.Name;
                    customFieldInfo.Type = customProp.CustomPropertyTypeName;
                    customFieldNames.push(customFieldInfo);
                }
                populateCustomFieldNames(customFieldNames, artifact);
                getUsers(project);
            }
        },
        error: function () {
            populateCustomFieldNames([{ "Name": "" }], artifact);
            getUsers(project);
        }
    });
}

function populateComponents(components){
    let newComponents = [];
    for (let component of components){
        newComponents.push([component.Name, component.ComponentId]);
        currentComponents[component.Name] = component.ComponentId;
    }
    return Excel.run(function (context) {
        let componentRange = "I3:J" + (components.length + 2);
        let sheet = context.workbook.worksheets.getItem("Lookups");
        let clearRange = sheet.getRange("I3:J10000").clear();
        let componentList = sheet.getRange(componentRange);
        if (components.length < 1){
            clearRange;
        }
        else{
            componentList.values = newComponents;
        }
        return context.sync();
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
        return context.sync();
    });
}

function populateReleases(releases){
    let releaseArray = [];
    for (let release of releases){
        releaseArray.push([release.VersionNumber, release.ReleaseId]);
        currentReleases[release.VersionNumber] = release.ReleaseId;
    }
    return Excel.run(function (context) {
        let releaseRange = "E3:F" + (releases.length + 2);
        let sheet = context.workbook.worksheets.getItem("Lookups");
        let clearRange = sheet.getRange("E3:F10000").clear();
        let releaseList = sheet.getRange(releaseRange);
        if (releases.length < 1){
            clearRange;
        }
        else{
            releaseList.values = releaseArray;
        }
        return context.sync();
    });

}

function populateUsers(userList){
    let userArrays = [];
    for (let user of userList){
        userArrays.push([user.FullName, user.UserId]);
        currentUsers[user.FullName] = user.UserId;
    }
    return Excel.run(function (context) {
        let userRange = "G3:H" + (userList.length + 2);
        let sheet = context.workbook.worksheets.getItem("Lookups");
        let clearRange = sheet.getRange("G3:H10000").clear();
        let userLookup = sheet.getRange(userRange);
        if (userList.length < 1){
            clearRange;
        }
        else{
            userLookup.values = userArrays;
        }
        return context.sync();
    });

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