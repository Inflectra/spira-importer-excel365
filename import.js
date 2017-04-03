function ajaxImport(artifact, objTemplate) {
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/'
        + artifact + userInfo.auth,
        success: function (data) {
            let valueArray = [];
            for (let i = 0; i < data.length; i++) {
                valueArray.push(jsonToArray(data[i], objTemplate));
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
        url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects/' + project + '/'
        + 'components?active_only=true&include_deleted=false&'
        + 'username=' + userInfo.username + '&api-key=' + userInfo.apikey,
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
        url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects/'
        + project + '/releases' + userInfo.auth,
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
        url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects/' + project + '/users' + userInfo.auth,
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
    for (let i = 0; i < Object.keys(objTemplate).length; i++) {
        objTemplate[Object.keys(objTemplate)[i]] = oldObj[Object.keys(objTemplate)[i]];
        valArray.push(objTemplate[Object.keys(objTemplate)[i]]);
    }
    return valArray;
}

function toExcel(artifact, newValues) {
    return Excel.run(function (context) {
        let sheetName = convertToSheetName(artifact);
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
    //in the future the following if statement will be a switch for
    //different artifact types and setting the num value
    if (artifact == "requirements") {
        artifactNum = 1;
    }
    disableButtons();

    //Get list of custom properties associated with the current project and artifact
    $.ajax({
        method: "GET",
        crossDomain: true,
        dataType: "json",
        url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects/' + project
        + '/custom-properties/' + artifactNum + userInfo.auth,
        success: function (data) {
            customFieldNames = [];
            if (data.length < 1) {
                populateCustomFieldNames([{ "Name": "" }], artifact);
                getUsers(project);
            }
            else {
                for (let i = 0; i < data.length; i++) {
                    let customFieldInfo = {};
                    customFieldInfo.Name = data[i].Name;
                    customFieldInfo.Type = data[i].CustomPropertyTypeName;
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
    for (let i = 0; i < components.length; i++){
        newComponents.push([components[i].Name, components[i].ComponentId]);
        currentComponents[components[i].Name] = components[i].ComponentId;
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
    for (let i = 0; i < Object.keys(cusObj).length; i++){
        newNames.push(cusObj[i].Name);
    }
    if (newNames.length < 30) {
        for (let i = newNames.length; i < 30; i++) {
            newNames.push("");
        }
    }
    return Excel.run(function (context) {
        let sheetName = convertToSheetName(artifact);
        let customFieldNameRange = columnRanges.customFieldRanges[artifact][0] + "2:" + columnRanges.customFieldRanges[artifact][1] + "2";
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let names = sheet.getRange(customFieldNameRange);
        names.values = [newNames];
        return context.sync();
    });
}

function populateReleases(releases){
    let releaseArray = [];
    for (let i = 0; i < releases.length; i++){
        releaseArray.push([releases[i].VersionNumber, releases[i].ReleaseId]);
        currentReleases[releases[i].VersionNumber] = releases[i].ReleaseId;
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
    for (let i = 0; i < userList.length; i++){
        userArrays.push([userList[i].FullName, userList[i].UserId]);
        currentUsers[userList[i].FullName] = userList[i].UserId;
        if (userList[i].UserName == userInfo.username){
            $('#current-user').html("Logged in as: " + userList[i].FullName);
        }
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
        let sheetName = convertToSheetName(artifact);
        let sheet = context.workbook.worksheets.getItem(sheetName);
        let IdCell = sheet.getCell(currentRow, 0);
        IdCell.values = [[newId]];
        return context.sync();
    });
}