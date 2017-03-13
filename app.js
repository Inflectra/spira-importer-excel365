'use strict';
var userInfo = {
  //temporary hard coded values
  "spiraUrl": "https://demo.spiraservice.net/rodrigo-pereira/",
  "username": "administrator",
  "apikey": "{AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
};

var newRequirement = {
  "RequirementId": null,
  "Name": "test",
  "Description": "test Description",
  "ReleaseVersionNumber": null,
  "RequirementTypeName": null,
  "ImportanceName": null,
  "StatusName": null,
  "EstimatePoints": null,
  "AuthorName": null,
  "OwnerName": null,
  "ComponentName": null
};

(function () {


  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#logIn').click(logIn);
      $('#testing').click(testing);
      $('#help-toggle').click(showHelp);
      $('#export').click(getRowAmount);
    })
  };

  function logIn() {
    //userInfo.spiraUrl = $("#url").val();
    //userInfo.username = $("#username").val();
    //userInfo.apikey = $("#apikey").val();

    if (userInfo.spiraUrl.charAt(userInfo.spiraUrl.length - 1) != "/") {
      userInfo.spiraUrl += "/";
    }
    if (userInfo.spiraUrl.charAt(4) !== "s") {
      $('#url').addClass("error");
      $('#error-message').text(" Invalid URL (must be https)").addClass("ms-Icon ms-Icon--Error");
    }
    else {
      getProjects();
    }
    console.log(userInfo);
    //$("#logInScreen").addClass("hidden");

  }// end of logIn

  function getProjects() {
    var projects = undefined;
    $.ajax({
      method: "GET",
      crossDomain: true,
      url: userInfo.spiraUrl
      + "services/v5_0/RestService.svc/projects?username="
      + userInfo.username
      + "&api-key="
      + userInfo.apikey,
      success: function (data, textStatus, response) {
        console.log(response.status);
        console.log(data);
        projects = data;
        for (let i = 0; i < projects.length; i++) {
          $('<option value="' + projects[i].ProjectId + '">' + projects[i].Name + '</option>').appendTo('#projects');
        }
        $('#logInScreen').addClass("hidden");
        $('#mainScreen').removeClass("hidden");
      },
      error: function () {
        $('#username').addClass("error");
        $('#apikey').addClass("error");
        $('#error-message').text(" Invalid Username or API key").addClass("ms-Icon ms-Icon--Error");
      }
    });
  }



  function getRowAmount() {
    return Excel.run(function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const sheetRange = sheet.getUsedRange();
      sheetRange.load();
      return context.sync()
        .then(function () {
          getReqValues(sheetRange.values.length);
        })
    });
  }

  function getReqValues(rows) {
    return Excel.run(function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const inputRange = "A3:K" + rows;
      const inputValues = sheet.getRange(inputRange);
      inputValues.load();
      return context.sync()
        .then(function () {
          buildRequirementObject(inputValues.values);
        })
    });
  }

  function buildRequirementObject(pulledValues) {
    for (let i = 0; i < pulledValues.length; i++) {
      let j = 0;
      for (let prop in newRequirement) {
        newRequirement[prop] = pulledValues[i][j];
        j++
      }
      postRequirement(cleanObject(newRequirement));
    }
  }

  function postRequirement(req) {
    console.log(req);
    $.ajax({
      async: true,
      method: "POST",
      crossDomain: true,
      contentType: "application/json",
      dataType: "json",
      url: userInfo.spiraUrl
      + "services/v5_0/RestService.svc/projects/"
      + $('#projects').val()
      + "/requirements?username="
      + userInfo.username
      + "&api-key="
      + userInfo.apikey,
      data: JSON.stringify(req)
    }).done(function (data) {
      console.log("sent", data);
    })
  }

  function cleanObject(Obj) {
    var cleaned = {};
    for (let prop in Obj) {
      if (Obj[prop] != "") {
        cleaned[prop] = Obj[prop];
      }
    }
    return cleaned;
  }

  function showHelp(){
    $('#chevron-icon').toggleClass("ms-Icon--ChevronRight");
    $('#chevron-icon').toggleClass("ms-Icon--ChevronDown");
    $('#help-text').toggleClass("hidden");
  }

  //for testing calls and functions with a temporary "test" button on index.html
  function testing() {
  }//end of testing
})();
