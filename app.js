'use strict';
var userInfo = {
  //temporary hard coded values
  "spiraUrl": "https://demo.spiraservice.net/rodrigo-pereira/",
  "username": "administrator",
  "apikey": "{AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
};

var projects = undefined;

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
        getProjects();
        $('#logIn').click(logIn);
        $('#testing').click(testing);
      });
    };

    function logIn() {

      userInfo.spiraUrl = $("#url").val();
      userInfo.username = $("#username").val();
      userInfo.apikey = $("#apikey").val();
      //$("#logInScreen").addClass("hidden");

    }// end of logIn

    function getProjects() {
      $.ajax({
        method: "GET",
        crossDomain: true,
        url: userInfo.spiraUrl
        + "services/v5_0/RestService.svc/projects?username="
        + userInfo.username
        + "&api-key="
        + userInfo.apikey,
      }).done(function (data) {
        projects = data;
        for (let i = 0; i < projects.length; i++) {
          $('<option value="' + projects[i].ProjectId + '">' + projects[i].Name + '</option>').appendTo('#projects');
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

    function buildRequirementObject(pulledValues){
      for (let i = 0; i < pulledValues.length; i++){
        let j = 0;
        for (let prop in newRequirement){
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

    function cleanObject(Obj){
      var cleaned = {};
      for (let prop in Obj){
        if (Obj[prop] != ""){
          cleaned[prop] = Obj[prop];
        }
      }
      return cleaned;
    }

    //for testing calls and functions with a temporary "test" button on index.html
    function testing() {
      getRowAmount();
    }//end of testing
  })();
