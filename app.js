'use strict';

var userInfo = {
  //temporary hard coded values
  "spiraUrl": "https://demo.spiraservice.net/rodrigo-pereira/",
  "username": "administrator",
  "apikey": "{AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
  "auth": "?username=administrator&api-key={AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
};


//template objects for correct order in sheet, currently requirement
//template only

var requirementObj = {
  "RequirementId": null,
  "Name": null,
  "Description": null,
  "ReleaseVersionNumber": null,
  "RequirementTypeName": null,
  "ImportanceName": null,
  "StatusName": null,
  "EstimatePoints": null,
  "AuthorName": null,
  "OwnerName": null,
  "ComponentName": null
};

//column ranges for different sheets, currently only requirements.
var columnRanges = {
  requirements : "A3:K"
};

function cleanObject(Obj) {
    var cleaned = {};
    for (let prop in Obj) {
      if (Obj[prop] != "") {
        cleaned[prop] = Obj[prop];
      }
    }
    return cleaned;
  }

function disableButtons() {
    $('button').attr('disabled', 'disabled');
    $('button').addClass('is-disabled');
  }

  function enableButtons() {
    $('button').removeClass('is-disabled');
    $('button').prop('disabled', false);
  }

  function toIdString(artifact){
    let newString = toSheetName(artifact);
    newString = newString.split(" ");
    newString = newString.join("");
    newString = newString.split("");
    newString[newString.length - 1] = "Id";
    newString = newString.join("");
    return newString;
  }

  function toSheetName(artifact) {
    let newString = artifact.split("-");
    for (let i in newString){
      newString[i] = newString[i].charAt(0).toUpperCase() + newString[i].substr(1);
    }
    newString = newString.join(" ");
    return newString;
  }

(function () {

  function logIn() {
    //userInfo.spiraUrl = $("#url").val();
    //userInfo.username = $("#username").val();
    //userInfo.apikey = $("#apikey").val();
    //userInfo.auth = "?username=" + $("#username").val() + "&api-key=" + $("#apikey").val();

    if (userInfo.spiraUrl.charAt(userInfo.spiraUrl.length - 1) != "/") {
      userInfo.spiraUrl += "/";
    }
    if (userInfo.spiraUrl.charAt(4) !== "s") {
      $('#url').addClass("error");
      $('#error-message').text(" Invalid URL (must be https)").addClass("ms-Icon ms-Icon--Error");
    } else {

      getProjects();
    };

  } // end of logIn

  function getProjects() {
    $.ajax({
      method: "GET",
      crossDomain: true,
      url: `${userInfo.spiraUrl}services/v5_0/RestService.svc/projects${userInfo.auth}`,
      success: function (data, textStatus, response) {
        for (let i = 0; i < data.length; i++) {
          $('<option value="' + data[i].ProjectId + '">' + data[i].Name + '</option>').appendTo('#projects');
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

  function showHelp() {
    $('#chevron-icon').toggleClass("ms-Icon--ChevronRight");
    $('#chevron-icon').toggleClass("ms-Icon--ChevronDown");
    $('#help-text').toggleClass("hidden");
  }

  //for testing calls and functions with a temporary "test" button on index.html
  function testing() { } //end of testing

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#logIn').click(logIn);
      $('#testing').click(testing);
      $('#help-toggle').click(showHelp);

      $('#export').click(() => {
        $('#projects').removeClass('error');
        disableButtons();
        var selectedProject = $('#projects').val();
        if (selectedProject != -1) {
          //getRowAmount();
          switch ($('#artifact').val()) {
          case "requirements":
            grabExcelValues(null, $('#artifact').val(), requirementObj);
            break;
          default:
            console.log("Could not export");
        }
        } else {
          $('#projects').addClass('error');
          enableButtons();
        }
      });

      $('#import').click(() => {
        disableButtons();
        switch ($('#artifact').val()) {
          case "requirements":
            ajaxImport($('#artifact').val(), requirementObj);
            break;
          default:
            console.log("Could not import");
        }
      });
    })
  };
})();