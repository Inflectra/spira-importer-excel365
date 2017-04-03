'use strict';

var userInfo = {
  //temporary hard coded values
  "spiraUrl": "https://demo.spiraservice.net/rodrigo-pereira/",
  "username": "administrator",
  "apikey": "{AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
  "auth": "?username=administrator&api-key={AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
};

var currentComponents = {

};

var currentUsers = {
  //relevant users will be stored here after selecting a project
};

var currentReleases = {
  //relevant releases will be stored here after selecting a project
};
//template objects for correct order in sheet, currently requirement
//template only

var requirementType = {
  "Package": -1,
  "Need": 1,
  "Feature": 2,
  "Use Case": 3,
  "User Story": 4,
  "Quality": 5,
  "Design Element": 6,
};

var requirementObj = {
  "RequirementId": null,
  "Name": null,
  "Description": null,
  "ReleaseId": null,
  "RequirementTypeId": null,
  "ImportanceId": null,
  "StatusId": null,
  "EstimatePoints": null,
  "AuthorId": null,
  "OwnerId": null,
  "ComponentId": null,
};

var reqStatus = {
  "Requested": 1,
  "Evaluated": 7,
  "Accepted": 5,
  "Rejected": 6,
  "Planned": 2,
  "In Progress": 3,
  "Developed": 4,
  "Obsolete": 8,
  "Tested": 9,
  "Completed": 10,
};

//column ranges for different sheets, currently only requirements.
var columnRanges = {
  requirements: "A3:K",
  customFieldRanges: {
    requirements: ["N", "AQ"],
  }
};

//array to hold custom field names
var customFieldNames = [];

function cleanObject(Obj) {
  var cleaned = {};
  for (let i = 0; i < Object.keys(Obj).length; i++) {
    if (Obj[Object.keys(Obj)[i]] != "") {
      cleaned[Object.keys(Obj)[i]] = Obj[Object.keys(Obj)[i]];
    }
  }
  return cleaned;
}

//converts from Excel days since 1/1/1990 to Spira milliseconds since 1/1/1970
function daysToMseconds(days) {
  days -= 2; //for some reason excel returns 2 extra days?
  const between = 25566;
  days -= between;
  let milliseconds = days *= 8.64e+7;
  return '/Date(' + milliseconds + ')/';
}

function disableButtons() {
  $('button').attr('disabled', 'disabled');
  $('button').addClass('is-disabled');
  $('select').attr('disabled', 'disabled');
  $('select').addClass('is-disabled');
}

function enableButtons() {
  $('button').removeClass('is-disabled');
  $('button').prop('disabled', false);
  $('select').removeClass('is-disabled');
  $('select').prop('disabled', false);
}

function multilistConvert(str) {
  function correct(str) {
    let newString = parseInt(str.trim());
    return newString;
  }
  let arr = str.split(",");
  arr = arr.map(correct);
  arr = arr.filter(Number.isInteger);
  return arr;
}

//converts artifact name into a string ending with "Id" to be used
//for accessing the correct keys in objects (ex. turns "requirements" into "RequirementId")
function convertToIdKey(artifactName) {
  let newString = convertToSheetName(artifactName);
  newString = newString.split(" ");
  newString = newString.join("");
  newString = newString.split("");
  newString[newString.length - 1] = "Id";
  newString = newString.join("");
  return newString;
}

//converts artifact name into the correct format for the Excel template sheet names
function convertToSheetName(artifactName) {
  let newString = artifactName.split("-");
  for (let i = 0; i < newString.length; i++) {
    newString[i] = newString[i].charAt(0).toUpperCase()
      + newString[i].substr(1);
  }
  newString = newString.join(" ");
  return newString;
}

(function () {

  function logIn() {
    disableButtons();
    //userInfo.spiraUrl = $("#url").val();
    //userInfo.username = $("#username").val();
    //userInfo.apikey = $("#apikey").val();
    //userInfo.auth = "?username=" + $("#username").val() + "&api-key=" + $("#apikey").val();

    if (userInfo.spiraUrl.charAt(userInfo.spiraUrl.length - 1) != "/") {
      userInfo.spiraUrl += "/";
      enableButtons();
    }
    if (userInfo.spiraUrl.charAt(4) !== "s") {
      $('#url').addClass("error");
      $('#error-message').html('<p class="ms-baseFont">Invalid URL (must be https)</p>');
      enableButtons();
    } else {

      getProjects();
    };

  } // end of logIn

  function getProjects() {
    $.ajax({
      method: "GET",
      crossDomain: true,
      url: userInfo.spiraUrl + 'services/v5_0/RestService.svc/projects' + userInfo.auth,
      success: function (data, textStatus, response) {
        for (let i = 0; i < data.length; i++) {
          $('<option value="' + data[i].ProjectId + '">' + data[i].Name + '</option>').appendTo('#projects');
        }
        $('#current-user').html("Logged in as: " + userInfo.username);
        $('#logInScreen').addClass("hidden");
        $('#mainScreen').removeClass("hidden");
        enableButtons();
      },
      error: function () {
        $('#username').addClass("error");
        $('#apikey').addClass("error");
        $('url').addClass("error");
        $('#error-message').html("Invalid Login Info");
        enableButtons();
      }
    });
  }

  function showHelp() {
    $('#chevron-icon').toggleClass("ms-Icon--ChevronRight");
    $('#chevron-icon').toggleClass("ms-Icon--ChevronDown");
    $('#help-text').toggleClass("hidden");
  }

  //for testing calls and functions with a temporary "test" button on index.html
  function testing() {
    return Excel.run(function (context) {
      let sheetName = convertToSheetName("requirements");
      let sheet = context.workbook.worksheets.getItem(sheetName);
      let testCell = sheet.getCell(3, 0);
      testCell.load();
      return context.sync()
        .then(function () {
          let val = testCell.values[0][0];
          val = multilistConvert(val);
        });
    });
  } //end of testing

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#logIn').click(logIn);
      $('#testing').click(testing);
      $('#help-toggle').click(showHelp);
      $('#clear-log').click(function () {
        $(this).addClass("hidden");
        $('#log-box').html('');
      });

      $('#export').click(function () {
        $('#clear-log').removeClass("hidden");
        disableButtons();
        $('#projects').removeClass('error');
        var selectedProject = $('#projects').val();
        if (selectedProject != -1) {
          switch ($('#artifact').val()) {
            case "requirements":
              grabExcelValues(null, $('#artifact').val(), requirementObj, columnRanges.customFieldRanges.requirements);
              break;
            default:
              console.log("Could not export");
          }
        } else {
          $('#projects').addClass('error');
          enableButtons();
        }
      });

      $('#import').click(function () {
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