'use strict';
var userInfo = {
    //temporary hard coded values
    "spiraUrl": "https://demo.spiraservice.net/rodrigo-pereira/",
    "username": "administrator",
    "apikey": "{AA50F584-BBC9-42A0-81BA-9F8A5CD8144A}",
};

var authParams = "";

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

var projects = undefined;

(function () {


    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#logIn').click(function () {
                if (logIn == true) {
                    $('#logInScreen').addClass("hidden");
                    $('#mainScreen').removeClass("hidden");
                }
            });
            $('#testing').click(testing);
            $('#help-toggle').click(showHelp);
            $('#export').click(getRowAmount);
        })
    };

    function logIn(userObj) {
        authParams = "?username="
        + userObj.username
        + "&api-key="
        + userObj.apikey;
        return getProjects(authParams);
    }
    
});