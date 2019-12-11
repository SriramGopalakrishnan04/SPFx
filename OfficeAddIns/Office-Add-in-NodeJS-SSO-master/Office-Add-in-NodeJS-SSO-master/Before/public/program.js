// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
	 $("#getGraphAccessTokenButton").click(function () {
                getOneDriveFiles();
            });
    });
}

var timesGetOneDriveFilesHasRun = 0;
var triedWithoutForceConsent = false;
var timesMSGraphErrorReceived = false;

function getOneDriveFiles() {
    timesGetOneDriveFilesHasRun++;
    triedWithoutForceConsent = true;
    getDataWithToken({ forceConsent: false });
}

function getData(relativeUrl, accessToken) {
    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
    .done(function (result) {
        showResult(result);
    })
    .fail(function (result) {
        handleServerSideErrors(result);
    });
}

function handleClientSideErrors(result) {

    switch (result.error.code) {

        // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
        //        prompt to provide a 2nd authentication factor.
        case 13001:
    getDataWithToken({ forceAddAccount: true });
    break;

        // TODO3: Handle the case where the user's sign-in or consent was aborted.
        case 13002:
    if (timesGetOneDriveFilesHasRun < 2) {
        showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
    } else {
        logError(result);
    }
    break;

        // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
        //        nor Microsoft Account.
        case 13003:
    showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
    break;

        // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
        //        the user has not granted the service permission to their `profile`.
        case 13005:
    getDataWithToken({ forceConsent: true });
    break;

        // TODO6: Handle an unspecified error from the Office host.
        case 13006:
    showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
    break;

        // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
        //        web service/application.
        case 13007:
    showResult(['That operation cannot be done at this time. Please try again later.']);
    break;

        // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
        //        before a previous call of it completed.
        case 13008:
    showResult(['Please try that operation again after the current operation has finished.']);
    break;

        // TODO9: Handle the case where the add-in does not support forcing consent.
        case 13009:
    if (triedWithoutForceConsent) {
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
    } else {
        getDataWithToken({ forceConsent: false });
    }
    break;

        // TODO10: Log all other client errors.
        default:
    logError(result);
    break;
    }
}

function handleServerSideErrors(result) {

    // TODO11: Handle the case where AAD asks for an additional form of authentication.
    if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 50076){
    getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
}

    // TODO12: Handle the case where consent has not been granted, or has been revoked.
    else if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 65001){
    getDataWithToken({ forceConsent: true });
}

    // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow
    else if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 70011){
    showResult(['The add-in is asking for a type of permission that is not recognized.']);
}

    // TODO14: Handle the case where the token that the add-in's client-side sends to its
    //         server-side is not valid because it is missing `access_as_user` scope (permission).
    else if (result.responseJSON.error.name
        && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
    showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
}

    // TODO15: Handle the case where the token sent to Microsoft Graph in the request for
    //         data is expired or invalid.
    else if (result.responseJSON.error.name
        && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
    if (!timesMSGraphErrorReceived) {
        timesMSGraphErrorReceived = true;
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    } else {
        logError(result);
    }
}

    // TODO16: Log all other server errors.
    else {
        logError(result);
    }
}

function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                accessToken = result.value;
                getData("/api/values", accessToken);
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }

// Displays the data, assumed to be an array.
function showResult(data) {

    // Note that in this sample, the data parameter is an array of OneDrive file/folder
    // names. Encoding/sanitizing to protect against Cross-site scripting (XSS) attacks
    // is not needed because there are restrictions on what characters can be used in 
    // OneDrive file and folder names. These restrictions do not necessarily apply 
    // to other kinds of data including other kinds of Microsoft Graph data. So, to 
    // make this method safely reusable in other contexts, it uses the jQuery text() 
    // method which automatically encodes values that are passed to it.
	$.each(data, function (i) {
        var li = $('<li/>').addClass('ms-ListItem').appendTo($('#file-list'));
        var outerSpan = $('<span/>').addClass('ms-ListItem-secondaryText').appendTo(li);
        $('<span/>').addClass('ms-fontColor-themePrimary').appendTo(outerSpan).text(data[i]);
      });
}

function logError(result) {

    // Error messages can have a variety of structures depending on the ultimate
    // ultimate source and how intervening code restructures it before relaying it.
    console.log("Status: " + result.status);
    if (result.error.code) {
        console.log("Code: " + result.error.code);
    }
    if (result.error.name) {
        console.log("Code: " + result.error.name);
    }
    if (result.error.message) {
        console.log("Code: " + result.error.message);
    }
    if (result.responseJSON.error.name) {
        console.log("Code: " + result.responseJSON.error.name);
    }
    if (result.responseJSON.error.name) {
        console.log("Code: " + result.responseJSON.error.name);
    }       
}
