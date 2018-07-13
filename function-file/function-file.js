Office.initialize = function () {
    $(document).ready(function () {
        var item = Office.context.mailbox.item;
    });
}

var clickEventGlobal;

function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon,
        message: text,
        persistent: false
    });
}

function statusUpdateError(text) {
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
        type: "errorMessage",
        message: text
    });
    if (typeof clickEventGlobal.completed !== "undefined") {
        clickEventGlobal.completed();
    }

}

/*
Office.initialize = function (reason) {
    $(document).ready(function () {
        var item = Office.context.mailbox.item;
    });
};
*/

function parseEvent(event, format) {
    try {
        var entryPoints = event.conferenceData.entryPoints;
        var returnData = "---";
        var lineBreak;
        var isMeet = true;
        if (format == "html") {
            lineBreak = "<br>";
            returnData += lineBreak + "<b>Please join the Google Hangouts session using the information below.</b>";
            entryPoints.forEach(function (entryPoint) {
                if (entryPoint.entryPointType == "video") {
                    if (entryPoint.label != undefined) {
                        url = "https://" + entryPoint.label;
                        returnData += lineBreak + "<b>Video: </b><a href='" + url + "'>" + url + "</a>" + lineBreak;
                    }
                    else{
                        returnData = "---<br>Error: This plugin only supports Google Hangouts Meet.  Please talk to your administrator about enabling this on your account.<br>--";
                    }
                } else if (entryPoint.entryPointType == "phone") {
                    returnData += "<b>Phone: </b><a href='tel:" + entryPoint.label + "'>" + entryPoint.label + "</a>" + lineBreak;
                    returnData += "<b>PIN: </b>" + entryPoint.pin + lineBreak;
                }
            });
        } else if (format == "text") {
            lineBreak = "\n";
            returnData += lineBreak + "Please join the Google Hangouts session using the information below.";
            entryPoints.forEach(function (entryPoint) {
                if (entryPoint.entryPointType == "video") {
                    url = "https://" + entryPoint.label;
                    returnData += lineBreak + "Video: " + url + lineBreak;
                } else if (entryPoint.entryPointType == "phone") {
                    returnData += "Phone: " + entryPoint.label + lineBreak;
                    returnData += "PIN: " + entryPoint.pin + lineBreak;
                }
            });
        }

        returnData += "---" + lineBreak + lineBreak + lineBreak + lineBreak;
        return returnData;
    } catch (err) {
        statusUpdateError(err.toString().substring(0, 150));
    }
}





function prependItemBody(item, event) {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    //item.body.append("Test", "test"); 



                    item.body.prependAsync(

                        parseEvent(event, "html"), {
                            //event.conferenceData.toString(), {    
                            //'<b>Greetings!</b>', {
                            coercionType: Office.CoercionType.Html,
                            asyncContext: {
                                var3: 1,
                                var4: 2
                            }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            } else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                } else {
                    // Body is of text type.
                    item.body.prependAsync(

                        parseEvent(event, "text"), {
                            coercionType: Office.CoercionType.Text,
                            asyncContext: {
                                var3: 1,
                                var4: 2
                            }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            } else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(appointment) {
    document.getElementById('appointment').innerText += appointment;
}

function makeid() {
    var text = "";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for (var i = 0; i < 10; i++)
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}
// Client ID and API key from the Developer Console
var CLIENT_ID = '<your_client_id_here>';
var API_KEY = '<your_api_key_here>';

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = "https://www.googleapis.com/auth/calendar";

/*var authorizeButton = document.getElementById('authorize-button');
var signoutButton = document.getElementById('signout-button');
var createButton = document.getElementById('create-button');
*/


/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
    gapi.load('client:auth2', initClient);
}

/**
 *  Initializes the API client library and sets up sign-in state
 *  listeners.
 */
function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        clientId: CLIENT_ID,
        discoveryDocs: DISCOVERY_DOCS,
        scope: SCOPES
    }).then(function () {
        // Listen for sign-in state changes.
        gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

        // Handle the initial sign-in state.
        updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());

        try {
            if (!gapi.auth2.getAuthInstance().isSignedIn.get()) {
                handleAuthClick();
                if (!gapi.auth2.getAuthInstance().isSignedIn.get()) {
                    console.log("Error: Not signed in to Google Hangouts.  Please sign in.");
                    //statusUpdateError("Error: Not signed in to Google Hangouts.  Please sign in.");
                    if (typeof clickEventGlobal.completed !== "undefined") {
                        clickEventGlobal.completed();
                    }
                } else {
                    CreateEvent(clickEventGlobal);
                }
            } else {
                CreateEvent(clickEventGlobal);
            }
        } catch (err) {
            console.log("Error inserting meeting.  Error: " + err.toString());
            statusUpdateError("Error inserting meeting.  Error: " + err.toString());

        }

    });
}

function updateSigninStatus(isSignedIn) {
    if (isSignedIn) {
        console.log("Signed in to Google Hangouts Meet");
    } else {
        //console.log("Error: Not signed in to Google Hangouts.  Please sign in.");
        //statusUpdateError("Error: Not signed in to Google Hangouts.  Please sign in.");
        if (typeof clickEventGlobal.completed !== "undefined") {
            clickEventGlobal.completed();
        }
    }
}

function handleAuthClick(event) {
    gapi.auth2.getAuthInstance().signIn();
    //statusUpdate("blue-icon-16", "Signed In to Google Hangouts Meet");

}

function handleSignoutClick(event) {

    gapi.auth2.getAuthInstance().signOut();
    gapi.auth2.getAuthInstance().disconnect();

    statusUpdate("blue-icon-16", "Signed Out of Google Hangouts Meet");

}

function CreateEvent(clickEvent) {

    var tmpStart = new Date();
    tmpStart.setHours(tmpStart.getHours() + 1);
    var tmpEnd = new Date();
    tmpEnd.setHours(tmpStart.getHours() + 2);

    var event = {
        //'summary': subject,
        'summary': 'Google Hangouts Meet Session',
        'location': 'Google Hangouts Meet',
        'description': 'Video/Phone conference via Google Hangouts Meet',
        'conferenceSolutionKey': {
            'type': 'hangoutsMeet'
        },
        'start': {
            //'dateTime': '2018-06-28T09:00:00-07:00',
            //'dateTime': start,
            'dateTime': tmpStart.toISOString(),
        },
        'end': {
            //'dateTime': '2018-06-28T17:00:00-07:00',
            //'dateTime': end,
            'dateTime': tmpEnd.toISOString(),
        },
        'conferenceData': {
            'createRequest': {
                requestId: makeid()
            }
        },
        'reminders': {
            'useDefault': false,
            'overrides': [{
                    'method': 'email',
                    'minutes': 24 * 60
                },
                {
                    'method': 'popup',
                    'minutes': 10
                }
            ]
        }
    };

    try {
        var request = gapi.client.calendar.events.insert({
            'calendarId': 'primary',
            'conferenceDataVersion': 1,
            'resource': event
        });
    } catch (err) {
        console.log(err);
    }

    var eventId;

    try {
        request.execute(function (event) {
            eventId = event.id;
            prependItemBody(Office.context.mailbox.item, event);

            var deleteRequest = gapi.client.calendar.events.delete({
                'calendarId': 'primary',
                'eventId': eventId
            });

            deleteRequest.execute(function (eventDeleted) {

                if (typeof clickEvent.completed !== "undefined") {
                    clickEvent.completed();
                }

            });

        });
    } catch (err) {
        console.log(err);
    }
}

function getMeetingDetails(item, clickEvent) {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                // Successfully got the subject, display it.
                var subject = asyncResult.value;
                var start;
                var end;
                item.start.getAsync(
                    function (asyncResult2) {
                        if (asyncResult2.status == Office.AsyncResultStatus.Failed) {
                            write(asyncResult2.error.message);
                        } else {
                            // Successfully got the start time, display it.
                            start = asyncResult2.value;
                            item.end.getAsync(
                                function (asyncResult3) {
                                    if (asyncResult3.status == Office.AsyncResultStatus.Failed) {
                                        write(asyncResult3.error.message);

                                    } else {
                                        // Successfully got the end time, display it.

                                        end = asyncResult3.value;
                                        CreateEvent(subject, start, end, clickEvent);
                                    }
                                });
                        }
                    });
            }
        });
}

function createOnClick(clickEvent) {
    clickEventGlobal = clickEvent;
    handleClientLoad();
    /*
    try{
        getMeetingDetails(Office.context.mailbox.item, clickEvent);
    } catch(err){
        console.log("Error inserting meeting.  Error: " + err.toString());
        statusUpdateError("Error inserting meeting.  Error: " + err.toString());

    }*/

}


//})();
