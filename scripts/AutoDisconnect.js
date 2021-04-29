///////////////////
//    README     //
///////////////////

/*
After you've joined a Meeting press Ctrl+Shift+I and paste the whole code (Ctrl+A) into the console section of the window and press enter.
The script will automatically press the disconnect button after the current number of people in the meeting is lower then the below set threshold

If you want to manually stop the script, write 'cancelAction()' then press enter in the console window.
If you want to restart the script, write 'action()' then press enter in the console window.
*/

///////////////////
//   SETTINGS   ///
///////////////////

// If people currently in meeting is less then this number, the script auto-disconnects.
var threshold = 8;

// Delay that the script checks count after. (in milliseconds)
var delay = 1000;

///////////////////

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

var dotask = true;
async function action() {
    dotask = true;
    console.clear()
    while(dotask) {
        var sum = 0;
        // Open sidepanel if necessary
        var button = document.getElementById("roster-button");
        if(button.getAttribute("track-outcome") == "15") {
            button.click();
        }
        // Count users (Sidepanel must be open) 
        var elements = document.getElementsByTagName("calling-roster-section");
        for (var i = 0; i < elements.length; i++) {
            if(elements[i].getAttribute("participants-list") == "ctrl.participantsInCall" || elements[i].getAttribute("participants-list") == "ctrl.attendeesInMeeting") {
                var childrens = elements[i].children[0].children[0].getElementsByTagName("button");
                for (var j = 0; j < childrens.length; j++) {
                    if(childrens[j].classList.contains("roster-list-title")) {
                        // Append the count into the sum.
                        sum += parseInt(childrens[j].children[2].innerHTML.replace("(","").replace(")",""));
                    }
                }
            }
        }
        // Log current users into the console window.
        console.log("Users in meeting: " + sum + "\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        if(sum <= threshold && sum != 0) {
            // Disconnect because the sum is lower then the set threshold
            document.getElementById("hangup-button").click();
            dotask = false;
            console.warn("Threshold is " + threshold + ", disconnecting at " + sum + " attendees...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        }
        // Wait
        await sleep(delay);
    }
}
async function cancelAction() {
    // Stop the script
    dotask = false;
}
action();
