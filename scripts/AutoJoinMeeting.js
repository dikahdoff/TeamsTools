///////////////////
//    README     //
///////////////////

/*
After you've joined a Meeting press Ctrl+Shift+I and paste the whole code (Ctrl+A) into the console section of the window and press enter.


If you want to manually stop the script, write 'cancelJoining()' then press enter in the console window.
If you want to restart the script, write 'joiningAction()' then press enter in the console window.
*/

///////////////////
//   SETTINGS   ///
///////////////////

// Delay that the script acts after. (in milliseconds)
var joinDelay = 1000;

// Delay to wait before joining. (in milliseconds)
var joinWait = 5000;

///////////////////

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Loop
var dojoining = true;
async function joiningAction() {
    dojoining = true;
    console.clear()
    while(dojoining) {
        console.log("Fetching meetings...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        // Get meeting button
        var element = document.getElementsByClassName("ts-sym ts-btn ts-btn-primary inset-border icons-call-jump-in ts-calling-join-button app-title-bar-button app-icons-fill-hover call-jump-in");
        if(element.length > 0) {
            element = element[0];
            // Validated, a meeting is running and the join button is available.
            // Wait for 'joinWait' and then click the join button
            console.log("Found a meeting, joining in " + joinWait + "ms\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            await sleep(joinWait);
            console.log("Joining...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            element.click();
            await sleep(50);
            // Check if the continue without camera or audio dialog opens
            var withoutBtn = document.getElementsByClassName("ts-btn ts-btn-fluent ts-btn-fluent-secondary-alternate");
            if(withoutBtn.length > 0) {
                // If it did, close it.
                withoutBtn[0].click();
                console.warn("Closed warning box.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            }
            await sleep(50);
            var toggleMic = document.getElementById("preJoinAudioButton");
            // Check if toggle mic exists
            if(toggleMic != null) {
                // Check if mic is on, then turn off if it is.
                if(toggleMic.getAttribute("telemetry-outcome") == "30") {
                    toggleMic.children[0].children[0].click();
                    console.warn("Turned microphone off.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                }
            }
            await sleep(50);
            // Check if toggle camera exists
            var toggles = document.getElementsByTagName("toggle-button");
            for(var i = 0; i < toggles.length; i++) {
                if(toggles[i].getAttribute("on-click") == "ctrl.toggleVideo()") {
                    var toggleCam = toggles[i];
                    if(toggleCam.length > 0) {
                        // Check if camera is on, then turn off if it is.
                        if(toggleCam.getAttribute("telemetry-outcome") == "26") {
                            toggleCam.children[0].children[0].click();
                            console.warn("Turned camera off.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                        }
                    }
                }
            }
            await sleep(50);
            // And finally, join the meeting.
            var joinBtn = document.getElementsByClassName("join-btn ts-btn inset-border ts-btn-primary");
            if(joinBtn.length > 0) {
                joinBtn[0].click();
                dojoining = false;
                console.log("Joined.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            } else {
                console.error("Failed to join meeting, can't find the join button.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            }
        }
        // Wait
        await sleep(joinDelay);
    }
}
async function cancelJoining() {
    // Stop the script
    dojoining = false;
}
joiningAction();
