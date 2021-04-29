function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
var dotask = true;
async function action() {
    dotask = true;
    console.clear()
    while(dotask) {
        var sum = 0;
        var button = document.getElementById("roster-button");
        if(button.getAttribute("track-outcome") == "15") {
            button.click();
        }
        var elements = document.getElementsByTagName("calling-roster-section");
        for (var i = 0; i < elements.length; i++) {
            if(elements[i].getAttribute("participants-list") == "ctrl.participantsInCall" || elements[i].getAttribute("participants-list") == "ctrl.attendeesInMeeting") {
                var childrens = elements[i].children[0].children[0].getElementsByTagName("button");
                for (var j = 0; j < childrens.length; j++) {
                    if(childrens[j].classList.contains("roster-list-title")) {
                        sum += parseInt(childrens[j].children[2].innerHTML.replace("(","").replace(")",""));
                    }
                }
            }
        }
        console.log("Users in meeting: " + sum + "\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        if(sum <= threshold && sum != 0) {
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
