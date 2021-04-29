///////////////////
//    README     //
///////////////////

/*
After you've joined a Meeting press Ctrl+Shift+I and paste the whole code (Ctrl+A) into the console section of the window and press enter.
Set the kicking target's name below using the 'targetName' variable. The script will search for this attendee and kick the person if able to.

If you want to manually stop the script, write 'cancelKicking()' then press enter in the console window.
If you want to restart the script, write 'kickingAction()' then press enter in the console window.
*/

///////////////////
//   SETTINGS   ///
///////////////////

// Delay that the script acts after. (in milliseconds)
var kickDelay = 150;

// User to target with the kick bot
var targetName = "John Smith";

///////////////////

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Loop
var dokicking = true;
async function kickingAction() {
    dokicking = true;
    console.clear()
    while(dokicking) {
        // Open sidepanel if necessary
        try {
            var button = document.getElementById("roster-button");
            if(button.getAttribute("track-outcome") == "15") {
                button.click();
            }
        } catch (error) { /* Don't care */}
        if(kickDelay < 50) {
            dokicking = false;
            console.error("The delay cannot be lower then 50ms, because Teams will freak out.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        } else {
            var elements = document.getElementsByTagName("calling-roster-section");
            for (var i = 0; i < elements.length; i++) {
                if(elements[i].getAttribute("participants-list") != "ctrl.participantsFromMeeting" && elements[i].getAttribute("participants-list") != "ctrl.participantsFromTeamSuggestions") {
                    var list = elements[i].children[0].children[1].children[0].children[0].children;
                    for (var j = 0; j < list.length; j++) {
                        if(list[j].getAttribute("role") == "listitem") {
                            var itemChildren = list[j].children;
                            for (var k = 0; k < itemChildren.length; k++) {
                                if(itemChildren[k].classList.contains("participantInfo")) {
                                    var partInfo = itemChildren[k].children;
                                    var thisPartOk = true;
                                    for (var l = 0; l < partInfo.length; l++) {
                                        if(partInfo[l].classList.contains("organizerStatus")) {
                                            // This participant is the organizer, cannot be kicked and shouldn't be checked.
                                            thisPartOk = false;
                                        }
                                    }
                                    for (var x= 0; x < itemChildren.length; x++) {
                                        if(itemChildren[x].classList.contains("stateActions")) {
                                            var statusChildren = itemChildren[x].children;
                                            for (var y= 0; y < statusChildren.length; y++) {
                                                if(statusChildren[y].classList.contains("state")) {
                                                    if(statusChildren[y].innerHTML != "") {
                                                        // This participant is currently leaving and shouldn't be checked.
                                                        thisPartOk = false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if(thisPartOk) {
                                        for (var l = 0; l < partInfo.length; l++) {
                                            if(partInfo[l].classList.contains("name")) {
                                                var cname = partInfo[l].children[0].innerHTML;
                                                if(cname.toLowerCase().contains(targetName.toLowerCase())) {
                                                    console.log("Found " + targetName + " as " + cname + ", kicking...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                                                    for (var m = 0; m < itemChildren.length; m++) {
                                                        if(itemChildren[m].classList.contains("stateActions") && itemChildren[m].classList.contains("hide-participant-state")) {
                                                            var btns = itemChildren[m].children;
                                                            for(var n = 0; n < btns.length; n++) {
                                                                if(btns[n].classList.contains("participant-menu") && btns[n].getAttribute("ng-if") == "ctrl.actionsMenuEnabled") {
                                                                    btns[n].children[0].children[0].click();
                                                                    await sleep(5);
                                                                    var dropMenu = document.getElementsByClassName("context-dropdown");
                                                                    if(dropMenu.length > 0) {
                                                                        var dropElements = dropMenu[0].children[0].children;
                                                                        for(var o = 0; o < dropElements.length; o++) {
                                                                            if(dropElements[o].children[0].children[1].getAttribute("translate-once") == "calling_roster_remove_paticipant_popup_item") {
                                                                                dropElements[o].children[0].click();
                                                                                console.log(cname + " has been kicked...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                                                                            }
                                                                        }
                                                                    } else {
                                                                        console.log("Please wait...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // Wait
            await sleep(kickDelay);
        }
    }
}
async function cancelKicking() {
    // Stop the script
    dokicking = false;
}
kickingAction();
