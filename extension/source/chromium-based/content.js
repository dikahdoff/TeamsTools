// Runs when a https://teams.microsoft.com page is loaded.
var parsed = null;
parsed = JSON.parse(httpGet("https://raw.githubusercontent.com/dikahdoff/TeamsUtils/main/scripts.json"));
//parsed = JSON.parse(httpGet("https://cdn.jsdelivr.net/gh/dikahdoff/TeamsUtils/scripts.json"));
var manifestData = chrome.runtime.getManifest();
//
var dokicking = false;
var dojoining = false;
var dodisconnect = false;
//
var didswitch = false;
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
function httpGet(theUrl)
{
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open( "GET", theUrl, false );
    xmlHttp.send( null );
    return xmlHttp.responseText;
}
function isNumeric(str) {
    if (typeof str != "string") return false
    return !isNaN(str) &&
           !isNaN(parseFloat(str))
}
function strToBool(string){
    switch(string.toLowerCase().trim()){
        case "true": case "yes": case "1": return true;
        case "false": case "no": case "0": case null: return false;
        default: return null;
    }
}
function isRunning(key) {
    switch (key) {
        case "AutoKick":
            return dokicking;
            break;
        case "AutoJoinMeeting":
            return dojoining;
            break;
        case "AutoDisconnect":
            return dodisconnect;
            break;
        default:
            return null;
            break;
    }
}
function openGenerator() {
    window.open(parsed.link);
}
function openPage() {
    window.open(parsed.link);
}
function workInProgress() {
    alert('Work in Progress');
}
function stopScript(key) {
    switch (key) {
        case "AutoKick":
            cancelKicking();
            break;
        case "AutoJoinMeeting":
            cancelJoining();
            break;
        case "AutoDisconnect":
            cancelDisconnect();
            break;
        default:
            alert("ERROR! Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            break;
    }
    closeMenu();
}
function getSettings(key) {
    var set = new Array;
    parsed.scripts.forEach(element => {
        if(element.key == key) {
            element.settings.forEach(settingsElement => {
                settingsElement.settings.forEach(settingsDetails => {
                    var typeStr = "";
                    switch (settingsDetails.type) {
                        case "int":
                            typeStr = "(Possible values: Numbers)";
                            break;
                        case "bool":
                            typeStr = "(Possible values: True/False)";
                            break;
                        case "str":
                            typeStr = "(Possible values: Text)";
                            break;
                        default:
                            break;
                    }
                    set[settingsDetails.key] = getInput(((settingsDetails.name == null) ? settingsElement.name : settingsDetails.name), settingsElement.description + " " + typeStr,settingsDetails.type, ((settingsDetails.default == null) ? "0" : settingsDetails.default));
                });
            });
        }
    });
    console.log(set);
    switch (key) {
        case "AutoKick":
            kickingAction(set['kickDelay'], set['targetName']);
            break;
        case "AutoJoinMeeting":
            joiningAction(set['joinDelay'], set['joinWait'], set['switchChannel'], set['switchTo'], set['switchToChannel']);
            break;
        case "AutoDisconnect":
            disconnectAction(set['threshold'], set['delay']);
            break;
        default:
            alert("ERROR! Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            break;
    }
    closeMenu();
}
function getInput(name, desc, type, def) {
    var taskNotDone = true;
    while(taskNotDone) {
        var input = prompt(name + "\n" + desc, def);
        if (input == null || input == "") {
            return def;
        } else {
            switch (type) {
                case "int":
                    if(isNumeric(input)) {
                        taskNotDone = false;
                        return parseInt(input);
                    } else {
                        alert('Entered value cannot be parsed as an integer.');
                    }
                    break;
                case "bool":
                    var parsedBool = strToBool(input);
                    if(parsedBool != null) {
                        taskNotDone = false;
                        return parsedBool;
                    } else {
                        alert('Entered value cannot be parsed as a boolean.');
                    }
                    break;
                case "str":
                    taskNotDone = false;
                    return input;
                    break;
                default:
                    break;
            }
        }
    }
}
function openMenu() {
    if(document.getElementById("teamsutils-settings") == null) {
        var menu = document.createElement("div");
        menu.classList.add("popover", "settings-dropdown", "app-default-menu", "am-fade", "top");
        menu.setAttribute("data-tid", "settingsDropdownMenu");
        menu.setAttribute("oc-lazy-load", "settings-dropdown");
        menu.setAttribute("lazy-load-scenario", "lazy_settings_dropdown");
        menu.setAttribute("lazy-load-on-success", "appHeaderBar.onSettingsDropDownLoad()");
        menu.setAttribute("acc-role-dom", "dialog");
        menu.setAttribute("role", "dialog");
        menu.setAttribute("aria-modal", "true");
        var settingsDropdown = document.createElement("settings-dropdown");
        var divSettingsDropdown = document.createElement("div");
        var divDivSettingsDropdown = document.createElement("div");
        divDivSettingsDropdown.setAttribute("ng-show", "sdc.state === sdc.DropdownStates.INITIAL");
        var title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "<b>" + parsed.title + "</b>";
        var mainList = document.createElement("ul");
        mainList.classList.add("dropdown-menu");
        mainList.id = "settings-dropdown-list";
        mainList.setAttribute("acc-role-dom", "menu dialog");
        mainList.setAttribute("kb-list", "");
        mainList.setAttribute("kb-cyclic", "");
        mainList.setAttribute("role", "menu");
        mainList.appendChild(title);
        // Create title
        title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "Version: " + ((manifestData.version == parsed.latestVersion) ? manifestData.version : manifestData.version + ", new version available: v" + parsed.latestVersion);
        mainList.appendChild(title);
        // End of Create Title
        // Create Button
        var btnElement = document.createElement("li");
        btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
        btnElement.setAttribute("data-prevent-trigger-refocus", "true");
        btnElement.setAttribute("acc-role-dom", "menu-item");
        btnElement.setAttribute("role", "menuitem");
        btnElement.setAttribute("kb-item", "");
        var btnBtnElement = document.createElement("button");
        btnBtnElement.classList.add("ts-sym");
        btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
        btnBtnElement.innerText = parsed.link;
        btnBtnElement.addEventListener('click', openPage);
        btnElement.appendChild(btnBtnElement);
        mainList.appendChild(btnElement);
        // End of Create Button
        var splitterElement = document.createElement("li");
        splitterElement.setAttribute("ng-if","::sdc.displaySettingsComponents");
        splitterElement.classList.add("divider");
        splitterElement.setAttribute("acc-role-dom", "menu-separator");
        mainList.appendChild(splitterElement);
        /*
        !!! NOT IMPLEMENTED YET, DO NOT UNCOMMENT !!!
        // Create title
        title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "Automatic execution";
        mainList.appendChild(title);
        // End of Create Title
        // Create Button
        var btnElement = document.createElement("li");
        btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
        btnElement.setAttribute("data-prevent-trigger-refocus", "true");
        btnElement.setAttribute("acc-role-dom", "menu-item");
        btnElement.setAttribute("role", "menuitem");
        btnElement.setAttribute("kb-item", "");
        var btnBtnElement = document.createElement("button");
        btnBtnElement.classList.add("ts-sym");
        btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
        btnBtnElement.innerText = "Create automation script";
        btnBtnElement.addEventListener('click', openGenerator);
        btnElement.appendChild(btnBtnElement);
        mainList.appendChild(btnElement);
        // End of Create Button
        // Create Button
        var btnElement = document.createElement("li");
        btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
        btnElement.setAttribute("data-prevent-trigger-refocus", "true");
        btnElement.setAttribute("acc-role-dom", "menu-item");
        btnElement.setAttribute("role", "menuitem");
        btnElement.setAttribute("kb-item", "");
        var btnBtnElement = document.createElement("button");
        btnBtnElement.classList.add("ts-sym");
        btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
        btnBtnElement.innerText = "Import automation script";
        btnBtnElement.addEventListener('click', workInProgress);
        btnElement.appendChild(btnBtnElement);
        mainList.appendChild(btnElement);
        // End of Create Button
        // Create Button
        var btnElement = document.createElement("li");
        btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
        btnElement.setAttribute("data-prevent-trigger-refocus", "true");
        btnElement.setAttribute("acc-role-dom", "menu-item");
        btnElement.setAttribute("role", "menuitem");
        btnElement.setAttribute("kb-item", "");
        var btnBtnElement = document.createElement("button");
        btnBtnElement.classList.add("ts-sym");
        btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
        btnBtnElement.innerText = "Start script";
        btnBtnElement.addEventListener('click', workInProgress);
        btnElement.appendChild(btnBtnElement);
        mainList.appendChild(btnElement);
        // End of Create Button
        splitterElement = document.createElement("li");
        splitterElement.setAttribute("ng-if","::sdc.displaySettingsComponents");
        splitterElement.classList.add("divider");
        splitterElement.setAttribute("acc-role-dom", "menu-separator");
        mainList.appendChild(splitterElement);
        */
        // Create title
        title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "Manual execution";
        mainList.appendChild(title);
        // End of Create Title
        var scripts = parsed.scripts;
        scripts.forEach(element => {
            // Create Button
            var btnElement = document.createElement("li");
            btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
            btnElement.setAttribute("data-prevent-trigger-refocus", "true");
            btnElement.setAttribute("acc-role-dom", "menu-item");
            btnElement.setAttribute("role", "menuitem");
            btnElement.setAttribute("kb-item", "");
            var btnBtnElement = document.createElement("button");
            btnBtnElement.classList.add("ts-sym");
            btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
            btnBtnElement.id = "triggerbtn-" + element.key;
            var isRun = (isRunning(element.key));
            btnBtnElement.addEventListener('click', function (){
                if(isRun) {
                    stopScript(element.key);
                } else {
                    getSettings(element.key);
                }
            });
            btnBtnElement.innerText = (isRun ? '✔️' : '❌') + " " + element.name;
            btnElement.appendChild(btnBtnElement);
            mainList.appendChild(btnElement);
            // End of Create Button
        });
        divSettingsDropdown.appendChild(divDivSettingsDropdown);
        settingsDropdown.appendChild(divSettingsDropdown);
        menu.id = "teamsutils-settings";
        menu.appendChild(settingsDropdown);
        menu.style="display: block; visibility: visible;";
        var closeBtn = document.createElement("button");
        closeBtn.classList.add("ts-btn", "ts-btn-primary", "admin-upgrade-button");
        closeBtn.type = "button";
        closeBtn.setAttribute("data-tid", "upgrade-button");
        closeBtn.innerText = "Close Menu";
        closeBtn.addEventListener('click', closeMenu);
        divDivSettingsDropdown.appendChild(mainList);
        splitterElement = document.createElement("li");
        splitterElement.setAttribute("ng-if","::sdc.displaySettingsComponents");
        splitterElement.classList.add("divider");
        splitterElement.setAttribute("acc-role-dom", "menu-separator");
        mainList.appendChild(splitterElement);
        menu.appendChild(closeBtn);
        document.body.appendChild(menu);
    } else {
        closeMenu();
    }
}
function closeMenu() {
    document.getElementById("teamsutils-settings").remove();
}
console.log("[TeamsUtils] Starting...");
window.addEventListener ("load", mainFunc, false);
async function mainFunc() {
    console.log("[TeamsUtils] Started.");
    console.log("[TeamsUtils] Initializing...");
    var foundBar = false;
    while(!foundBar) {
        var bar = document.getElementsByClassName("powerbar-profile fadeable");
        if(bar.length > 0) {
            foundBar = true;
            console.log("[TeamsUtils] Injecting...");
            var btn = document.createElement("a");
            btn.innerText = "TeamsUtils Menu";
            btn.href = "#";
            btn.addEventListener('click', openMenu);
            bar[0].appendChild(btn);
            console.log("[TeamsUtils] Injected.");
        }
        await sleep(100);
    }
}
/////////////////////
/* !!! SCRIPTS !!! */
/////////////////////
/* !!! AutoKick.js !!! */
async function kickingAction(kickDelay, targetName) {
    dokicking = true;
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
                                                if(cname.toString().toLowerCase().toString().includes(targetName.toString().toLowerCase().toString())) {
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
/* !!! AutoJoinMeeting.js !!! */
async function joiningAction(joinDelay, joinWait, switchChannel, switchTo, switchToChannel) {
    dojoining = true;
    while(dojoining) {
        if(switchChannel && !didswitch) {
            console.warn("Switching to " + switchTo + " before fetching...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            var teams = document.getElementsByClassName("app-left-rail-width");
            if(teams.length > 0) {
                teams = teams[0].children[0].children[0].children[1].children;
                if(teams.length > 0) {
                    for(var i = 0; i < teams.length; i++) {
                        if(teams[i].classList.contains("team")) {
                            var cTeamName = teams[i].children[0].children[0].children[0].children[1].children[0].children[0].innerHTML;
                            if(cTeamName.toLowerCase().contains(switchTo.toLowerCase())) {
                                console.log("Found " + switchTo + " team, opening if necessary...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                                // Check if teams is collapsed, if it is, open it.
                                if(teams[i].getAttribute("aria-expanded") == "false") {
                                    console.log(teams[i]);
                                    teams[i].children[0].children[0].click();
                                }
                                await sleep(50);
                                // Click on the channel.
                                var channels = teams[i].children[0].children[1].children[0].children[0].children;
                                for(var j = 0; j < channels.length; j++) {
                                    var channelName = channels[j].children[0].children[0].children[0].innerHTML;
                                    if(channelName.toLowerCase().contains(switchToChannel.toLowerCase())) {
                                        console.log("Found " + switchToChannel + " channel, opening if necessary...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                                        channels[j].children[0].click();
                                        didswitch = true;
                                    }
                                }
                            }
                        }
                    }
                } else {
                    console.error("You don't have any teams to join to.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
                }
            } else {
                console.error("Couldn't find teams list.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            }
            await sleep(joinDelay);
        }
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
            await sleep(100);
            // Check if the continue without camera or audio dialog opens
            var withoutBtn = document.getElementsByClassName("ts-btn ts-btn-fluent ts-btn-fluent-secondary-alternate");
            if(withoutBtn.length > 0) {
                // If it did, close it.
                withoutBtn[0].click();
                console.warn("Closed warning box.\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
            }
            await sleep(100);
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
/* !!! AutoDisconnect.js !!! */
async function disconnectAction(disconnectThreshold, disconnectDelay) {
    dodisconnect = true;
    while(dodisconnect) {
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
        if(sum <= disconnectThreshold && sum != 0) {
            // Disconnect because the sum is lower then the set disconnectThreshold
            document.getElementById("hangup-button").click();
            dodisconnect = false;
            console.warn("disconnectThreshold is " + disconnectThreshold + ", disconnecting at " + sum + " attendees...\nDownload this script: https://github.com/dikahdoff/TeamsUtils");
        }
        // Wait
        await sleep(disconnectDelay);
    }
}
async function cancelDisconnect() {
    // Stop the script
    dodisconnect = false;
}