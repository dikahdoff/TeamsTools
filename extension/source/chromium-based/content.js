var parsed = null;

// Replace scripts.json with offline version if server is unavailable.
onlinever = httpGet("https://raw.githack.com/dikahdoff/TeamsTools/master/scripts.json")
if(onlinever == false) {
    console.warn("TeamsTools server offline, trying to load cached version");
    chrome.storage.sync.get('tscripts', function(result) {
        scripts = result.tscripts;
        if ((typeof scripts === "undefined") || scripts == null ) {
            panic("TeamsTools server offline and no cached version of config. TeamsTools can't load. Sorry for the inconvenience");
        } else {
            console.log("TeamsTools server offline, loading cached version...");
            parsed = scripts;
            parsed.title += " [OFFLINE]";
            preInit();
        }
    });
} else {
    console.log("TeamsTools server online, loading that...");
    parsed = JSON.parse(onlinever);
    value = parsed;
    chrome.storage.sync.set({tscripts: value}, function() {
        log("Scripts saved to cache.", false);
    });
    preInit();
}

var manifestData = chrome.runtime.getManifest();

var settings = {}
var defaultSettings = {
    firstSetup: false,
    doSeeMore: true,
    doDarkMode: true,
    doRemoveAnnoy: true,
};
var dokicking, dojoining, dodisconnect = false;

/* Initialization */
function preInit() {
    log("Starting...", false);
    // Inject dark mode CSS beforehand because it's better to not get blinded, then switch back to flashbang mode (if the user wants to), then flashbanging those users, who don't want to be flashbanged
    log("Patching Dark Mode...", false);
    addStyle(`body {
        background-color: black;
        color: white;
    }
    .app-loading {
        background-color: black;
        color: white;
    }
    .guest-license-error, .guest-license-error-footer {
        background: black;
    }
    .guest-license-error-dropdown-options {
        color: black;
    }`);
    log("Patched Dark Mode.", false);
    // Try loading in the settings, if failed, load default settings
    try {
        log("Loading settings...", false);
        chrome.storage.sync.get('tsettings', function(result) {
            settings = result.tsettings;
            if (typeof settings === "undefined") {
                value = JSON.stringify(defaultSettings);
                try {
                    chrome.storage.sync.set({tsettings: value}, function() {
                        settings = defaultSettings;
                        log("Settings loaded from defaultSettings and got saved to syncsave", false);
                        init();
                    });
                } catch (error) {
                    console.warn(error);
                    log("Settings loaded from temp defaultSettings due to an error", false);
                    settings = defaultSettings;
                    init();
                }
            } else {
                settings = JSON.parse(settings);
                log("Settings loaded from syncsave", false);
                init();
            }
        });
    } catch (error) {
        console.warn(error);
        value = JSON.stringify(defaultSettings);
        try {
            chrome.storage.sync.set({tsettings: value}, function() {
                settings = defaultSettings;
                log("Settings loaded from defaultSettings and got saved to syncsave", false);
                init();
            });
        } catch (error) {
            console.warn(error);
            log("Settings loaded from temp defaultSettings due to an error", false);
            settings = defaultSettings;
            init();
        }
    }
}

// Inject more CSS and unpatch dark mode if required
function init() {
    log("Started.",false);
    if(settings.doSeeMore) {
        seeMoreAction();
    }
    if(settings.doDarkMode) {
        doDarkMode();
    }
    if(settings.doRemoveAnnoy) {
        doRemoveAnnoy();
    }
    log("Initializing...",false);
    addStyle(`#tutils-donation, #tutils-donation-long {
        font-size: 16px;
        font-weight: bold;
        background: linear-gradient(to right, #6666ff, #0099ff , #00ff00, #ff3399, #6666ff);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
        animation: rainbow_animation 6s ease-in-out infinite;
        background-size: 400% 100%;
    }
    #tutils-donation-long {
        font-size: 20px;
    }
    #tutils-donation:hover, #tutils-donation-long:hover {
        text-decoration: underline;
    }
    @keyframes rainbow_animation {
        0%,100% {
            background-position: 0 0;
        }
        50% {
            background-position: 100% 0;
        }
    }`);
    if(!settings.doDarkMode) {
        log("Unpatching Dark Mode...",false);
        addStyle(`.app-loading {
            background-color: white;
            color: black;
        }
        .guest-license-error, .guest-license-error-footer {
            background: #f0f0f0;
        }
        .guest-license-error-dropdown-options {
            color: white;
        }`);
    }
    //window.addEventListener ("load", mainFunc, false);
    mainFunc();
}
// Inject TeamsTools button to the header when it's found
async function mainFunc() {
    var foundBar = false;
    while(!foundBar) {
        var bar = document.getElementsByClassName("powerbar-profile fadeable");
        if(bar.length > 0) {
            foundBar = true;
            log("Injecting...",false);
            var btn = document.createElement("button");
            btn.type = "button";
            btn.classList.add("ts-sym", "me-profile");
            btn.href = "#";
            // Append image
            var userInfoBtn = document.createElement("div");
            userInfoBtn.classList.add("user-information-button");
            userInfoBtn.setAttribute("data-tid", "userInformation");
            var profileImgParent = document.createElement("div");
            profileImgParent.classList.add("profile-img-parent");
            var profilePictureItem = document.createElement("profile-picture");
            profilePictureItem.setAttribute("css-class", "user-picture");
            var imgProfPic = document.createElement("img");
            imgProfPic.src = parsed.teamsBtn;
            imgProfPic.style = "object-fit: cover;";
            profilePictureItem.appendChild(imgProfPic);
            profileImgParent.appendChild(profilePictureItem);
            userInfoBtn.appendChild(profileImgParent);
            btn.appendChild(userInfoBtn);
            btn.addEventListener('click', openMenu);
            var dlink = document.createElement("a");
            dlink.innerHTML = "Donate                               ";
            dlink.target = "_blank";
            dlink.href = parsed.donationLink;
            dlink.id = "tutils-donation-long";
            var waffleheader = document.getElementsByClassName("waffle-header");
            if(waffleheader.length > 0) {
                waffleheader[0].appendChild(dlink);
            } else {
                waffleheader = document.getElementsByClassName("powerbar-profile fadeable");
                if(waffleheader.length > 0) {
                    dlink.innerHTML = "Donate";
                    waffleheader[0].insertBefore(dlink, waffleheader[0].firstChild);
                }
            }
            bar[0].appendChild(btn);
            // Remove annoying Desktop app download button
            if(settings.doRemoveAnnoy) {
                document.getElementById("get-app-button").parentElement.parentElement.remove();
            }
            log("Injected.",false);
            //if(!settings.firstSetup) {
                //log("Welcome, newbie!")
            //} else {
                log("Welcome!");
            //}
        }
        await sleep(100);
    }
}

/* Functions */
// Send async threads to sleep for given time period
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Send HTTP GET request and get the content
function httpGet(url)
{
    try {
        var xmlHttp = new XMLHttpRequest();
        xmlHttp.open("GET", url, false);
        xmlHttp.send(null);
        return xmlHttp.responseText;
    } catch (error) {
        console.warn(error);
        return false;
    }
}

// Inject CSS into the webpage
function addStyle(styleString) {
    const style = document.createElement('style');
    style.textContent = styleString;
    document.head.append(style);
}

// Manage settings I/O
function manageSettings(write = false, settingsJson = settings) {
    if(write) {
        value = JSON.stringify(settingsJson);
        try {
            chrome.storage.sync.set({tsettings: value}, function() {
                return true;
            });
        } catch (error) {
            console.warn(error);
            return false;
        }
    } else {
        try {
            chrome.storage.sync.get(['tsettings'], function(result) {
                settings = JSON.parse(result.tsettings);
                return true;
            });
        } catch (error) {
            console.warn(error);
            return manageSettings(true, defaultSettings);
        }
    }
}

// Determine if a String can be parsed as a numeric variable or not
function isNumeric(str) {
    if (typeof str != "string") return false
    return !isNaN(str) &&
           !isNaN(parseFloat(str))
}

// Convert a String to a boolean variable if possible
function strToBool(string){
    switch(string.toLowerCase().trim()){
        case "true": case "yes": case "1": return true;
        case "false": case "no": case "0": case null: return false;
        default: return null;
    }
}

// Get input from the user through input boxes
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
                        var out = parseInt(input);
                        if(0 < out < 32766) {
                            taskNotDone = false;
                            return out;
                        } else {
                            alert('Entered value is out of acceptable range.');
                        }
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
                    if(input.length < 512) {
                        taskNotDone = false;
                        return input;
                    } else {
                        alert('Entered value is too long.');
                    }
                    break;
                case "array_str":
                    var parsedArrStr = input.split(";");
                    if(parsedArrStr != null) {
                        taskNotDone = false;
                        return parsedArrStr;
                    } else {
                        alert("Entered value can't be parsed as a String Array.");
                    }
                    break;
                default:
                    break;
            }
        }
    }
}

// Check if selected action is currently running
function isRunning(key) {
    switch (key) {
        case "AutoKick":
            return dokicking;
        case "AutoJoinMeeting":
            return dojoining;
        case "AutoDisconnect":
            return dodisconnect;
        case "AutoSeeMore":
            return settings.doSeeMore;
        case "DoDarkMode":
            return settings.doDarkMode;
        case "DoRemoveAnnoy":
            return settings.doRemoveAnnoy;
        default:
            return null;
    }
}

// Start selected script
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
    log("Running " + key + "...");
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
        case "AutoSeeMore":
            seeMoreAction(100);
            break;
        case "DoDarkMode":
            doDarkMode();
            break;
        case "DoRemoveAnnoy":
            doRemoveAnnoy(100);
            break;
        default:
            log("[ERROR] Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            alert("ERROR! Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            break;
    }
    closeMenu();
}

// Stop selected script
function stopScript(key) {
    log("Stopped " + key + ".");
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
        case "AutoSeeMore":
            cancelSeeMore();
            break;
        case "DoDarkMode":
            cancelDarkMode();
            break;
        case "DoRemoveAnnoy":
            cancelRemoveAnnoy();
            break;
        default:
            log("[ERROR] Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            alert("ERROR! Action isn't implemented in this version of the client. Please update the extension: " + parsed.link);
            break;
    }
    closeMenu();
}

// Links or pop-ups
function openGenerator() {
    window.open(parsed.generatorLink);
}
function openDonation() {
    window.open(parsed.donationLink);
}
function openPage() {
    window.open(parsed.link);
}
function openChangelog() {
    window.open(parsed.changelog);
}
function openStorePage() {
    window.open(parsed.storeLink);
}
function workInProgress() {
    alert('Work in Progress');
}

// Log to console and to GUI if possible
function log(logStr, gui = true) {
    var nowTime = new Date().toLocaleTimeString();
    console.log("(" + nowTime + ") " + parsed.prefix + "" + logStr + "\nDownload this script: " + parsed.link + "");
    if(gui) {
        var outputTxt = document.getElementById("control-input");
        if(outputTxt != null) {
            outputTxt.setAttribute("placeholder", "(" + nowTime + ") " + parsed.teamsTitle + "> " + logStr);
        } else {
            console.warn(parsed.prefix + "Failed to log to GUI, because selected element was not found.");
        }
    }
}

// Send panic message to user
function panic(panicStr, gui = true) {
    var nowTime = new Date().toLocaleTimeString();
    panicStr = "(" + nowTime + ") [TeamsTools Panic] " + panicStr;
    console.error(panicStr);
    if(gui) {
        alert(panicStr);
    }
}

// Title modifier observer task
var config = {childList: true};
var target = document.querySelector('title');
var observer = new MutationObserver(function(mutations) {
    mutations.forEach(function(mutation) {
        if(document.title != null && !document.title.includes(parsed.teamsTitle)) {
            document.title = document.title.replace("Microsoft Teams", parsed.teamsTitle);
            try {
                document.getElementsByClassName("teams-title")[0].innerText = parsed.teamsTitle + " - ";
            } catch (error) { /* Cope */ }
        }
    });
});

// Function ran when the user opens the TeamsTools menu
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
        title.innerHTML = "Version: " + ((manifestData.version == parsed.latestVersion) ? manifestData.version : manifestData.version + ", new version available: " + parsed.latestVersion);
        mainList.appendChild(title);
        // End of Create Title
        if(parsed.donationLink != null) {
            // Create Button
            var btnElement = document.createElement("li");
            btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
            btnElement.setAttribute("data-prevent-trigger-refocus", "true");
            btnElement.setAttribute("acc-role-dom", "menu-item");
            btnElement.setAttribute("role", "menuitem");
            btnElement.setAttribute("kb-item", "");
            var btnBtnElement = document.createElement("button");
            btnBtnElement.classList.add("ts-sym", "tutils-donation");
            btnBtnElement.id = "tutils-donation";
            btnBtnElement.setAttribute("ng-click", "sdc.gotoManage(); sdc.hide()");
            btnBtnElement.innerText = parsed.style.button + " Help fund the project!";
            btnBtnElement.title = "Opens " + parsed.donationLink;
            btnBtnElement.addEventListener('click', openDonation);
            btnElement.appendChild(btnBtnElement);
            mainList.appendChild(btnElement);
            // End of Create Button
        }
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
        btnBtnElement.innerText = parsed.style.button + " Visit project link";
        btnBtnElement.title = "Opens " + parsed.link;
        btnBtnElement.addEventListener('click', openPage);
        btnElement.appendChild(btnBtnElement);
        mainList.appendChild(btnElement);
        // End of Create Button
        if(parsed.storeLink != null) {
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
            btnBtnElement.innerText = parsed.style.button + " Visit store link";
            btnBtnElement.title = "Opens " + parsed.storeLink;
            btnBtnElement.addEventListener('click', openStorePage);
            btnElement.appendChild(btnBtnElement);
            mainList.appendChild(btnElement);
            // End of Create Button
        }
        if(parsed.changelog != null) {
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
            btnBtnElement.innerText = parsed.style.button + " Changelog";
            btnBtnElement.title = "Opens " + parsed.changelog;
            btnBtnElement.addEventListener('click', openChangelog);
            btnElement.appendChild(btnBtnElement);
            mainList.appendChild(btnElement);
            // End of Create Button
        }
        var splitterElement = document.createElement("li");
        splitterElement.setAttribute("ng-if","::sdc.displaySettingsComponents");
        splitterElement.classList.add("divider");
        splitterElement.setAttribute("acc-role-dom", "menu-separator");
        mainList.appendChild(splitterElement);

        limit = parsed.characterLimit;
        // Create title
        title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "Settings";
        mainList.appendChild(title);
        // End of Create Title
        var settings = parsed.settings;
        settings.forEach(element => {
            description = element.description;
            if(description.length > limit) {
                description = description.slice(0,limit) + "...";
            }
            // Create Button
            var btnElement = document.createElement("li");
            btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
            btnElement.setAttribute("data-prevent-trigger-refocus", "true");
            btnElement.setAttribute("acc-role-dom", "menu-item");
            btnElement.setAttribute("role", "menuitem");
            btnElement.setAttribute("kb-item", "");
            btnElement.title = element.description;
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
            btnBtnElement.innerHTML = (isRun ? parsed.style.running : parsed.style.stopped) + " " + element.name;
            btnElement.appendChild(btnBtnElement);
            var btnDescElement = document.createElement("div");
            btnDescElement.classList.add("settings-dropdown-header");
            btnDescElement.setAttribute("data-tid", "product-name-and-type");
            btnDescElement.setAttribute("ng-bind-html", "::sdc.teamsProductName");
            btnDescElement.setAttribute("acc-role-dom", "menu-item");
            btnDescElement.setAttribute("kb-item-role", "menuitem");
            btnDescElement.setAttribute("style", "font-size: 10px;");
            btnDescElement.innerHTML = description;
            btnElement.appendChild(btnDescElement);
            mainList.appendChild(btnElement);
            // End of Create Button
        });
        var splitterElement = document.createElement("li");
        splitterElement.setAttribute("ng-if","::sdc.displaySettingsComponents");
        splitterElement.classList.add("divider");
        splitterElement.setAttribute("acc-role-dom", "menu-separator");
        mainList.appendChild(splitterElement);
        // Create title
        title = document.createElement("div");
        title.classList.add("settings-dropdown-header");
        title.setAttribute("data-tid", "product-name-and-type");
        title.setAttribute("ng-bind-html", "::sdc.teamsProductName");
        title.setAttribute("acc-role-dom", "menu-item");
        title.setAttribute("kb-item-role", "menuitem");
        title.innerHTML = "Toggles";
        mainList.appendChild(title);
        // End of Create Title
        var scripts = parsed.scripts;
        scripts.forEach(element => {
            description = element.description;
            if(description.length > limit) {
                description = description.slice(0,limit) + "...";
            }
            // Create Button
            var btnElement = document.createElement("li");
            btnElement.setAttribute("ng-if", "::(sdc.displaySettingsComponents && sdc.isFreemiumAdmin)");
            btnElement.setAttribute("data-prevent-trigger-refocus", "true");
            btnElement.setAttribute("acc-role-dom", "menu-item");
            btnElement.setAttribute("role", "menuitem");
            btnElement.setAttribute("kb-item", "");
            btnElement.title = element.description;
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
            btnBtnElement.innerHTML = (isRun ? parsed.style.running : parsed.style.stopped) + " " + element.name;
            btnElement.appendChild(btnBtnElement);
            var btnDescElement = document.createElement("div");
            btnDescElement.classList.add("settings-dropdown-header");
            btnDescElement.setAttribute("data-tid", "product-name-and-type");
            btnDescElement.setAttribute("ng-bind-html", "::sdc.teamsProductName");
            btnDescElement.setAttribute("acc-role-dom", "menu-item");
            btnDescElement.setAttribute("kb-item-role", "menuitem");
            btnDescElement.setAttribute("style", "font-size: 10px;");
            btnDescElement.innerHTML = description;
            btnElement.appendChild(btnDescElement);
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
        closeBtn.title = "Close this menu manually";
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

// Function runs when user closes the menu or clicking an action does so
function closeMenu() {
    document.getElementById("teamsutils-settings").remove();
}

/* Scripts/actions */
// AutoKick
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
            log("The delay cannot be lower then 50ms, because Teams will freak out. Please stop the script and set a higher delay.");
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
                                                    log("Found " + targetName + " as " + cname + ", kicking...");
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
                                                                                log(cname + " has been kicked...");
                                                                            }
                                                                        }
                                                                    } else {
                                                                        log("Please wait...");
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

// AutoJoinMeeting
var didswitch = false;
async function joiningAction(joinDelay, joinWait, switchChannel, switchTo, switchToChannel) {
    dojoining = true;
    while(dojoining) {
        if(switchChannel && !didswitch) {
            log("Switching to " + switchTo + " before fetching...");
            var teams = document.getElementsByClassName("app-left-rail-width");
            if(teams.length > 0) {
                teams = teams[0].children[0].children[0].children[1].children;
                if(teams.length > 0) {
                    for(var i = 0; i < teams.length; i++) {
                        if(teams[i].classList.contains("team")) {
                            var cTeamName = teams[i].children[0].children[0].children[0].children[1].children[0].children[0].innerHTML;
                            if(cTeamName.toLowerCase().contains(switchTo.toLowerCase())) {
                                log("Found " + switchTo + " team, opening if necessary...");
                                // Check if teams is collapsed, if it is, open it.
                                if(teams[i].getAttribute("aria-expanded") == "false") {
                                    teams[i].children[0].children[0].click();
                                }
                                await sleep(50);
                                // Click on the channel.
                                var channels = teams[i].children[0].children[1].children[0].children[0].children;
                                for(var j = 0; j < channels.length; j++) {
                                    var channelName = channels[j].children[0].children[0].children[0].innerHTML;
                                    if(channelName.toLowerCase().contains(switchToChannel.toLowerCase())) {
                                        log("Found " + switchToChannel + " channel, opening if necessary...");
                                        channels[j].children[0].click();
                                        didswitch = true;
                                    }
                                }
                            }
                        }
                    }
                } else {
                    log("[ERROR] You don't have any teams to join to.");
                }
            } else {
                log("[ERROR] Couldn't find teams list.");
            }
            await sleep(joinDelay);
        }
        log("Fetching meetings...");
        // Get meeting button
        var element = document.getElementsByClassName("ts-sym ts-btn ts-btn-primary inset-border icons-call-jump-in ts-calling-join-button app-title-bar-button app-icons-fill-hover call-jump-in");
        if(element.length > 0) {
            element = element[0];
            // Validated, a meeting is running and the join button is available.
            // Wait for 'joinWait' and then click the join button
            log("Found a meeting, joining in " + joinWait + "ms");
            await sleep(joinWait);
            log("Joining...");
            element.click();
            await sleep(100);
            // Check if the continue without camera or audio dialog opens
            var withoutBtn = document.getElementsByClassName("ts-btn ts-btn-fluent ts-btn-fluent-secondary-alternate");
            if(withoutBtn.length > 0) {
                // If it did, close it.
                withoutBtn[0].click();
                log("Closed warning box.");
            }
            await sleep(100);
            var toggleMic = document.getElementById("preJoinAudioButton");
            // Check if toggle mic exists
            if(toggleMic != null) {
                // Check if mic is on, then turn off if it is.
                if(toggleMic.getAttribute("telemetry-outcome") == "30") {
                    toggleMic.children[0].children[0].click();
                    log("Turned microphone off.");
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
                            log("Turned camera off.");
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
                log("Joined.");
            } else {
                log("[ERROR] Failed to join meeting, can't find the join button.");
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

// AutoDisconnect
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
        log("Users in meeting: " + sum);
        if(sum <= disconnectThreshold && sum != 0) {
            // Disconnect because the sum is lower then the set disconnectThreshold
            document.getElementById("hangup-button").click();
            dodisconnect = false;
            log("Disconnect Threshold is " + disconnectThreshold + ", disconnecting at " + sum + " attendees...");
        }
        // Wait
        await sleep(disconnectDelay);
    }
}
async function cancelDisconnect() {
    // Stop the script
    dodisconnect = false;
}

// AutoSeeMore
async function seeMoreAction(seeMoreCheckDelay) {
    settings.doSeeMore = true;
    manageSettings(true);
    while(settings.doSeeMore) {
        var sum = 0;
        // Count See More buttons
        var buttons = document.getElementsByClassName("ts-sym ts-see-more-button ts-see-more-fold");
        if(buttons.length > 0) {
            for (var i = 0; i < buttons.length; i++) {
                if(buttons[i].getAttribute("ng-click") == "toggleSeeMore()") {
                    buttons[i].click();
                    sum++;
                }
            }

            if(sum > 0) {
                var seeLessBtns = document.getElementsByClassName("ts-sym ts-see-more-button");
                for (var i = 0; i < seeLessBtns.length; i++) {
                    if(seeLessBtns[i].getAttribute("track-default") == "toggleSeeLess" || seeLessBtns[i].getAttribute("track-default") == "toggleSeeMore") {
                        seeLessBtns[i].remove();
                    }
                }
            }
        }
        // Wait
        await sleep(seeMoreCheckDelay);
    }
}
async function cancelSeeMore() {
    // Stop the script
    settings.doSeeMore = false;
    manageSettings(true);
}

// DoDarkMode
async function doDarkMode() {
    settings.doDarkMode = true;
    manageSettings(true);
}
async function cancelDarkMode() {
    // Stop the script
    settings.doDarkMode = false;
    manageSettings(true);
}

// DoRemoveAnnoy
async function doRemoveAnnoy(removeAnnoyDelay) {
    settings.doRemoveAnnoy = true;
    manageSettings(true);
    var annoyances = ["download-app-button", "download-mobile-app-button"];
    while(settings.doRemoveAnnoy) {
        try {
            annoyances.forEach(annoyance => {
                var elem = document.getElementById(annoyance);
                if(!(typeof elem === undefined) && elem != null) {
                    elem.parentElement.remove();
                }
            });
        } catch (error) {
            console.warn(error);
        }
        var dlAppBanner = document.getElementsByClassName("app-notification-banner banner-show promo-banner accent-banner button-banner engagement-surface");
        if(dlAppBanner.length > 0) {
            dlAppBanner[0].remove();
        }
        // Wait
        await sleep(removeAnnoyDelay);
    }
}
async function cancelRemoveAnnoy() {
    // Stop the script
    settings.doRemoveAnnoy = false;
    manageSettings(true);
}

// Start title modifier observer task at last
observer.observe(target, config);