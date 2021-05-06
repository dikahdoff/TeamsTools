async function getCurrentTab() {
    let queryOptions = { active: true, currentWindow: true };
    let [tab] = await chrome.tabs.query(queryOptions);
    return tab;
}
function httpGet(theUrl)
{
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open( "GET", theUrl, false );
    xmlHttp.send( null );
    return xmlHttp.responseText;
}
async function popupTask() {
    var parsed = JSON.parse(httpGet("https://raw.githubusercontent.com/dikahdoff/TeamsUtils/main/scripts.json"));
    if(parsed != null) {
        document.getElementById("popup_title").innerText = parsed.title;
        document.getElementById("popup_header_logo").src = parsed.icon;
        var subtitle = document.getElementById("popup_subtitle");
        var manifestData = chrome.runtime.getManifest();
        subtitle.innerText = parsed.subtitle + "\n(v" + ((manifestData.version == parsed.latestVersion) ? manifestData.version : manifestData.version + ", new version available: v" + parsed.latestVersion) + ")";
        subtitle.setAttribute("href", "repo.html");
        subtitle.setAttribute("target", "_blank");
        var currentTab = await getCurrentTab();
        if(currentTab.url.toString().startsWith(parsed.teamsLink)) {
            var scripts = parsed.scripts;
            var tutorial = document.createElement("img");
			var tutorialTxt = document.createElement("p");
			tutorialTxt.innerHTML = "Use this menu to access the extension's functions.";
			document.getElementById("popup_utils").appendChild(tutorialTxt);
			tutorial.src = "tutorial.png";
			tutorial.setAttribute("style", "width:100%");
			document.getElementById("popup_utils").appendChild(tutorial);
        } else {
            var errorItem = document.createElement("h3");
            errorItem.innerHTML = "Please navigate to <a href='teams.html' target='_blank'>"+parsed.teamsLink+"</a> to use this extension.";
            document.getElementById("popup_utils").appendChild(errorItem);
        }
    } else {
        console.error("Failed to load scripts.json.");
    }
}
popupTask();