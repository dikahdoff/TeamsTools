function httpGet(theUrl)
{
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open( "GET", theUrl, false );
    xmlHttp.send( null );
    return xmlHttp.responseText;
}
var parsed = JSON.parse(httpGet("https://raw.githubusercontent.com/dikahdoff/TeamsUtils/main/scripts.json"));
var elem = document.createElement("meta");
elem.setAttribute("http-equiv","refresh");
elem.setAttribute("content", "0; URL=" + parsed.link);
document.getElementById("appendThis").appendChild(elem);