// ==UserScript==
// @name			Office Online Viewer
// @description		View office documents in your browser with Microsoft Office Online
// @namespace		ogzergin
// @version        	0.1
// @include        	*
// @exclude        	http*://view.officeapps.live.com/*
// @exclude        	http*://docs.google.com/*
// @exclude        	http*://mail.google.com/*
// @exclude        	http*://viewer.zoho.com/*
// @exclude        	http*://office.live.com/*
// ==/UserScript==

var pageLinks=document.links;
var fileTypes=["doc","docx","xls","xlsx","ppt","pps","pptx"];

//https://view.officeapps.live.com/op/view.aspx?src=
var strOfficeHost ="view.officeapps.live.com";
var strViewOfficeUrl = "https://" + strOfficeHost + "/op/view.aspx?src=";

parseLinks();

addDebouncedEventListener(document, 'DOMNodeInserted', function(evt) {
    parseLinks();
}, 1000);

function endsWith(str, suffix) {  //  check if string has suffix 
    return str.indexOf(suffix, str.length - suffix.length) !== -1;
}

function stripQueryString(str) {
  return str.protocol + '//' + str.hostname + str.pathname;
}

function parseLinks(){
	
	for(var i=0;i<pageLinks.length;i++){
		if(pageLinks[i].isParsed != true && pageLinks[i].host != strOfficeHost){
			 parseLink(pageLinks[i]);
			 pageLinks[i].isParsed=true;
		} 
	}
}

function parseLink(link){

	var url = stripQueryString(link);

	if(checkType(url)){
		addOfficeLink(link);
	}
		
}

function checkType(str){
	for(var i=0; i<fileTypes.length; i++){
		if(endsWith(str, '.'+fileTypes[i]))
			return true;
	}
	return false;
}
	//word => http://i.imgur.com/Wce14X6.png
	//powerpoint => http://i.imgur.com/iQn5mKp.png
	// excel => http://i.imgur.com/xPTt5d6.png
function addOfficeLink(link){
	var officeLink = document.createElement('a');
	officeLink.href = strViewOfficeUrl + encodeURI(stripQueryString(link));
	officeLink.isParsed=true;
	officeLink.target="_blank";
	
	var ico = document.createElement("img");

	if(endsWith(officeLink.href, ".doc")||endsWith(officeLink.href, ".docx"))
		ico.src = "http://i.imgur.com/Wce14X6.png";
	else if(endsWith(officeLink.href, ".xls")|| endsWith(officeLink.href, ".xlsx"))
		ico.src = "http://i.imgur.com/xPTt5d6.png";
	else
		ico.src = "http://i.imgur.com/iQn5mKp.png";
		
	
	ico.style.marginLeft = "5px";
	officeLink.appendChild(ico);
	link.parentNode.insertBefore(officeLink, link.nextSibling);
	
}

function addDebouncedEventListener(obj, eventType, listener, delay) {
    var timer;

    obj.addEventListener(eventType, function(evt) {
        if (timer) {
            window.clearTimeout(timer);
        }
        timer = window.setTimeout(function() {
            timer = null;
            listener.call(obj, evt);
        }, delay);
    }, false);
}
