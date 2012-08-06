// ==UserScript==
// @name MP Automatic User Search
// @description takes the a parameter from another page and uses that to fill in the field to search
// @include http://caseviewer.codegear.net/MaintenanceReport/Default.aspx?*
// ==/UserScript==

	var thisURL = document.URL;
	var extractedEmail = thisURL.substring(62,thisURL.length);
	var atChar = extractedEmail.replace('%40','@');
		
	var bValue = document.getElementById('ctl00_ContentPlaceHolder1_EmailAddressBox');
		bValue.value = atChar;
		

	callEvent();
	
	var newURL = document.URL;
	
	if(newURL.indexOf("%") != -1)
		window.stop();
	
function callEvent(){

	var A = document.createEvent("MouseEvent");
		A.initMouseEvent("click", true, true, window,0, 0, 0, 0, 0, false, false, false, false, 0, null);

	var buttonGo = document.getElementById('ctl00_ContentPlaceHolder1_Button1');
		buttonGo.dispatchEvent(A); //Comment this line to disable automatic button click.
}

	
