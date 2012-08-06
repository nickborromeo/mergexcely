// ==UserScript==
// @name sancOrderList
// @description scripts that modify the Order Fulfillment list
// @include https://license-admin.codegear.com/srs6/admin/el/order_list.jsp*
// ==/UserScript==

var pTableBody = document.getElementsByTagName("tbody");
var pTRow = pTableBody[1].getElementsByTagName("tr");
//Create the Storage and Display DIVS and Gets the URL to be passed from the parent page

for(var j=2; j<pTRow.length; j++){
    
    var sRow = document.createElement("td");
        //sRow.setAttribute("width", "600")

    var storageDiv = document.createElement("div");
        storageDiv.setAttribute("id","storage"+j);
        storageDiv.setAttribute("style", "display:none");
    
    
    var displayDiv = document.createElement("div");
        displayDiv.setAttribute("id","display"+j);
        //displayDiv.innerHTML = "DISPLAY";
                
    sRow.appendChild(displayDiv);
    sRow.appendChild(storageDiv);
    pTRow[j].appendChild(sRow);
    //pTRow[j].appendChild(storageDiv);

    
    var pTData = pTRow[j].getElementsByTagName("td");
    //URL of other page with info to grab
    var pURL =  pTData[0].getElementsByTagName("a")[0].href.trim();
    //alert("before the div");
    
    accessByDOM(pURL,j);
    //alert("passed the URL");
}

//gets the information from the accessed hidden page - to be displayed in parent



function getCertInformation(){
    
    
    return itemList;
}




/*===============AJAX====================*/*
function createXHR() 
{
    var request = false;
        try {
            request = new ActiveXObject('Msxml2.XMLHTTP');
        }
        catch (err2) {
            try {
                request = new ActiveXObject('Microsoft.XMLHTTP');
            }
            catch (err3) {
		try {
			request = new XMLHttpRequest();
		}
		catch (err1) 
		{
			request = false;
		}
            }
        }
    return request;
}


function getBody(content) 
{
   test = content.toLowerCase();    // to eliminate case sensitivity
   var x = test.indexOf("<body");
   if(x == -1) return "";

   x = test.indexOf(">", x);
   if(x == -1) return "";

   var y = test.lastIndexOf("</body>");
   if(y == -1) y = test.lastIndexOf("</html>");
   if(y == -1) y = content.length;    // If no HTML then just grab everything till end

   return content.slice(x + 1, y);   
} 

function loadHTML(url, fun, storage, param)
{
	var xhr = createXHR();
	xhr.onreadystatechange=function()
	{ 
		if(xhr.readyState == 4)
		{
			//if(xhr.status == 200)
			{
				storage.innerHTML = getBody(xhr.responseText);
				fun(storage, param);
			}
		} 
	}; 

	xhr.open("GET", url , true);
	xhr.send(null); 

} 

function processByDOM(responseHTML, target)
{
        var tableBody = responseHTML.getElementsByTagName("tbody");
        var tRow = tableBody[3].getElementsByTagName("tr");
        
        var itemList = "";
        for(var i=2; i<tRow.length; i++){
            var tData = tRow[i].getElementsByTagName("td");
            //these variables get the certificate number and the description
            var certNum = tData[0].childNodes[3].nextSibling.nodeValue.trim();
            var description = "";
            var PID = "";
            if(tData[0].getElementsByTagName("b")[4].innerHTML == "Product Description:")
                description =  tData[0].getElementsByTagName("b")[4].nextSibling.nodeValue.trim();
            else if(tData[0].getElementsByTagName("b")[5].innerHTML == "Product Description:")
                description =  tData[0].getElementsByTagName("b")[5].nextSibling.nodeValue.trim();
            else if(tData[0].getElementsByTagName("b")[6].innerHTML == "Product Description:")
                description =  tData[0].getElementsByTagName("b")[6].nextSibling.nodeValue.trim();
            else if(tData[0].getElementsByTagName("b")[7].innerHTML == "Product Description:")
                description =  tData[0].getElementsByTagName("b")[7].nextSibling.nodeValue.trim();
            //itemList += ;
            
            if(tData[0].getElementsByTagName("b")[3].innerHTML == "Pid:")
                PID =  tData[0].getElementsByTagName("b")[3].nextSibling.nodeValue.trim();
            
            target.innerHTML += certNum +" - "+description+" | PID: <a href=https://na7.salesforce.com/search/SearchResults?searchType=2&str="+PID+"&search=Search&sen=02i target=_blank>"+PID+ "</a><br/>";
        }
        
}
	
	
function accessByDOM(url,index)
{
        //var responseHTML = document.createElement("body");	// Bad for opera
        var responseHTML = document.getElementById("storage"+index);
        var y = document.getElementById("display"+index);
        loadHTML(url, processByDOM, responseHTML, y);
}
//*/