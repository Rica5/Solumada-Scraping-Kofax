var period = document.getElementById("period");
var datestart = document.getElementById("datestart");
var dateend = document.getElementById("dateend");
var searchit = document.getElementById("searchit");
var take_date = false;
const queryString = window.location.search;
			const urlParams = new URLSearchParams(queryString);
			const ssearch = urlParams.get('ssearch');
			const speriod = urlParams.get('speriod');
			const sdatestart = urlParams.get('sdatestart');
			const sdateend = urlParams.get('sdateend');
			if (queryString.length != 0){
				if (speriod == "spec"){
                    document.getElementById("specific").style.display = "block";
                    period.style.display = "none";
                    take_date = true;
                    datestart.value = sdatestart;
                    dateend.value = sdateend;
                    searchit.value = ssearch;
                }
                else{
                    document.getElementById("specific").style.display = "none";
                    period.style.display = "block";
                    period.value = speriod;
                    searchit.value = ssearch;
                }
			}
function closing(){
    document.getElementById("specific").style.display = "none";
    period.style.display = "block";
    period.value = "";
    take_date = false;
}
function change_period(){
    if (period.value == "spec"){
        document.getElementById("specific").style.display = "block";
        period.style.display = "none";
        take_date = true;
    }
    else{
        document.getElementById("specific").style.display = "none";
        period.style.display = "block";
        take_date = false;
    }
}
function go_filter(){
    document.getElementById("filter").style.display = "none";
    document.getElementById("sanim").style.display = "block";
    if (take_date){
        filterspec();
    }
    else{
        filter();
    }
    
}
function filter(){
    var http = new XMLHttpRequest();
    http.open("POST", "/filter", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/";
          }
          else{

            window.location = "/details?speriod="+period.value+"&ssearch="+searchit.value;
          }
          
      }
    };
    http.send("period="+period.value+"&searchit="+searchit.value.replace(/\s/g, ''));
}
function filterspec(){
    var http = new XMLHttpRequest();
    http.open("POST", "/filter", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
              window.location = "/";
          }
          else{
            window.location = "/details?sdatestart="+datestart.value+"&sdateend="+dateend.value+"&ssearch="+searchit.value+"&speriod=spec";
          }
      }
    };
    http.send("period="+"spec"+"&datestart=" + datestart.value + "&dateend="+dateend.value + "&searchit="+searchit.value.replace(/\s/g, ''));
}