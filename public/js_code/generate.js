var loading = document.getElementById("loading");
var btnd = document.getElementById("download");
var btng = document.getElementById("generate_excel");
btnd.disabled =true;
function sendRequest_generate(url) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        loading.style.display = "none";
        btnd.disabled = false;
        btng.disabled = true;
      }
    };
    http.send();
  }
  function generate(){
    loading.style.display = "block";
      sendRequest_generate("/generate");
  }
  function downloads(){
    btnd.disabled =true;
  }