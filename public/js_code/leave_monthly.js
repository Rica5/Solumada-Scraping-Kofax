var datestart = document.getElementById("datestart");
var dateend = document.getElementById("dateend");
var loading = document.getElementById("sanim");
var btng = document.getElementById("generate_excel");
var download = document.getElementById("download");
var loading_excel  = document.getElementById("loading");
download.disabled = true;
var leave_id;
function go_filter(){
    loading.style.display = "block";
    send_leave(datestart.value,dateend.value);
}
function send_leave(d1,d2){
    var http = new XMLHttpRequest();
    http.open("POST", "/monthly_leave", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            loading.style.display = "none";
            window.location = "/leavelist";
          }
          
      }
    };
    http.send("datestart="+d1 + "&dateend=" + d2);
}
function generate(){
    loading.style.display = "block";
    var http = new XMLHttpRequest();
    http.open("POST", "/leave_report", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            loading.style.display = "none";
            download.disabled = false;
          }
          
      }
    };
    http.send();
}
function downloads(){
    btnd.disabled =true;
  }
  function send_delete(){
    var http = new XMLHttpRequest();
    http.open("POST", "/delete_leave", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          window.location = "/leavelist";
      }
    }
    http.send("id="+leave_id);
  }
 
  function delete_leave(id,nom){
    leave_id = id;
    document.getElementById("del_leave").innerHTML = "Supprimer le congé de " + nom;
  }

