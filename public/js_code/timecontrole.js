var first_point = "";
var time_today ="";
function senddata1(){
    var http = new XMLHttpRequest();
    http.open("POST", "/startwork", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            first_point= this.responseText;
            time_today = "Commencer à " + first_point + " => " + "Devrait partir à " +moment(first_point,"HH:mm").add(parseInt(document.getElementById("hour").value),"hours").format("HH:mm");
            info.innerHTML = "Lieu de travail : "+ chloc.value + " " + hour_today.value + " heures <br>" + time_today;
          }
          
      }
    };
    http.send("locaux="+document.getElementById("locaux").value+"&timework="+document.getElementById("hour").value);
}
function senddata2(locauxverif){
    var http = new XMLHttpRequest();
    http.open("POST", "/leftwork", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            window.location = "/session_end";
          }
          else{
            window.location = "/";
          }
          
      }
    };
    http.send("locaux="+locauxverif);
}
function senddata2_choice(locauxverif,choice){
  var http = new XMLHttpRequest();
  http.open("POST", "/handlework", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "error"){
          window.location = "/session_end";
        }
        else{
          window.location = "/";
        }
        
    }
  };
  http.send("locaux="+locauxverif+"&choice="+choice);
}
function senddata3(activity){
  var http = new XMLHttpRequest();
  http.open("POST", "/activity", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "error"){
          window.location = "/session_end";
        }
        else{
        
        }
        
    }
  };
  http.send("activity="+activity);
}
function send_changing_hour(){
  var http = new XMLHttpRequest();
  http.open("POST", "/changing", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
        if (this.responseText == "error"){
          window.location = "/session_end";
        }
        else{
          hour_today.value = ch_heure.value;
           time_today = "Commencer à " + first_point + " => " + "Devrait partir à " +moment(first_point,"HH:mm").add(parseInt(ch_heure.value),"hours").format("HH:mm");
          info.innerHTML = "Lieu de travail : "+ chloc.value + " " + hour_today.value + " heures <br>" + time_today;
          ch_heure.value = "";
        }
        
    }
  };
  http.send("ch_hour="+ch_heure.value);
}
function send_rem(m_code,rem){
  var http = new XMLHttpRequest();
  http.open("POST", "/rem", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      thanks();
    }
  };
  http.send("m_code="+m_code+"&rem="+rem);
}