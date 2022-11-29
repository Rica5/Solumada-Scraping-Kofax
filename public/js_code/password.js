var email = document.getElementById("email");
var code = document.getElementById("code");
var pw1 = document.getElementById("pw1");
var pw2 = document.getElementById("pw2");
var dmail = document.getElementById("mail");
var dcode = document.getElementById("verification");
var dpass = document.getElementById("pass");
var denied = document.getElementById("denie");
var info = document.getElementById("info");
var p1 = document.getElementById("p1");
var p2 = document.getElementById("p2");
var p_done = false;
var p1_done = false;
var p2_done = false;
var r = document.getElementById("r");
var v = document.getElementById("v");
var c = document.getElementById("c");
var time = 6;
var redirection = document.getElementById("redirection");

function checkmail(){
    r.innerHTML = "VERIFICATION...";
    var http = new XMLHttpRequest();
    http.open("POST", "/checkmail", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
              denied.style.display = "block";
             denied.innerHTML = "L'email que vous avez saisi est non enregistrer";
             r.innerHTML = "REINITIALISER";
          }
          else{
            denied.style.display = "none";
              dmail.style.display = "none";
            dcode.style.display = "block";
            info.style.display = "block";
            info.innerHTML = "Email verifier, un code vous a été envoyé";
          }
          
      }
    };
    http.send("email=" + email.value.replace(/\s/g, ''));
}
function checkcode(){
    v.innerHTML = "VERIFICATION...";
    var http = new XMLHttpRequest();
    http.open("POST", "/checkcode", true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
            denied.style.display = "block";
             denied.innerHTML = "Desole le code est faux";
             v.innerHTML = "VERIFIER";
          }
          else{
            denied.style.display = "none";
              dcode.style.display = "none";
              dpass.style.display = "block";
              info.style.display = "block";
              info.innerHTML = "Code verifier";
          }
          
      }
    };
    http.send("code=" + code.value.replace(/\s/g, ''));
}
function change_pass(){
    if (verify_all()){
        var http = new XMLHttpRequest();
        http.open("POST", "/changepass", true);
        http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        http.onreadystatechange = function () {
          if (this.readyState == 4 && this.status == 200) {
              if (this.responseText == "error"){
                denied.style.display = "block";
                 denied.innerHTML = "Le mot de passe ne peut pas être le meme qu'avant";
              }
              else{
                denied.style.display = "none";
                info.style.display = "block";
                  info.innerHTML = "Mot de passe changer avec success";
                  clockRunner();
              }
              
          }
        };
        http.send("pass=" + pw1.value);
    }
   
}
function verify_pass1(){
    if (pw1.value.length < 8){
        p1.style.display = "block";
        p1.innerHTML = "Le mot de passe doit avoir 8 caracteres minimum";
        p1_done = false;
    }
    else if (pw1.value.replace(/\s/g, '') == ""){
        p1.style.display = "block";
        p1.innerHTML = "Le champ mot de passe est requis";
        p1_done = false;
    }
    else{
        p1.style.display = "none";
        p1_done = true;
    }
    verify_pass2();
    
}
function verify_pass2(){
    if (pw2.value != pw1.value){
        p2.style.display = "block";
        p2.innerHTML = "Mot de passe non identique";
        p2_done = false;
    }
    else{
        p2.style.display = "none";
        p2_done = true;
    }
}
function verify_all(){
    if (p1_done && p2_done){
        return true;
    }
    else{
        return false;
    }
}
function clockRunner () {
    time--;
    redirection.innerHTML = "Redirection dans " + time;
    if (time<=0){
        window.location = "/";
    }
    else{
        setTimeout(clockRunner, 1000);

    }
    
  }

