var email = document.getElementById("email");
var fname = document.getElementById("fname");
var lname = document.getElementById("lname");
var occup = document.getElementById("occup");
var m_code = document.getElementById("m_code");
var num_agent = document.getElementById("num_agent");
var shift = document.getElementById("shift");
var em = document.getElementById("em");
var fn = document.getElementById("fn");
var ln = document.getElementById("ln");
var oc = document.getElementById("oc");
var co = document.getElementById("co");
var nu = document.getElementById("nu");
var sh = document.getElementById("sh");
var email_done = false;
var fname_done = false;
var lname_done = false;
var occup_done = false;
var code_done = false;
var agent_done = false;
var passed_done = true;
var shift_done = false;
var btnsave = document.getElementById("save");
var success = document.getElementById("success");
var denied = document.getElementById("denie");

function verify_email(){
    if(email.value.includes("@") && email.value.replace(/\s/g, '') != ""){
        email_done = true;
        email.removeAttribute("style");
        em.style.display = "none";
    }
    else{
        em.style.display = "block";
        email_done = false;
        passed_done = false;
    }
    verify_all();
}
function verify_fname(){
    if(fname.value.replace(/\s/g, '') != ""){
        fname_done = true;
        fn.style.display = "none";
    }
    else{
        fn.style.display = "block";
        fname_done = false;
        passed_done = false;
    }
    verify_all();
}
function verify_lname(){
    if(lname.value.replace(/\s/g, '') != ""){
        lname_done = true;
        ln.style.display = "none";
    }
    else{
        ln.style.display = "block";
        lname_done = false;
        passed_done = false;
    }
    verify_all();
}
function verify_occup(){
    if(occup.value.replace(/\s/g, '') != ""){
        if (occup.value == "Admin" || occup.value == "Checker"){
            m_code.disabled = true;
            m_code.value = "N/A";
            num_agent.disabled = true;
            num_agent.value = "N/A";
            co.style.display = "none";
            nu.style.display = "none";
        }
        else{
            m_code.disabled = false;
            num_agent.disabled = false;
        }
        occup_done = true;
        oc.style.display = "none";
    }
    else{
        
        oc.style.display = "block";
        occup_done = false;
        passed_done = false;
    }
    verify_all();
}
function verify_code(){
    if(m_code.value.replace(/\s/g, '') != ""){
        code_done = true;
        co.style.display = "none";
    }
    else{
        co.style.display = "block";
        code_done = false;
        passed_done = false;
    }
    verify_all();
}
function verify_agent(){
    if(num_agent.value.replace(/\s/g, '') != ""){
        agent_done = true;
        nu.style.display = "none";
    }
    else{
        nu.style.display = "block";
        agent_done = false;
        passed_done = false;
    }
    verify_all();
}

function verify_all(){
    if(email_done && fname_done && lname_done && occup_done && code_done && agent_done && shift_done){
        passed_done = true;
    }
    else{
        passed_done = false;
    }
}
function verify_shift(){
    if (shift.value == ""){
        sh.style.display = "block";
        passed_done = false;
        shift_done = false;
    }
    else{
        sh.style.display = "none";
        shift_done = true;
    }
}
function save(){
    passed_done = true;
    verify_email();verify_fname();verify_lname();verify_occup();verify_code();verify_agent();verify_shift();
    if (occup.value == "Admin" || occup.value == "Checker"){
        code_done = true;
        agent_done = true;
        shift_done = true;
        verify_all();
    }
    if(passed_done){
       sendRequest("/addemp",email.value,m_code.value,num_agent.value,fname.value,lname.value,occup.value,shift.value);
    }
   
}
function sendRequest(url, mail,code,agent,firsts,lasts,occupation,shifts) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          if (this.responseText == "error"){
             success.style.display = "none";
             denied.style.display = "block";
             denied.innerHTML = "L'utilisateur est déja enregistrés";
          }
	      else if (this.responseText == "retour"){
		      window.location = "/";
	      }
          else{
            success.style.display = "block";
            denied.style.display = "none";
            success.innerHTML = "Utilisateur "+ this.responseText + " enregistré";
            email.value = "";
             fname.value = "";
             lname.value ="";
             occup.value = "";
            m_code.value ="";
            num_agent.value = "";
            shift.value = "";
            email_done = false;fname_done = false;lname_done = false;occup_done = false;code_done = false;agent_done = false;passed_done = true;shift_done = false;
          }
          
      }
    };
    http.send("email=" + mail + "&mcode=" + code + "&num_agent=" + agent +"&first_name="+firsts+"&last_name="+lasts+"&occupation="+occupation+"&shift="+shifts);
  }
