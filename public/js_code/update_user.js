//Add User
var add_email = document.getElementById("add_email");
var add_name = document.getElementById("add_nom");
var add_last_name = document.getElementById("add_prenom");
var add_usuel = document.getElementById("add_usuel");
var add_mcode = document.getElementById("add_mcode");
var add_num_agent = document.getElementById("add_num_agent");
var add_matricule = document.getElementById("add_matricule");
var add_function = document.getElementById("add_fonction");
var add_occupation = document.getElementById("add_occup");
var add_cin = document.getElementById("add_cin");
var add_sexe = document.getElementById("add_sexe");
var add_situation = document.getElementById("add_situation");
var add_adresse = document.getElementById("add_adresse");
var add_cnaps = document.getElementById("add_cnaps");
var add_class = document.getElementById("add_class");
var add_contrat = document.getElementById("add_contrat");
var add_embauche = document.getElementById("add_embauche");
var btn_add = document.getElementById("btn_add");
var add_element = document.querySelectorAll(".add_user");
function verify_add_input() {
  var deactivate = 0;
  for (ae = 0; ae < add_element.length; ae++) {
    if (add_element[ae].value == "") {
      deactivate++;
      add_element[ae].style = "border-color:red";
    }
    else {
      add_element[ae].style = "";
    }
  }
  if (deactivate != 0) {
    btn_add.disabled = true;
  }
  else {
    btn_add.disabled = false;
  }
}

function enregistrer() {
  sendRequest("/addemp",
    add_email.value,
    add_name.value,
    add_last_name.value,
    add_usuel.value,
    add_mcode.value,
    add_num_agent.value,
    add_matricule.value,
    add_function.value,
    add_occupation.value,
    add_embauche.value,
    add_cin.value,
    add_sexe.value,
    add_situation.value,
    add_adresse.value,
    add_cnaps.value,
    add_class.value,
    add_contrat.value,
  )
}
function sendRequest(url, email, nom, prenom, usuel, mcode, num_agent, matricule, fonction, occ, embauche, cin, sexe, situation, adresse, cnapsnum, classify, contrat) {
  var http = new XMLHttpRequest();
  http.open("POST", url, true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      if (this.responseText.includes("already")) {
        document.getElementById("notif").setAttribute("style", "background-color:red");
        showNotif("Utilisateur déja enregistrés avec le même M-code/email/numéro agent/matricule");
      }
      else if (this.responseText == "error") {
        showNotif("Une erreur s'est produite dans le serveur, reesayez ou contacter un technicien");
      }
      else {
        document.getElementById("notif").setAttribute("style", "background-color:limeagreen");
        showNotif("Utilisateur " + this.responseText + " enregistré");
      }
    }
  };
  http.send("email=" + email +
    "&name=" + nom +
    "&last_name=" + prenom
    + "&usuel=" + usuel +
    "&mcode=" + mcode +
    "&num_agent=" + num_agent +
    "&matricule=" + matricule +
    "&function_choosed=" + fonction +
    "&occupation=" + occ +
    "&enter_date=" + embauche +
    "&cin=" + cin +
    "&gender=" + sexe +
    "&situation=" + situation +
    "&location=" + adresse +
    "&num_cnaps=" + cnapsnum +
    "&classification=" + classify +
    "&contrat=" + contrat);
}














//Updating _user
var up_email = document.getElementById("up_email");
var up_name = document.getElementById("up_nom");
var up_last_name = document.getElementById("up_prenom");
var up_usuel = document.getElementById("up_usuel");
var up_mcode = document.getElementById("up_mcode");
var up_num_agent = document.getElementById("up_num_agent");
var up_matricule = document.getElementById("up_matricule");
var up_function = document.getElementById("up_fonction");
var up_occupation = document.getElementById("up_occup");
var up_cin = document.getElementById("up_cin");
var up_sexe = document.getElementById("up_sexe");
var up_situation = document.getElementById("up_situation");
var up_adresse = document.getElementById("up_adresse");
var up_cnaps = document.getElementById("up_cnaps");
var up_class = document.getElementById("up_class");
var up_contrat = document.getElementById("up_contrat");
var up_embauche = document.getElementById("up_embauche");
var btn_up = document.getElementById("btn_up");
var up_element = document.querySelectorAll(".up_user");
var all_id = ["up_email", "up_nom", "up_prenom", "up_usuel", "up_mcode", "up_num_agent", "up_matricule", "up_fonction", "up_occup", "up_embauche", "up_cin",
  "up_sexe", "up_situation", "up_adresse", "up_cnaps", "up_class", "up_contrat"];
var select_field = ["up_fonction", "up_occup", "up_sexe", "up_situation"];
var ids = "";
function verify_up_input() {
  var deactivate = 0;
  for (up = 0; up < up_element.length; up++) {
    if (up_element[up].value == "") {
      deactivate++;
      up_element[up].style = "border-color:red";
    }
    else {
      up_element[up].style = "";
    }
  }
  if (deactivate != 0) {
    btn_up.disabled = true;
  }
  else {
    btn_up.disabled = false;
  }
}
function change_identifiant() {
  up_mcode.value = "N/A";
  up_num_agent.value = "N/A";
  up_matricule.value = "N/A";
}
function getdata(url, id) {
  var http = new XMLHttpRequest();
  http.open("POST", url, true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      var data = this.responseText.split(",");
      for (put = 0; put < data.length; put++) {
        if (select_field.includes(all_id[put])) {
          document.getElementById(all_id[put]).value = data[put];
          document.getElementById("select2-" + all_id[put] + "-container").innerHTML = data[put];
        }
        else {
          document.getElementById(all_id[put]).value = data[put];
        }
      }
      ids = id;
      console.log(ids);
    }
  };
  http.send("id=" + id);
}

function modifier() {
  update_user("/updateuser"
    , ids,
    up_email.value,
    up_name.value,
    up_last_name.value,
    up_usuel.value,
    up_mcode.value,
    up_num_agent.value,
    up_matricule.value,
    up_function.value,
    up_occupation.value,
    up_embauche.value,
    up_cin.value,
    up_sexe.value,
    up_situation.value,
    up_adresse.value,
    up_cnaps.value,
    up_class.value,
    up_contrat.value,
  );
}
function update_user(url, id, email, nom, prenom, usuel, mcode, num_agent, matricule, fonction, occ, embauche, cin, sexe, situation, adresse, cnapsnum, classify, contrat) {
  var http = new XMLHttpRequest();
  http.open("POST", url, true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      if (this.responseText.includes("already")) {
        document.getElementById("notif").setAttribute("style", "background-color:red");
        showNotif("L'utilisateur existe déja  (m-code/numbering agent/email/matricule)");
      }
      else if (this.responseText == "error") {
        showNotif("Erreur dans le modification");
      }
      else {
        document.getElementById("notif").setAttribute("style", "background-color:limeagreen");
        showNotif("La modification s'est fait avec succés");
      }
    }
  };
  http.send("id=" + id +
    "&email=" + email +
    "&name=" + nom +
    "&last_name=" + prenom
    + "&usuel=" + usuel +
    "&mcode=" + mcode +
    "&num_agent=" + num_agent +
    "&matricule=" + matricule +
    "&function_choosed=" + fonction +
    "&occupation=" + occ +
    "&enter_date=" + embauche +
    "&cin=" + cin +
    "&gender=" + sexe +
    "&situation=" + situation +
    "&location=" + adresse +
    "&num_cnaps=" + cnapsnum +
    "&classification=" + classify +
    "&contrat=" + contrat);
}
function showNotif(text) {
  const notif = document.querySelector('.notification');
  notif.innerHTML = text;
  notif.style.display = 'block';
  setTimeout(() => {
    notif.style.display = 'none';
    window.location = "/userlist";
  }, 2000);
}

function delete_user(user) {
  textwarn.innerHTML = "Are you sure to delete user <b>" + user + "</b>";
  textwarn.setAttribute("style", "color:aliceblue");
  del = user;
}
function confirm_del() {
  drop_user("/dropuser", del);
}
function drop_user(url, fname) {
  var http = new XMLHttpRequest();
  http.open("POST", url, true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      if (this.responseText == "error") {
        window.location = "/";
      }
      else {
        showNotif(this.responseText);
      }
    }
  };
  http.send("fname=" + fname);
}
function update_project_user(m_c) {
  var options = document.getElementById(m_c).selectedOptions;
  var values = Array.from(options).map(({ value }) => value);
  send_project(m_c, values);
}
function send_project(owner, project_choose) {
  var http = new XMLHttpRequest();
  http.open("POST", "/update_project", true);
  http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
  http.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
      if (this.responseText == "error") {
        window.location = "/";
      }
      else {

      }
    }
  };
  http.send("owner=" + owner + "&choice=" + project_choose);
}
