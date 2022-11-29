var btn_next = document.getElementById("next");
var btn_prev = document.getElementById("prev");
var act_number = 0;
var users;
var all_users;
var list_div = document.getElementById("list_name");
var search_text = document.getElementById("text");
var datestart = document.getElementById("start");
var dateend = document.getElementById("end");
var ouvrables = document.getElementById("ouvrable");
var ouvres = document.getElementById("ouvre");
var calendaire = document.getElementById("calendaire");
var result_page = document.getElementById("result_page");
(function() {
	function Init() {
		var fileSelect = document.getElementById('file-upload'),
			fileDrag = document.getElementById('file-drag'),
			submitButton = document.getElementById('submit-button');

		fileSelect.addEventListener('change', fileSelectHandler, false);

		// Is XHR2 available?
		var xhr = new XMLHttpRequest();
		if (xhr.upload) 
		{
			// File Drop
			fileDrag.addEventListener('dragover', fileDragHover, false);
			fileDrag.addEventListener('dragleave', fileDragHover, false);
			fileDrag.addEventListener('drop', fileSelectHandler, false);
		}
	}

	function fileDragHover(e) {
		var fileDrag = document.getElementById('file-drag');

		e.stopPropagation();
		e.preventDefault();
		
		fileDrag.className = (e.type === 'dragover' ? 'hover' : 'modal-body file-upload');
	}

	function fileSelectHandler(e) {
		// Fetch FileList object
		var files = e.target.files || e.dataTransfer.files;

		// Cancel event and hover styling
		fileDragHover(e);

		// Process all File objects
		for (var i = 0, f; f = files[i]; i++) {
			parseFile(f);
			uploadFile(f);
		}
	}

	function output(msg) {
		var m = document.getElementById('messages');
		m.innerHTML = msg;
	}

	function parseFile(file) {
		output(
			'<ul>'
			+	'<li>Name: <strong>' + encodeURI(file.name) + '</strong></li>'
			+	'<li>Type: <strong>' + file.type + '</strong></li>'
			+	'<li>Size: <strong>' + (file.size / (1024 * 1024)).toFixed(2) + ' MB</strong></li>'
			+ '</ul>'
		);
	}

	function setProgressMaxValue(e) {
		var pBar = document.getElementById('file-progress');

		if (e.lengthComputable) {
			pBar.max = e.total;
		}
	}

	function updateFileProgress(e) {
		var pBar = document.getElementById('file-progress');

		if (e.lengthComputable) {
			pBar.value = e.loaded;
		}
	}

	function uploadFile(file) {

		var xhr = new XMLHttpRequest(),
			fileInput = document.getElementById('class-roster-file'),
			pBar = document.getElementById('file-progress'),
			fileSizeLimit = 1024;	// In MB
		if (xhr.upload) {
			// Check if file is less than x MB
			if (file.size <= fileSizeLimit * 1024 * 1024) {
				// Progress bar
				pBar.style.display = 'inline';
				xhr.upload.addEventListener('loadstart', setProgressMaxValue, false);
				xhr.upload.addEventListener('progress', updateFileProgress, false);

				// File received / failed
				xhr.onreadystatechange = function(e) {
					if (xhr.readyState == 4) {
						if (this.responseText.includes("Erreur")){
							document.getElementById("error_process").innerHTML = this.responseText;
							document.getElementById("error_process").style.display = "block";
						}
						else{
							showing_result(this.responseText);
						}
						
					}
				};

				// Start upload
				xhr.open('POST', "paie", true);
				var filedata = new FormData();
				filedata.append("fileup",file);
				filedata.append("start",datestart.value);
				filedata.append("end",dateend.value);
				filedata.append("ouvrable",ouvrables.value);
				filedata.append("ouvre",ouvres.value);
				filedata.append("calendaire",calendaire.value);
				document.getElementById("show").style.display = "none";
				xhr.send(filedata);
			} else {
				output('Please upload a smaller file (< ' + fileSizeLimit + ' MB).');
			}
		}
	}

	// Check for the various File API support.
	if (window.File && window.FileList && window.FileReader) {
		Init();
	} else {
		document.getElementById('file-drag').style.display = 'none' ;
	}
})();
function showing_result(us){
	users = JSON.parse(us);
						all_users = users;
						document.getElementById("show").style.display = "block";
						rendu();
						change_color(users[0][1],users[0][0]);
						result_page.style.display = "block";
}
function next(){
		act_number = act_number+1;
		change_color(users[act_number][1],users[act_number][0]);
		btn_prev.disabled = false;
}
function prev(){
		act_number = act_number - 1 ;
		change_color(users[act_number][1],users[act_number][0]);
		btn_next.disabled = false;
}
function search(){
	if (search_text.value == ""){
		users = all_users;
		rendu();
		change_color(users[0][1],users[0][0]);
	}
	else {
		users = [];
		parcours();
	}
}
function rendu(){
	remove();
	for(i=0;i<users.length;i++){
		var h4 = document.createElement("h4");
		var texte = document.createTextNode(users[i][1]);
		h4.appendChild(texte);
		h4.setAttribute("class","list text-center name");
		h4.setAttribute("onclick","change_color('"+users[i][1]+"','"+users[i][0]+"')");
		list_div.appendChild(h4);
	}
}
function parcours(){
	for(p=0;p<all_users.length;p++){
		if (all_users[p][1].toLowerCase().includes(search_text.value.toLowerCase()) || all_users[p][0].toLowerCase().includes(search_text.value.toLowerCase())){
			users.push(all_users[p]);
		}
	}
	remove();
	rendu();
}
function remove(){
	var name_element = document.querySelectorAll(".name");
	for (r=0;r<name_element.length;r++){
		name_element[r].remove();
	}
}
function change_color(name,code){
	var name_element = document.querySelectorAll(".name");
	for (e=0;e<name_element.length;e++){
	  if (name == name_element[e].textContent){
		act_number = e;
		name_element[e].className = "clicked text-center name";
		document.getElementById("show").setAttribute("data",'/Paie/'+code+'.pdf');
		document.getElementById("navigator").innerHTML = (e+1) + " / " + name_element.length;
		if (act_number == users.length - 1){
			btn_next.disabled = true;
		}
		else{
			btn_next.disabled = false;
		}
		if (act_number  == 0){
			btn_prev.disabled = true;
		}
		else{
			btn_prev.disabled = false;
		}
	  }
	  else {
		name_element[e].className = "list text-center name";
	  }
	}
   
  }
  function vider(){
	window.location = "/empty";
  }