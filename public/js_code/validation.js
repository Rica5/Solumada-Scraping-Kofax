//Valide if true
function validate_true(validate){
    sendRequest_true('/validate',validate);
}
//Denied if false
function validate_false(ids){
    sendRequest_false('/denied',ids);
}
const queryString = window.location.search;
			const urlParams = new URLSearchParams(queryString);
			const search = urlParams.get('search');
//listener filter
			setTimeout(()=>{
				filter = document.querySelector("input[type='search']");
				var table = $('#zero_config').DataTable();
				if (search != null){
					table.search( search ).draw();
				}
				
			},1000)

function sendRequest_true(url,id) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
          window.location = "/validelate?search="+filter.value;
      }
    };
    http.send("id="+id);
  }
  function sendRequest_false(url,id) {
    var http = new XMLHttpRequest();
    http.open("POST", url, true);
    http.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    http.onreadystatechange = function () {
      if (this.readyState == 4 && this.status == 200) {
        window.location = "/validelate?search="+filter.value;
      }
    };
    http.send("id="+id);
  }