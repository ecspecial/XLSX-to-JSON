// Method to upload a valid excel file
function upload() {
    var files = document.getElementById('file_upload').files;
    if (files.length == 0) {
        alert("Пожалуйста, выберите файл...");
        return
    }
    var fileName = files[0].name;
    var extension = fileName.substring(fileName.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    }else{
        alert("Пожалуйста, выберите excel файл нужного формата.")
    }
}

// Method to read excel file and convert it into JSON 
function excelFileToJSON(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
            
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
            });
            var result = {};
            workbook.SheetNames.forEach(function (sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });
            // Displaying the json result
            // var resultEle=document.getElementById("json-result");
            // resultEle.value=JSON.stringify(result, null, 4);
            // resultEle.style.display='block';
            console.log(result);
        }
    }catch(e){
        console.error(e)
    }
}