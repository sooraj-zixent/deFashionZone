
var ExcelParser = function(){
  this.parseFile = function(callback){

    /* set up XMLHttpRequest */
    var url = "dataFiles/item-list.xlsx";
    var oReq = new XMLHttpRequest();
    oReq.open("GET", url, true);
    oReq.responseType = "arraybuffer";
    oReq.onload = function(e) {
      var arraybuffer = oReq.response;

      /* convert data to binary string */
      var data = new Uint8Array(arraybuffer);
      var arr = new Array();
      for(var i = 0; i != data.length; ++i) {
        arr[i] = String.fromCharCode(data[i]);
      }
      var bstr = arr.join("");

      /* Call XLSX */
      var workbook = XLSX.read(bstr, {type:"binary"});
      /* DO SOMETHING WITH workbook HERE */
      var first_sheet_name = workbook.SheetNames[0];
      /* Get worksheet */
      var worksheet = workbook.Sheets[first_sheet_name];
      var json = XLSX.utils.sheet_to_json(worksheet,{raw:true})
      callback({errorMap:validateDataFile(json), data : json});
    }
     oReq.send();

    function validateDataFile(dataFile){
        var errorMessage = function(row,columnName,msg) {
          return ret = {
            message : msg,
            column  : columnName, 
            rowNumber: row
          }
        }
        var idList = [];
        var codeList = [];
        var urlList = [];
        var itemCategories = ["kurtis","saree"];
        var errorList = [];
        var i = 0;
        dataFile.forEach(function(data) {
          i++; 
          // Validate ID
          if(data.id == undefined || isNaN(data.id) || data.id < 0 ){
            errorList.push(new errorMessage(i,"id","ID invalid "+data.id));
          }else if(idList.indexOf(data.id) >=0) {
            errorList.push(new errorMessage(i,"id","Duplicate id :"+data.id));
          }else{
              idList.push(data.id);
          }

          // Validate code
          if(data.code == undefined || data.code.length < 4  ){
            errorList.push(new errorMessage(i,"code","Code invalid "+data.code));
          }else if(codeList.indexOf(data.code) >= 0) {
            errorList.push(new errorMessage(i,"code","Duplicate Code :"+data.code));
          }else{
              codeList.push(data.code);
          }

          // Validate itemUrl
          if(data.itemUrl == undefined || data.itemUrl.length < 4 || codeList.indexOf(data.itemUrl)>0 ){
            errorList.push(new errorMessage(i,"itemUrl","Url invalid "+data.itemUrl));
          }else if(urlList.indexOf(data.itemUrl) >= 0) {
            errorList.push(new errorMessage(i,"itemUrl","Duplicate Item Url :"+data.itemUrl));
          }else{
              urlList.push(data.itemUrl);
          }

          // Validate itemCategory
          if(data.itemCategory == undefined || data.itemCategory.length < 4 || itemCategories.indexOf(data.itemCategory) < 0 ){
            errorList.push(new errorMessage(i,"itemCategory","Category invalid "+data.itemCategory));
          }
          // Validate name
          if(data.name == undefined || data.name.length < 5 ){
            errorList.push(new errorMessage(i,"name","Name invalid "+data.name));
          }
          // Validate shortDescription
          if(data.shortDescription == undefined || data.shortDescription.length < 10 ){
            errorList.push(new errorMessage(i,"shortDescription","longDescription invalid "+data.shortDescription));
          }
          // Validate longDescription
          if(data.longDescription == undefined || data.longDescription.length < 10 ){
            errorList.push(new errorMessage(i,"longDescription","longDescription invalid "+data.longDescription));
          }
          // Validate longDescription
          if(data.originalPrice == undefined || isNaN(data.originalPrice) || data.originalPrice < 0  ){
            errorList.push(new errorMessage(i,"originalPrice","originalPrice invalid "+data.originalPrice));
          }
          // Validate sellingPrice
          if(data.sellingPrice == undefined || isNaN(data.sellingPrice) || data.sellingPrice < 0 || data.sellingPrice > data.originalPrice  ){
            errorList.push(new errorMessage(i,"sellingPrice","sellingPrice invalid "+data.sellingPrice));
          } 

          if(data.numberOfImages == undefined || isNaN(data.numberOfImages) || data.numberOfImages < 1 ||  data.numberOfImages > 10  ){
            errorList.push(new errorMessage(i,"numberOfImages","numberOfImages invalid"));
          }else{
            for(var j = 1; j <= data.numberOfImages; j++){
                var imageData = data["image"+j];
                if(imageData == undefined || imageData.length < 5){
                    errorList.push(new errorMessage(i,"image"+j,"image"+j+" invalid "));
                }
            }
          }
        
      });
      return errorList;
    }
  }
}