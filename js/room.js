function readExcel1() {
  let currentDt = currentDate();
  let testDt = "02/28/2021";
  console.log(currentDt);
  let input = event.target;
  let reader = new FileReader();
  reader.onload = function () {
    let data = reader.result;
    let workBook = XLSX.read(data, { type: 'binary' });
    workBook.SheetNames.forEach(function (sheetName) {
      if(sheetName.split(" ")[1] != "Departure"){
        alert("Departure 문서가 아닙니다.");
        return false;
      }
      let rows = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName]);
      const datas = rows.map(parent => {
        return Object.keys(parent).reduce((acc, key) => ({
          ...acc,
          [key.replace(/\s/g, "")]: parent[key],
        }), {});
      });
      datas.forEach(row => {
        if(testDt == row.OrgDepDate){
          var roomUId = document.getElementById("s_"+row.RmNo);
          if(roomUId != null){
            roomUId.innerHTML = "C/O";
            roomUId.style.color = "red";
          }
        }
      });
    });
  };
  reader.readAsBinaryString(input.files[0]);
}

function readExcel2() {
  let input = event.target;
  let reader = new FileReader();
  reader.onload = function () {
    let data = reader.result;
    let workBook = XLSX.read(data, { type: 'binary' });
    workBook.SheetNames.forEach(function (sheetName) {
      if(sheetName.split(" ")[2] != "Summary"){
        alert("Summary 문서가 아닙니다.");
        return false;
      }
      let rows = XLSX.utils.sheet_to_json(workBook.Sheets[sheetName]);
      const datas = rows.map(parent => {
        return Object.keys(parent).reduce((acc, key) => ({
          ...acc,
          [key.replace(/\s/g, "")]: parent[key],
        }), {});
      });
      

      datas.forEach(row => {
        console.log(JSON.stringify(row));
        if((row.RoomStatus == "Vacant" && row.CleanStatus == "Cleaned")
          || row.RoomStatus == "Out Of Order"){
          let roomStatus = "V";
          let color = "black";
          if(row.RoomStatus == "Out Of Order"){
            roomStatus = "O.O"
            color = "black";
          }
          delete row.RoomStatus;
          delete row.CleanStatus;
          for(key in row){
            var roomNo = row[key].split(" ")[0];
            var roomUId = document.getElementById("s_"+roomNo);
            var roomSId = document.getElementById("s_"+roomNo);
            if(roomUId != null){
              roomSId.innerHTML = roomStatus;
              roomSId.style.color=color;
            }
          }
        }
      });
    });
  };
  reader.readAsBinaryString(input.files[0]);
}

function reRoomCheck(){
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  for(var i = 3 ; i <= 16 ; i++){
    switch (i) {
      case 3 : 
        startRoomNo = 5;
        endRoomNo = 22;
        break;
      case 4 : 
        endRoomNo = 34;
      case 5 : 
        endRoomNo = 35;
      case 16 : 
        endRoomNo = 6;
    }
    if(i < 10){
      floorNo = "0" + i;
    }
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(i < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + j;
      }
      var roomSId = document.getElementById("s_"+ roomNo);
      if(roomSId == null) {
        
      }
    }
  }
}

function currentDate(){
  var date = new Date(); 
  var year = date.getFullYear(); 
  var month = new String(date.getMonth()+1); 
  var day = new String(date.getDate()); 

  if(month.length == 1){ 
    month = "0" + month; 
  } 
  if(day.length == 1){ 
    day = "0" + day; 
  } 
  var currentDt = day + "/" + month + "/" + year
  return currentDt
}

function printPage(){
  console.log(1);
  var initBody;
  window.onbeforeprint = function(){
   initBody = document.body.innerHTML;
   document.body.innerHTML =  document.getElementById('content').innerHTML;
  };
  window.onafterprint = function(){
   document.body.innerHTML = initBody;
  };
  window.print();
  return false;
}