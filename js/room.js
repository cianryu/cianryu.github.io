let fileCnt = 0;
let vCnt = 0;
let cCnt = 0;
let ooCnt = 0;
window.addEventListener("load", function(event) {
  let current = new Date();
  let year = current.getFullYear();
  let month = current.getMonth()+1;
  let day = current.getDate();
  let currentDt = year + "년 " + month + "월 " + day + "일";
  document.getElementById("currentDt").innerHTML = currentDt;
});

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
          var roomSId = document.getElementById("s_"+row.RmNo);
          if(roomSId != null){
            roomSId.innerHTML = "C";
            roomSId.style.color = "red";
            ++cCnt;
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
        let roomStatus = "V";
        let color = "black";
        let size = "small";
        if((row.RoomStatus == "Vacant" && row.CleanStatus == "Cleaned")){
          delete row.RoomStatus;
          delete row.CleanStatus;
          for(key in row){
            var roomNo = row[key].split(" ")[0];
            var roomSId = document.getElementById("s_"+roomNo);
            if(roomSId != null){
              roomSId.innerHTML = roomStatus;
              roomSId.style.color=color;
              roomSId.style.fontSize=size;
              ++vCnt;
            }
          }
        }
        if(row.RoomStatus == "Out Of Order"){
          roomStatus = "O.O"
          color = "red";
          size="xx-small";
          delete row.RoomStatus;
          delete row.CleanStatus;
          for(key in row){
            var roomNo = row[key].split(" ")[0];
            var roomSId = document.getElementById("s_"+roomNo);
            if(roomSId != null){
              roomSId.innerHTML = roomStatus;
              roomSId.style.color=color;
              roomSId.style.fontSize=size;
              ++ooCnt;
            }
          }
        }
      });
    });
  };
  reader.readAsBinaryString(input.files[0]);
}

function reRoomCheck(){
  if(vCnt == 0){
    alert("Summary를 업로드 후 진행해주시지 바랍니다.");
    return;
  }else if(cCnt == 0){
    alert("Departure를 업로드 후 진행해주시지 바랍니다.");
    return;
  }
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  for(var i = 3 ; i <= 16 ; i++){
    if(i == 13){
      continue;
    }
    switch (i) {
      case 3 : 
        startRoomNo = 5;
        endRoomNo = 22;
        break;
      case 4 : 
        endRoomNo = 34;
        break;
      case 5 : 
        endRoomNo = 35;
        break;
      case 16 : 
        endRoomNo = 6;
        break;
    default:
      startRoomNo = 1;
      endRoomNo = 31;
      break;
    }
    if(i < 10){
      floorNo = "0" + i;
    }else{
      floorNo = i;
    }
    let roomStatus = "O";
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(j < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + "" + j;
      }
      var roomSId = document.getElementById("s_"+ roomNo);
      console.log(roomNo + ":" + roomSId.innerHTML);
      if(roomSId.innerHTML == "") {
        console.log(roomNo);
        roomSId.innerHTML = roomStatus;
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