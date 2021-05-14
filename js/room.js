let fileCnt = 0;
let notCleaning = 0;
let vCnt = 0;
let cCnt = 0;
let oCnt = 0;
let ooCnt = 0;
let totalCnt = 0;

const a_type = ["3A"
              , "4A"
              , "5A"
              , "6A"
              , "7A"
              , "8A"
              , "9A"
              , "10A"
              , "11A"
              , "12A"
              , "14A"
              , "15A"
              , "16A"]
let a_staff = []

const b_type = ["4B"
              , "5B"
              , "6B"
              , "7B"
              , "8B"
              , "9B"
              , "10B"
              , "11B"
              , "12B"
              , "14B"
              , "15B"]
let b_staff = []

let roomTypeList = ["V"
                  , "C"
                  , "O"
                  , "O.O"
]


window.addEventListener("load", function(event) {
  let current = new Date();
  let year = current.getFullYear();
  let month = current.getMonth()+1;
  let day = current.getDate();
  let currentDt = year + "년 " + month + "월 " + day + "일";
  document.getElementById("currentDt").innerHTML = currentDt;

  let roomTypeAll = document.getElementsByClassName("roomType");
  for(var i = 0 ; i < roomTypeAll.length ; i++){
    roomTypeAll[i].addEventListener("click", fn_change_room_type, false);
  }
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
        //if(testDt == row.OrgDepDate){
          var roomSId = document.getElementById("s_"+row.RmNo);
          if(roomSId != null){
            roomSId.innerHTML = "C";
            roomSId.style.color = "red";
            ++cCnt;
          }
        }
      //}
      );
      fn_totalCnt();
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
        let size = "medium";
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
          size="small";
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
        fn_totalCnt();
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
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
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
      if(roomSId.innerHTML == "") {
        roomSId.innerHTML = roomStatus;
        ++oCnt;
      }
    }
  }
  fn_totalCnt();
}

function fn_totalCnt(){
  totalCnt = vCnt + cCnt + oCnt + ooCnt;
  document.getElementById("total_v").innerHTML = vCnt;
  document.getElementById("total_c").innerHTML = cCnt;
  document.getElementById("total_o").innerHTML = oCnt;
  document.getElementById("total_oo").innerHTML = ooCnt;
  document.getElementById("total").innerHTML = totalCnt;
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
  printSetData();
  
  var initBody = document.body.innerHTML;
  window.onbeforeprint = function(){
    document.body.innerHTML = document.getElementById('content').innerHTML;
  }
  window.onafterprint = function(){
    document.body.innerHTML = initBody;
  }
  window.print();
  totalRoomChk("update");
  fn_floor_staff_re_set();
}

function fn_floor_staff(){
  for(let i in a_type){
    var aStaff = document.getElementById(a_type[i]).value;
    if(aStaff != undefined){
      var aRStaff = document.getElementsByClassName("staff"+a_type[i]);
      let aRoomNoClass = "floor" + a_type[i];
      let aRoomNo = document.getElementsByClassName(aRoomNoClass);
      for(var j = 0 ; j < aRStaff.length ; j++){
        if(aStaff == ""){
          aRoomNo[j].style.backgroundColor = "#fbffad";
        }else{
          aRoomNo[j].style.backgroundColor = "";
          aRStaff[j].value = aStaff;
        }
      }
      a_staff[i] = aStaff;
    }
  }
  for(i in b_type){
    var bStaff = document.getElementById(b_type[i]).value;
    if(bStaff != undefined){
      var bRStaff = document.getElementsByClassName("staff"+b_type[i]);
      let bRoomNoClass = "floor" + b_type[i];
      let bRoomNo = document.getElementsByClassName(bRoomNoClass);
      for(var j = 0 ; j < bRStaff.length ; j++){
        if(bStaff == ""){
          bRoomNo[j].style.backgroundColor = "#fbffad";
        }else{
          bRoomNo[j].style.backgroundColor = "";
          bRStaff[j].value = bStaff;
        }
      }
      b_staff[i] = bStaff;
    }
  }
  if(totalCnt == 372){
    fn_notCleaning();
  }
}

function fn_floor_staff_re_set(){
  for(var i = 0 ; i < 13 ; i++){
    if(a_staff[i] != null && a_staff[i] != undefined){
      document.getElementById(a_type[i]).value = a_staff[i];
    }
  }
  for(var i = 0 ; i < 11 ; i++){
    if(b_staff[i] != null && b_staff[i] != undefined){
      document.getElementById(b_type[i]).value = b_staff[i];
    }
  }
}

function printSetData(){
  totalRoomChk("print");
}

function totalRoomChk(type){
  let startRoomNo = 1;
  let endRoomNo = 31;
  let floorNo = "";
  let roomNo = "";
  let input_html = "";
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
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
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
    for(var j = startRoomNo ; j <= endRoomNo ; j++){
      if(j < 10){
        roomNo = floorNo + "0" + j;
      }else{
        roomNo = floorNo + "" + j;
      }

      var roomSId_input = document.getElementById("u_"+ roomNo).firstElementChild;
      var roomSId = document.getElementById("u_"+ roomNo);
      if(type == "update" && roomSId != null && roomSId != "") {
        input_html = fn_set_input(i, j);
        
        var roomSId_text = roomSId.innerHTML;
        roomSId.innerHTML = input_html;
        roomSId.firstElementChild.value = roomSId_text;
      }else if(type == "print" && roomSId_input != null && roomSId_input != "") {
        roomSId_input.parentElement.innerHTML = roomSId_input.value;
      }
    }
  }
}

function fn_set_input(i, j){
  switch (i) {
    case 3 : 
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      break;
    case 4 : 
      if((j > 5 && j < 11) || (j > 21 && j < 28)){
        input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      }else if(j > 10 && j < 22){
        input_html = '<input type="text" class="staff'+i+'B" value=""/>'
      }else{
        input_html = '<input type="text" value=""/>';
      }
      break;
    case 5 : 
      if((j > 5 && j < 11) || (j > 22 && j < 29)){
        input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      }else if(j > 10 && j < 23){
        input_html = '<input type="text" class="staff'+i+'B" value=""/>'
      }else{
        input_html = '<input type="text" value=""/>';
      }
      break;
    case 16 : 
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
      break;
  default:
    if((j > 5 && j < 11) || (j > 21 && j < 27)){
      input_html = '<input type="text" class="staff'+i+'A" value=""/>'
    }else if(j > 10 && j < 22){
      input_html = '<input type="text" class="staff'+i+'B" value=""/>'
    }else{
      input_html = '<input type="text" value=""/>';
    }
    break;
  }
  return input_html;
}

function fn_change_room_type(){
  let roomType = this.innerHTML;
  let nextRoomType;
  switch (roomType) {
    case roomTypeList[0] : 
      this.style.color = "red";
      nextRoomType = roomTypeList[1];
      vCnt--;
      cCnt++;
      break;
    case roomTypeList[1] : 
      this.style.color = "black";
      nextRoomType = roomTypeList[2];
      cCnt--;
      oCnt++;
      break;
    case roomTypeList[2] :
      this.style.color = "red";
      this.style.fontSize = "small";
      nextRoomType = roomTypeList[3];
      oCnt--;
      ooCnt++;
      break;
    case roomTypeList[3] : 
      this.style.color = "black";
      this.style.fontSize = "medium";
      nextRoomType = roomTypeList[0];
      ooCnt--;
      vCnt++;
      break;
    case "" : 
      this.style.color = "black";
      nextRoomType = roomTypeList[0];
      vCnt++;
      break;
  }
  this.innerHTML = nextRoomType;
  fn_totalCnt();
  if(totalCnt == 372){
    fn_notCleaning();
  }
}

function fn_notCleaning(){
  notCleaning = 0;
  if(totalCnt != 372){
    alert("Occupied 처리를 해주시기 바랍니다.");
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
        startRoomNo = 1;
        endRoomNo = 34;
        break;
      case 5 : 
        startRoomNo = 1;
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
      var roomSId = document.getElementById("s_" + roomNo);
      var staffNm = document.getElementById("u_" + roomNo).firstElementChild.value;
      console.log(roomNo);
      console.log(document.getElementById("u_" + roomNo).firstElementChild.innerHTML);
      if((roomSId.innerHTML == "C" || roomSId.innerHTML == "O") &&
          staffNm == "") {
        roomSId.style.backgroundColor = "#fca8ff";
        notCleaning++;
      }else{
        roomSId.style.backgroundColor = "";
      }
    }
  }
  document.getElementById("notCleaningCnt").innerHTML = notCleaning;
}