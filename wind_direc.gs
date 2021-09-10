// Wind Direction Converter, wind_direc.gs
// Version 1.0
//


function allClear() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("data");

  //start row
  var sr = 6;

  sheet.getRange(sr, 1, 2000 + sr, 10).clear();
  sheet.getRange('F1').clearContent();
  sheet.getRange('F3').clearContent();

}


function windDirecCal() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("data");
  //a:16-dir Kanji
  //b:16-dir Alphabetical
  //c:Angle (deg.)

  var dtype = sheet.getRange('F1').getValue();
  if (dtype == '') {
    Browser.msgBox("風向元データの種類をプルダウンで選択してください。");
    return;
  }

  if (dtype == '漢字') {
    var atype = 1;
  }
  if (dtype == 'アルファベット') {
    var atype = 2;
  }
  if (dtype == '角度（度）') {
    var atype = 3;
  }

  //start row
  var sr = 6;

  for (var i = 1; i <= 3; i++) {
    if (i !== atype) {
      sheet.getRange(sr, i, 2000 + sr, 1).clear();
    }
  }

  if (atype == 1) {
    //read 1 a=1,b=2,c=3
    var arrin = sheet.getRange(sr, 1, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //1 to 2
    var brrouttrs = [];//1-dimensional
    brrouttrs = windDirec_12(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 2, arrintrs.length, 1).setValues(brrout);

    //2 to 3
    var crrouttrs = [];//1-dimensional
    crrouttrs = windDirec_23(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 3, brrouttrs.length, 1).setValues(crrout);
    //console.log(arrintrs);
  }

  if (atype == 2) {
    //read 2 a=2, b=3, c=1
    var arrin = sheet.getRange(sr, 2, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //2 to 3
    var brrouttrs = [];//1-dimensional
    brrouttrs = windDirec_23(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 3, arrintrs.length, 1).setValues(brrout);

    //3 to 1
    var crrouttrs = [];//1-dimensional
    crrouttrs = windDirec_31(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 1, brrouttrs.length, 1).setValues(crrout);
    //console.log(brrouttrs);
  }

  if (atype == 3) {
    //read 3 a=3, b=1, c=2
    var arrin = sheet.getRange(sr, 3, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //3 to 1
    var brrouttrs = [];//1-dimensional
    brrouttrs = windDirec_31(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 1, arrintrs.length, 1).setValues(brrout);

    //1 to 2
    var crrouttrs = [];//1-dimensional
    crrouttrs = windDirec_12(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 2, brrouttrs.length, 1).setValues(crrout);
    //console.log(crrout);
  }


}

function windSpeedCal() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("data");
  //a:m/s
  //b:knot
  //c:km/h

  var dtype = sheet.getRange('F3').getValue();
  if (dtype == '') {
    Browser.msgBox("風速元データの種類をプルダウンで選択してください。");
    return;
  }

  if (dtype == 'm/s') {
    var atype = 4;
  }
  if (dtype == 'knot') {
    var atype = 5;
  }
  if (dtype == 'km/h') {
    var atype = 6;
  }

  //start row
  var sr = 6;

  for (var i = 4; i <= 6; i++) {
    if (i !== atype) {
      sheet.getRange(sr, i, 2000 + sr, 1).clear();
    }
  }

  if (atype == 4) {
    //read 4 a=4,b=5,c=6
    var arrin = sheet.getRange(sr, 4, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //4 to 5
    var brrouttrs = [];//1-dimensional
    brrouttrs = windSpd_45(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 5, arrintrs.length, 1).setValues(brrout);

    //5 to 6
    var crrouttrs = [];//1-dimensional
    crrouttrs = windSpd_56(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 6, brrouttrs.length, 1).setValues(crrout);
    //console.log(arrintrs);
  }

  if (atype == 5) {
    //read 5 a=5, b=6, c=4
    var arrin = sheet.getRange(sr, 5, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //5 to 6
    var brrouttrs = [];//1-dimensional
    brrouttrs = windSpd_56(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 6, arrintrs.length, 1).setValues(brrout);

    //6 to 4
    var crrouttrs = [];//1-dimensional
    crrouttrs = windSpd_64(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 4, brrouttrs.length, 1).setValues(crrout);
    //console.log(brrouttrs);
  }

  if (atype == 6) {
    //read 6 a=6, b=4, c=5
    var arrin = sheet.getRange(sr, 6, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
    var _ = Underscore.load();
    var arrintr = _.zip.apply(_, arrin);//transpose
    var arrintrs = arrintr[0];//1-dimensional

    //6 to 4
    var brrouttrs = [];//1-dimensional
    brrouttrs = windSpd_64(arrintrs);//function //1-dimensional
    var brrouttr = [brrouttrs];//2-dimensional
    var brrout = _.zip.apply(_, brrouttr);//transpose
    sheet.getRange(sr, 4, arrintrs.length, 1).setValues(brrout);

    //4 to 5
    var crrouttrs = [];//1-dimensional
    crrouttrs = windSpd_45(brrouttrs);//function //1-dimensional
    var crrouttr = [crrouttrs];//2-dimensional
    var crrout = _.zip.apply(_, crrouttr);//transpose
    sheet.getRange(sr, 5, brrouttrs.length, 1).setValues(crrout);
    //console.log(crrout);
  }


}

function windVector() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("data");

  //start row
  var sr = 6;
  var _ = Underscore.load();

  sheet.getRange(sr, 8, 2000 + sr, 3).clear();

  //read wind direction
  var arrin = sheet.getRange(sr, 3, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
  var arrintr = _.zip.apply(_, arrin);//transpose
  var wd = arrintr[0];//1-dimensional

  //wind motion direction
  var uvd = [];
  for (let i = 0; i < wd.length; i++) {
    if (wd[i] == '---') {
      uvd[i] = '---';
    } else {
      uvd[i] = (wd[i] + 180) % 360;
    }
  }

  //read wind speed
  var arrin = sheet.getRange(sr, 4, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
  var arrintr = _.zip.apply(_, arrin);//transpose
  var ws = arrintr[0];//1-dimensional

  // wind vector
  var u = [];
  var v = [];

  for (let i = 0; i < ws.length; i++) {
    if(uvd[i] == '---'){
      u[i] = 0;
      v[i] = 0;
    }else{
      u[i] = ws[i] * Math.sin(uvd[i] / 180 * Math.PI);
      v[i] = ws[i] * Math.cos(uvd[i] / 180 * Math.PI);
    }
  }

  // output uvd,u,v
  var arrouttr = [uvd, u, v];//2-dimensional
  var arrout = _.zip.apply(_, arrouttr);//transpose
  sheet.getRange(sr, 8, uvd.length, 3).setValues(arrout);

}

function windVectorToDS() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("data");

  //start row
  var sr = 6;
  var _ = Underscore.load();

  sheet.getRange(sr, 1, 2000 + sr, 8).clear();

  //read wind u,v
  var arrin = sheet.getRange(sr, 9, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
  var arrintr = _.zip.apply(_, arrin);//transpose
  var u = arrintr[0];//1-dimensional

  var arrin = sheet.getRange(sr, 10, sheet.getLastRow() - sr + 1).getValues();//2-dimensional
  var arrintr = _.zip.apply(_, arrin);//transpose
  var v = arrintr[0];//1-dimensional

  //wind motion direction uvd
  var uvd = [];
  for (let i = 0; i < u.length; i++) {
    if (v[i] > 0) {
      uvd[i] = (Math.atan(u[i] / v[i]) / Math.PI * 180 + 360) % 360;
    } else {
      if( u[i] == 0 && v[i] == 0){
        uvd[i] = '---';
      }else{
        uvd[i] = (Math.atan(u[i] / v[i]) / Math.PI * 180 + 180 + 360) % 360;
      }
    }
  }

  //wind direction wd
  var wd = [];
  for (let i = 0; i < uvd.length; i++) {
    if(uvd[i] == '---'){
      wd[i] = '---';
    }else{
      wd[i] = (uvd[i] + 180) % 360;
    }
  }

  //read wind speed
  var ws = [];
  for (let i = 0; i < u.length; i++) {
    ws[i] = Math.sqrt(u[i] ** 2 + v[i] ** 2);
  }


  // output uvd
  var arrouttr = [uvd];//2-dimensional
  var arrout = _.zip.apply(_, arrouttr);//transpose
  sheet.getRange(sr, 8, uvd.length, 1).setValues(arrout);
  // output wd
  var arrouttr = [wd];//2-dimensional
  var arrout = _.zip.apply(_, arrouttr);//transpose
  sheet.getRange(sr, 3, wd.length, 1).setValues(arrout);
  // output ws
  var arrouttr = [ws];//2-dimensional
  var arrout = _.zip.apply(_, arrouttr);//transpose
  sheet.getRange(sr, 4, ws.length, 1).setValues(arrout);

  //3 to 1
  var brrouttrs = [];//1-dimensional
  brrouttrs = windDirec_31(wd);//function //1-dimensional
  var brrouttr = [brrouttrs];//2-dimensional
  var brrout = _.zip.apply(_, brrouttr);//transpose
  sheet.getRange(sr, 1, wd.length, 1).setValues(brrout);

  //1 to 2
  var crrouttrs = [];//1-dimensional
  crrouttrs = windDirec_12(brrouttrs);//function //1-dimensional
  var crrouttr = [crrouttrs];//2-dimensional
  var crrout = _.zip.apply(_, crrouttr);//transpose
  sheet.getRange(sr, 2, brrouttrs.length, 1).setValues(crrout);
  //console.log(crrout);

  //4 to 5
  var brrouttrs = [];//1-dimensional
  brrouttrs = windSpd_45(ws);//function //1-dimensional
  var brrouttr = [brrouttrs];//2-dimensional
  var brrout = _.zip.apply(_, brrouttr);//transpose
  sheet.getRange(sr, 5, ws.length, 1).setValues(brrout);

  //5 to 6
  var crrouttrs = [];//1-dimensional
  crrouttrs = windSpd_56(brrouttrs);//function //1-dimensional
  var crrouttr = [crrouttrs];//2-dimensional
  var crrout = _.zip.apply(_, crrouttr);//transpose
  sheet.getRange(sr, 6, brrouttrs.length, 1).setValues(crrout);

}


function windDirec_12(arr) {
  var warrd = ['北', '北北東', '北東', '東北東', '東', '東南東', '南東', '南南東', '南', '南南西', '南西', '西南西', '西', '西北西', '北西', '北北西'];
  var warrn = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW'];
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    var d = warrd.indexOf(arr[i]);
    if (d === -1) {
      arrout[i] = '---';
    } else {
      arrout[i] = warrn[d];
    }
  }

  return arrout;
}


function windDirec_23(arr) {
  var warrd = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW'];
  var warrn = [0, 22.5, 45, 67.5, 90, 112.5, 135, 157.5, 180, 202.5, 225, 247.5, 270, 292.5, 315, 337.5];
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    var d = warrd.indexOf(arr[i]);
    if (d === -1) {
      arrout[i] = '---';
    } else {
      arrout[i] = warrn[d];
    }
  }

  return arrout;
}

function windDirec_31(arr) {
  var warrd = [0, 11.25, 33.75, 56.25, 78.75, 101.25, 123.75, 146.25, 168.75, 191.25, 213.75, 236.25, 258.75, 281.25, 303.75, 326.25, 348.75, 360];//n=18
  var warrn = ['北', '北北東', '北東', '東北東', '東', '東南東', '南東', '南南東', '南', '南南西', '南西', '西南西', '西', '西北西', '北西', '北北西', '北'];//n=17
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] == '---') {
      arrout[i] = '---';
    } else {
      var a = arr[i] % 360;
      if (a < 0) {
        var a = a + 360;
      }
      //if (arr[i] == '') {
      //  arrout[i] = '';
      //} else {
      for (let j = 0; j < 17; j++) {
        if (a >= warrd[j] && a < warrd[j + 1]) {
          arrout[i] = warrn[j];
          break;
        }
      }
    }
  }

  return arrout;
}

function windSpd_45(arr) {
  // m/s to knot
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    arrout[i] = arr[i] / 0.5144;
  }

  return arrout;
}

function windSpd_56(arr) {
  // knot to km/h
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    arrout[i] = arr[i] * 1.852;
  }

  return arrout;
}

function windSpd_64(arr) {
  // km/h to m/s
  var arrout = [];
  for (let i = 0; i < arr.length; i++) {
    arrout[i] = arr[i] * 1000 / 3600;
  }

  return arrout;
}

