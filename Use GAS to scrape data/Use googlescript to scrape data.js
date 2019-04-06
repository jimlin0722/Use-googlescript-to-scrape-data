
/* select 'Tools' -> 'Macros' -> 'Import' to add the function you want to the sheet  */

function createTriggers() {
  var weekday = new Date().getDay()
  if (weekday != 0 && weekday != 6){
    var trigger_update = ScriptApp.newTrigger('update')
        .timeBased()
        .everyMinutes(1)
        .create();
    var triggerID = trigger_update.getUniqueId()
    var url = 'YOUR_SPREADSHEET_URL'
    var name = 'YOUR_SHEET_NAME'
    var SpreadSheet = SpreadsheetApp.openByUrl(url);
    var SheetName = SpreadSheet.getSheetByName(name)
    SheetName.getRange(2,2,1,1).setValue(triggerID)
  }
}


function deleteTrigger() {
  var weekday = new Date().getDay()
  if (weekday != 0 && weekday != 6){
    var url = 'YOUR_SPREADSHEET_URL'
    var name = 'YOUR_SHEET_NAME'
    var SpreadSheet = SpreadsheetApp.openByUrl(url)
    var SheetName = SpreadSheet.getSheetByName(name)
    var allTriggers = ScriptApp.getProjectTriggers()
    for (var i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getUniqueId() == SheetName.getSheetValues(2,2,1,1)) {
        ScriptApp.deleteTrigger(allTriggers[i])
        break;
      }
    }
    SheetName.getRange(2,2,1,1).setValue("");
  }
}



function update(){
  var url = 'YOUR_SPREADSHEET_URL'
  var name = 'YOUR_SHEET_NAME'
  var SpreadSheet = SpreadsheetApp.openByUrl(url)
  var SheetName = SpreadSheet.getSheetByName(name)
  var lastRow = SheetName.getLastRow()
  var nextRow = parseInt(lastRow)+1
  var stock = SheetName.getSheetValues(2,1,1,1)
  var data = send_request(stock)
  var now = new Date()
  next_A = SheetName.getRange("A" + nextRow.toString())
  next_A.setValue(convert(now/1000))
  next_B = SheetName.getRange("B" + nextRow.toString())
  next_B.setValue(data.currentprice)
  next_C = SheetName.getRange("C" + nextRow.toString())
  if (now.getMinutes() == 0 && now.getHours() == 9){
    next_C.setValue(data.totalvol)
  }
  else if (now.getHours() < 9 || now.getHours() > 13){
    next_C.setValue(0)
  }
  else{
    next_C.setValue(data.totalvol-SheetName.getSheetValues(lastRow,4,1,1))
  }
  next_D = SheetName.getRange("D" + nextRow.toString())
  next_D.setValue(data.totalvol)
  next_E = SheetName.getRange("E" + nextRow.toString())
  next_E.setValue(data.best5saleprice)
  next_F = SheetName.getRange("F" + nextRow.toString())
  next_F.setValue(data.best5salevol)
  next_G = SheetName.getRange("G" + nextRow.toString())
  next_G.setValue(data.best5buyprice)
  next_H = SheetName.getRange("H" + nextRow.toString())
  next_H.setValue(data.best5buyvol)
  SheetName.insertRowAfter(nextRow)
}

function send_request(code) {
  var options_step1 = {
    "medthod": "get",
    "headers": {
    "User-Agent": "Mozilla/5.0"
     },
  }
  var re_step1 = UrlFetchApp.fetch("http://mis.twse.com.tw/stock/fibest.jsp?stock"+ code, options_step1)
  var cookie = re_step1.getHeaders()
  var headers = {
    "Cookie" : cookie,
    "User-Agent": "Mozilla/5.0"
  }
  var options_step2 = {
    "method" : "get",
    "headers" : headers
  }
  var now = Date.now()
  var url = "http://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_" + code + ".tw&json=1&delay=0&_=" + now.toString()
  var re_step2 = UrlFetchApp.fetch(url, options_step2)
  var data = JSON.parse(re_step2.getContentText())
  var stockinfo = new FORMAT()
      stockinfo.currentprice   = data.msgArray[0].z
      stockinfo.tempvol = data.msgArray[0].tv
      stockinfo.totalvol = data.msgArray[0].v
      stockinfo.best5saleprice = data.msgArray[0].a
      stockinfo.best5salevol = data.msgArray[0].f
      stockinfo.best5buyprice = data.msgArray[0].b
      stockinfo.best5buyvol = data.msgArray[0].g
      stockinfo.crawltime = data.msgArray[0].t
      stockinfo.datatime = data.msgArray[0].tlong
  return stockinfo
}

function FORMAT() {
  this.currentprice = 0.0
  this.tempvol = 0
  this.totalvol = 0
  this.best5saleprice = []
  this.best5salevol = []
  this.best5buyprice = []
  this.best5buyvol = []
  this.crawltime = 0
  this.datatime = 0
}

function convert(timestamp){

 var unixtimestamp = timestamp
 var date = new Date(unixtimestamp*1000)
 var year = date.getFullYear()
 var month = date.getMonth()+1
 var day = date.getDate()
 var hours = date.getHours()
 var minutes = "0" + date.getMinutes()
 var seconds = "0" + date.getSeconds()
  // return date time in y/m/d h:m:s format
 var convdataTime = year+'/'+month+'/'+day+' '+hours + ':' + minutes.substr(-2) + ':' + seconds.substr(-2)

 return convdataTime;

}

function clean_table(){

  var url = 'YOUR_SPREADSHEET_URL'
  var name = 'YOUR_SHEET_NAME'
  var SpreadSheet = SpreadsheetApp.openByUrl(url)
  var SheetName = SpreadSheet.getSheetByName(name)
  var lastRow = parseInt(SheetName.getLastRow());
  var nextRow = lastRow+1
  var out_tradetime = []
  for (var i =lastRow-390 ; i < nextRow ; i++){
    var cell_value = SheetName.getSheetValues(i, 1, 1, 1)
    var hours = new Date(cell_value).getHours()
    var minutes = new Date(cell_value).getMinutes()
    var weekday = new Date(cell_value).getDay()
    if (hours < 9 || hours > 13){
      out_tradetime.push(i)
    }
    else if (hours == 13 && minutes > 31){
      out_tradetime.push(i)
    }
    else if (weekday == 6 || weekday == 0){
      out_tradetime.push(i)
    }
  }
  for (var i = 0 ; i < out_tradetime.length; i ++){
     SheetName.deleteRows(out_tradetime[i]-i, 1)
  }
}
