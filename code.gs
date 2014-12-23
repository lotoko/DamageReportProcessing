// 位置情報
function LatLng(){
  this.lat;
  this.lng;
  return this;
}
// ポータル情報
function Portal(){
  this.name;
  this.owner;
  this.adress;
  this.LatLng = LatLng();
  return this;
}
Portal.prototype = {
  GetBasicInfos : function( wordsArray, pos){
    this.name = wordsArray[pos++];
    var loc = wordsArray[pos++].match(/https\:\/\/www\.ingress\.com\/intel\?ll=([0-9]+\.[0-9]+),([0-9]+\.[0-9]+)\&/);
    this.LatLng.lat = loc[1]
    this.LatLng.lng = loc[2];
    this.adress = wordsArray[pos++];
    return pos;
  },
  GetIntelURL : function(){
    return "https://www.ingress.com/intel?ll=" + this.LatLng.lat + "," + this.LatLng.lng + "&z=19";
  }
};
// エージェント情報
function Agent(){
  this.name;
  this.faction;
  this.level;
  return this;
}
// ダメージ情報
function DamageInfo(){
  this.sourceAG = Agent();
  this.attackerName;
  this.date;
  this.time;
  this.portal = new Portal();
  this.type;
  this.damage;
  this.reso;
  this.helth;
  this.status;
  this.affectedPortals = [];
  return this;
}
// 
function ExtractDamageInfos( date, body, infoArray){
  const timeDiff = 9;
  const keywords_body01 = ["Agent Name:", "Faction:", "Current Level:", "DAMAGE REPORT"];
  // IntelMapのURL抜き出し
  body = body.replace(/<a href="https:\/\/www.ingress.com\/intel\?\l\l=([0-9\.&am;plz=,]+)"("[^"]*"|'[^']*'|[^'">])*>/g,"|https://www.ingress.com/intel?ll="+"$1"+"|");
  // htmlのタグ削除
  body = body.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g,"|");
  // 区切りを|に統一
  body = body.replace(/[\n\r]/g,"|");
  body = body.replace(/[|]+/g,"|");
  body = body.replace(/[\s]+/g," ");
  var bodyArray = body.split("|");
  bodyArray.splice(0, 1);
  bodyArray.splice(bodyArray.length-1, 1);
  var info = new DamageInfo();
  for( var i = 0, j = 0; i < keywords_body01.length; i++, j+=2){
    if( keywords_body01[i] =! bodyArray[j] ){
      Browser.msgBox("ERROR!!@68");//ERROR
    }
  }
  info.sourceAG.name = bodyArray[1];
  info.sourceAG.faction = bodyArray[3];
  info.sourceAG.level = bodyArray[5];
  info.date = date;
  var p = 7;
  do{
    p = info.portal.GetBasicInfos(bodyArray, p);
    if( bodyArray[p] == "LINKS DESTROYED" || bodyArray[p] == "LINK DESTROYED" ){
      p++;
      var count = 1;
      while( bodyArray[p] != "DAMAGE:" ){
        if(count > 8){
          Browser.msgBox("ERROR!!@83");//ERROR
          break;
        }
        info.affectedPortals.push(new Portal());
        p = info.affectedPortals[info.affectedPortals.length-1].GetBasicInfos(bodyArray, p);
        count ++;
      }
    }
    if( bodyArray[p] == "DAMAGE:"){
      p++;
     do{
        var array = bodyArray[p].match(/([0-9]+) (Resonator[s]*|Link[s]*|Mod[s]*) destroyed by/);
        info.damage = array[1];
        info.type = array[2];
        p++;
        info.attackerName = bodyArray[p++];
        var time = bodyArray[p++].match(/at ([0-9]+):([0-9]+) hrs GMT/);
        info.time = info.date;
        info.time.setHours( time[1] + timeDiff );
        info.time.setMinutes(time[2]);
        info.time.setSeconds(0);
      }while( bodyArray[p].indexOf("destroyed by") != -1 );
      //if( bodyArray[p].match(/No remaining Resonators detected on this Portal./)){
      //  info.reso = 0;
      //}
      if( bodyArray[p].indexOf("No remaining Resonators detected") == -1){
        info.reso = bodyArray[p++].match(/([0-9]+) Resonator[s]* remaining on this Portal./)[1]
      }else{
        info.reso = 0;
        p++;
      }
      p++;
      info.status = bodyArray[p++];
      info.health = bodyArray[p++].match(/([0-9]+)/)[1];
      if( bodyArray[p++].indexOf("[uncaptured]") == -1 ){
        info.portal.owner = bodyArray[p++];
      }else{
        info.portal.owner = "[uncaptured]";
      }
    }else{
      Browser.msgBox("ERROR!!@123");//ERROR
    }
    infoArray.push(info);
  }while( p < bodyArray.length);
}
function ExtractDamageInfosFromTrheads( threads, damageThreads){
  for (var i = 0; i < threads.length; i++) {
    var nMessage = threads[i].getMessageCount();
    var fs_subject = threads[i].getFirstMessageSubject();
    var mArray = threads[i].getMessages();
    var damageMessages = [];
    for(var j = 0; j< mArray.length; j++){
      var date = mArray[j].getDate();
      var subject = mArray[j].getSubject();
      var body = mArray[j].getBody();
      ExtractDamageInfos( date, body, damageMessages);
    }
    damageThreads.push(damageMessages);
  }
}
function TEST(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("シート1");
  var label = GmailApp.getUserLabelByName("Ingress/preDamageReport");
  var threads = label.getThreads(0, 50);
  var damageThreads = [];
  ExtractDamageInfosFromTrheads( threads, damageThreads)
  //-------------------
  var countRow = 1;
  var countCol = 1;
  const Items = ["Date", "Attacker", "PortalName", "Owner", "Lat", "Lng", "Type", "Damage", "Reso", "Helth", "Status"];
  const ItemsDict = {"Date":0, "Attacker":1, "PortalName":2, "Owner":3, "Lat":4, "Lng":5, "Type":6, "Damage":7, "Reso":8, "Helth":9, "Status":10};
  var data = sheet.getDataRange().getValues();
  countRow = data.length + 1;
  if( countRow == 1){
    for( var i = 0; i < Items.length; i++){
      sheet.getRange(countRow, countCol+i).setValue(Items[i]);
    }
    countRow += 1;
  }
  for( var i = 0; i < damageThreads.length; i++){
    for( var j = 0; j < damageThreads[i].length; j++){
      var dinfo = damageThreads[i][j];
      sheet.getRange(countRow, countCol+ItemsDict["Date"]).setValue(dinfo.time);
      sheet.getRange(countRow, countCol+ItemsDict["Attacker"]).setValue(dinfo.attackerName);
      sheet.getRange(countRow, countCol+ItemsDict["PortalName"]).setValue(dinfo.portal.name);
      sheet.getRange(countRow, countCol+ItemsDict["Owner"]).setValue(dinfo.portal.owner);
      sheet.getRange(countRow, countCol+ItemsDict["Lat"]).setValue(dinfo.portal.LatLng.lat);
      sheet.getRange(countRow, countCol+ItemsDict["Lng"]).setValue(dinfo.portal.LatLng.lng);
      sheet.getRange(countRow, countCol+ItemsDict["Type"]).setValue(dinfo.type);
      sheet.getRange(countRow, countCol+ItemsDict["Damage"]).setValue(dinfo.damage);
      sheet.getRange(countRow, countCol+ItemsDict["Reso"]).setValue(dinfo.reso);
      sheet.getRange(countRow, countCol+ItemsDict["Helth"]).setValue(dinfo.health);
      sheet.getRange(countRow, countCol+ItemsDict["Status"]).setValue(dinfo.status);
      countRow++;
    }
  }
  //sheet.getRange(1, 1).setValue(n);
}
