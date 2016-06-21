var AliasList = {};
var RoomList = {};

var Sheets = SpreadsheetApp.getActive().getSheets();
var CalendarID = Sheets[0].getRange(1, 2).getValue();
var LogSpSheetID = Sheets[0].getRange(2, 2).getValue();
for (var i = 1; Sheets[1].getRange(i, 1).getValue(); i++) {
  var key = Sheets[1].getRange(i, 1).getValue();
  var val = Sheets[1].getRange(i, 2).getValue();
  if (val.match(/^https?:\/\//)) {
    AliasList[key] = val;
  }
  else {
    AliasList[key] = val;
  }
}

function setLog(room, text) {
  var sheet = SpreadsheetApp.openById(LogSpSheetID).getSheets()[0];
  sheet.appendRow([new Date, room, text]);
}

function postIdobata(room, text) {
  if (!text) { 
    text = room;
    room = AliasList['DEFAULT'];
  }
  else {
    if ( room in AliasList ) { room = AliasList[room]; }
    if ( !(room in RoomList) ) { room = AliasList['DEFAULT']; }
  }
  var url = AliasList[room];
  var params = {
    payload: { source: text },
    method: "post"
  };
  UrlFetchApp.fetch(url, params);
  setLog(room, text);
}

function getEventsToday() {
  var calendar = CalendarApp.getCalendarById(CalendarID);
  var today = new Date();
  return calendar.getEventsForDay(today);
}

function mainFunction() {
  getEventsToday().forEach(function(event) {
    var message = '';
    var description = event.getDescription();
    var room = AliasList['DEFAULT'];
    if ( description.match(/^r(?:oom)?: *([\w]+)/m) ) {
        room = RegExp.$1;
    }
    if ( description.match(/^u(?:sers?)?: *(@?[\w]+(?:, *@?[\w]+)*)/m) ) {
      message += RegExp.$1.split(' ').join('').split(',').map(function (name) {
        if (name.match(/^@/)) { return name; }
        return '@' + name;
      }).join(' ') + ' ';
    }
    message += "本日は";
    var time = event.getStartTime();
    var h = time.getHours();
    if (h > 0) {
      var m = time.getMinutes();
      m = m < 10 ? '0' + m : m;
      message += "" + h + "時" + m + "分より";
    }
    var locate = event.getLocation();
    if (locate) {
      message += locate + "にて";
    }
    var title = event.getTitle();
    message += "『" + title + "』の予定があります。";
    if ( description.match(/^c(?:omment)?: *([^\n\r]*)/m) ) {
        message += '\n' + RegExp.$1;
    }
    Logger.log(room);
    Logger.log(message);
    Logger.log('==============');
    postIdobata(room, message);
  });
}
