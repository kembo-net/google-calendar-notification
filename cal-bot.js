var AliasList = {};
var RoomList = {};

var Sheets = SpreadsheetApp.getActive().getSheets();
var CalendarID = Sheets[0].getRange(1, 2).getValue();
var MyCalendar = CalendarApp.getCalendarById(CalendarID);
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

var debug = false;

Date.prototype.setNextDate = function() { return this.setDate(this.getDate() + 1); };

function setLog(room, text) {
  if (debug) {
    Logger.log(room);
    Logger.log(text);
    Logger.log('==============');
  }
  else {
    var sheet = SpreadsheetApp.openById(LogSpSheetID).getSheets()[0];
    sheet.appendRow([new Date, room, text]);
  }
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
  if (!debug) { UrlFetchApp.fetch(url, params); }
  setLog(room, text);
}

function genMessageFromEvent(event) {
  var message = '';
  var description = event.getDescription();
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
  return message;
}

function mainFunction() {
  var date = new Date();
  var flag = false;
  var events = MyCalendar.getEventsForDay(date);
  if (events.length == 0) {
    postIdobata('皆さまおはようございます。\n本日登録されている特別なイベントはございません。');
  }
  else {
    postIdobata('皆さまおはようございます。');
    events.forEach(function(event) {
      var message = genMessageFromEvent(event);
      var room = AliasList['DEFAULT'];
      if ( event.getDescription().match(/^r(?:oom)?: *([\w]+)/m) ) {
        room = RegExp.$1;
      }
      postIdobata(room, message);
      flag = flag || (room == AliasList['DEFAULT']);
    });
  }
  if (date.getDay() == 1) {
    var day_str = '日月火水木金土';
    var messages = [];
    do {
      date.setNextDate();
      events = MyCalendar.getEventsForDay(date);
      if (events.length > 0) {
        var message = day_str[date.getDay()] + '曜日には' + events.map(
          function(event) { return '『' + event.getTitle() + '』'; }
        ).join('') + 'の予定があります。';
        messages.push(message);
      }
    } while (date.getDay() != 0);
    if (messages.length == 0) {
      postIdobata('今週予定されている特別なイベントはございません。');
    }
    else {
      if (flag) {
        postIdobata('続いて今週の予定です。');
      }
      else {
        postIdobata('今週の予定です。');
      }
      messages.forEach(postIdobata);
    }
  }
}

function getEventsMonday() {
  var next_monday = new Date();
  while (next_monday.getDay() != 1) {
    next_monday.setNextDate();
  }
  return MyCalendar.getEventsForDay(next_monday);
}

function remindNextMonday() {
  var message = '皆さま今週もお疲れ様でした。\n来週月曜';
  var events = getEventsMonday().map(
    function(event) { return '『' + event.getTitle() + '』'; });
  if (events.length == 0) {
    message += 'に現在登録されている予定はございません。';
  }
  else {
    message += 'は' + events.join('') + 'の予定がございます。';
  }
  postIdobata(message);
}
