// ID of the settings document in JSON format 
const SETTINGS_ID = "1kBitZVMZ6BGRKXUcdQ6WQX5nwywLTSpr";

const BOLD_STYLE = SpreadsheetApp.newTextStyle().setBold(true).build();
const NO_BOLD_STYLE = SpreadsheetApp.newTextStyle().setBold(false).build();
const NO_UNDERLINE_STYLE = SpreadsheetApp.newTextStyle().setUnderline(false).build();
const BLACK_STYLE = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').build();
const RED_STYLE = SpreadsheetApp.newTextStyle().setForegroundColor('#FF0000').build();
const PAUSE = 500;
const MAX_COUNT_RETRY = 10;
const TIMEZONE_S = 'UTC+3';

class Member {
  constructor(login, showed_name) {
    this.login = login;
    this.showed_name = showed_name;
    this.rating = 0;
    this.rating_before_deadline = 0;
    this.cdf_rating = undefined;
    this.total_groups = new Map();
    this.total_groups_before_deadline = new Map();
    this.subs = [];
    this.updateAttempts();
  }

  updateAttempts() {
    this.allAttempts = 0;
    this.successfulAttempts = 0;
    this.successfulBeforeDeadlineAttempts = 0;
    this.subs.forEach(function (sub) {
      this.allAttempts++;
      if (sub.status() == 1) {
        this.successfulAttempts++;
        if (sub.ok_before_dealine) {
          this.successfulBeforeDeadlineAttempts++;
        }
      }
    }, this);
  }

  addSubmission(sub) {
    if (sub.verdict == 'OK') {
      this.rating += sub.task.rating;
      if (!this.total_groups.has(sub.task.group)) {
        this.total_groups.set(sub.task.group, 0);
      }
      this.total_groups.set(sub.task.group, this.total_groups.get(sub.task.group) + sub.task.rating);
      if (sub.ok_before_dealine) {
        this.rating_before_deadline += sub.task.rating;
        if (!this.total_groups_before_deadline.has(sub.task.group)) {
          this.total_groups_before_deadline.set(sub.task.group, 0);
        }
        this.total_groups_before_deadline.set(sub.task.group, this.total_groups_before_deadline.get(sub.task.group) + sub.task.rating);
      }
    }
  }

  toString() {
    this.updateAttempts();
    var text = '';
    if (this.showed_name) {
      text = this.showed_name + '\n';
    }
    text += `(${this.login})`;
    text += `\n${this.successfulBeforeDeadlineAttempts} | ${this.successfulAttempts} | ${this.allAttempts}`
    return text;
  }

  color() {
    if (this.cdf_rating === undefined)
      return '#000000';
    if (this.cdf_rating < 1200)
      return '#808080';
    if (this.cdf_rating < 1400)
      return '#00FF00';
    if (this.cdf_rating < 1600)
      return '#03A89E';
    if (this.cdf_rating < 1900)
      return '#0000FF';
    if (this.cdf_rating < 2100)
      return '#AA00AA';
    if (this.cdf_rating < 2400)
      return '#FF8C00';
    return '#FF0000';
  }
}

class GroupTask {
  constructor(groop_name, show, row) {
    this.name = groop_name;
    this.show = show;
    this.row = row;
    this.total_calc = false;
    this.total_calc_before_deadline = false;
    this.show_langs = false;
    this.show_date = false;
    this.show_time = false;
    this.show_params = false;
    this.last_or_first = 'last';
    this.deadline = undefined;
  }

  toString() {
    var text = this.name;
    if (this.deadline !== undefined) {
      text += `\nДедлайн: ${this.deadline.toLocaleString("ru-RU")} ${TIMEZONE_S}`;
    }
    return text;
  }
}

class Task {
  constructor(contestId, index, row, group) {
    this.contestId = contestId;
    this.index = index;
    this.row = row;
    this.name = null;
    this.rating = null;
    this.points = null;
    this.group = group;
    this.allAttempts = new Set();
    this.successfulAttempts = new Set();
    this.successfulBeforeDeadlineAttempts = new Set();
  }

  toString() {
    var text = '' + this.contestId + this.index;
    if (this.rating !== null) {
      text += ' (' + this.rating + ')';
    }
    if (this.name !== null) {
      text += '\n' + this.name;
    }
    text += `\n${this.successfulBeforeDeadlineAttempts.size} | ${this.successfulAttempts.size} | ${this.allAttempts.size}`;
    return text;
  }
}

class Sumbission {
  constructor(task, id, verdict, member, creationTimeSeconds) {
    this.task = task;
    this.programmingLanguages = new Set();
    this._verdict = verdict;
    this.passedTestCount = 0;
    this.timeConsumedMillis = 0;
    this.memoryConsumedBytes = 0;
    this.creationTimeMillis = creationTimeSeconds * 1000;
    this.creationTime = new Date(this.creationTimeMillis);
    this.id = id;
    this.number_of_failures = 0;
    this.ok_before_dealine = this.status() == 1 && ((this.task.group.deadline === undefined) || (this.task.group.deadline >= this.creationTime));

    if (this.status() != 1) {
      this.number_of_failures = 1;
    } else {
      this.task.successfulAttempts.add(member.login);
      if (this.ok_before_dealine) {
        this.task.successfulBeforeDeadlineAttempts.add(member.login);
      }
    }
    this.task.allAttempts.add(member.login);
  }

  langs() {
    return Array.from(this.programmingLanguages).join(', ');
  }

  date() {
    return this.creationTime.toLocaleDateString("ru-RU")
  }

  time() {
    return this.creationTime.toLocaleTimeString("ru-RU")
  }

  timeConsumed() {
    return this.timeConsumedMillis + ' мс';
  }

  memoryConsumed() {
    if (this.memoryConsumedBytes < 1024) {
      return this.memoryConsumedBytes + ' б';
    }
    this.memoryConsumedBytes /= 1024.0;
    if (this.memoryConsumedBytes < 1024) {
      return Math.round(this.memoryConsumedBytes * 10) / 10 + ' Кб';
    }
    this.memoryConsumedBytes /= 1024.0;
    if (this.memoryConsumedBytes < 1024) {
      return Math.round(this.memoryConsumedBytes * 10) / 10 + ' Мб';
    }
    this.memoryConsumedBytes /= 1024.0;
    if (this.memoryConsumedBytes < 1024) {
      return Math.round(this.memoryConsumedBytes * 10) / 10 + ' Гб';
    }
  }

  get verdict() {
    switch (this._verdict) {
      case 'OK':
        return 'OK';
      case 'FAILED':
        return 'FAIL ' + this.passedTestCount;
      case 'PARTIAL':
        return 'PART ' + this.passedTestCount;
      case 'COMPILATION_ERROR':
        return 'CE';
      case 'RUNTIME_ERROR':
        return 'RE ' + this.passedTestCount;
      case 'WRONG_ANSWER':
        return 'WA ' + this.passedTestCount;
      case 'PRESENTATION_ERROR':
        return 'PE ' + this.passedTestCount;
      case 'TIME_LIMIT_EXCEEDED':
        return 'TL ' + this.passedTestCount;
      case 'MEMORY_LIMIT_EXCEEDED':
        return 'ML ' + this.passedTestCount;
      case 'INPUT_PREPARATION_CRASHED':
      case 'CRASHED':
        return 'CRASH ' + this.passedTestCount;
      case 'SKIPPED':
        return 'SKIP';
      case 'IDLENESS_LIMIT_EXCEEDED':
        return 'IL';
      case 'TESTING':
        return 'TEST';
      default:
        return this.verdict + ' ' + this.passedTestCount;
    }
  }

  status() {
    switch (this._verdict) {
      case 'OK':
        return 1;
      case 'PARTIAL':
        return 2;
      case 'FAILED':
      case 'COMPILATION_ERROR':
      case 'RUNTIME_ERROR':
      case 'WRONG_ANSWER':
      case 'PRESENTATION_ERROR':
      case 'TIME_LIMIT_EXCEEDED':
      case 'MEMORY_LIMIT_EXCEEDED':
      case 'INPUT_PREPARATION_CRASHED':
      case 'CRASHED':
      case 'IDLENESS_LIMIT_EXCEEDED':
        return 3;
      default:
        return 4;
    }
  }

  color() {
    switch (this.status()) {
      case 0:
        return '#000000';
      case 1:
        return '#0f7913';
      case 2:
        return '#72ED80';
      case 3:
        return '#92140C';
    }
  }

  backColor() {
    if (this.ok_before_dealine || this.status() != 1) {
      return '#FFFFFF';
    }
    return '#E5E5E5';
  }

  toString() {
    var text = '';
    if (this.status() == 1 && this.number_of_failures > 0) {
      text = `'+${this.number_of_failures} `;
    } else if (this.status() > 1 && this.number_of_failures > 1) {
      text = `'-${this.number_of_failures} `;
    }
    text += this.verdict;
    var group = this.task.group;
    if (group.show_date || group.show_time) {
      text += `\n${this.task.group.last_or_first}: `;
      if (group.show_date) {
        text += this.date();
      }
      if (group.show_date && group.show_time) {
        text += ' ';
      }
      if (group.show_time) {
        text += this.time();
      }
      text += ' ' + TIMEZONE_S;
    }
    if (group.show_langs) {
      text += '\n' + this.langs();
    }
    if (group.show_params) {
      text += '\n' + this.timeConsumed() + '   ' + this.memoryConsumed();
    }
    return text;
  }
}

function loadJsonData(url) {
  var count_retry = 0;
  var data;
  do {
    data = JSON.parse(UrlFetchApp.fetch(url).getAs("application/json").getDataAsString());
    count_retry++;
  } while (data["status"] != 'OK' && count_retry < MAX_COUNT_RETRY);
  return data;
}

function update() {
  var settings_file = DriveApp.getFileById(SETTINGS_ID);
  var settings = JSON.parse(settings_file.getAs("application/json").getDataAsString());
  var sheet = SpreadsheetApp.openById(settings['sheet_id']).getSheets()[0];
  var members = [];
  var tasks = new Map();
  settings['members'].forEach(function(item) {
    members.push(new Member(item['login'], item['showed_name']));
  });
  sheet.setFrozenRows(settings['total_calc'] ? 2 : 1);
  sheet.setFrozenColumns(1);

  // Печать общего рейтинга
  var offset_rows = 1;
  if (settings['total_calc'] || settings['total_calc_before_deadline']) {
    var text;
    if (settings['total_calc'] && settings['total_calc_before_deadline']) {
      text = 'Общий рейтинг\n(до дедлайна)';
    } else if (settings['total_calc']) {
      text = 'Общий рейтинг';
    } else {
      text = 'Общий рейтинг до дедлайна';
    }
    offset_rows++;
    sheet.getRange(2, 1)
      .setValue(text)
      .setTextStyle(BOLD_STYLE)
      .setTextStyle(BLACK_STYLE)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null)
      .setBorder(null, null, true, null, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setBackgroundRGB(255, 255, 255);
  } 
  
  settings['task_groups'].forEach(function(item) {
    // Парсинг групп заданий
    var group = new GroupTask(item['groop_name'], item['show'], offset_rows + 1);
    group.total_calc = item['total_calc'];
    group.total_calc_before_deadline = item['total_calc_before_deadline'];
    group.show_langs = item['show_langs'] && settings['show_langs'] === null || settings['show_langs'];
    group.show_time = item['show_time'] && settings['show_time'] === null || settings['show_time'];
    group.show_date = item['show_date'] && settings['show_date'] === null || settings['show_date'];
    group.show_params = item['show_params'] && settings['show_params'] === null || settings['show_params'];
    group.last_or_first = settings['last_or_first'];
    if (item['deadline'] !== undefined) {
      group.deadline = new Date(Date.parse(item['deadline']));
    }
    if (group.last_or_first === null) {
      group.last_or_first = item['last_or_first'];
    }
    // Заполнение названий групп заданий
    if (group.show) {
      offset_rows++;
      sheet.getRange(offset_rows, 1)
        .setValue('' + group)
        .setTextStyle(BOLD_STYLE)
        .setTextStyle(BLACK_STYLE)
        .setHorizontalAlignment('left');
      sheet.getRange(offset_rows, 1, 1, members.length + 1)
        .setBorder(null, null, null, true, null, null)
        .setBackgroundRGB(87, 203, 235);
      sheet.getRange(offset_rows, 2, 1, members.length)
        .setValue('');
    }
    sheet.getRange(offset_rows, 1, 1, members.length + 1)
      .setBorder(null, null, true, null, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    item['tasks'].forEach(function(task) {
      offset_rows++;
      var new_task = new Task(task['contestId'], task['index'], offset_rows, group);
      tasks.set('' + task['contestId'] + task['index'], new_task);
      sheet.getRange(offset_rows, 2, 1, members.length)
        .setValue('')
        .setBorder(null, null, null, false, null, null)
        .setBackground('#FFFFFF');
    });
    sheet.getRange(offset_rows, 1, 1, members.length + 1)
      .setBorder(null, null, true, null, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  var start_time = new Date();
  try {
    var url = 'https://codeforces.com/api/user.info?handles=';
    members.forEach(function(member, i) {
      if (i > 0) {
        url += ';';
      }
      url += member.login;
      
    });
    contest_data = loadJsonData(url);
    contest_data['result'].forEach(function(member, i) {
      members[i].cdf_rating = member['rating'];
    }); 
  } catch (e) {
    console.log('ERROR: ' + e)
  } finally {
    var curDate = null;
    do {
      curDate = new Date();
    } while(curDate - start_time < PAUSE);
  }

  members.forEach(function(member, i) {
    var start_time = new Date();
    try {
      // Загружаем API
      user_subs = loadJsonData(`https://codeforces.com/api/user.status?lang=ru&handle=${member.login}&from=1`)
      console.log(member.login);
      user_subs = user_subs['result'].reverse();
      user_subs.forEach(function(user_sub, i) {
        var contestId = user_sub['contestId'];
        var problemInfdex = user_sub['problem']['index'];
        var key = '' + contestId + problemInfdex;
        if (tasks.has(key)) {
          var task = tasks.get(key)
          var new_sub = new Sumbission(task, user_sub['id'], user_sub['verdict'], member, user_sub['creationTimeSeconds']);
          new_sub.passedTestCount = user_sub['passedTestCount'] + 1;
          new_sub.timeConsumedMillis = user_sub['timeConsumedMillis'];
          new_sub.memoryConsumedBytes = user_sub['memoryConsumedBytes'];
          new_sub.programmingLanguages.add(user_sub['programmingLanguage']);
          task.name = user_sub['problem']['name'];
          task.rating = user_sub['problem']['rating'];
          task.points = user_sub['problem']['points'];
          if (member.subs[task.row] === undefined) {
            member.subs[task.row] = new_sub;
          } else {
            var old_status = member.subs[task.row].status();
            var new_status = new_sub.status();
            if (old_status >= 3 && new_status < 3) { // ОК или PARTIAL после неверных попыток
              new_sub.number_of_failures = member.subs[task.row].number_of_failures;
              member.subs[task.row] = new_sub;
            } else if (old_status < 3 && new_status >= 3) { // Неверная попытка после удачных
              // Игнор
            } else {
              new_sub.ok_before_dealine |= member.subs[task.row].ok_before_dealine;
              new_sub.number_of_failures = member.subs[task.row].number_of_failures;
              if (new_status != 1) {
                new_sub.number_of_failures += 1;
              }
              if (new_status < 3) {
                new_sub.memoryConsumedBytes = Math.min(new_sub.memoryConsumedBytes, member.subs[task.row].memoryConsumedBytes);
                new_sub.timeConsumedMillis = Math.min(new_sub.timeConsumedMillis, member.subs[task.row].timeConsumedMillis);
              }
              if (new_sub.task.last_or_first == 'first') {
                new_sub.creationTimeMillis = member.subs[task.row].creationTimeMillis;
              }
              for (var lang of member.subs[task.row].programmingLanguages) {
                new_sub.programmingLanguages.add(lang);
              }
              member.subs[task.row] = new_sub;
            }
          }
        }
      });
      member.subs.forEach(function (sub) {
        member.addSubmission(sub);
      });
    } catch (e) {
      console.log(member.login + ' ERROR: ' + e)
    } finally {
      var curDate = null;
      do {
        curDate = new Date();
      } while(curDate - start_time < PAUSE);
    }
  });

  if (settings['sorting'] !== null && settings['sorting']['arg'] !== null) {
    var arg = settings['sorting']['arg'];
    if (settings['sorting']['direction'] == 'ASC') {
      members.sort((a,b)=> (a[arg] > b[arg] ? 1 : -1));
    }
    if (settings['sorting']['direction'] == 'DESC') {
      members.sort((a,b)=> (a[arg] < b[arg] ? 1 : -1));
    }
  }

  // Печать результатов
  members.forEach(function(member, i) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText('' + member)
      .setLinkUrl('https://codeforces.com/profile/' + member.login)
      .setTextStyle(BOLD_STYLE)
      .setTextStyle(NO_UNDERLINE_STYLE)
      .build();
    sheet.getRange(1, i + 2)
      .setRichTextValue(richText)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null)
      .setBorder(null, null, true, null, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setFontColor(member.color())
      .setBackgroundRGB(181, 230, 210);

    member.subs.forEach(function(sub, row) {
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('' + sub)
        .setLinkUrl(`https://codeforces.com/contest/${sub.task.contestId}/submission/${sub.id}`)
        .setTextStyle(NO_BOLD_STYLE)
        .setTextStyle(NO_UNDERLINE_STYLE)
        .build();
      sheet.getRange(row, i + 2)
        .setRichTextValue(richText)
        .setHorizontalAlignment('center')
        .setFontColor(sub.color())
        .setBackground(sub.backColor());
    });

    for (var [group, value] of member.total_groups) {
      if (group.show) {
        var text = '';
        if (group.total_calc) {
          text += value;
        }
        if (group.total_calc_before_deadline) {
          if (member.total_groups_before_deadline.has(group)) {
            if (text == '') {
              text += member.total_groups_before_deadline.get(group);
            } else {
              text += `\n(${member.total_groups_before_deadline.get(group)})`;
            }
          } else {
            if (text != '') {
              text += '\n(0)';
            }
          }
        }
        sheet.getRange(group.row, i + 2)  
          .setValue(text)
          .setTextStyle(BLACK_STYLE)
          .setTextStyle(BOLD_STYLE)
          .setHorizontalAlignment('center');
      }
    }

    if (settings['total_calc'] || settings['total_calc_before_deadline']) {
      var text;
      if (settings['total_calc'] && settings['total_calc_before_deadline']) {
        text = `${member.rating}\n(${member.rating_before_deadline})`;
      } else if (settings['total_calc']) {
        text = `${member.rating}`;
      } else {
        text = `${member.rating_before_deadline}`;
      }
      sheet.getRange(2, i + 2)
        .setValue(text)
        .setTextStyle(RED_STYLE)
        .setTextStyle(BOLD_STYLE)
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, null, null)
        .setBorder(null, null, true, null, null, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        .setBackgroundRGB(255, 255, 255);
    }
  });
  
  // Поиск названий заданий, для которых они ещё не получены
  var contest_for_parsing = new Set();
  for (var [_, task] of tasks) {
    if (task.name === null) {
      contest_for_parsing.add(task.contestId);
    }
  };
  for (var contestId of contest_for_parsing) {
    var start_time = new Date();
    try {
      contest_data = loadJsonData('https://codeforces.com/api/contest.standings?lang=ru&from=1&count=1&contestId=' + contestId);
      contest_data['result']['problems'].forEach(function(problem) {
        var key = problem['contestId'] + problem['index'];
        var task = tasks.get(key);
        if (task !== undefined) {
          task.name = problem['name'];
          task.rating = problem['rating'];
          task.points = problem['points'];
        }
      });
    } catch (e) {
      console.log('Contest ' + contestId + ' ERROR: ' + e)
    } finally {
      var curDate = null;
      do {
        curDate = new Date();
      } while(curDate - start_time < PAUSE);
    }
  }

  // Печать названий заданий
  for (var [_, task] of tasks) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText('' + task)
      .setLinkUrl(`https://codeforces.com/problemset/problem/${task['contestId']}/${task['index']}`)
      .setTextStyle(BOLD_STYLE)
      .setTextStyle(NO_UNDERLINE_STYLE)
      .setTextStyle(BLACK_STYLE)
      .build();
    sheet.getRange(task.row, 1)
      .setRichTextValue(richText)
      .setHorizontalAlignment('left')
      .setBorder(null, null, null, true, true, null)
      .setBackgroundRGB(255, 255, 255);
    sheet.getRange(task.row, 1 + members.length)
      .setBorder(null, null, null, true, true, null)
  }

  var date = new Date();
  var options = {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    weekday: 'long',
    timezone: 'UTC',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric'
  };
  sheet.getRange(1, 1)
    .setValue(`${date.toLocaleString("ru", options)} ${TIMEZONE_S}`);
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.openById(settings['sheet_id']);
  var resource = {"requests": [{"autoResizeDimensions": {"dimensions": {
    "dimension": "COLUMNS",
    "sheetId": ss.getActiveSheet().getSheetId(),
    "startIndex": 0,
    "endIndex": members.length,
  }}}]};
  Sheets.Spreadsheets.batchUpdate(resource, ss.getId());
}
