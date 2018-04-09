/* グローバル変数 */
var _spreadSheet;      // マスタースプレッドシートオブジェクト
var _calendar;         // カレンダーオブジェクト
var _lunchGroup = [];  // ランチグループ配列
var _teamMembers = [];   // チームメンバー一覧

/* 定数 */
var PERSONAL_INFO_MASTER = '個人情報マスター' // 個人情報マスター
var TEAM_MASTER          = 'チームマスター'  // チームマスター
var CONFIG_SHEET         = '設定'         // 設定
var CONFIRM_SHEET        = '予定確認'      // 予定確認用シート
var LUNCH_TIME;                          // ランチ設定月とランチタイム範囲(マスターから取得)
var PRE_LUNCH_COL;                       // 前回ランチに行ったグループの列番号
var GROUP_NUM;                           // ランチグループ数(マスターから取得)
var IS_FORCE_REGISTRATION;               // ランチの予定を強制的に入れるかどうか(マスターから取得)
var FORCE_LUNCH_DATE_RANGE;              // 強制予定登録する営業日日数(月末から何営業日か？ マスターから取得)
var SEARCH_ORDER;                        // 空き予定を月初から探すか月末から探すか(マスターから取得)
var OVERLAPPING_NUM        = 2;          // グループ内の職種重複可能人数(途中で緩和する可能性あり)
var GROUP_CONFIRM_ROW      = 2;          // グループ最終確認のために出力するシートの行番号
var PERSONAL_INFO_ROW_OFFS = 1;          // 人に割り当てられているIDと行番号とのオフセット
var NEW_LINE = String.fromCharCode(10);  // シート出力時の改行コード

/**
 * ランチグループのメンバー構成
 */
var LunchComp = function() {
  this.shinkis = [];    // 新規
  this.kizons = [];     // 運用
}

/**
 * ランチに行くグループクラス
 * @detail
 *    ランチメンバー、実行日、グループに含まれる最も多いグループを管理する
 */
var LunchGroup = function(id, list) {
  this.id = id;                                     // グループID
  this.members = new LunchComp();                   // ランチメンバー
  this.execDate = [];                               // ランチ実行日時(start, end)
  this.jobCategoryList = list;                      // 会社に存在するチーム名一覧(key:ジョブ名　value:グループに所属しているジョブ数)
  this.mostJobCategory;                             // グループ所属メンバー内で最も多いジョブ名
}

/**
 * グループメンバーのジョブリストとグループ内で最も多いジョブを更新する
 * @detail
 *    グループで同じジョブが被らないようにするため。
 *    チームに新規で参加する際に呼ばれる。
 * @param jobCategory 新規でグループに参加するメンバーのジョブ名
 */
LunchGroup.prototype.updateJobCategoryList = function(jobCategory) {
  this.jobCategoryList[jobCategory]++;

  var mostNum = 0;
  for (var job in this.jobCategoryList) {
    if (this.jobCategoryList[job] > mostNum) {
      mostNum = this.jobCategoryList[job];
      this.mostJobCategory = job;
    }
  }
}

/**
 * グループに参加するメンバーのジョブが、グループ内で最も多いジョブでないかどうか判定する
 * @detail
 *    重複可能人数(OVERLAPPING_NUM)を超えていなければ参加可能
 * @param jobCategory 参加を検討しているメンバーのジョブ名
 * @return 最も多いチームであれば true 、そうでなければ false を返す
 */
LunchGroup.prototype.isMostJobCategory = function(jobCategory) {
  return this.jobCategoryList[jobCategory] > OVERLAPPING_NUM && this.mostJobCategory == jobCategory;
}

LunchGroup.prototype.getLeaderPreGroup = function(type) {
  var preGroup = 0;
  if (type == "新規") {
    if (this.members.shinkis.length != 0) {
      preGroup = this.members.shinkis[0].preLunchGroup;
    }
  } else if (type == "運用") {
    if (this.members.kizons.length != 0) {
      preGroup = this.members.kizons[0].preLunchGroup;
    }
  }
  return preGroup;
}

/**
 * グループメンバーを１つの配列にして返す
 * @return グループのメンバーリストを返す　メンバーが0人であれば空の配列を返す
 */
LunchGroup.prototype.getMemberArray = function() {
  var array = [];
  this.members.shinkis.forEach(function(shinki) {
    array.push(shinki);
  });
  this.members.kizons.forEach(function(kizon) {
    array.push(kizon);
  });
  return array;
}

/**
 * グループメンバーのメールアドレスを1つの配列にして返す
 * @return グループメンバーのアドレスリストを返す　メンバーが0人であれば空の配列を返す
 */
LunchGroup.prototype.getMemberAddresses = function() {
  var addresses = [];
  this.getMemberArray().forEach(function(mem) {
    addresses.push(mem.address);
  });
  return addresses;
}

/**
 * マスタースプレッドシートに今回参加するグループのIDを書き込む
 * @detail
 *    PERSONAL_INFO_ROW_OFFS: 人に割り当てられているIDと行番号とのオフセット
 *    PRE_LUNCH_COL: 前回ランチに行ったグループの列番号
 */
LunchGroup.prototype.setLunchGroupIdInSheet= function() {
  var self = this;
  var setGroupId = function(personId) {
    setValueInSheet(PERSONAL_INFO_MASTER, personId + PERSONAL_INFO_ROW_OFFS, PRE_LUNCH_COL + 2, self.id);
  }
  this.getMemberArray().forEach(function(mem) {
    setGroupId(mem.id);
  });
}

/**
 * メンバー情報
 * @detail
 *   スプレッドシートのDBを元に構築される
 */
var Person = function(array) {
  this.id            = array[0];     // メンバーID
  this.name          = array[1];     // 名前
  this.address       = array[3];     // メールアドレス
  this.jobCategory   = array[6];     // 職種
  this.type          = array[7];     // タイプ
  this.isShuffable   = array[8] == '◯' ? true : false;                     // シャッフル対象かどうか
  this.preLunchGroup = array[PRE_LUNCH_COL] ? array[PRE_LUNCH_COL] : 0;    // 前回ランチグループID(前回未参加なら0)
  if (this.isShuffable) { CalendarApp.subscribeToCalendar(this.address); }  // マイカレンダーに他人のカレンダーを登録する
  this.calendar      = CalendarApp.getCalendarById(this.address);          // Googleカレンダーオブジェクト
  this.lunchTimeEvents = [];         // ランチタイム時の予定一覧(key:日 value:Eventオブジェクト)
};

/**
 * ランチタイム時のイベントを取得し、連想配列に格納する
 * @detail
 *    - 予定が入っていても問題ないと判断されるパターンが存在する
 *      - 予定が一日の時間を超えている場合(日で入力されている予定は丸一日予定が抑えられていることになっている)
 *      - okWordsのいずれかがイベント名に入っている場合
 *      - okPeopleに名前が入っている人の場合
 *    - 上記で問題ないとされても予定ありと判断するパターンが存在する
 *      - ngWordsのいずれかがイベント名に入っている場合
 *    - okWords,ngWords,okPeopleはマスタースプレッドシートから取得する
 *    - IS_FORCE_REGISTRATION: 強制予定登録フラグ
 * @param days ランチ実行可能日一覧
 */
Person.prototype.setLunchTimeEvents = function(days) {
  var self = this;

  if (IS_FORCE_REGISTRATION) {
    return;
  }

  var formatReg = function(str) { return "[" + str + "]"; }

  var configDB = _spreadSheet.getSheetByName(CONFIG_SHEET);
  var okPeople = getValuesInSheet(configDB, 11, 2, 1, configDB.getDataRange().getLastColumn() - 1)[0].filter(Boolean);
  var okWords = getValuesInSheet(configDB, 9, 2, 1, configDB.getDataRange().getLastColumn() - 1)[0].filter(Boolean);
  var ngWords = getValuesInSheet(configDB, 10, 2, 1, configDB.getDataRange().getLastColumn() - 1)[0].filter(Boolean);
  var okPRegex = new RegExp(okPeople.join('|'));
  var okWRegex = new RegExp(okWords.join('|'));
  var ngWRegex = new RegExp(ngWords.join('|'));

  var validateEvent = function(event) {
    var isNG = false;

    // 1.ランチ時間内に予定が終了する  2.ランチ時間内に予定が開始する  3.ランチ時間内に予定が終わらない
    var lunchStartTime = new Date(_yyyymmdd_hhmmss(event.getStartTime(), LUNCH_TIME.start));
    var lunchEndTime   = new Date(_yyyymmdd_hhmmss(event.getStartTime(), LUNCH_TIME.end));
    if ((lunchStartTime < event.getEndTime()   && event.getEndTime()   < lunchEndTime) ||
        (lunchStartTime < event.getStartTime() && event.getStartTime() < lunchEndTime) ||
        (lunchStartTime >= event.getStartTime() && event.getEndTime()  >= lunchEndTime)) {
      isNG = true;
    } else {
      isNG = false;
    }

    if (event.getEndTime().getTime() - event.getStartTime().getTime() >= 86400000) {  // 1日分(60 * 60 * 24 * 1000)
      isNG = false;
    }

    if (self.name.match(okPRegex)) { isNG = false; }
    if (event.getTitle().match(okWRegex)) { isNG = false; }
    if (event.getTitle().match(ngWRegex)) { isNG = true; }
    return isNG;
  };

  var events;
  if (SEARCH_ORDER == 'ASC') {
    events = self.calendar.getEvents(days[0], days[days.length - 1]);
  } else if (SEARCH_ORDER == 'DESC') {
    events = self.calendar.getEvents(days[days.length - 1], days[0]);
  }
  events.forEach(function(event) {
    if (validateEvent(event)) {
      self.lunchTimeEvents[event.getStartTime().getDate()] = event;    // ランチ先約の予定を登録する
    }
  });
}

/**
 * メンバーの予定が空いているかを判定する
 * @param day 予定を確認する日
 * @return 予定が入っていなければ true 、そうでなければ false を返す
 */
Person.prototype.isScheduleFree = function(day) {
  var event = this.lunchTimeEvents[day.getDate()];
  if (!event) {
    return true;
  } else {
    Logger.log(this.name + " さんに予定が入っています。\n  イベント:" + event.getTitle() + " " + event.getStartTime() + " " + event.getEndTime());
    return false;
  }
}

/**
 * ランチ設定月とランチタイム範囲
 * @detail
 *   スプレッドシートのDBを元に構築される。
 */
var LunchTime = function(month, start, end) {
  this.month = month - 1;  // ランチ設定月(0-11の範囲)
  this.start = start;      // start time
  this.end   = end;        // end time
}

/* ユーティリティ */
/**
 * Dateクラスのインスタンスを生成するために(yyyy/mm/dd hh:mm:ss)形式に変更する
 * @param date 日時(Dateインスタンス)
 * @param time 時間(Dateインスタンス)
 * @return yyyy/mm/dd hh:mm:ssの文字列にして返す
 */
var _yyyymmdd_hhmmss = function(date, time) {
  var time = time || date;
  var ret = "";
  var month = date.getMonth() + 1;
  ret += date.getFullYear() + "/" + month + "/" + date.getDate();
  ret += " " + time.getHours() + ":" + time.getMinutes() + ":" + time.getSeconds();
  return ret;
};

/**
 * 0-nの範囲でランダムに整数値を返す
 * @param n ランダム値の上限
 * @return 0-nの範囲でランダムに整数値を返す
 */
function rand(n) {
  return Math.floor(Math.random() * n);
}

/**
 * arrayの中からランダムで1つを選出する
 * @param array 母集団
 * @return 母集団からランダムに選出されたオブジェクトを返す
 */
function getRandomElement(array) {
  var n = rand(array.length);
  return array[n];
}

/**
 * ランチグループの中からランダムで1つのグループを選出する
 * @return ランチグループの中からランダムで選出されたグループのIDを返す
 */
function getRandomGroupId() {
  return rand(GROUP_NUM) + 1;
}

/* シート値のアクセッサ */
function setValueInSheet(sheetName, row, col, value) {
  _spreadSheet.getSheetByName(sheetName).getRange(row, col).setValue(value);
}

function getValueInSheet(sheetName, row, col) {
  return _spreadSheet.getSheetByName(sheetName).getRange(row, col).getValue();
}

function getValuesInSheet(sheet, row, col, numRows, numCols) {
  numRows = numRows || 1;
  numCols = numCols || 1;
  return values = sheet.getRange(row, col, numRows, numCols).getValues();
}

/**
 * グループ分け実行関数
 */
function execGrouping() {
  initSpreadSheet();

  var isExec = getValueInSheet(CONFIG_SHEET, 1, 2);
  if (!isExec) {
    return;
  }

  initForGrouping();
  createPersonalData();
  groupMembers();

  // ランチ実行日を決定
  _lunchGroup.forEach(function(group) {
    var exec;
    if (IS_FORCE_REGISTRATION) {
      exec = getForceRandomLunchExecDate();
    } else {
      exec = searchLunchExecDate(group);
    }
    if (Math.round(exec.getTime() / 1000) != 0) {
      group.execDate['start'] = new Date(_yyyymmdd_hhmmss(exec, LUNCH_TIME.start));
      group.execDate['end']   = new Date(_yyyymmdd_hhmmss(exec, LUNCH_TIME.end));
    }
  });

  setLunchGroupInfo();
  setValueInSheet(CONFIG_SHEET, 1, 2, 'FALSE');
}

/**
 * スプレッドシート初期化
 */
function initSpreadSheet() {
  _spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * グローバル変数初期化、ランチグループを作成する
 */
function initForGrouping() {
  GROUP_NUM = getValueInSheet(CONFIG_SHEET, 4, 2);
  PRE_LUNCH_COL = getValueInSheet(CONFIG_SHEET, 5, 2);

  LUNCH_TIME = new LunchTime(getValueInSheet(CONFIG_SHEET, 6, 2),
                             getValueInSheet(CONFIG_SHEET, 7, 2),
                             getValueInSheet(CONFIG_SHEET, 8, 2));
  IS_FORCE_REGISTRATION = getValueInSheet(CONFIG_SHEET, 12, 2);
  FORCE_LUNCH_DATE_RANGE = getValueInSheet(CONFIG_SHEET, 13, 2);
  SEARCH_ORDER = getValueInSheet(CONFIG_SHEET, 14, 2);

  var createJobCategoryList = function() {
    var ret = [];
    var teamDB = _spreadSheet.getSheetByName(TEAM_MASTER);
    var JobCategoryList = getValuesInSheet(teamDB, 2, 1, teamDB.getDataRange().getLastRow() - 1);
    JobCategoryList.forEach(function(jobCategory) {
      ret[jobCategory] = 0;
    });
    return ret;
  }

  // ランチグループを生成
  for (var i = 1; i <= GROUP_NUM; i++) {
      _lunchGroup[i] = new LunchGroup(i, createJobCategoryList());
  }
}

/**
 * マスタースプレッドシートからPersonオブジェクトを生成し、条件に合う人を配列に格納する
 */
function createPersonalData() {
  var lunchDB = _spreadSheet.getSheetByName(PERSONAL_INFO_MASTER);
  var personalData = getValuesInSheet(lunchDB, 2, 1, lunchDB.getDataRange().getLastRow() - 1, lunchDB.getDataRange().getLastColumn());
  personalData.forEach(function(data) {
    var person = new Person(data);
    if (person.isShuffable)　{
      _teamMembers.push(person);
    }
  });
}

/**
 * チームメンバーをランチグループに振り分ける
 */
function groupMembers() {
  var errorCount = 0;

  // 新規
  var shinkis = [];
  _teamMembers.forEach(function(value) {
    if (value.type == "新規") {
      shinkis.push(value);
    }
  });
  var MAX_MEMBER = Math.floor(shinkis.length / GROUP_NUM);
  while (shinkis.length != 0) {
    var shinki = getRandomElement(shinkis);
    var groupId = getRandomGroupId();
    if (!_lunchGroup[groupId].isMostJobCategory(shinki.jobCategory) &&
        _lunchGroup[groupId].members.shinkis.length < MAX_MEMBER    &&
        groupId != _lunchGroup[groupId].getLeaderPreGroup(shinki.type)) {
       _lunchGroup[groupId].members.shinkis.push(shinki);
       _lunchGroup[groupId].updateJobCategoryList(shinki.jobCategory);
       shinkis.splice(shinkis.indexOf(shinki), 1);
       _teamMembers.splice(_teamMembers.indexOf(shinki), 1);
       errorCount = 0;
    } else {
      errorCount++;
      if (errorCount == 1000) {
        MAX_MEMBER++;
        OVERLAPPING_NUM++;
        errorCount = 0;
      }
    }
  }

  // 運用
  var kizons = [];
  _teamMembers.forEach(function(value) {
    if (value.type == "運用") {
      kizons.push(value);
    }
  });
  var MAX_MEMBER = Math.floor(kizons.length / GROUP_NUM);
  while (kizons.length != 0) {
    var kizon = getRandomElement(kizons);
    var groupId = getRandomGroupId();
    if (!_lunchGroup[groupId].isMostJobCategory(kizon.jobCategory) &&
        _lunchGroup[groupId].members.kizons.length < MAX_MEMBER    &&
        groupId != _lunchGroup[groupId].getLeaderPreGroup(kizon.type)) {
       _lunchGroup[groupId].members.kizons.push(kizon);
       _lunchGroup[groupId].updateJobCategoryList(kizon.jobCategory);
       kizons.splice(kizons.indexOf(kizon), 1);
       _teamMembers.splice(_teamMembers.indexOf(kizon), 1);
       errorCount = 0;
    } else {
      errorCount++;
      if (errorCount == 1000) {
        MAX_MEMBER++;
        OVERLAPPING_NUM++;
        errorCount = 0;
      }
    }
  }
}

/**
 * グループ内のメンバーのカレンダーから指定時間が空いているかを検索する
 * @param group 予定を決めるグループ
 * @return ランチ実行日 予定が全く空いていない場合は unixTime=0を返す
 */
function searchLunchExecDate(group) {
  var canLunchDays = getLunchAvailableDays();
  var members = group.getMemberArray();

  members.forEach(function(member) {
    member.setLunchTimeEvents(canLunchDays);
  });

  for (var i = 0; i < canLunchDays.length; i++) {
    var isAllOk = true;
    Logger.log("******** " + canLunchDays[i].toDateString() + " ********");
    members.forEach(function(member) {
      if (!member.isScheduleFree(canLunchDays[i])) {
        isAllOk = false;
      }
    });

    if (isAllOk) {
      Logger.log("グループID:" + group.id + " ALL OK!!");
      return canLunchDays[i];
    }
  }

  Logger.log("グループID:" + group.id + " NG!!");
  return new Date(0);
}

/**
 * グループのランチ予定を月末５営業日でランダムに振り分ける
 * @detail
 *    - 予定強制登録フラグがtrueの場合に実行される
 * @param group 予定を決めるグループ
 * @return ランチ実行日 予定が全く空いていない場合は unixTime=0を返す
 */
function getForceRandomLunchExecDate() {
  var ret = new Date(0);

  var canLunchDays = getLunchAvailableDays(FORCE_LUNCH_DATE_RANGE);
  var maxGroupInDate = Math.floor(GROUP_NUM / canLunchDays.length);
  maxGroupInDate = maxGroupInDate || 1;

  var dateCounter = function(date) {
    var c = 0;
    _lunchGroup.forEach(function(g) {
      if (typeof g.execDate['start'] !== "undefined" &&
          g.execDate['start'].getDate() == date.getDate()) {
        c++;
      }
    });
    return c;
  };
  var isNG = true;
  var error = 0;
  while(isNG) {
    var randomLunchDay = getRandomElement(canLunchDays);
    if (dateCounter(randomLunchDay) < maxGroupInDate) {
      ret = randomLunchDay;
      isNG = false;
    } else {
      error++;
      if (error > 50) {
        maxGroupInDate++;
        error = 0;
      }
    }
  }
  return ret;
}

/**
 * ランチ可能日を配列にして返す
 * @detail
 *   ランチ可能日: 休日、祝日を除いた日、つまり平日
 * @param range 月末からのrange営業日だけ取得したい場合に指定 default:0
 * @return ランチ可能日の配列を返す
 */
function getLunchAvailableDays(range) {
    range = range || 0;
    var isBusinessDay = function(date) {
    if (date.getDay() == 0 || date.getDay() == 6) {
      return false;
    }
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if(calJa.getEventsForDay(date).length > 0){
      return false;
    }
    return true;
  }

  var days = [];
  var tmpDate = new Date();

  // ランチ可能営業日の開始を指定(当月なら次の日、来月なら月初)
  var sDate;
  if (tmpDate.getMonth() == LUNCH_TIME.month) {
    sDate = new Date(tmpDate.getFullYear(), tmpDate.getMonth(), tmpDate.getDate() + 1);
  } else {
    sDate = new Date(tmpDate.getFullYear(), LUNCH_TIME.month, 1);
  }

  // 終了を指定(開始日の月末)
  var eDate;
  eDate = new Date(sDate.getFullYear(), sDate.getMonth() + 1, 0);

  if (SEARCH_ORDER == 'ASC') {
    for (var d = sDate; d <= eDate; d.setDate(d.getDate() + 1)) {
      if (isBusinessDay(d)) {
        days.push(new Date(d));
      }
    }
  } else if (SEARCH_ORDER == 'DESC') {
    for (var d = eDate; d >= sDate; d.setDate(d.getDate() - 1)) {
      if (isBusinessDay(d)) {
        days.push(new Date(d));
      }
    }
  }

  if (range) {
    var remainingDays = Math.abs(days[0].getDate() - days[days.length - 1].getDate());
    if (range < remainingDays) {
      days.splice(range, days.length - range);
    }
  }

  return days;
}

/**
 * ランチグループ情報を整形し、マスタースプレッドシートに書き込む
 * @detail
 *   書き込んだ結果を見て問題ないようならカレンダーに登録するボタンを押してもらう
 */
function setLunchGroupInfo() {
  var C_GROUP_ID   = 1;
  var C_START_TIME = 2;
  var C_SHINKI     = 3;
  var C_KIZON      = 4;

  var C_SYS_START_TIME = 8;
  var C_SYS_END_TIME   = 9;
  var C_SYS_GUESTS     = 10;

  var lunchGroupStr = [];
  _lunchGroup.forEach(function(group) {
    var ROW = group.id + 2;

    // ランチグループ情報をマスタースプレッドシートに反映
    setValueInSheet(CONFIRM_SHEET, ROW, C_GROUP_ID, group.id);
    setValueInSheet(CONFIRM_SHEET, ROW, C_START_TIME, group.execDate['start'] || '');
    var members = group.members;

    var getNameArray = function(members) {
      var nameArray = [];
      members.forEach(function(member) {
        nameArray.push(member.name);
      });
      return nameArray.join(NEW_LINE);
    };
    setValueInSheet(CONFIRM_SHEET, ROW, C_SHINKI, getNameArray(members.shinkis));
    setValueInSheet(CONFIRM_SHEET, ROW, C_KIZON, getNameArray(members.kizons));

    // カレンダー作成に必要な情報をマスタースプレッドシートに反映
    setValueInSheet(CONFIRM_SHEET, ROW, C_SYS_START_TIME, group.execDate['start'] || '');
    setValueInSheet(CONFIRM_SHEET, ROW, C_SYS_END_TIME, group.execDate['end'] || '');
    setValueInSheet(CONFIRM_SHEET, ROW, C_SYS_GUESTS, group.getMemberAddresses().join());

    // 今回のランチグループをマスタースプレッドシートに反映
    group.setLunchGroupIdInSheet();
  });

  // 最終更新日を反映
  setValueInSheet(CONFIRM_SHEET, 1, 2, new Date());
}

/**
 * カレンダー用の情報を集約するクラス
 */
var CalendarInfo = function(id, info){
  this.groupId       = id + 1;             // GroupId
  this.title         = info[0];            // カレンダータイトル
  this.description   = info[1];　　　       // カレンダー説明
  this.startTime     = new Date(info[2]);  // 予定開始時間
  this.endTime       = new Date(info[3]);  // 予定終了時間
  this.guests        = info[4];            // ゲスト登録するメンバーのアドレス
  this.isSendInvites = info[5];            // ゲストに招待メールを送るかどうか
}

/**
 * カレンダー初期化
 */
function initCalendar() {
  var address = Session.getActiveUser().getEmail();
  CalendarApp.subscribeToCalendar(address);
  _calendar = CalendarApp.getCalendarById(address);
}

/**
 * スケジュール設定実行関数
 * @detail
 *    グルーピングされた状態で実行すると、予定がメンバーのカレンダーに書き込まれる
 */
function setCalendarAtGroupCalendar() {
  initSpreadSheet();
  initCalendar();

  var calendarInfoList = [];
  var confirmSheet = _spreadSheet.getSheetByName(CONFIRM_SHEET);
  var calInfoData = getValuesInSheet(confirmSheet, 3, 6, confirmSheet.getDataRange().getLastRow() - 2, confirmSheet.getDataRange().getLastColumn() - 5);
  for (var i = 0; i < calInfoData.length; i++) {
    calendarInfoList.push(new CalendarInfo(i, calInfoData[i]));
  }

  var ROW_OFFS = 2;
  var COLUMN   = 12;
  calendarInfoList.forEach(function(info) {
    var calEvent = _calendar.createEvent(info.title, info.startTime, info.endTime,
                                            {description: info.description,
                                             guests:      info.guests,
                                             sendInvites: info.isSendInvites});
    calEvent.setGuestsCanModify(true);
    setValueInSheet(CONFIRM_SHEET, info.groupId + ROW_OFFS, COLUMN, calEvent.getId());
  });
}

/**
 * カレンダー削除実行関数
 */
function deleteCalendar() {
  initSpreadSheet();
  initCalendar();
  var confirmSheet = _spreadSheet.getSheetByName(CONFIRM_SHEET);
  var calInfoData = getValuesInSheet(confirmSheet, 3, 12, confirmSheet.getDataRange().getLastRow() - 2, 2);
  var ROW_OFFS = 3;
  for (var i = 0; i < calInfoData.length; i++) {
    var isDelete = calInfoData[i][1];
    if (isDelete) {
      var event = _calendar.getEventById(calInfoData[i][0]);
      event.deleteEvent();
      setValueInSheet(CONFIRM_SHEET, i + ROW_OFFS, 12, '');
      setValueInSheet(CONFIRM_SHEET, i + ROW_OFFS, 13, '');
    }
  }
}
