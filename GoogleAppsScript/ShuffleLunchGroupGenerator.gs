/* グローバル変数 */
var _spreadSheet;      // マスタースプレッドシートオブジェクト
var _calendar;         // カレンダーオブジェクト
var _lunchGroup = [];  // ランチグループ配列
var _directors = [];   // ディレクター一覧

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
var MIN_REQUIRED_NUM       = 2;          // グループに最低でも必要な要素数 leader+neuron=2
var OVERLAPPING_NUM        = 2;          // グループ内のチーム重複可能人数(途中で緩和する可能性あり)
var GROUP_CONFIRM_ROW      = 2;          // グループ最終確認のために出力するシートの行番号
var PERSONAL_INFO_ROW_OFFS = 2;          // 人に割り当てられているIDと行番号とのオフセット
var NEW_LINE = String.fromCharCode(10);  // シート出力時の改行コード

/**
 * ディレクターランチグループのメンバー構成
 */
var DirectorLunchComp = function() {
  this.leader;             // リーダー
  this.neuron;             // ニューロン(仕切り役)
  this.shinsotsus = [];    // 新卒たち
  this.people = [];        // その他の人たち
}

/**
 * ランチに行くグループクラス
 * @detail
 *    ランチメンバー、実行日、グループに含まれる最も多いグループを管理する
 */
var LunchGroup = function(id, list) {
  this.id = id;                              // グループID
  this.members = new DirectorLunchComp();    // ランチメンバー
  this.execDate = [];                        // ランチ実行日時(start, end)
  this.teamList = list;                      // 会社に存在するチーム名一覧(key:チーム名　value:グループに所属しているチーム数)
  this.mostTeam;                             // グループ所属メンバー内で最も多いチーム名
}

/**
 * グループメンバーが属するチームのリストとグループ内で最も多いチームを更新する
 * @detail
 *    グループで同じチームが被らないようにするため。
 *    チームに新規で参加する際に呼ぶ。
 * @param team 新規でグループに参加するメンバーのチーム名
 */
LunchGroup.prototype.updateTeamList = function(team) {
  this.teamList[team]++;

  var mostNum = 0;
  for (var t in this.teamList) {
    if (this.teamList[t] > mostNum) {
      mostNum = this.teamList[t];
      this.mostTeam = t;
    }
  }
}

/**
 * グループに参加するメンバーが属するチームが、グループ内で最も多いチームでないかどうか判定する
 * @detail
 *    重複可能人数(OVERLAPPING_NUM)を超えていなければ参加可能
 * @param team 参加を検討しているメンバーのチーム名
 * @return 最も多いチームであれば true 、そうでなければ false を返す
 */
LunchGroup.prototype.isMostTeam = function(team) {
  return this.teamList[team] > OVERLAPPING_NUM && this.mostTeam == team;
}

/**
 * グループメンバーを１つの配列にして返す
 * @return グループのメンバーリストを返す　メンバーが0人であれば空の配列を返す
 */
LunchGroup.prototype.getMemberArray = function() {
  var array = [];
  array.push(this.members.leader);
  array.push(this.members.neuron);
  this.members.shinsotsus.forEach(function(shinsotsu) {
    array.push(shinsotsu);
  });
  this.members.people.forEach(function(person) {
    array.push(person);
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
  this.team          = array[7];     // 所属チーム
  this.address       = array[9];     // メールアドレス
  this.jobCategory   = array[8];     // 職種(ex: ディレクター、エンジニア)
  this.type          = array[13];    // タイプ(ex: PL、ベテラン、新卒)
  this.hasAccount    = array[10] == '◯' ? true : false;                   // アカウントがあるかどうか
  this.isNeuron      = array[12] == 'ニューロン' ? true : false;              // ニューロンかどうか
  this.preLunchGroup = array[PRE_LUNCH_COL] ? array[PRE_LUNCH_COL] : 0;    // 前回ランチグループID(前回未参加なら0)
  if (this.hasAccount) { CalendarApp.subscribeToCalendar(this.address); }  // マイカレンダーに他人のカレンダーを登録する
  this.calendar      = CalendarApp.getCalendarById(this.address);          // Googleカレンダーオブジェクト
  this.lunchTimeEvents = [];         // ランチタイム時の予定一覧(key:日 value:Eventオブジェクト)
};

/**
 * リーダーかどうかを判定する
 * @return リーダーと定義されるタイプであれば true 、そうでなければ false を返す
 */
Person.prototype.isLeader = function() {
  return this.type == 'GM' || this.type == 'PL' || this.type == 'ベテラン';
}

/**
 * 新卒かどうかを判定する
 * @return 新卒と定義されるタイプであれば true 、そうでなければ false を返す
 */
Person.prototype.isShinsotsu = function() {
  return this.type == '新卒';
}

/**
 * ランチタイム時のイベントを取得し、連想配列に格納する
 * @detail
 *    - 1ヶ月の予定を一気に取得するか、日ごとに取得するか迷ったが今は日ごとに取得している
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

  var events = self.calendar.getEvents(days[0], days[days.length - 1]);
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
  groupDirectorMembers();

  // ランチ実行日を決定
  _lunchGroup.forEach(function(group) {
    var exec;
    if (IS_FORCE_REGISTRATION) {
      exec = getForceRandomLunchExecDate(group);
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

  var createTeamList = function() {
    var ret = [];
    var teamDB = _spreadSheet.getSheetByName(TEAM_MASTER);
    var teamList = getValuesInSheet(teamDB, 2, 1, teamDB.getDataRange().getLastRow() - 1);
    teamList.forEach(function(team) {
      ret[team] = 0;
    });
    return ret;
  }

  // ランチグループを生成
  for (var i = 1; i <= GROUP_NUM; i++) {
      _lunchGroup[i] = new LunchGroup(i, createTeamList());
  }
}

/**
 * マスタースプレッドシートからPersonオブジェクトを生成し、ディレクターのみを配列に格納する
 */
function createPersonalData() {
  var lunchDB = _spreadSheet.getSheetByName(PERSONAL_INFO_MASTER);
  var personalData = getValuesInSheet(lunchDB, 3, 1, lunchDB.getDataRange().getLastRow() - 2, lunchDB.getDataRange().getLastColumn());
  personalData.forEach(function(data) {
    var person = new Person(data);
    if (person.hasAccount && person.jobCategory == 'ディレクター')　{
      _directors.push(person);
    }
  });
}

/**
 * ディレクターメンバーをランチグループに振り分ける
 * @detail
 *    - リーダー、ニューロン、新卒、その他メンバーを順番に振り分ける
 */
function groupDirectorMembers() {
  // leader選択
  var leaders = [];
  _directors.forEach(function(value) {
    if (value.isLeader()) {
      leaders.push(value);
    }
  });
  var cntL = 0;
  while (cntL < GROUP_NUM) {
    var leader = getRandomElement(leaders);
    var groupId = getRandomGroupId();
    if (!_lunchGroup[groupId].members.leader) {
      _lunchGroup[groupId].members.leader = leader;
      _lunchGroup[groupId].updateTeamList(leader.team);
      _directors.splice(_directors.indexOf(leader), 1);
      leaders.splice(leaders.indexOf(leader), 1);
      cntL++;
    }
  }

  // ニューロン選択
  var neurons = [];
  _directors.forEach(function(value) {
    if (value.isNeuron) {
      neurons.push(value);
    }
  });
  var cntN = 0;
  var neuronsCopy = neurons.slice(0, neurons.length);
  var deleteNeurons = [];
  while (cntN < GROUP_NUM) {
    if (neurons.length == 0) {
      neurons = neuronsCopy.slice(0, neuronsCopy.length);
    }
    var neuron = getRandomElement(neurons);
    var groupId = getRandomGroupId();
    if (!_lunchGroup[groupId].members.neuron) {
      _lunchGroup[groupId].members.neuron = neuron;
      _lunchGroup[groupId].updateTeamList(neuron.team);
      deleteNeurons[neuron.id] = neuron;
      neurons.splice(neurons.indexOf(neuron), 1);
      cntN++;
    }
  }
  deleteNeurons.forEach(function(value) {
    _directors.splice(_directors.indexOf(value), 1);
  });

  // 新卒メンバーとその他メンバーで使用する
  var errorCount = 0;
  var isConditionRelaxation = false;

  // 新卒選択
  var shinsotsus = [];
  _directors.forEach(function(value) {
    if (value.isShinsotsu()) {
      shinsotsus.push(value);
    }
  });
  var MAX_SHINSOTSU_MEMBER = Math.floor(shinsotsus.length / GROUP_NUM);
  while (shinsotsus.length != 0) {
    var shinsotsu = getRandomElement(shinsotsus);
    var groupId = getRandomGroupId();
    /** 新卒メンバー選出条件
     * - 前回のリーダーと同じグループではない(但し、この条件で詰まるようなら緩和する)
     * - グループで一番多いチームメンバーではない(3人以上同じチームがいないこと)
     * - グループの人数がオーバーしていない         **/
    if ((groupId != _lunchGroup[groupId].members.leader.preLunchGroup || isConditionRelaxation) &&
        !_lunchGroup[groupId].isMostTeam(shinsotsu.team)                                        &&
        _lunchGroup[groupId].members.shinsotsus.length < MAX_SHINSOTSU_MEMBER) {
       _lunchGroup[groupId].members.shinsotsus.push(shinsotsu);
       _lunchGroup[groupId].updateTeamList(shinsotsu.team);
       shinsotsus.splice(shinsotsus.indexOf(shinsotsu), 1);
       _directors.splice(_directors.indexOf(shinsotsu), 1);
       errorCount = 0;
    } else {
      errorCount++;
      if (errorCount == 500) {
        isConditionRelaxation = true;
      } else if (errorCount == 1000) {
        MAX_SHINSOTSU_MEMBER++;
        OVERLAPPING_NUM++;
        errorCount = 0;
      }
    }
  }

  // メンバー選択
  errorCount = 0;
  isConditionRelaxation = false;
  OVERLAPPING_NUM = 2;
  var MAX_MEMBER = Math.floor(_directors.length / GROUP_NUM);
  while (_directors.length != 0) {
    var mem = getRandomElement(_directors);
    var groupId = getRandomGroupId();
    /** メンバー選出条件
     * - 前回のリーダーと同じグループではない(但し、この条件で詰まるようなら緩和する)
     * - グループで一番多いチームメンバーではない(3人以上同じチームがいないこと)
     * - グループの人数がオーバーしていない         **/
    if ((groupId != _lunchGroup[groupId].members.leader.preLunchGroup || isConditionRelaxation) &&
        !_lunchGroup[groupId].isMostTeam(mem.team) &&
        _lunchGroup[groupId].members.people.length < MAX_MEMBER) {
       _lunchGroup[groupId].members.people.push(mem);
       _lunchGroup[groupId].updateTeamList(mem.team);
      _directors.splice(_directors.indexOf(mem), 1);
      errorCount = 0;
    } else {
      errorCount++;
      if (errorCount == 500) {
        isConditionRelaxation = true;
      } else if (errorCount == 1000) {
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

    if (isSchedulesDuplicated(group.members.neuron, canLunchDays[i])) {
      Logger.log("ニューロンの " + group.members.neuron.name + " さんが同じ日に別のランチ予定が入っています。");
      isAllOk = false;
    }

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
 *    - ニューロンだけは被らないように調整する
 * @param group 予定を決めるグループ
 * @return ランチ実行日 予定が全く空いていない場合は unixTime=0を返す
 */
function getForceRandomLunchExecDate(group) {
  var ret = new Date(0);

  var canLunchDays = getLunchAvailableDays(FORCE_LUNCH_DATE_RANGE);
  var maxGroupInDate = Math.floor(GROUP_NUM / canLunchDays.length);
  maxGroupInDate = maxGroupInDate || 1;

  // (グループ数 / 営業日日数) > ニューロン数 だった場合は無限ループなので検知してunixTime:0を返す
  var getUniqueNeuronNum = function() {
    var uniqueNeuron = [];
    _lunchGroup.forEach(function(group) {
      uniqueNeuron.push(group.members.neuron.id);
    });
    return uniqueNeuron.filter(function (x, i, self) { return self.indexOf(x) !== i; }).length;
  }
  if (Math.ceil(GROUP_NUM / canLunchDays.length) > getUniqueNeuronNum()) {
    return ret;
  }

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
    if (dateCounter(randomLunchDay) < maxGroupInDate &&
        !isSchedulesDuplicated(group.members.neuron, randomLunchDay)) {
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
 * 既に所属しているグループのランチ日と被っていないか判定する
 * @detail
 *    基本的にはニューロンの予定被り判定に使う
 * @param member 判定したいメンバーのオブジェクト
 * @param date   ランチ予定日
 * @return 既に所属しているグループのランチ日と被っていればtrue 、そうでなければ false を返す
 */
function isSchedulesDuplicated(member, date) {
  for (var i = 1; i < _lunchGroup.length; i++) {
    var group = _lunchGroup[i];
    if (group.members.neuron.id == member.id &&
        typeof group.execDate['start'] !== "undefined" &&
        group.execDate['start'].getDate() == date.getDate()) {
      return true;
    }
  }
  return false;
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
    var remainingDays = days[0].getDate() - days[days.length - 1].getDate();
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
  var C_LEADER     = 3;
  var C_NEURON     = 4;
  var C_MEMBER     = 5;
  var C_SHINSOTSU  = 6;

  var C_SYS_START_TIME = 10;
  var C_SYS_END_TIME   = 11;
  var C_SYS_GUESTS     = 12;

  var lunchGroupStr = [];
  _lunchGroup.forEach(function(group) {
    var ROW = group.id + 2;

    // ランチグループ情報をマスタースプレッドシートに反映
    setValueInSheet(CONFIRM_SHEET, ROW, C_GROUP_ID, group.id);
    setValueInSheet(CONFIRM_SHEET, ROW, C_START_TIME, group.execDate['start'] || '');
    var members = group.members;
    setValueInSheet(CONFIRM_SHEET, ROW, C_LEADER, members.leader.name);
    setValueInSheet(CONFIRM_SHEET, ROW, C_NEURON, members.neuron.name);
    var getNameArray = function(members) {
      var nameArray = [];
      members.forEach(function(member) {
        nameArray.push(member.name);
      });
      return nameArray.join(NEW_LINE);
    };
    setValueInSheet(CONFIRM_SHEET, ROW, C_MEMBER, getNameArray(members.people));
    setValueInSheet(CONFIRM_SHEET, ROW, C_SHINSOTSU, getNameArray(members.shinsotsus));

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
  var calInfoData = getValuesInSheet(confirmSheet, 3, 8, confirmSheet.getDataRange().getLastRow() - 2, confirmSheet.getDataRange().getLastColumn() - 7);
  for (var i = 0; i < calInfoData.length; i++) {
    calendarInfoList.push(new CalendarInfo(i, calInfoData[i]));
  }

  var ROW_OFFS = 2;
  var COLUMN   = 14;
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
  var calInfoData = getValuesInSheet(confirmSheet, 3, 14, confirmSheet.getDataRange().getLastRow() - 2, 2);
  var ROW_OFFS = 3;
  for (var i = 0; i < calInfoData.length; i++) {
    var isDelete = calInfoData[i][1];
    if (isDelete) {
      var event = _calendar.getEventById(calInfoData[i][0]);
      event.deleteEvent();
      setValueInSheet(CONFIRM_SHEET, i + ROW_OFFS, 14, '');
      setValueInSheet(CONFIRM_SHEET, i + ROW_OFFS, 15, '');
    }
  }
}
