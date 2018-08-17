var PLAYLIST_SHEET_NAME = 'PlayListMaster';
var OUTPUT_SHEET_NAMES = ['ViewCount', 'LikeCount', 'DislikeCount', 'CommentCount'];
var TRENDING_OUTPUT_SHEET_NAME = 'Trending'
var CHANNELLIST_SHEET_NAME = 'ChannelListMaster';
var CHANNEL_OUTPUT_SHEET_NAMES = ['ChannelSubscriberCount'];
var MAX_RESEARCH_LIST_SIZE = 50;

var _spreadSheet;        // マスタースプレッドシートオブジェクト
var _movieList = [];     // 動画リスト
var _channelList = [];   // チャンネルリスト

/* シート値のアクセッサ */
function setValueInSheet(sheetName, row, col, value) {
  _spreadSheet.getSheetByName(sheetName).getRange(row, col).setValue(value);
}
function getValueInSheet(sheetName, row, col) {
  return _spreadSheet.getSheetByName(sheetName).getRange(row, col).getValue();
}
function getValuesInSheet(sheetName, startRow, startCol, numRow, numCol) {
  numCol = (typeof numCol !== 'undefined') ? numCol : 1;
  return _spreadSheet.getSheetByName(sheetName).getRange(startRow, startCol, numRow, numCol).getValues();
}

/* シートの終端の取得 */
function getNewLineRow(sheetName) {
  return _spreadSheet.getSheetByName(sheetName).getLastRow() + 1;
}
function getNewLineCol(sheetName) {
  return _spreadSheet.getSheetByName(sheetName).getLastColumn() + 1;
}

function toInt(num) {
  return (typeof num !== 'undefined') ? num : 0;
};

/**
 * 出力形式(動画リサーチ)
 */
var ResearchResult = function() {
  this.videoId;    　　// 動画ID
  this.title;      　　// 動画タイトル
  this.viewCount;  　　// 再生回数
  this.likeCount;      // いいね数
  this.dislikeCount;   // よくないね数
  this.commentCount;   // コメント数
}

/**
 * 出力形式(チャンネルリサーチ)
 */
var ChannelResearchResult = function() {
  this.channelId;    　　// チャンネルID
  this.category;         // チャンネルカテゴリ(user or channel)
  this.title;      　　　// チャンネルタイトル
  this.viewCount;  　　　// チャンネル再生回数
  this.commentCount;     // チャンネルコメント数
  this.subscriberCount;  // チャンネル登録者数
}

/**
 * スプレッドシート初期化
 */
function initSpreadSheet() {
  _spreadSheet = SpreadsheetApp.openById('<my-spreadsheet-id>');
}

function initMovieTitleList() {
  var values = getValuesInSheet(OUTPUT_SHEET_NAMES[0], 1, 1, getNewLineRow(OUTPUT_SHEET_NAMES[0]) - 1);
  values.forEach(function(value) {
    _movieList.push(value[0]);
  });
}

function getResearchVideoIds(researchListId) {
  var results = YouTube.PlaylistItems.list('id, snippet', {
    playlistId: researchListId,
    maxResults: MAX_RESEARCH_LIST_SIZE,
  });
  var videoIds = [];
  results.items.forEach(function(item) {
    videoIds.push(item.snippet.resourceId.videoId);
  });
  return videoIds;
}

function getResearchListIds() {
  var list = [];
  var startRow = 2;
  var values = getValuesInSheet(PLAYLIST_SHEET_NAME, startRow, 1, getNewLineRow(PLAYLIST_SHEET_NAME) - startRow);
    values.forEach(function(value) {
    list.push(value[0]);
  });
  return list;
}

function research() {
  initSpreadSheet();
  initMovieTitleList();

  var output = [];
  var researchListIds = getResearchListIds();
  researchListIds.forEach(function(researchListId) {
    var videoIds = getResearchVideoIds(researchListId);
    var results = YouTube.Videos.list('snippet, statistics', {
      id: videoIds.join(','),
      maxResults: MAX_RESEARCH_LIST_SIZE,
    });

    var items = results.items;
    for (var i = 0; i < items.length; i++) {
      var data = new ResearchResult();
      data.videoId       = videoIds[i];
      data.title         = items[i].snippet.title;
      data.viewCount     = toInt(items[i].statistics.viewCount);
      data.likeCount     = toInt(items[i].statistics.likeCount);
      data.dislikeCount  = toInt(items[i].statistics.dislikeCount);
      data.commentCount  = toInt(items[i].statistics.commentCount);

      output.push(data);
    }
  });

  writeResearchResult(output);
}

function writeResearchResult(itemList) {
  var date = new Date();
  var newLineRow = getNewLineRow(OUTPUT_SHEET_NAMES[0]);
  var newLineCol = getNewLineCol(OUTPUT_SHEET_NAMES[0]);

  var getTitleRow = function(videoId) {
    return _movieList.indexOf(videoId) + 1;
  };
  var write = function(sheetName, row, col, item, value) {
    setValueInSheet(sheetName, 1, col, date);
    setValueInSheet(sheetName, row, 1, item.videoId);
    setValueInSheet(sheetName, row, 2, item.title);
    setValueInSheet(sheetName, row, col, value);
  };

  itemList.forEach(function(item) {
    var row = getTitleRow(item.videoId);
    if (row == 0) {
      row = newLineRow;
      newLineRow = newLineRow + 1;
    }
    write(OUTPUT_SHEET_NAMES[0], row, newLineCol, item, item.viewCount);
    write(OUTPUT_SHEET_NAMES[1], row, newLineCol, item, item.likeCount);
    write(OUTPUT_SHEET_NAMES[2], row, newLineCol, item, item.dislikeCount);
    write(OUTPUT_SHEET_NAMES[3], row, newLineCol, item, item.commentCount);
  });
}

function researchTrending() {
  initSpreadSheet();

  var output = [];
  var researchListId = 'PLuXL6NS58Dyztg3TS-kJVp58ziTo5Eeck';  // 急上昇プレイリスト
  var videoIds = getResearchVideoIds(researchListId);
  var results = YouTube.Videos.list('snippet, statistics', {
    id: videoIds.join(','),
    maxResults: MAX_RESEARCH_LIST_SIZE,
  });

  var items = results.items;
  for (var i = 0; i < items.length; i++) {
    var data = new ResearchResult();
    data.videoId       = videoIds[i];
    data.title         = items[i].snippet.title;
    data.viewCount     = toInt(items[i].statistics.viewCount);
    data.likeCount     = toInt(items[i].statistics.likeCount);
    data.dislikeCount  = toInt(items[i].statistics.dislikeCount);
    data.commentCount  = toInt(items[i].statistics.commentCount);

    output.push(data);
  }

  writeTrendingResearchResult(output);
}

function writeTrendingResearchResult(itemList) {
  _spreadSheet.getSheetByName(TRENDING_OUTPUT_SHEET_NAME).getRange('A3:G50').clear();

  var date = new Date();
  setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, 1, 7, date);

  var row = 3;
  for (var i = 0; i < itemList.length; i++) {
    var item = itemList[i];
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 1, i + 1);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 2, item.title);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 3, item.viewCount);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 4, item.likeCount);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 5, item.dislikeCount);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 6, item.commentCount);
    setValueInSheet(TRENDING_OUTPUT_SHEET_NAME, row, 7, 'https://www.youtube.com/watch?v=' + item.videoId);
    row = row + 1;
  }
}

function initChannelTitleList() {
  var values = getValuesInSheet(CHANNEL_OUTPUT_SHEET_NAMES[0], 1, 1, getNewLineRow(CHANNEL_OUTPUT_SHEET_NAMES[0]) - 1);
  values.forEach(function(value) {
    _channelList.push(value[0]);
  });
}

var ChannelList = function() {
  this.id;
  this.category;
};

function getResearchChannelList() {
  var list = [];
  var startRow = 2;
  var values = getValuesInSheet(CHANNELLIST_SHEET_NAME, startRow, 1, getNewLineRow(CHANNELLIST_SHEET_NAME) - startRow, 2);
  values.forEach(function(value) {
    var data = new ChannelList();
    data.id = value[0];
    data.category = value[1];
    list.push(data);
  });
  return list;
}

function researchChannel() {
  initSpreadSheet();
  initChannelTitleList();

  var output = [];
  var researchList = getResearchChannelList();

  researchList.forEach(function(channel) {
    var results;
    if (channel.category == 'user') {
      results = YouTube.Channels.list('snippet, statistics', {
        forUsername: channel.id,
      });
    } else {
      results = YouTube.Channels.list('snippet, statistics', {
        id: channel.id,
      });
    }

    var items = results.items;
    for (var i = 0; i < items.length; i++) {
      var data = new ChannelResearchResult();
      data.channelId       = channel.id;
      data.category        = channel.category;
      data.title           = items[i].snippet.title;
      data.viewCount       = toInt(items[i].statistics.viewCount);
      data.commentCount    = toInt(items[i].statistics.commentCount);
      data.subscriberCount = toInt(items[i].statistics.subscriberCount);

      output.push(data);
    }
  });

  writeChannelResearchResult(output);
}

function writeChannelResearchResult(itemList) {
  var date = new Date();
  var newLineRow = getNewLineRow(CHANNEL_OUTPUT_SHEET_NAMES[0]);
  var newLineCol = getNewLineCol(CHANNEL_OUTPUT_SHEET_NAMES[0]);

  var getTitleRow = function(channelId) {
    return _channelList.indexOf(channelId) + 1;
  };
  var write = function(sheetName, row, col, item, value) {
    var linkUrl = 'https://www.youtube.com/' + item.category + '/' + item.channelId;
    var LINK_TEXT = '=HYPERLINK("' + linkUrl + '", "LINK")';
    setValueInSheet(sheetName, 1, col, date);
    setValueInSheet(sheetName, row, 1, item.channelId);
    setValueInSheet(sheetName, row, 2, LINK_TEXT);
    setValueInSheet(sheetName, row, 3, item.title);
    setValueInSheet(sheetName, row, col, value);
  };

  itemList.forEach(function(item) {
    var row = getTitleRow(item.channelId);
    if (row == 0) {
      row = newLineRow;
      newLineRow = newLineRow + 1;
    }
    write(CHANNEL_OUTPUT_SHEET_NAMES[0], row, newLineCol, item, item.subscriberCount);
  });
}
