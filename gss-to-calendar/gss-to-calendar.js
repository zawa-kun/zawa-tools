// ========================================================================
// Googleスプレッドシート & Googleカレンダー連携設定
// ========================================================================
const CONFIG = {
  CALENDAR_ID: 'primary',
  SHEET_NAME: 'kokochan',
  // カラム番号（A列=1, B列=2...）
  COLUMNS: {
    COMPANY: 2,    //  企業名
    START: 3,      //  開始日時
    END: 4,        //  終了日時
    STATUS: 5,     //  選考状況
    DESC: 6,       //  選考詳細
    LOCATION: 7,   //  選考会場
    EVENT_ID: 12,   // Googleカレンダーイベントの一意ID（新規か既存か確認用）
    PREV_STATUS: 13 //  前回ステータス(選考移行検知用)
  },

  // カレンダー登録対象となる選考フェーズ
  TRIGGER_PHASES: [
    '説明会', 'ES', 'WEBテスト', 'ES＆WEBテスト', 
    '1次面接', '2次面接', '3次面接', '最終面接'
  ]
};


// ========================================================================
// メイン処理:  スプレッドシート編集時にカレンダーへ同期
// ========================================================================
/**
 * スプレッドシート編集イベントをトリガーにカレンダーと同期
 * - 新規作成: イベントIDが無い場合
 * - 選考移行: ステータスが変更された場合（古いイベントは残す）
 * - 内容更新: 同一ステータス内での時間・場所等の変更
 * 
 * @param {Object} e - onEdit トリガーイベントオブジェクト
 */
function syncToCalendar(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const range = e.range;
  const row = range.getRow();
  console.log(row); // debug

  // -----------------------------------------------------------------------
  // 1. 早期リターン: 処理対象外の編集をフィルタリング
  // -----------------------------------------------------------------------
  
  // ヘッダー行（1行目）や対象外シートの編集は無視
  if (row < 2 || range.getSheet().getName() !== CONFIG.SHEET_NAME) return;
  

  // -----------------------------------------------------------------------
  // 2. データ取得 & バリデーション
  // -----------------------------------------------------------------------

  const maxCol = Math.max(...Object.values(CONFIG.COLUMNS));
  const rowData = sheet.getRange(row, 1, 1, maxCol).getValues()[0];
  const eventData = {
    company:     rowData[CONFIG.COLUMNS.COMPANY - 1],
    start:      rowData[CONFIG.COLUMNS.START - 1],
    end:        rowData[CONFIG.COLUMNS.END - 1],
    location:   rowData[CONFIG.COLUMNS. LOCATION - 1] || '', 
    desc:       rowData[CONFIG.COLUMNS.DESC - 1] || '',
    status:     rowData[CONFIG.COLUMNS.STATUS - 1],
    eventId:    rowData[CONFIG.COLUMNS.EVENT_ID - 1],
    prevStatus: rowData[CONFIG.COLUMNS.PREV_STATUS - 1]
  };
  console.log(eventData); // debug

  // 必須フィールドの存在確認
  if (!eventData.company || !eventData.start || !eventData.status) return;
  
  // 日付型の確認
  if (!(eventData.start instanceof Date) || isNaN(eventData.start)) {
    console.log(`行${row}:  開始日時が無効な形式です`);
    return;
  }

  // 終了日時が存在する場合は型チェック
  if (eventData.end && (!(eventData.end instanceof Date) || isNaN(eventData.end))) {
    console.log(`行${row}: 終了日時が無効な形式です`);
    return;
  }

  // 登録対象の選考フェーズかチェック
  if (!CONFIG.TRIGGER_PHASES.includes(eventData.status)) return;

  // イベント名の作成
  let eventTitle = "";
  if(eventData.location == ""){
    // 場所無し：[選考状況]社名
    eventTitle = `[${eventData.status}]${eventData.company}`;

  } else if(eventData.location == "オンライン") { // オンラインのみは[オ]に省略
    // オンライン：[オ][選考状況]社名
    if(eventData.status == "説明会") { 
      eventTitle = `[オ][説]${eventData.company}`;  // 説明会のみは[説]に省略
    }else{
      eventTitle = `[オ][${eventData.status}]${eventData.company}`
    }

  } else {
    // 場所無しでもオンラインでもない
    // [対面][選考状況]社名
    if(eventData.status == "説明会") { 
      eventTitle = `[対面][説]${eventData.company}`;
    }else{
      eventTitle = `[対面][${eventData.status}]${eventData.company}`
    }

  }
  console.log(eventTitle); // debug

  // -----------------------------------------------------------------------
  // 3. カレンダーイベントの生存確認
  // -----------------------------------------------------------------------
  
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  
  var existingEvent = null; // カレンダー側にスプレッドシートから取得したEventIDがあれば
  if(eventData.eventId){
    try {
      existingEvent = calendar.getEventById(eventData.eventId);
    } catch {
      console.log(`イベントID ${eventData.eventId} がカレンダーに見つかりません:  ${error.message}`);
      existingEvent = null;
    }
  }
  
  // -----------------------------------------------------------------------
  // 4. モード判定 & 処理実行
  // -----------------------------------------------------------------------

  if (!existingEvent || !eventData.eventId) {
    // ケースA: 新規作成
    // - イベントIDが未登録
    // - または、IDはあるがカレンダーから削除されている
    createEvent(sheet, row, calendar, eventTitle, eventData);
    
  } else if (eventData.status !== eventData.prevStatus) {
    // ケースB:  選考移行（新規イベント作成）
    // - 前回と異なる選考ステータス → 古いイベントはそのまま残す
    createEvent(sheet, row, calendar, eventTitle, eventData);
    
  } else {
    // ケースC:  既存イベントの更新
    // - 同一ステータス内での時間・場所・詳細の変更
    updateEvent(existingEvent, eventTitle, eventData);
    
    // 念のため前回ステータスを同期（通常は変更なしだが堅牢性向上のため）
    sheet.getRange(row, CONFIG.COLUMNS. PREV_STATUS).setValue(eventData.status);
  }
}

// ==========================================
// イベント作成関数
// ==========================================
function createEvent(sheet, row, calendar, title, data) {
  let newEvent;
  
  // 終日判定: 開始が0:00 または 終了日時がない
  const isAllDay = isMidNight(data.start) || !data.end;

  const options = {
    location: data.location,
    description: data.desc
  };

  if (isAllDay) {
    newEvent = calendar.createAllDayEvent(title, data.start, options);
  } else {
    // 終了時間が未定なら開始+1時間
    const endTime = data.end instanceof Date ? data.end : new Date(data.start.getTime() + 60 * 60 * 1000);
    newEvent = calendar.createEvent(title, data.start, endTime, options);
  }

  // 通知設定 (1時間前)
  newEvent.addPopupReminder(60);

  // IDとステータスをシートに書き戻す
  sheet.getRange(row, CONFIG.COLUMNS.EVENT_ID).setValue(newEvent.getId());
  sheet.getRange(row, CONFIG.COLUMNS.PREV_STATUS).setValue(data.status);
}

// ==========================================
// イベント更新関数
// ==========================================
function updateEvent(event, title, data) {
  event.setTitle(title);
  event.setLocation(data.location);
  event.setDescription(data.desc);

  // 日時の更新
  const isAllDay = isMidNight(data.start) || !data.end;
  
  if (isAllDay) {
    // setAllDayDateを使うと既存が時間指定でも終日に変更される
    event.setAllDayDate(data.start);
  } else {
    const endTime = data.end instanceof Date ? data.end : new Date(data.start.getTime() + 60 * 60 * 1000);
    // 時間指定に変更（setTimeは開始と終了を同時にセット）
    event.setTime(data.start, endTime);
  }
}

// ==========================================
// ユーティリティ
// ==========================================
// 時間が00:00:00かどうか判定
function isMidNight(dateObj) {
  if (!(dateObj instanceof Date)) return false;
  return dateObj.getHours() === 0 && dateObj.getMinutes() === 0;
}