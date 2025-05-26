/**
 * ICSファイルからGoogleカレンダーにイベントをインポートするGAS
 */

// 設定
const CONFIG = {
  ICS_FILE_ID: 'hoge',                    // Googleドライブ内のICSファイルID
  CALENDAR_ID: 'example@gmail.com',       // カレンダーID
  ERROR_SHEET_NAME: 'ICS_Import_Errors',  // エラー記録シート名
  BATCH_SIZE: 10,                         // 一度に処理する件数
  SLEEP_INTERVAL: 1000                    // API呼び出し間隔（ミリ秒）
};

/**
 * メイン実行関数
 */
function importIcsToCalendar() {
  try {
    Logger.log('=== ICSインポート開始 ===');

    // 1. ICSファイルを読み込み・解析
    const events = loadAndParseIcsFile();
    Logger.log(`ICSファイルから ${events.length} 件のイベントを取得`);

    // 2. 重複を除去（ICSファイル内の重複）
    const uniqueEvents = removeDuplicateEvents(events);
    Logger.log(`重複除去後: ${uniqueEvents.length} 件のユニークイベント`);

    // 3. 既存カレンダーイベントを取得
    const existingEvents = getExistingCalendarEvents(uniqueEvents);
    Logger.log(`カレンダーから ${existingEvents.size} 件の既存イベントを取得`);

    // 4. 登録が必要なイベントをフィルタリング
    const eventsToCreate = filterEventsToCreate(uniqueEvents, existingEvents);
    Logger.log(`登録対象: ${eventsToCreate.length} 件`);

    // 5. 前回のエラーイベントを取得・追加
    const errorEvents = getErrorEvents();
    const allEventsToProcess = [...eventsToCreate, ...errorEvents];
    Logger.log(`エラー再試行を含む総処理対象: ${allEventsToProcess.length} 件`);

    // 6. バッチ処理でイベントを作成
    const result = processEventsInBatches(allEventsToProcess);

    // 7. 結果を出力
    logFinalResults(result);

  } catch (error) {
    Logger.log(`致命的エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * ICSファイルの読み込みと解析
 */
function loadAndParseIcsFile() {
  const file = DriveApp.getFileById(CONFIG.ICS_FILE_ID);
  const content = file.getBlob().getDataAsString();
  return parseIcsContent(content);
}

/**
 * ICSコンテンツの解析
 */
function parseIcsContent(content) {
  const events = [];
  const lines = content.split(/\r?\n/);
  let currentEvent = null;
  let inEvent = false;

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();

    // 行の継続処理
    while (i + 1 < lines.length && (lines[i + 1].startsWith(' ') || lines[i + 1].startsWith('\t'))) {
      line += lines[++i].substring(1);
    }

    if (line === 'BEGIN:VEVENT') {
      currentEvent = {};
      inEvent = true;
    }
    else if (line === 'END:VEVENT' && inEvent) {
      if (currentEvent?.title && currentEvent?.start) {
        events.push(currentEvent);
      }
      currentEvent = null;
      inEvent = false;
    }
    else if (inEvent && currentEvent) {
      parseEventProperty(line, currentEvent);
    }
  }

  return events;
}

/**
 * イベントプロパティの解析
 */
function parseEventProperty(line, event) {
  if (line.startsWith('SUMMARY:')) {
    event.title = line.substring(8).trim();
  }
  else if (line.startsWith('DTSTART;VALUE=DATE:')) {
    event.isAllDay = true;
    event.start = parseIcsDate(line.substring(19));
  }
  else if (line.startsWith('DTSTART:')) {
    event.isAllDay = false;
    event.start = parseIcsDateTime(line.substring(8));
  }
  else if (line.startsWith('DTEND;VALUE=DATE:')) {
    event.end = parseIcsDate(line.substring(17));
  }
  else if (line.startsWith('DTEND:')) {
    event.end = parseIcsDateTime(line.substring(6));
  }
  else if (line.startsWith('DESCRIPTION:')) {
    event.description = line.substring(12);
  }
  else if (line.startsWith('LOCATION:')) {
    event.location = line.substring(9);
  }
}

/**
 * ICS日付解析（YYYYMMDD）
 */
function parseIcsDate(dateStr) {
  const year = parseInt(dateStr.substring(0, 4));
  const month = parseInt(dateStr.substring(4, 6)) - 1;
  const day = parseInt(dateStr.substring(6, 8));
  return new Date(year, month, day);
}

/**
 * ICS日時解析（YYYYMMDDTHHMMSSZ）
 */
function parseIcsDateTime(dateTimeStr) {
  const [datePart, timePart] = dateTimeStr.split('T');

  const year = parseInt(datePart.substring(0, 4));
  const month = parseInt(datePart.substring(4, 6)) - 1;
  const day = parseInt(datePart.substring(6, 8));

  const hour = parseInt(timePart.substring(0, 2));
  const minute = parseInt(timePart.substring(2, 4));
  const second = parseInt(timePart.substring(4, 6));

  if (dateTimeStr.endsWith('Z')) {
    return new Date(Date.UTC(year, month, day, hour, minute, second));
  } else {
    return new Date(year, month, day, hour, minute, second);
  }
}

/**
 * ICSファイル内の重複イベントを除去
 * 同日・同タイトルのイベントは1つだけ残す
 */
function removeDuplicateEvents(events) {
  const uniqueEvents = [];
  const seenKeys = new Set();

  events.forEach(event => {
    const key = generateEventKey(event);
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      uniqueEvents.push(event);
    }
  });

  return uniqueEvents;
}

/**
 * イベントの一意キーを生成
 */
function generateEventKey(event) {
  const date = formatDateKey(event.start);
  const title = event.title.toLowerCase().trim();
  return `${date}|${title}`;
}

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 */
function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 既存のカレンダーイベントを取得
 */
function getExistingCalendarEvents(events) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const existingKeys = new Set();

  // イベントの日付範囲を取得
  const dateRange = getEventDateRange(events);
  if (!dateRange.start || !dateRange.end) {
    return existingKeys;
  }

  Logger.log(`既存イベント取得: ${formatDateKey(dateRange.start)} ～ ${formatDateKey(dateRange.end)}`);

  // 期間内の既存イベントを取得
  const existingEvents = calendar.getEvents(dateRange.start, dateRange.end);

  existingEvents.forEach(event => {
    const eventDate = event.isAllDayEvent() ? event.getAllDayStartDate() : event.getStartTime();
    const key = `${formatDateKey(eventDate)}|${event.getTitle().toLowerCase().trim()}`;
    existingKeys.add(key);
  });

  return existingKeys;
}

/**
 * イベントの日付範囲を取得
 */
function getEventDateRange(events) {
  if (events.length === 0) return { start: null, end: null };

  let minDate = new Date(events[0].start);
  let maxDate = new Date(events[0].start);

  events.forEach(event => {
    const eventDate = new Date(event.start);
    if (eventDate < minDate) minDate = eventDate;
    if (eventDate > maxDate) maxDate = eventDate;
  });

  // 範囲を少し広げる
  const start = new Date(minDate);
  start.setDate(start.getDate() - 1);
  start.setHours(0, 0, 0, 0);

  const end = new Date(maxDate);
  end.setDate(end.getDate() + 1);
  end.setHours(23, 59, 59, 999);

  return { start, end };
}

/**
 * 登録が必要なイベントをフィルタリング
 */
function filterEventsToCreate(events, existingKeys) {
  return events.filter(event => {
    const key = generateEventKey(event);
    return !existingKeys.has(key);
  });
}

/**
 * エラーシートから前回のエラーイベントを取得
 */
function getErrorEvents() {
  try {
    const sheet = getOrCreateErrorSheet();
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return []; // ヘッダーのみ

    const errorEvents = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) { // タイトルと日付がある
        try {
          const event = {
            title: row[0],
            start: new Date(row[1]),
            isAllDay: row[3] === 'TRUE',
            description: row[4] || '',
            location: row[5] || '',
            rowIndex: i + 1 // スプレッドシートの行番号
          };

          if (row[2]) { // 終了日時がある場合
            event.end = new Date(row[2]);
          }

          errorEvents.push(event);
        } catch (parseError) {
          Logger.log(`エラーイベント解析失敗 行${i + 1}: ${parseError.toString()}`);
        }
      }
    }

    Logger.log(`エラーシートから ${errorEvents.length} 件の再試行対象を取得`);
    return errorEvents;

  } catch (error) {
    Logger.log(`エラーイベント取得エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * エラー記録シートの取得または作成
 */
function getOrCreateErrorSheet() {
  let spreadsheet;

  // 既存のスプレッドシートを検索
  const files = DriveApp.getFilesByName(CONFIG.ERROR_SHEET_NAME);
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.openById(files.next().getId());
  } else {
    // 新規作成
    spreadsheet = SpreadsheetApp.create(CONFIG.ERROR_SHEET_NAME);
  }

  const sheet = spreadsheet.getActiveSheet();

  // ヘッダーの設定（初回のみ）
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([[
      'Title', 'Start Date', 'End Date', 'Is All Day', 'Description', 'Location', 'Error Message'
    ]]);
  }

  return sheet;
}

/**
 * バッチ処理でイベントを作成
 */
function processEventsInBatches(events) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const result = { created: 0, skipped: 0, errors: 0 };
  const errorSheet = getOrCreateErrorSheet();

  // エラーシートをクリア（再試行のため）
  if (errorSheet.getLastRow() > 1) {
    errorSheet.getRange(2, 1, errorSheet.getLastRow() - 1, 7).clear();
  }

  for (let i = 0; i < events.length; i += CONFIG.BATCH_SIZE) {
    const batch = events.slice(i, i + CONFIG.BATCH_SIZE);
    Logger.log(`バッチ処理: ${i + 1} ～ ${Math.min(i + CONFIG.BATCH_SIZE, events.length)} / ${events.length}`);

    batch.forEach(event => {
      try {
        // イベント作成
        createCalendarEvent(calendar, event);
        result.created++;

        Logger.log(`✅ 作成成功: ${event.title} (${formatDateKey(event.start)})`);

      } catch (error) {
        result.errors++;

        Logger.log(`❌ 作成エラー: ${event.title} - ${error.toString()}`);

        // エラーをスプレッドシートに記録
        recordError(errorSheet, event, error.toString());
      }
    });

    // バッチ間のスリープ
    if (i + CONFIG.BATCH_SIZE < events.length) {
      Utilities.sleep(CONFIG.SLEEP_INTERVAL);
    }
  }

  return result;
}

/**
 * カレンダーイベントの作成
 */
function createCalendarEvent(calendar, event) {
  const options = {
    description: event.description || '',
    location: event.location || ''
  };

  if (event.isAllDay) {
    const endDate = event.end || new Date(event.start.getTime() + 24 * 60 * 60 * 1000);
    calendar.createAllDayEvent(event.title, event.start, endDate, options);
  } else {
    const endTime = event.end || new Date(event.start.getTime() + 60 * 60 * 1000);
    calendar.createEvent(event.title, event.start, endTime, options);
  }
}

/**
 * エラーをスプレッドシートに記録
 */
function recordError(sheet, event, errorMessage) {
  try {
    const row = [
      event.title,
      event.start,
      event.end || '',
      event.isAllDay ? 'TRUE' : 'FALSE',
      event.description || '',
      event.location || '',
      errorMessage
    ];

    sheet.appendRow(row);

  } catch (error) {
    Logger.log(`エラー記録失敗: ${error.toString()}`);
  }
}

/**
 * 最終結果のログ出力
 */
function logFinalResults(result) {
  Logger.log('========== 処理完了 ==========');
  Logger.log(`✅ 作成成功: ${result.created} 件`);
  Logger.log(`⏭️ スキップ: ${result.skipped} 件`);
  Logger.log(`❌ エラー: ${result.errors} 件`);

  const total = result.created + result.skipped + result.errors;
  if (total > 0) {
    const successRate = ((result.created / total) * 100).toFixed(1);
    Logger.log(`📊 成功率: ${successRate}%`);
  }

  if (result.errors > 0) {
    Logger.log(`⚠️ エラーが発生しました。${CONFIG.ERROR_SHEET_NAME} シートを確認してください。`);
    Logger.log('再実行すると、エラーイベントの登録を再試行します。');
  }
}

/**
 * エラーシートの内容を確認する関数（デバッグ用）
 */
function checkErrorSheet() {
  try {
    const sheet = getOrCreateErrorSheet();
    const data = sheet.getDataRange().getValues();

    Logger.log('=== エラーシート内容 ===');
    if (data.length <= 1) {
      Logger.log('エラーはありません');
      return;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`${i}. ${row[0]} (${row[1]}) - ${row[6]}`);
    }

    Logger.log(`総エラー件数: ${data.length - 1}`);

  } catch (error) {
    Logger.log(`エラーシート確認エラー: ${error.toString()}`);
  }
}

/**
 * エラーシートをクリアする関数（手動実行用）
 */
function clearErrorSheet() {
  try {
    const sheet = getOrCreateErrorSheet();

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clear();
      Logger.log('エラーシートをクリアしました');
    } else {
      Logger.log('エラーシートは既に空です');
    }

  } catch (error) {
    Logger.log(`エラーシートクリアエラー: ${error.toString()}`);
  }
}

/**
 * 特定の日付範囲のイベントを確認する関数（デバッグ用）
 */
function checkEventsInDateRange(startDateStr, endDateStr) {
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);

    const events = calendar.getEvents(startDate, endDate);

    Logger.log(`=== ${startDateStr} ～ ${endDateStr} のイベント ===`);
    Logger.log(`総件数: ${events.length}`);

    const eventsByDate = {};

    events.forEach(event => {
      const eventDate = event.isAllDayEvent() ?
        formatDateKey(event.getAllDayStartDate()) :
        formatDateKey(event.getStartTime());

      if (!eventsByDate[eventDate]) {
        eventsByDate[eventDate] = [];
      }

      eventsByDate[eventDate].push({
        title: event.getTitle(),
        isAllDay: event.isAllDayEvent()
      });
    });

    // 日付順に表示
    Object.keys(eventsByDate).sort().forEach(date => {
      Logger.log(`\n📅 ${date}:`);
      eventsByDate[date].forEach(event => {
        const type = event.isAllDay ? '[終日]' : '[時間指定]';
        Logger.log(`  ${type} ${event.title}`);
      });
    });

  } catch (error) {
    Logger.log(`イベント確認エラー: ${error.toString()}`);
  }
}

/**
 * ICSファイルの内容を確認する関数（デバッグ用）
 */
function checkIcsFileContent() {
  try {
    const events = loadAndParseIcsFile();

    Logger.log('=== ICSファイル内容 ===');
    Logger.log(`総イベント数: ${events.length}`);

    // 日付別にグループ化
    const eventsByDate = {};

    events.forEach(event => {
      const date = formatDateKey(event.start);
      if (!eventsByDate[date]) {
        eventsByDate[date] = [];
      }
      eventsByDate[date].push(event);
    });

    // 日付順に表示
    Object.keys(eventsByDate).sort().forEach(date => {
      Logger.log(`\n📅 ${date} (${eventsByDate[date].length}件):`);
      eventsByDate[date].forEach(event => {
        const type = event.isAllDay ? '[終日]' : '[時間指定]';
        Logger.log(`  ${type} ${event.title}`);
      });
    });

    // 重複チェック
    const uniqueEvents = removeDuplicateEvents(events);
    const duplicateCount = events.length - uniqueEvents.length;

    if (duplicateCount > 0) {
      Logger.log(`\n⚠️ ICSファイル内重複: ${duplicateCount} 件`);
    }

  } catch (error) {
    Logger.log(`ICSファイル確認エラー: ${error.toString()}`);
  }
}

/**
 * 設定確認関数
 */
function checkConfiguration() {
  Logger.log('=== 設定確認 ===');
  Logger.log(`ICSファイルID: ${CONFIG.ICS_FILE_ID}`);
  Logger.log(`カレンダーID: ${CONFIG.CALENDAR_ID}`);
  Logger.log(`エラーシート名: ${CONFIG.ERROR_SHEET_NAME}`);
  Logger.log(`バッチサイズ: ${CONFIG.BATCH_SIZE}`);
  Logger.log(`スリープ間隔: ${CONFIG.SLEEP_INTERVAL}ms`);

  // ファイルアクセス確認
  try {
    const file = DriveApp.getFileById(CONFIG.ICS_FILE_ID);
    Logger.log(`✅ ICSファイル: ${file.getName()}`);
  } catch (error) {
    Logger.log(`❌ ICSファイルアクセスエラー: ${error.toString()}`);
  }

  // カレンダーアクセス確認
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    Logger.log(`✅ カレンダー: ${calendar.getName()}`);
  } catch (error) {
    Logger.log(`❌ カレンダーアクセスエラー: ${error.toString()}`);
  }
}

/**
 * 統計情報を表示する関数
 */
function showStatistics() {
  try {
    Logger.log('=== 統計情報 ===');

    // ICSファイルの統計
    const events = loadAndParseIcsFile();
    const uniqueEvents = removeDuplicateEvents(events);

    Logger.log(`📄 ICSファイル:`);
    Logger.log(`  - 総イベント数: ${events.length}`);
    Logger.log(`  - ユニークイベント数: ${uniqueEvents.length}`);
    Logger.log(`  - 重複数: ${events.length - uniqueEvents.length}`);

    // カレンダーの統計
    const existingEvents = getExistingCalendarEvents(uniqueEvents);
    const eventsToCreate = filterEventsToCreate(uniqueEvents, existingEvents);

    Logger.log(`📅 カレンダー:`);
    Logger.log(`  - 既存イベント数: ${existingEvents.size}`);
    Logger.log(`  - 新規登録対象: ${eventsToCreate.length}`);

    // エラーの統計
    const errorEvents = getErrorEvents();
    Logger.log(`❌ エラー:`);
    Logger.log(`  - 前回エラー件数: ${errorEvents.length}`);

    Logger.log(`\n📊 処理予定:`);
    Logger.log(`  - 総処理件数: ${eventsToCreate.length + errorEvents.length}`);
    Logger.log(`  - 推定処理時間: ${Math.ceil((eventsToCreate.length + errorEvents.length) / CONFIG.BATCH_SIZE * (CONFIG.SLEEP_INTERVAL / 1000))} 秒`);

  } catch (error) {
    Logger.log(`統計情報取得エラー: ${error.toString()}`);
  }
}

/**
 * テスト実行関数（少数のイベントで動作確認）
 */
function testImport() {
  try {
    Logger.log('=== テスト実行開始 ===');

    // 設定を一時的に変更
    const originalBatchSize = CONFIG.BATCH_SIZE;
    CONFIG.BATCH_SIZE = 3; // テストでは3件ずつ処理

    // ICSファイルから最初の5件だけ取得
    const allEvents = loadAndParseIcsFile();
    const testEvents = allEvents.slice(0, 5);

    Logger.log(`テスト対象: ${testEvents.length} 件`);

    // 重複除去
    const uniqueEvents = removeDuplicateEvents(testEvents);

    // 既存イベントチェック
    const existingEvents = getExistingCalendarEvents(uniqueEvents);
    const eventsToCreate = filterEventsToCreate(uniqueEvents, existingEvents);

    Logger.log(`テスト登録対象: ${eventsToCreate.length} 件`);

    // 処理実行
    if (eventsToCreate.length > 0) {
      const result = processEventsInBatches(eventsToCreate);
      logFinalResults(result);
    } else {
      Logger.log('テスト登録対象がありません');
    }

    // 設定を元に戻す
    CONFIG.BATCH_SIZE = originalBatchSize;

    Logger.log('=== テスト実行完了 ===');

  } catch (error) {
    Logger.log(`テスト実行エラー: ${error.toString()}`);
  }
}