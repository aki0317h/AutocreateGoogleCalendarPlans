function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📅 カスタムメニュー')
    .addItem('予定をカレンダーに追加', 'addEventsFromSheet')
    .addToUi();
}

function addEventsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const calendar = CalendarApp.getCalendarById(''); // カレンダーID

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const rawDate = row[0];        // A列：日付
    const rawStartTime = row[1];   // B列：開始時間
    const rawEndTime = row[2];     // C列：終了時間
    const title = row[3];          // D列：タイトル
    const location = row[4] || ''; // E列：場所（任意）
    const description = row[5] || ''; // F列：メモ（任意）

    // 必須項目がなければスキップ
    if (!rawDate || !rawStartTime || !rawEndTime || !title) continue;

    // 日付オブジェクトを安全に取得
    const dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) {
      Logger.log(`不正な日付形式: ${rawDate}`);
      continue;
    }

    // 開始・終了時刻もDateとして扱い、"HH:mm" に整形
    const startTimeStr = Utilities.formatDate(new Date(rawStartTime), 'Asia/Tokyo', 'HH:mm');
    const endTimeStr = Utilities.formatDate(new Date(rawEndTime), 'Asia/Tokyo', 'HH:mm');

    // 日付を "yyyy-MM-dd" に整形
    const dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd');

    // "yyyy-MM-ddTHH:mm" → ISO形式で Date オブジェクト化
    const startDateTime = new Date(`${dateStr}T${startTimeStr}`);
    const endDateTime = new Date(`${dateStr}T${endTimeStr}`);

    // 無効な日付ならスキップ
    if (isNaN(startDateTime.getTime()) || isNaN(endDateTime.getTime())) {
      Logger.log(`${title} Invalid Date`);
      continue;
    }

    // 🔁 仕事イベントなら前後に移動時間を追加
    if (title === "仕事") {
      const moveTitle = "移動";
      const moveStartTime = new Date(startDateTime.getTime() - 30 * 60 * 1000);
      const moveEndTime = new Date(endDateTime.getTime() + 30 * 60 * 1000);

      calendar.createEvent(moveTitle, moveStartTime, startDateTime, {
        location: location,
        description: description
      });

      calendar.createEvent(moveTitle, endDateTime, moveEndTime, {
        location: location,
        description: description
      });
    }

    // 🔁 本体イベント（仕事含むすべて）
    calendar.createEvent(title, startDateTime, endDateTime, {
      location: location,
      description: description
    });

    // ✅ 実行ログ（任意）
    Logger.log(`追加: ${title} ${startDateTime.toLocaleString()} ～ ${endDateTime.toLocaleString()}`);

    // ✅ 追加後、該当行（A〜F列）をクリア
    sheet.getRange(i + 1, 1, 1, 6).clearContent();
  }
}
