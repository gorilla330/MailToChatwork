/**
 * ============ 設定 ============
 * 1) Chatwork: ROOM_ID, TO_IDS_CSV（担当者のaccount_idをカンマ区切り。未指定なら未アサイン）
 * 2) Spreadsheet: ログ＆サマリーを書き込む先のシート
 * 3) スクリプトプロパティ: CHATWORK_TOKEN を保存
 */
const CONFIG = {
  INFO_ADDRESS: 'info@ascope.net',
  // 件名条件：完全一致(お問い合わせがありました) or 含有一致(＜至急＞<B肝新規相談会申込>)
  SUBJECT_HOOK: /^(お問い合わせがありました)$|＜至急＞<B肝新規相談会申込>/,
  SUBJECT_QUERIES: ['お問い合わせがありました', '＜至急＞<B肝新規相談会申込>'], // Gmail検索（OR）
  URGENT_KEYWORD: '至急',                        // 至急判定キーワード（件名/本文）
  DONE_LABEL: 'CWタスク化済',                 // Gmailの重複防止ラベル
  // 投稿先（テスト/本番）
  MODE: 'prod',                               // 'test' | 'prod'
  TEST_ROOM_ID: '5564141',                    // マイチャット
  PROD_ROOM_ID: '18848369',                   // 全体スレ
  TO_IDS_CSV: '<<11111,22222>>',             // 例: "11111,22222"。通知では未使用（必要なら本文先頭に宛先文言を追加可）
  TOKEN_PROP_KEY: 'CHATWORK_TOKEN',
  TITLE_LINE: '新規問い合わせがあります！',
  INCLUDE_FULL_BODY: true,
  MAX_BODY_LEN: 8000,
  DUE: { enable: true, daysFromNow: 2, hour: 18, minute: 0, tz: 'Asia/Tokyo' },
  // 検索期間（日）: 時間主導トリガーによるポーリング用。短いほど高速。
  SEARCH_DAYS: 2,
  // カテゴリ→投稿先ルームID（本番）
  ROOM_MAP: {
    general: '33559798',     // 一般案件管理スレ（トップページ含む）
    labor: '204109750',      // 労働LP+支店ルート三田希望
    bkan: '23509676',        // B肝スレ
    asbestos: '172962669',   // アスベスト案件・全体スレ
  },

  // ===== カテゴリ振り分け（件名で判定）=====
  CATEGORIES: [
    { name: '労働',        pattern: /労働[ＬL]Ｐ|労働/ },
    { name: 'アスベスト',  pattern: /アスベスト|石綿/ },
    { name: 'B型肝炎',      pattern: /B型肝炎訴訟|B型肝炎給付金|B型肝炎|B肝/ },
    { name: 'トップページ',  pattern: /ascope\.net・トップページ|トップページ/ },
  ],
  DEFAULT_CATEGORY: '一般案件',

  // ===== スプレッドシート =====
  SPREADSHEET_ID: '1cb6AGC7YOZNeu0orwYC-3N38_bzgSneEQ_nmL-p8Tqs',
  SHEET_LOG: 'log',          // データログ
  SHEET_SUMMARY: 'summary',  // 集計
};

/** メイン：対象メール（メッセージ単位）→Chatworkタスク作成→ログ記録→ラベル付与→サマリー更新 */
function createTasksWithLogging() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet = ensureLogSheet_(ss);
  const summarySheet = ensureSummarySheet_(ss);

  // 既処理の messageId セット（スプシ側ガード）
  const processedMsgIds = buildProcessedMsgSet_(logSheet);

  // Gmail検索（期間で絞る。ラベル除外はしない＝同一スレッドの新着も拾う）
  const doneLabel = GmailApp.getUserLabelByName(CONFIG.DONE_LABEL) || GmailApp.createLabel(CONFIG.DONE_LABEL);
  // 件名は OR 条件で検索
  const subjectOr = (CONFIG.SUBJECT_QUERIES || []).map(q => `"${q}"`).join(' OR ');
  const gmailSearchQuery = `deliveredto:${CONFIG.INFO_ADDRESS} newer_than:${CONFIG.SEARCH_DAYS}d subject:(${subjectOr})`;
  const threads = GmailApp.search(gmailSearchQuery);
  // 期間しきい値（メッセージ日時）
  const thresholdDate = new Date(Date.now() - CONFIG.SEARCH_DAYS * 24 * 3600 * 1000);

  let newLogRows = [];
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(msg => {
      try {
        // 期間外のメッセージはスキップ
        const msgDate = msg.getDate && msg.getDate();
        if (msgDate && msgDate < thresholdDate) return;

        const subject = msg.getSubject() || '';
        if (!CONFIG.SUBJECT_HOOK.test(subject)) return;

        const gmailMessageId = String(msg.getId());
        if (processedMsgIds.has(gmailMessageId)) return;

        const threadId = String(thread.getId());
        const category = detectCategory_(subject);
        const plainBodyForDetect = (msg.getPlainBody && msg.getPlainBody()) || msg.getBody() || '';
        const isUrgent = (subject.indexOf(CONFIG.URGENT_KEYWORD) !== -1) || (plainBodyForDetect.indexOf(CONFIG.URGENT_KEYWORD) !== -1);

        const body = isUrgent ? buildUrgentBody_(msg) : buildTaskBody_(msg);
        const dueUnix = CONFIG.DUE.enable ? computeDueUnix_(CONFIG.DUE) : '';

        // --- Chatwork 通知（メッセージ投稿）---
        // 1) すべてマイチャットへ
        const myChatMessageId = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
        // 2) 本番モードのときはカテゴリ対応ルームへも通知
        let chatworkMessageId = myChatMessageId;
        if (CONFIG.MODE === 'prod') {
          const destRoomId = isUrgent ? CONFIG.PROD_ROOM_ID : mapCategoryToRoomId_(category);
          const postedId = cwPostMessage_(destRoomId, body);
          if (postedId) chatworkMessageId = postedId;
        }

        // --- ログ行を追加 ---
        const now = new Date();
        newLogRows.push([
          now,                 // A: processed_at
          threadId,            // B: gmail_thread_id
          gmailMessageId,      // C: gmail_message_id（重複判定のキー）
          subject,             // D: subject
          msg.getFrom() || '', // E: from
          category,            // F: category
          chatworkMessageId || '', // G: chatwork_message_id
          dueUnix || '',       // H: due_unix
          'created'            // I: status
        ]);

        // --- ラベル（視認用。GmailAppではスレッドに付与）---
        try {
          thread.addLabel(doneLabel);
        } catch (labelErr) {
          console.warn('Failed to add label to thread:', labelErr);
        }

      } catch (err) {
        console.error(err);
        GmailApp.sendEmail(Session.getActiveUser().getEmail(), '【GAS】Chatworkタスク作成エラー', String(err));
      }
    });
  });

  // まとめてログ追加（1回のappendで高速化）
  if (newLogRows.length) logSheet.getRange(logSheet.getLastRow()+1, 1, newLogRows.length, newLogRows[0].length).setValues(newLogRows);

  // 集計を更新
  updateSummary_(logSheet, summarySheet);
}

/** Chatwork：メッセージ投稿 → message_id を返す */
function cwPostMessage_(roomId, body) {
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.TOKEN_PROP_KEY);
  if (!token) throw new Error(`Script Properties に ${CONFIG.TOKEN_PROP_KEY} が未設定です。`);

  const payload = { body };
  const res = UrlFetchApp.fetch(`https://api.chatwork.com/v2/rooms/${roomId}/messages`, {
    method: 'post',
    headers: { 'X-ChatWorkToken': token },
    payload,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(`Chatworkメッセージ投稿エラー ${code}: ${res.getContentText()}`);

  // 期待レス例: {"message_id":"1234567890"}
  try {
    const json = JSON.parse(res.getContentText());
    const id = json.message_id ? String(json.message_id) : '';
    return id;
  } catch (_) {
    return '';
  }
}

/** 現在のモードに応じて投稿先ルームIDを返す */
function getActiveRoomId_() {
  return CONFIG.MODE === 'prod' ? CONFIG.PROD_ROOM_ID : CONFIG.TEST_ROOM_ID;
}

/** カテゴリ名→本番の投稿先ルームID */
function mapCategoryToRoomId_(categoryName) {
  if (categoryName === '労働') return CONFIG.ROOM_MAP.labor;
  if (categoryName === 'アスベスト') return CONFIG.ROOM_MAP.asbestos;
  if (categoryName === 'B型肝炎') return CONFIG.ROOM_MAP.bkan;
  // トップページやその他は一般案件へ
  return CONFIG.ROOM_MAP.general;
}

/** 件名→カテゴリ名 */
function detectCategory_(subject) {
  for (const c of CONFIG.CATEGORIES) if (c.pattern.test(subject)) return c.name;
  return CONFIG.DEFAULT_CATEGORY;
}

/** 通知本文（Chatworkの [info][title]... 形式） */
function buildTaskBody_(msg) {
  const link = 'https://mail.google.com/mail/u/0/#inbox/' + msg.getThread().getId();
  const title = '★新規案件のお知らせ★';
  const guidance = [
    '問い合わせフォームから下記の新規案件の問い合わせが来ています。',
    'メールワイズでも見ることができます。',
    '',
    '担当者は対応をお願いします。',
    '',
    'ーーーー＜以下本文＞ーーーー',
    ''
  ].join('\n');

  const plain = (msg.getPlainBody && msg.getPlainBody()) || msg.getBody() || '';
  const meta = [`件名: ${msg.getSubject()}`, `From: ${msg.getFrom()}`, `Gmail: ${link}`, ''].join('\n');

  // 本文の取り込み（長文はプレーン部分のみを切り詰める）
  let bodyPart = CONFIG.INCLUDE_FULL_BODY ? plain : plain.slice(0, 600);
  const fixedPart = `[info][title]${title}[/title]\n${guidance}${meta}`;
  const maxAllowedForBody = Math.max(0, CONFIG.MAX_BODY_LEN - fixedPart.length - 20);
  if (bodyPart.length > maxAllowedForBody) {
    bodyPart = bodyPart.slice(0, maxAllowedForBody) + '\n…(長文のため省略)';
  }

  return `${fixedPart}${bodyPart}[/info]`;
}

/** 至急通知本文（Chatworkの [info][title]... 形式） */
function buildUrgentBody_(msg) {
  const link = 'https://mail.google.com/mail/u/0/#inbox/' + msg.getThread().getId();
  const title = '★「至急」カテゴライズ通知のお知らせ★';
  const guidance = [
    '問い合わせフォームから「至急」にカテゴライズされた下記の通知が来ています。',
    '',
    '担当者は対応をお願いします。',
    '',
    'ーーー＜以下本文＞ーーー',
    ''
  ].join('\n');

  const plain = (msg.getPlainBody && msg.getPlainBody()) || msg.getBody() || '';
  const meta = [`件名: ${msg.getSubject()}`, `From: ${msg.getFrom()}`, `Gmail: ${link}`, ''].join('\n');

  let bodyPart = CONFIG.INCLUDE_FULL_BODY ? plain : plain.slice(0, 600);
  const fixedPart = `[info][title]${title}[/title]\n${guidance}${meta}`;
  const maxAllowedForBody = Math.max(0, CONFIG.MAX_BODY_LEN - fixedPart.length - 20);
  if (bodyPart.length > maxAllowedForBody) {
    bodyPart = bodyPart.slice(0, maxAllowedForBody) + '\n…(長文のため省略)';
  }

  return `${fixedPart}${bodyPart}[/info]`;
}

/** 期日(UNIX秒)を計算 */
function computeDueUnix_(opt) {
  const tz = opt.tz || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();
  const dt = new Date(now.getTime());
  dt.setDate(dt.getDate() + (opt.daysFromNow || 0));
  dt.setHours(opt.hour ?? 18, opt.minute ?? 0, 0, 0);
  // 厳密なTZ変換が必要なら Utilities.formatDate を使ってもOK
  return Math.floor(dt.getTime() / 1000);
}

/** ログシート（C列: gmail_message_id）から既処理 messageId セットを作る */
function buildProcessedMsgSet_(logSheet) {
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return new Set();
  const values = logSheet.getRange(2, 3, lastRow - 1, 1).getValues(); // C列のみ
  const set = new Set();
  values.forEach(r => { const v = r[0]; if (v) set.add(String(v)); });
  return set;
}

/** ログシートを用意（ヘッダ含む） */
function ensureLogSheet_(ss) {
  const headers = ['processed_at','gmail_thread_id','gmail_message_id','subject','from','category','chatwork_message_id','due_unix','status'];
  let sh = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_LOG);
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  }
  return sh;
}

/** サマリーシートを用意 */
function ensureSummarySheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_SUMMARY);
  return sh;
}

/** 集計：カテゴリ別の累計/最近30日/最近7日/本日 を更新 */
function updateSummary_(logSheet, summarySheet) {
  const lastRow = logSheet.getLastRow();
  const result = { total:{}, last30:{}, last7:{}, today:{} };

  if (lastRow >= 2) {
    const vals = logSheet.getRange(2,1,lastRow-1,6).getValues(); // A..F: processed_at, thread_id, msg_id, subject, from, category
    const now = new Date();
    const d0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    vals.forEach(r => {
      const dt = new Date(r[0]);
      const cat = r[5] || CONFIG.DEFAULT_CATEGORY;
      // total
      result.total[cat] = (result.total[cat]||0)+1;
      // last30
      if ((now - dt) <= 30*24*3600*1000) result.last30[cat] = (result.last30[cat]||0)+1;
      // last7
      if ((now - dt) <= 7*24*3600*1000) result.last7[cat] = (result.last7[cat]||0)+1;
      // today
      if (dt >= d0) result.today[cat] = (result.today[cat]||0)+1;
    });
  }

  // シートをクリアして書き直し
  summarySheet.clear();
  const sections = [
    ['カテゴリ別 合計', result.total],
    ['カテゴリ別 過去30日', result.last30],
    ['カテゴリ別 過去7日', result.last7],
    ['カテゴリ別 本日', result.today],
  ];

  let row = 1;
  sections.forEach(([title, obj]) => {
    summarySheet.getRange(row,1).setValue(title).setFontWeight('bold');
    row++;
    summarySheet.getRange(row,1,1,2).setValues([['カテゴリ','件数']]).setFontWeight('bold');
    row++;

    const cats = Object.keys(obj).sort();
    const data = cats.length ? cats.map(k => [k, obj[k]]) : [['(データなし)', 0]];
    summarySheet.getRange(row,1,data.length,2).setValues(data);
    row += data.length + 1;
  });
}

/** （任意）担当者ID確認：ルームメンバーの account_id をログ出力 */
function cwListMembers() {
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.TOKEN_PROP_KEY);
  const res = UrlFetchApp.fetch(`https://api.chatwork.com/v2/rooms/${getActiveRoomId_()}/members`, {
    method: 'get',
    headers: { 'X-ChatWorkToken': token },
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

/** 手動テスト：本番ルームにテスト通知を投稿（マイチャットにも送信） */
function testPostToProd() {
  const body = buildInfoMessageForTest_();
  const posted = [];

  // マイチャット（確認用）
  try {
    const mid = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, message_id: mid });
  } catch (e) {
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, error: String(e) });
  }

  // 本番の各カテゴリルーム
  const targets = [
    { key: 'general',   roomId: CONFIG.ROOM_MAP.general },
    { key: 'labor',     roomId: CONFIG.ROOM_MAP.labor },
    { key: 'bkan',      roomId: CONFIG.ROOM_MAP.bkan },
    { key: 'asbestos',  roomId: CONFIG.ROOM_MAP.asbestos },
  ];

  targets.forEach(t => {
    try {
      const mid = cwPostMessage_(t.roomId, body);
      posted.push({ room: t.key, roomId: t.roomId, message_id: mid });
    } catch (e) {
      posted.push({ room: t.key, roomId: t.roomId, error: String(e) });
    }
  });

  Logger.log(JSON.stringify({ posted }, null, 2));
}

/** テスト通知本文（シンプルな info/title ブロック） */
function buildInfoMessageForTest_() {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  return [
    '[info][title]★新規案件のお知らせ（テスト）★[/title]',
    `本番ルーム投稿の接続確認です（${now}）。`,
    'このメッセージはテストです。',
    '[/info]'
  ].join('\n');
}

/** 手動テスト：「至急」ルートで本番全体スレに通知（マイチャットにも送信） */
function testPostUrgent() {
  const body = buildUrgentMessageForTest_();
  const posted = [];

  // マイチャット
  try {
    const mid = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, message_id: mid });
  } catch (e) {
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, error: String(e) });
  }

  // 全体スレ（至急の宛先）
  try {
    const mid = cwPostMessage_(CONFIG.PROD_ROOM_ID, body);
    posted.push({ room: 'prod_all', roomId: CONFIG.PROD_ROOM_ID, message_id: mid });
  } catch (e) {
    posted.push({ room: 'prod_all', roomId: CONFIG.PROD_ROOM_ID, error: String(e) });
  }

  Logger.log(JSON.stringify({ posted }, null, 2));
}

/** テスト用：「至急」メッセージ本文 */
function buildUrgentMessageForTest_() {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const title = '★「至急」カテゴライズ通知のお知らせ★';
  const guidance = [
    '問い合わせフォームから「至急」にカテゴライズされた下記の通知が来ています。',
    '',
    '担当者は対応をお願いします。',
    '',
    'ーーー＜以下本文＞ーーー',
    ''
  ].join('\n');
  const meta = [`件名: （テスト）至急対応の確認`, `From: test@example.com`, `Gmail: （テスト）リンクなし`, ''].join('\n');
  const body = '（テスト本文）このメッセージは至急ルートの投稿テストです。';
  return `[info][title]${title}[/title]\n${guidance}${meta}${body}\n[/info]\n（送信時刻: ${now}）`;
}