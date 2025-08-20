/**
 * ============ 設定 ============
 * 1) Chatwork: ROOM_ID, TO_IDS_CSV（担当者のaccount_idをカンマ区切り。未指定なら未アサイン）
 * 2) Spreadsheet: ログ＆サマリーを書き込む先のシート
 * 3) スクリプトプロパティ: CHATWORK_TOKEN を保存
 */
const CONFIG = {
  INFO_ADDRESS: 'info@ascope.net',
  // 「Fwd:」や「【…】問合せフォームからのお問い合わせがありました」に対応
  SUBJECT_HOOK: /お問い合わせがありました|(?:お問い合わせ|お問合せ|お問合わせ|お問い合せ|問合せ|問い合わせ)フォームからのお問い合わせがありました|＜至急＞<B肝新規相談会申込>/,
  SUBJECT_QUERIES: [
    'お問い合わせがありました',
    '＜至急＞<B肝新規相談会申込>',
    'RAW:(お問い合わせ 回答 ありました)',
    'RAW:(問合せ フォーム お問い合わせ ありました)'
  ], // Gmail検索（OR）
  URGENT_KEYWORD: '至急',                        // 至急判定キーワード（件名/本文）
  DONE_LABEL: 'CWタスク化済',                 // Gmailの重複防止ラベル
  // 投稿先（テスト/本番）
  MODE: 'prod',                               // 'test' | 'prod'
  TEST_ROOM_ID: '406362955',                  // ASCOPEのbotくんのマイチャット
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

  // ===== 本文フィルタ（通知スキップ） =====
  BODY_SKIP_PATTERNS: [
    '「その他」のお問い合わせ（簡単な内容）'
  ],
  BODY_SKIP_ONLY_FOR_TOP: true,
  SKIP_ADD_DONE_LABEL: true,
  DEBUG: true,
};

/** 共通ログ関数：Logger.log と console.log の両方に出力 */
function debugLog(...args) {
  if (!CONFIG.DEBUG) return;
  try { Logger.log(args.map(x => (typeof x === 'string' ? x : JSON.stringify(x))).join(' ')); } catch(_) {}
  try { console.log(...args); } catch(_) {}
}

/** [ADD] SUBJECT_QUERIES から subject:() 句を生成（RAW: はクォートしない） */
function buildSubjectClause_() { // [ADD]
  const parts = (CONFIG.SUBJECT_QUERIES || []).map(q => {
    if (/^RAW:/.test(q)) return q.replace(/^RAW:/, ''); // 例: RAW:(問合せ フォーム お問い合わせ ありました)
    return `"${q}"`;
  });
  // 日本語の分かち書き・記号差異でも拾えるように保険（広め）
  parts.push('(お問い合わせ ありました)');
  parts.push('(問合せ ありました)');
  return `subject:(${parts.join(' OR ')})`;
}

/** [ADD] Gmail検索式を構築（堅牢化 + ログ） */
function buildGmailSearchQuery_() { // [ADD]
  const addr = CONFIG.INFO_ADDRESS;
  const addrClause = `(deliveredto:${addr} OR to:${addr})`;
  const subjectClause = buildSubjectClause_();
  const q = `in:anywhere ${addrClause} newer_than:${CONFIG.SEARCH_DAYS}d ${subjectClause}`;
  return q;
}

/** スレッド＆全メッセージに DONE ラベル付与（保険付き） */
function applyDoneLabel_(thread) {
  const doneLabel = GmailApp.getUserLabelByName(CONFIG.DONE_LABEL) || GmailApp.createLabel(CONFIG.DONE_LABEL);
  try { thread.addLabel(doneLabel); } catch (e) { debugLog('[WARN] thread.addLabel failed', e); }
  try {
    const msgs = thread.getMessages();
    for (const m of msgs) {
      try { m.addLabel(doneLabel); } catch (e) { debugLog('[WARN] message.addLabel failed', e); }
    }
  } catch (e) {
    debugLog('[WARN] fetch messages failed', e);
  }
}

/** メイン：対象メール（メッセージ単位）→Chatworkタスク作成→ログ記録→ラベル付与→サマリー更新 */
function createTasksWithLogging() {
  debugLog('[HB] BEGIN createTasksWithLogging', new Date());
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const logSheet = ensureLogSheet_(ss);
  const summarySheet = ensureSummarySheet_(ss);

  const processedMsgIds = buildProcessedMsgSet_(logSheet);

  // ===== Gmail検索（改良版） =====
  let gmailSearchQuery = buildGmailSearchQuery_(); // [CHG]
  debugLog('[HB] QUERY', gmailSearchQuery);
  let threads = GmailApp.search(gmailSearchQuery);
  debugLog('[HB] THREADS(found)', threads.length);

  // 0件なら「宛先制限なし」のフォールバックでも探す（ログだけ残す）[ADD]
  if (!threads.length) {
    const fallbackQ = `in:anywhere newer_than:${CONFIG.SEARCH_DAYS}d ${buildSubjectClause_()}`;
    debugLog('[HB] FALLBACK QUERY', fallbackQ);
    const fb = GmailApp.search(fallbackQ);
    debugLog('[HB] FALLBACK THREADS(found)', fb.length);
    if (fb.length) {
      debugLog('[INFO] 取りこぼし推定: deliveredto/to 条件で弾かれている可能性が高い');
      threads = fb; // 実際に処理も走らせる（取りこぼし防止）
    }
  }

  const thresholdDate = new Date(Date.now() - CONFIG.SEARCH_DAYS * 24 * 3600 * 1000);

  let newLogRows = [];
  let processedCount = 0, skippedCount = 0;
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(msg => {
      try {
        const msgDate = msg.getDate && msg.getDate();
        if (msgDate && msgDate < thresholdDate) return;

        const subject = msg.getSubject() || '';
        const subjectHookOk = CONFIG.SUBJECT_HOOK.test(subject);
        debugLog('[HB] MSG',
          Utilities.formatDate(msgDate || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
          `subject="${subject}"`,
          `hook=${subjectHookOk}`
        );
        if (!subjectHookOk) return;

        const gmailMessageId = String(msg.getId());
        if (processedMsgIds.has(gmailMessageId)) return;

        const threadId = String(thread.getId());
        const category = detectCategory_(subject);
        const plainBodyForDetect = (msg.getPlainBody && msg.getPlainBody()) || msg.getBody() || '';
        const isUrgent = (subject.indexOf(CONFIG.URGENT_KEYWORD) !== -1) || (plainBodyForDetect.indexOf(CONFIG.URGENT_KEYWORD) !== -1);

        // === 本文フィルタ：通知スキップ判定 ===
        const skipHit = bodySkipHit_(category, plainBodyForDetect, subject);
        if (skipHit) debugLog('[HB] BODY_SKIP_HIT', skipHit);
        if (skipHit) {
          const now = new Date();
          newLogRows.push([
            now,                 // processed_at
            threadId,            // gmail_thread_id
            gmailMessageId,      // gmail_message_id
            subject,             // subject
            msg.getFrom() || '', // from
            category,            // category
            '',                  // chatwork_message_id
            '',                  // due_unix
            `skipped:${skipHit}` // status
          ]);
          if (CONFIG.SKIP_ADD_DONE_LABEL) applyDoneLabel_(thread);
          skippedCount++;
          return;
        }

        const body = isUrgent ? buildUrgentBody_(msg) : buildTaskBody_(msg);
        const dueUnix = CONFIG.DUE.enable ? computeDueUnix_(CONFIG.DUE) : '';

        // --- Chatwork 通知（堅牢化）---
        // 1) マイチャットへ（確認用・必ず試行）
        let chatworkMessageId = '';
        let status = 'created';
        let cwError = '';

        const myChatMessageId = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
        chatworkMessageId = myChatMessageId;

        // 2) 本番モード: カテゴリ対応ルームへ
        if (CONFIG.MODE === 'prod') {
          const destRoomId = isUrgent ? CONFIG.PROD_ROOM_ID : mapCategoryToRoomId_(category);
          try { // [ADD] 個別 try/catch
            const postedId = cwPostMessage_(destRoomId, body);
            if (postedId) chatworkMessageId = postedId;
          } catch (e) {
            cwError = String(e);
            status = `created_with_error:${String(e).slice(0,200)}`;
            debugLog('[ERR] Chatwork post failed for dest room', destRoomId, e);

            // 任意: 一般案件にフォールバック投稿（必要なければコメントアウト）
            try {
              const fbBody = `[info][title]【自動振替】カテゴリ「${category}」宛の投稿に失敗しました（確認をお願いします）[/title]\n${body}`;
              const fbId = cwPostMessage_(CONFIG.ROOM_MAP.general, fbBody);
              if (fbId) chatworkMessageId = fbId;
              debugLog('[INFO] Fallback posted to general', CONFIG.ROOM_MAP.general);
            } catch (_) { /* フォールバック失敗は握りつぶし */ }
          }
        }

        // --- ログ行を追加（失敗でも必ず記録） ---
        const now = new Date();
        newLogRows.push([
          now,                 // processed_at
          threadId,            // gmail_thread_id
          gmailMessageId,      // gmail_message_id（重複判定のキー）
          subject,             // subject
          msg.getFrom() || '', // from
          category,            // category
          chatworkMessageId || '', // chatwork_message_id
          dueUnix || '',       // due_unix
          status               // status
        ]);

        // --- ラベル（失敗でも付ける：次回の迷子防止） [CHG] ---
        applyDoneLabel_(thread);
        processedCount++;

      } catch (err) {
        // ここに落ちるのは想定外の例外（ラベリング前）。メールで通知。
        console.error(err);
        GmailApp.sendEmail(Session.getActiveUser().getEmail(), '【GAS】Chatworkタスク作成エラー', String(err));
      }
    });
  });

  // まとめてログ追加
  if (newLogRows.length) logSheet.getRange(logSheet.getLastRow()+1, 1, newLogRows.length, newLogRows[0].length).setValues(newLogRows);

  // 集計を更新
  updateSummary_(logSheet, summarySheet);
  debugLog('[HB] END createTasksWithLogging', {processedCount, skippedCount, appended: newLogRows.length});
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

  try {
    const json = JSON.parse(res.getContentText());
    const id = json.message_id ? String(json.message_id) : '';
    return id;
  } catch (_) {
    return '';
  }
}

function getActiveRoomId_() { return CONFIG.MODE === 'prod' ? CONFIG.PROD_ROOM_ID : CONFIG.TEST_ROOM_ID; }

/** カテゴリ名→本番の投稿先ルームID */
function mapCategoryToRoomId_(categoryName) {
  if (categoryName === '労働') return CONFIG.ROOM_MAP.labor;
  if (categoryName === 'アスベスト') return CONFIG.ROOM_MAP.asbestos;
  if (categoryName === 'B型肝炎') return CONFIG.ROOM_MAP.bkan;
  return CONFIG.ROOM_MAP.general;
}

/** 件名→カテゴリ名 */
function detectCategory_(subject) {
  for (const c of CONFIG.CATEGORIES) if (c.pattern.test(subject)) return c.name;
  return CONFIG.DEFAULT_CATEGORY;
}

/** 本文スキップ判定（ヒットしたパターン文字列 or 空文字） */
function bodySkipHit_(category, plainBody, subject) {
  if (!plainBody) return '';
  if (CONFIG.BODY_SKIP_ONLY_FOR_TOP) {
    const subj = subject || '';
    const isTop = /トップページ/.test(subj) || category === 'トップページ';
    if (!isTop) return '';
    // トップページ案件のみフィルタ: OK
  }
  for (const p of CONFIG.BODY_SKIP_PATTERNS) {
    let re = p;
    if (typeof p === 'string') {
      re = new RegExp(p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    }
    try {
      if (re.test(plainBody)) return (typeof p === 'string') ? p : re.toString();
    } catch (_) {}
  }
  return '';
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

/** 集計：カテゴリ別の累計/最近30日/最近7日/本日 を更新（skipped を除外） */
function updateSummary_(logSheet, summarySheet) {
  const lastRow = logSheet.getLastRow();
  const result = { total:{}, last30:{}, last7:{}, today:{} };

  if (lastRow >= 2) {
    const vals = logSheet.getRange(2,1,lastRow-1,9).getValues();
    const now = new Date();
    const d0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    vals.forEach(r => {
      const status = (r[8] || '').toString();
      if (status.indexOf('skipped:') === 0) return;

      const dt = new Date(r[0]);
      const cat = r[5] || CONFIG.DEFAULT_CATEGORY;
      result.total[cat] = (result.total[cat]||0)+1;
      if ((now - dt) <= 30*24*3600*1000) result.last30[cat] = (result.last30[cat]||0)+1;
      if ((now - dt) <= 7*24*3600*1000) result.last7[cat] = (result.last7[cat]||0)+1;
      if (dt >= d0) result.today[cat] = (result.today[cat]||0)+1;
    });
  }

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

  try {
    const mid = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, message_id: mid });
  } catch (e) {
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, error: String(e) });
  }

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

  try {
    const mid = cwPostMessage_(CONFIG.TEST_ROOM_ID, body);
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, message_id: mid });
  } catch (e) {
    posted.push({ room: 'mychat', roomId: CONFIG.TEST_ROOM_ID, error: String(e) });
  }

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

/** 検証用：ワイド検索で件名をダンプ（検索が0件でも必ず件名をダンプできる） */
function debugDumpSubjects_() {
  const q = `in:anywhere newer_than:${CONFIG.SEARCH_DAYS}d`; // deliveredto/subject を外す
  const threads = GmailApp.search(q);
  debugLog('[DUMP] query=', q, 'threads=', threads.length);
  threads.slice(0, 30).forEach((t, i) => {
    const last = t.getMessages().pop();
    const ts = last && last.getDate() ? Utilities.formatDate(last.getDate(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : 'N/A';
    debugLog(`[${i}]`, ts, last ? last.getSubject() : '(no subject)');
  });
}

/** [ADD] 任意のクエリで対象を確認するデバッグ */
function debugSearchQuery_(q) { // [ADD]
  const threads = GmailApp.search(q);
  debugLog('[DEBUG SEARCH] q=', q, 'threads=', threads.length);
  threads.slice(0, 20).forEach((t, i) => {
    const last = t.getMessages().pop();
    const ts = last && last.getDate() ? Utilities.formatDate(last.getDate(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : 'N/A';
    debugLog(`[${i}]`, ts, last ? last.getSubject() : '(no subject)', 'threadId=' + t.getId());
  });
}

/** 過去の skipped 行に後付けでラベルを貼るユーティリティ（1回実行でOK） */
function replayApplyDoneLabelForSkipped() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rows = sh.getRange(2, 1, lastRow - 1, 9).getValues(); // A..I
  const doneLabel = GmailApp.getUserLabelByName(CONFIG.DONE_LABEL) || GmailApp.createLabel(CONFIG.DONE_LABEL);
  let ok = 0, ng = 0;

  for (const r of rows) {
    const status = String(r[8] || '');
    if (!status.startsWith('skipped:')) continue;
    const threadId = String(r[1] || '');
    if (!threadId) continue;
    try {
      const th = GmailApp.getThreadById(threadId);
      th.addLabel(doneLabel);
      th.getMessages().forEach(m => { try { m.addLabel(doneLabel); } catch (_) {} });
      ok++;
    } catch (e) { ng++; debugLog('[WARN] replay label failed', threadId, e); }
  }
  debugLog('[INFO] replayApplyDoneLabelForSkipped_', { ok, ng });
}


function debugSearch_current(){            // 現行の検索式（改修版が有効か判定）
  debugSearchQuery_(buildGmailSearchQuery_());
}
function debugSearch_subjectOnly7d(){      // 宛先条件なしフォールバック（7日）
  const q = 'in:anywhere newer_than:7d ' + buildSubjectClause_();
  debugSearchQuery_(q);
}
function debugSearch_fromProp(){           // プロパティに書いた任意クエリを実行
  const q = PropertiesService.getScriptProperties().getProperty('DEBUG_Q');
  if (!q) throw new Error('Script property DEBUG_Q が未設定です');
  debugSearchQuery_(q);
}