/**
 * 旅のしおり共有アプリ - Google Apps Script バックエンド
 *
 * データベースとして、このスクリプトが紐付いている（または SHEET_ID で指定した）
 * Google スプレッドシートを使用します。
 *
 * シート構成（Itineraries シート、シート名は SHEET_NAME で指定）:
 *   A: bookmark_id   しおりの一意なID（UUID）
 *   B: title         旅行タイトル
 *   C: dates         日程（自由記述テキスト）
 *   D: timeline      タイムラインデータ（JSON文字列）
 *   E: items         持ち物データ（JSON文字列）
 *   F: updated_at    最終更新日時（ISO文字列）
 *   G: owner_id      作成者（会員）のID。編集・削除にはこの会員としてのログインが必要
 *   H: shared_with    公開相手（招待済みメールアドレス）のJSON配列文字列
 *
 * 会員認証用シート（Members シート、シート名は MEMBERS_SHEET_NAME で指定）:
 *   A: member_id             会員の一意なID（UUID）
 *   B: nickname              ニックネーム
 *   C: email                 メールアドレス（ログインID）
 *   D: password_hash         パスワードハッシュ（SHA-256、salt+pepper付き）
 *   E: salt                  パスワードごとのソルト
 *   F: must_change_password  初回ログイン（仮パスワード）後の強制変更フラグ
 *   G: failed_attempts       ログイン連続失敗回数
 *   H: locked_until          ロック解除日時（ISO文字列、ロック中でなければ空文字）
 *   I: created_at            登録日時（ISO文字列）
 *   J: updated_at            最終更新日時（ISO文字列）
 *
 * アクセス権限の考え方:
 *   - 閲覧（GET）はログイン不要。URLを知っていれば誰でも見られる。
 *   - 作成・保存・削除（POST）にはログインが必須。しおりの owner_id が空（未ログイン時代の
 *     データ等）であれば、最初に保存したログイン中の会員がそのまま所有者になる。
 *     owner_id が設定済みの場合、その本人（同じ会員IDでログイン中）以外は保存・削除できない。
 *
 * デプロイ方法は同梱の DEPLOY.md を参照してください。
 */

// ==== バージョン ====
// デプロイのたびに更新すると、画面下部の表示で「新しいGASが反映されたか」を確認できる。
const GAS_VERSION = '1.3.0 (2026-11-05)';

// ==== 設定 ====
// スクリプトを特定のスプレッドシートに紐付けない場合（スタンドアロンスクリプトの場合）は
// ここにスプレッドシートIDを設定してください。
// コンテナバインド型スクリプト（スプレッドシートの「拡張機能 > Apps Script」から作成した場合）は
// 空文字のままで、SpreadsheetApp.getActiveSpreadsheet() が自動的に使われます。
const SHEET_ID = '';
const SHEET_NAME = 'Itineraries';
const HEADER = ['bookmark_id', 'title', 'dates', 'timeline', 'items', 'updated_at', 'owner_id', 'shared_with'];

// AIによるリスト作成支援（タイムライン・持ち物）で使うGeminiモデル。
// 個別バージョン名は廃止されることがあるため、常に最新安定版を指すエイリアスを使う。
const GEMINI_MODEL = 'gemini-flash-latest';
const MAX_SUGGESTIONS = 10; // 1回の提案リクエストで返す最大件数
const MAX_SHARED_WITH = 30; // しおり1件あたりの公開相手（招待メールアドレス）の最大件数

// ==== 予定表の読み取り（画像・PDF・Excel/CSV → タイムライン）の設定 ====
const SCHEDULE_IMPORT_MAX_FILES = 5;        // 1回に送れるファイル数
const SCHEDULE_IMPORT_MAX_TEXT = 30000;     // Excel/CSV/貼り付けテキストの最大文字数
const SCHEDULE_IMPORT_MAX_ENTRIES = 200;    // タイムラインの上限と合わせる
const SCHEDULE_IMPORT_ALLOWED_MIME = [
  'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'
];

// ==== 会員認証の設定 ====
const MEMBERS_SHEET_NAME = 'Members';
const MEMBERS_HEADER = [
  'member_id', 'nickname', 'email', 'password_hash', 'salt',
  'must_change_password', 'failed_attempts', 'locked_until', 'created_at', 'updated_at'
];
const SESSION_TTL_SECONDS = 6 * 60 * 60; // 6時間（CacheServiceの上限）
const MAX_LOGIN_ATTEMPTS = 5;
const LOCK_MINUTES = 15;

/**
 * 対象のシートを取得する。シートが存在しない場合はヘッダー付きで新規作成する。
 * 既存シート（owner_id / shared_with 列追加前に作られたもの）にはヘッダーを補完する。
 */
function getSheet_() {
  const ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  } else {
    ensureExtraColumns_(sheet);
  }
  return sheet;
}

/**
 * 既存シートに owner_id（G列）/ shared_with（H列）が無ければヘッダーを補完する。
 * 会員機能を後から追加した既存デプロイとの互換性のため。
 */
function ensureExtraColumns_(sheet) {
  for (let col = 1; col <= HEADER.length; col++) {
    const cell = sheet.getRange(1, col);
    if (!cell.getValue()) {
      cell.setValue(HEADER[col - 1]);
    }
  }
}

/**
 * bookmark_id に一致する行番号（1始まり、ヘッダー込み）を返す。
 * 見つからない場合は -1。
 */
function findRowIndexById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      return i + 2; // ヘッダー行 + 0始まりオフセット補正
    }
  }
  return -1;
}

/**
 * 1行分の値配列をレスポンス用オブジェクトに変換する。
 * sharedWith は所有者本人にのみ返す想定（doGet側でフィルタする）。
 */
function rowToObject_(rowValues) {
  return {
    id: rowValues[0],
    title: rowValues[1] || '',
    dates: rowValues[2] || '',
    timeline: safeParseJson_(rowValues[3], []),
    items: safeParseJson_(rowValues[4], []),
    updated_at: rowValues[5] || '',
    sharedWith: safeParseJson_(rowValues[7], [])
  };
}

function safeParseJson_(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch (e) {
    return fallback;
  }
}

/**
 * JSON文字列を作ってレスポンスを返す共通関数。
 * ContentService は任意のレスポンスヘッダーを追加できないため、
 * CORS プリフライトを発生させない「シンプルリクエスト」で通信することを前提にしている。
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET /?id=xxxx&token=xxxx（token省略可）
 * 指定した bookmark_id のしおりデータを返す。閲覧自体はログイン不要。
 * token（セッション）が有効かつ本人が所有者（または owner_id 未設定）の場合のみ
 * data.canEdit を true にし、公開相手（shared_with）の情報も含める。
 * id が指定されていない、または該当データが無い場合は success:false を返す。
 *
 * GET /?action=me&token=xxxx
 * セッションが有効な場合に会員情報を返す。
 *
 * GET /?action=myItineraries&token=xxxx
 * ログイン中の会員が作成したしおりの一覧（id/title/dates/updated_at）を返す。
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'me') {
      return handleMe_(e);
    }
    if (action === 'myItineraries') {
      return handleMyItineraries_(e);
    }
    if (action === 'version') {
      return jsonResponse_({ success: true, version: GAS_VERSION });
    }

    const id = e && e.parameter && e.parameter.id;
    if (!id) {
      return jsonResponse_({ success: false, error: 'id が指定されていません。' });
    }

    const sheet = getSheet_();
    const rowIndex = findRowIndexById_(sheet, id);
    if (rowIndex === -1) {
      return jsonResponse_({ success: false, error: 'not_found' });
    }

    const rowValues = sheet.getRange(rowIndex, 1, 1, HEADER.length).getValues()[0];
    const ownerId = rowValues[6] || '';
    const session = getSessionMember_(e.parameter && e.parameter.token);
    const canEdit = !!session && (!ownerId || ownerId === session.id);

    const data = rowToObject_(rowValues);
    data.canEdit = canEdit;
    if (!canEdit) {
      delete data.sharedWith; // 編集権限がない閲覧者には公開相手リストを見せない
    }

    return jsonResponse_({ success: true, data: data });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST body (JSON, text/plain として送信されることを想定):
 * {
 *   id: "xxxx",
 *   title: "沖縄旅行",
 *   dates: "2026/08/01 - 2026/08/03",
 *   timeline: [{ id, date, time, content }, ...],
 *   items: [{ id, text, checked }, ...],
 *   token: "xxxx"   // ログイン中であればセッショントークン（省略可）
 * }
 *
 * 既存の bookmark_id があれば上書き更新、無ければ新規行を追加（アップサート）する。
 * 新規作成時にログイン中であれば owner_id を紐付ける（更新時は元の owner_id を維持する）。
 */
/**
 * POST body (JSON, text/plain として送信されることを想定):
 * {
 *   id: "xxxx",
 *   title: "沖縄旅行",
 *   dates: "2026/08/01 - 2026/08/03",
 *   timeline: [{ id, date, time, content }, ...],
 *   items: [{ id, text, checked }, ...],
 *   sharedWith: ["a@example.com", ...],  // 公開相手（招待済みメールアドレス）
 *   token: "xxxx"   // ログイン中のセッショントークン（保存には必須）
 * }
 *
 * 既存の bookmark_id があれば上書き更新、無ければ新規行を追加（アップサート）する。
 * 保存にはログインが必須。新規作成時はログイン中の会員が owner_id として記録され、
 * 既存のしおりは owner_id が空（未ログイン時代のデータ）でない限り、本人以外は保存できない。
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ success: false, error: 'リクエストボディがありません。' });
    }

    const body = JSON.parse(e.postData.contents);

    // action:'suggest' の場合はしおりの保存ではなく、AIによる提案生成を行う。
    if (body.action === 'suggest') {
      return handleSuggest_(body);
    }

    // action:'parseSchedule' の場合は予定表（画像・PDF・表データ）をAIで読み取り、
    // タイムライン形式に変換して返す（保存はクライアント側で確認・修正後に行う）。
    if (body.action === 'parseSchedule') {
      return handleParseSchedule_(body);
    }

    // action:'delete' の場合はしおりの保存ではなく、削除を行う。
    if (body.action === 'delete') {
      return handleDelete_(body);
    }

    // action:'invite' の場合はしおりの保存ではなく、公開相手へのURLメール送信を行う。
    if (body.action === 'invite') {
      return handleInvite_(body);
    }

    // 会員機能関連のアクション。
    if (body.action === 'register') {
      return handleRegister_(body);
    }
    if (body.action === 'login') {
      return handleLogin_(body);
    }
    if (body.action === 'logout') {
      return handleLogout_(body);
    }
    if (body.action === 'changePassword') {
      return handleChangePassword_(body);
    }
    if (body.action === 'forgotPassword') {
      return handleForgotPassword_(body);
    }

    const id = body.id;
    if (!id) {
      return jsonResponse_({ success: false, error: 'id が指定されていません。' });
    }

    const session = getSessionMember_(body.token);
    if (!session) {
      return jsonResponse_({ success: false, error: 'しおりを保存するにはログインが必要です。' });
    }

    const sheet = getSheet_();
    const now = new Date().toISOString();

    // 同時編集による行の奪い合いを避けるためロックを取得する
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rowIndex = findRowIndexById_(sheet, id);

      if (rowIndex !== -1) {
        const existingOwner = sheet.getRange(rowIndex, 7, 1, 1).getValue() || '';
        if (existingOwner && existingOwner !== session.id) {
          return jsonResponse_({ success: false, error: 'このしおりを編集する権限がありません。' });
        }
      }

      const sharedWith = sanitizeEmailList_(body.sharedWith);
      const rowValues = [
        id,
        body.title || '',
        body.dates || '',
        JSON.stringify(body.timeline || []),
        JSON.stringify(body.items || []),
        now,
        session.id,
        JSON.stringify(sharedWith)
      ];

      if (rowIndex === -1) {
        sheet.appendRow(rowValues);
      } else {
        sheet.getRange(rowIndex, 1, 1, HEADER.length).setValues([rowValues]);
      }
    } finally {
      lock.releaseLock();
    }

    return jsonResponse_({ success: true, updated_at: now });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST body:
 * {
 *   action: 'delete',
 *   id: 'xxxx',
 *   token: 'xxxx'
 * }
 * 指定した bookmark_id の行を削除する。ログインが必須で、owner_id が設定済みの
 * しおりは本人（同じ会員IDでログイン中）以外は削除できない。
 */
function handleDelete_(body) {
  const id = body.id;
  if (!id) {
    return jsonResponse_({ success: false, error: 'id が指定されていません。' });
  }

  const session = getSessionMember_(body.token);
  if (!session) {
    return jsonResponse_({ success: false, error: 'しおりを削除するにはログインが必要です。' });
  }

  const sheet = getSheet_();

  // 保存処理と競合しないよう、削除時もロックを取得する。
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rowIndex = findRowIndexById_(sheet, id);
    if (rowIndex === -1) {
      return jsonResponse_({ success: false, error: 'not_found' });
    }
    const existingOwner = sheet.getRange(rowIndex, 7, 1, 1).getValue() || '';
    if (existingOwner && existingOwner !== session.id) {
      return jsonResponse_({ success: false, error: 'このしおりを削除する権限がありません。' });
    }
    sheet.deleteRow(rowIndex);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse_({ success: true });
}

/* ======================================================================
 * 会員認証（登録・ログイン・パスワード管理）
 * ====================================================================== */

/**
 * GET ?action=me&token=xxxx のハンドラ。
 * セッションが有効であれば会員情報を返す。
 */
function handleMe_(e) {
  const token = e && e.parameter && e.parameter.token;
  const session = getSessionMember_(token);
  if (!session) {
    return jsonResponse_({ success: false, error: 'not_logged_in' });
  }
  return jsonResponse_({ success: true, member: session });
}

/**
 * GET ?action=myItineraries&token=xxxx のハンドラ。
 * ログイン中の会員が作成したしおりの一覧を、更新日時の新しい順で返す。
 */
function handleMyItineraries_(e) {
  const token = e && e.parameter && e.parameter.token;
  const session = getSessionMember_(token);
  if (!session) {
    return jsonResponse_({ success: false, error: 'not_logged_in' });
  }

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse_({ success: true, itineraries: [] });
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const list = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i][6] === session.id) {
      list.push({
        id: values[i][0],
        title: values[i][1] || '',
        dates: values[i][2] || '',
        updated_at: values[i][5] || '',
        sharedWithCount: safeParseJson_(values[i][7], []).length
      });
    }
  }
  list.sort((a, b) => (b.updated_at || '').localeCompare(a.updated_at || ''));

  return jsonResponse_({ success: true, itineraries: list });
}

/**
 * POST body: { action: 'register', nickname: 'ケン', email: 'x@example.com' }
 * ニックネームとメールアドレスで会員登録し、仮パスワードをメール送信する。
 */
function handleRegister_(body) {
  try {
    const nickname = (body.nickname || '').toString().trim().slice(0, 50);
    const email = (body.email || '').toString().trim().toLowerCase().slice(0, 200);

    if (!nickname) {
      return jsonResponse_({ success: false, error: 'ニックネームを入力してください。' });
    }
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return jsonResponse_({ success: false, error: 'メールアドレスを正しく入力してください。' });
    }

    const sheet = getMembersSheet_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (findMemberRowByEmail_(sheet, email) !== -1) {
        return jsonResponse_({
          success: false,
          error: 'このメールアドレスは既に登録されています。ログインするか、パスワードの再発行をご利用ください。'
        });
      }

      const memberId = Utilities.getUuid();
      const tempPassword = generateTempPassword_();
      const salt = generateSalt_();
      const hash = hashPassword_(tempPassword, salt);
      const now = new Date().toISOString();

      sheet.appendRow([memberId, nickname, email, hash, salt, true, 0, '', now, now]);

      try {
        MailApp.sendEmail({
          to: email,
          subject: '【旅のしおり】会員登録完了・仮パスワードのお知らせ',
          body:
            nickname + ' 様\n\n' +
            '「旅のしおり」への会員登録が完了しました。\n\n' +
            '仮パスワード: ' + tempPassword + '\n\n' +
            'こちらの仮パスワードでログイン後、必ず新しいパスワードへの変更をお願いいたします。\n\n' +
            '※このメールに心当たりがない場合は破棄してください。'
        });
      } catch (mailErr) {
        // メール送信に失敗しても登録自体は成立させる（送信枠超過など）
        return jsonResponse_({
          success: true,
          warning: '登録は完了しましたが、確認メールの送信に失敗しました。しばらくしてからパスワード再発行をお試しください。'
        });
      }

      return jsonResponse_({ success: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST body: { action: 'login', email: 'x@example.com', password: '...' }
 * ログイン試行5回連続失敗で15分ロックする。
 */
function handleLogin_(body) {
  try {
    const email = (body.email || '').toString().trim().toLowerCase();
    const password = (body.password || '').toString();

    if (!email || !password) {
      return jsonResponse_({ success: false, error: 'メールアドレスとパスワードを入力してください。' });
    }

    const sheet = getMembersSheet_();
    const rowIndex = findMemberRowByEmail_(sheet, email);
    if (rowIndex === -1) {
      return jsonResponse_({ success: false, error: 'メールアドレスまたはパスワードが正しくありません。' });
    }

    const row = sheet.getRange(rowIndex, 1, 1, MEMBERS_HEADER.length).getValues()[0];
    const member = memberRowToObject_(row);

    if (member.lockedUntil && new Date(member.lockedUntil).getTime() > Date.now()) {
      const remainMin = Math.max(1, Math.ceil((new Date(member.lockedUntil).getTime() - Date.now()) / 60000));
      return jsonResponse_({
        success: false,
        error: 'ログイン試行回数が上限に達したため、あと約' + remainMin + '分ロックされています。'
      });
    }

    const inputHash = hashPassword_(password, member.salt);
    if (inputHash !== member.passwordHash) {
      const attempts = member.failedAttempts + 1;
      let lockedUntil = '';
      if (attempts >= MAX_LOGIN_ATTEMPTS) {
        lockedUntil = new Date(Date.now() + LOCK_MINUTES * 60000).toISOString();
      }
      sheet.getRange(rowIndex, 7).setValue(attempts);
      sheet.getRange(rowIndex, 8).setValue(lockedUntil);

      if (lockedUntil) {
        return jsonResponse_({
          success: false,
          error: 'パスワードが正しくありません。試行回数が上限に達したため' + LOCK_MINUTES + '分間ロックされます。'
        });
      }
      return jsonResponse_({ success: false, error: 'メールアドレスまたはパスワードが正しくありません。' });
    }

    // ログイン成功: 失敗回数・ロックをリセット
    sheet.getRange(rowIndex, 7).setValue(0);
    sheet.getRange(rowIndex, 8).setValue('');

    const token = createSession_(member.id, member.nickname, member.email, member.mustChangePassword);

    return jsonResponse_({
      success: true,
      token: token,
      member: { id: member.id, nickname: member.nickname, email: member.email, mustChangePassword: member.mustChangePassword }
    });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST body: { action: 'logout', token: 'xxxx' }
 */
function handleLogout_(body) {
  destroySession_(body.token);
  return jsonResponse_({ success: true });
}

/**
 * POST body: {
 *   action: 'changePassword', token: 'xxxx',
 *   currentPassword: '...',   // 強制変更（仮パスワード後の初回）の場合は不要
 *   newPassword: '...'
 * }
 */
function handleChangePassword_(body) {
  try {
    const session = getSessionMember_(body.token);
    if (!session) {
      return jsonResponse_({ success: false, error: 'セッションが切れました。再度ログインしてください。' });
    }

    const currentPassword = (body.currentPassword || '').toString();
    const newPassword = (body.newPassword || '').toString();
    if (!newPassword || newPassword.length < 8) {
      return jsonResponse_({ success: false, error: '新しいパスワードは8文字以上で入力してください。' });
    }

    const sheet = getMembersSheet_();
    const rowIndex = findMemberRowById_(sheet, session.id);
    if (rowIndex === -1) {
      return jsonResponse_({ success: false, error: '会員情報が見つかりません。' });
    }

    const row = sheet.getRange(rowIndex, 1, 1, MEMBERS_HEADER.length).getValues()[0];
    const member = memberRowToObject_(row);

    // 強制変更（仮パスワードでのログイン直後）でなければ現在のパスワードを確認する
    if (!member.mustChangePassword) {
      if (!currentPassword || hashPassword_(currentPassword, member.salt) !== member.passwordHash) {
        return jsonResponse_({ success: false, error: '現在のパスワードが正しくありません。' });
      }
    }

    const newSalt = generateSalt_();
    const newHash = hashPassword_(newPassword, newSalt);
    const now = new Date().toISOString();

    sheet.getRange(rowIndex, 4).setValue(newHash);
    sheet.getRange(rowIndex, 5).setValue(newSalt);
    sheet.getRange(rowIndex, 6).setValue(false);
    sheet.getRange(rowIndex, 10).setValue(now);

    // セッション上の情報も更新しておく
    CacheService.getScriptCache().put(
      'session_' + body.token,
      JSON.stringify({ id: member.id, nickname: member.nickname, email: member.email, mustChangePassword: false }),
      SESSION_TTL_SECONDS
    );

    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST body: { action: 'forgotPassword', email: 'x@example.com' }
 * 該当会員がいれば仮パスワードを再発行してメール送信する。
 * メールアドレスの存在有無が外部に漏れないよう、見つからない場合も success:true を返す。
 */
function handleForgotPassword_(body) {
  try {
    const email = (body.email || '').toString().trim().toLowerCase();
    if (!email) {
      return jsonResponse_({ success: false, error: 'メールアドレスを入力してください。' });
    }

    const sheet = getMembersSheet_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rowIndex = findMemberRowByEmail_(sheet, email);
      if (rowIndex === -1) {
        return jsonResponse_({ success: true });
      }

      const nickname = sheet.getRange(rowIndex, 2).getValue();
      const tempPassword = generateTempPassword_();
      const salt = generateSalt_();
      const hash = hashPassword_(tempPassword, salt);
      const now = new Date().toISOString();

      sheet.getRange(rowIndex, 4).setValue(hash);
      sheet.getRange(rowIndex, 5).setValue(salt);
      sheet.getRange(rowIndex, 6).setValue(true);
      sheet.getRange(rowIndex, 7).setValue(0);
      sheet.getRange(rowIndex, 8).setValue('');
      sheet.getRange(rowIndex, 10).setValue(now);

      try {
        MailApp.sendEmail({
          to: email,
          subject: '【旅のしおり】仮パスワード再発行のお知らせ',
          body:
            nickname + ' 様\n\n' +
            'パスワード再発行のリクエストを受け付けました。\n\n' +
            '仮パスワード: ' + tempPassword + '\n\n' +
            'こちらの仮パスワードでログイン後、必ず新しいパスワードへの変更をお願いいたします。\n\n' +
            '※このリクエストに心当たりがない場合はこのメールを破棄してください。'
        });
      } catch (mailErr) {
        return jsonResponse_({ success: false, error: 'メール送信に失敗しました。時間をおいて再度お試しください。' });
      }

      return jsonResponse_({ success: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * 文字列の配列から、有効なメールアドレスだけを重複除去・上限件数で正規化して返す。
 * 不正な形式や重複は静かに除外する（呼び出し側でエラー表示は行わない）。
 */
function sanitizeEmailList_(raw) {
  if (!Array.isArray(raw)) return [];
  const seen = new Set();
  const result = [];
  for (let i = 0; i < raw.length && result.length < MAX_SHARED_WITH; i++) {
    const email = (raw[i] || '').toString().trim().toLowerCase();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;
    if (seen.has(email)) continue;
    seen.add(email);
    result.push(email);
  }
  return result;
}

/**
 * POST body: {
 *   action: 'invite', token: 'xxxx', id: 'xxxx',
 *   emails: ['a@example.com', ...], shareUrl: 'https://.../?id=xxxx'
 * }
 * ログイン中の所有者が、指定した公開相手（メールアドレス）へ共有URLをメール送信する。
 * 送信した相手は shared_with 列に記録される（既存の記録とマージ）。
 * 受信者はURLを知っていればログインなしで閲覧できる（別途、閲覧APIはログイン不要）。
 */
function handleInvite_(body) {
  try {
    const session = getSessionMember_(body.token);
    if (!session) {
      return jsonResponse_({ success: false, error: 'ログインが必要です。' });
    }

    const id = body.id;
    if (!id) {
      return jsonResponse_({ success: false, error: 'id が指定されていません。' });
    }

    const shareUrl = (body.shareUrl || '').toString().trim().slice(0, 500);
    if (!shareUrl) {
      return jsonResponse_({ success: false, error: '共有URLが取得できませんでした。' });
    }

    const emails = sanitizeEmailList_(body.emails);
    if (emails.length === 0) {
      return jsonResponse_({ success: false, error: '送信先のメールアドレスを入力してください。' });
    }

    const sheet = getSheet_();
    const rowIndex = findRowIndexById_(sheet, id);
    if (rowIndex === -1) {
      return jsonResponse_({ success: false, error: 'not_found' });
    }

    const rowValues = sheet.getRange(rowIndex, 1, 1, HEADER.length).getValues()[0];
    const ownerId = rowValues[6] || '';
    if (ownerId && ownerId !== session.id) {
      return jsonResponse_({ success: false, error: 'このしおりを編集する権限がありません。' });
    }

    const title = rowValues[1] || '（無題のしおり）';
    let sentCount = 0;
    const failed = [];

    emails.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: '【旅のしおり】「' + title + '」が共有されました',
          body:
            (session.nickname || '') + ' さんから旅のしおりが共有されました。\n\n' +
            'タイトル: ' + title + '\n\n' +
            '以下のURLから、ログインなしで閲覧できます。\n' + shareUrl + '\n\n' +
            '※このメールに心当たりがない場合は破棄してください。'
        });
        sentCount++;
      } catch (mailErr) {
        failed.push(email);
      }
    });

    // 送信できた相手を shared_with に記録する（既存の記録とマージ・重複除去、上限あり）
    const existing = safeParseJson_(rowValues[7], []);
    const merged = sanitizeEmailList_(existing.concat(emails.filter(email => failed.indexOf(email) === -1)));
    sheet.getRange(rowIndex, 8).setValue(JSON.stringify(merged));

    if (sentCount === 0) {
      return jsonResponse_({ success: false, error: 'メール送信に失敗しました。' });
    }

    return jsonResponse_({ success: true, sent: sentCount, failed: failed, sharedWith: merged });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}
function getMembersSheet_() {
  const ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
    sheet.appendRow(MEMBERS_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * email に一致する会員の行番号（1始まり、ヘッダー込み）を返す。大文字小文字・前後空白は無視する。
 * 見つからない場合は -1。
 */
function findMemberRowByEmail_(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const normalized = (email || '').toString().trim().toLowerCase();
  const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if ((emails[i][0] || '').toString().trim().toLowerCase() === normalized) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * member_id に一致する会員の行番号（1始まり、ヘッダー込み）を返す。
 * 見つからない場合は -1。
 */
function findMemberRowById_(sheet, memberId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === memberId) return i + 2;
  }
  return -1;
}

/**
 * Members シートの1行分の値配列を会員オブジェクトに変換する。
 */
function memberRowToObject_(row) {
  return {
    id: row[0],
    nickname: row[1],
    email: row[2],
    passwordHash: row[3],
    salt: row[4],
    mustChangePassword: row[5] === true || row[5] === 'TRUE' || row[5] === 'true',
    failedAttempts: Number(row[6]) || 0,
    lockedUntil: row[7] || '',
    createdAt: row[8],
    updatedAt: row[9]
  };
}

/**
 * スクリプトプロパティに設定した pepper（アプリ全体共通の秘密文字列）を取得する。
 * 未設定の場合はエラーを投げる（DEPLOY.md参照）。
 */
function getPepper_() {
  const pepper = PropertiesService.getScriptProperties().getProperty('AUTH_PEPPER');
  if (!pepper) {
    throw new Error('AUTH_PEPPER が設定されていません。スクリプトプロパティに設定してください（DEPLOY.md参照）。');
  }
  return pepper;
}

/**
 * パスワードごとのランダムなソルトを生成する。
 */
function generateSalt_() {
  return Utilities.getUuid() + Utilities.getUuid();
}

/**
 * パスワードを SHA-256（salt + pepper付き）でハッシュ化する。
 */
function hashPassword_(password, salt) {
  const pepper = getPepper_();
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + ':' + salt + ':' + pepper,
    Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/**
 * 読みやすい8桁の仮パスワードを生成する（紛らわしい文字 0/O/1/I/L 等は除外）。
 */
function generateTempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let out = '';
  for (let i = 0; i < 8; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

/**
 * セッションを作成し、トークンを返す。セッション情報は CacheService に保存する
 * （GASのWebアプリはCookieを扱いにくいため、トークンはクライアント側で保持し、
 * 以後のリクエストに明示的に付与してもらう方式にしている）。
 */
function createSession_(memberId, nickname, email, mustChangePassword) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    'session_' + token,
    JSON.stringify({ id: memberId, nickname: nickname, email: email, mustChangePassword: !!mustChangePassword }),
    SESSION_TTL_SECONDS
  );
  return token;
}

/**
 * トークンからセッション情報（会員情報）を取得する。無効・期限切れの場合は null。
 */
function getSessionMember_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

/**
 * セッションを破棄する（ログアウト）。
 */
function destroySession_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('session_' + token);
}

/* ======================================================================
 * AIによるリスト作成支援（タイムライン・持ち物）
 * ====================================================================== */

/**
 * POST body:
 * {
 *   action: 'suggest',
 *   type: 'timeline' | 'items',
 *   title: '沖縄旅行',
 *   dates: '2026/08/01 - 2026/08/03',
 *   existing: ['ホテルチェックイン', ...]   // 既存の内容（重複回避のため）
 * }
 */
function handleSuggest_(body) {
  try {
    const type = body.type;
    const title = (body.title || '').toString().trim().slice(0, 200);
    const dates = (body.dates || '').toString().trim().slice(0, 100);
    const existing = Array.isArray(body.existing)
      ? body.existing.map(v => (v || '').toString().trim()).filter(Boolean).slice(0, 200)
      : [];

    if (!title) {
      return jsonResponse_({ success: false, error: '旅行タイトル（行き先）を入力してから提案を使ってください。' });
    }

    let suggestions;
    if (type === 'timeline') {
      suggestions = generateTimelineSuggestions_(title, dates, existing);
    } else if (type === 'items') {
      suggestions = generateItemSuggestions_(title, dates, existing);
    } else {
      return jsonResponse_({ success: false, error: '不明な提案タイプです。' });
    }

    return jsonResponse_({ success: true, suggestions: suggestions });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

function generateTimelineSuggestions_(title, dates, existingContents) {
  const prompt = [
    'あなたは旅行プランナーです。以下の旅行にふさわしいタイムライン（予定表）の項目を提案してください。',
    '',
    '旅行タイトル: ' + title,
    '日程: ' + (dates || '未入力'),
    'すでに登録済みの予定: ' + (existingContents.length ? existingContents.join('、') : 'なし'),
    '',
    '条件:',
    '- 現実的で具体的な予定を6〜8件、時刻とセットで提案してください。',
    '- 時刻は24時間表記（例: 09:00）にしてください。',
    '- 内容は20文字程度の簡潔な表現にしてください（例: ホテルチェックイン、〇〇観光、△△で昼食）。',
    '- 移動・観光・食事・休憩などをバランス良く含めてください。',
    '- 登録済みの予定と内容が重複するものは含めないでください。',
    '- 日程が複数日にわたる場合は、代表的な1日分の行程としてください。'
  ].join('\n');

  const schema = {
    type: 'ARRAY',
    items: {
      type: 'OBJECT',
      properties: {
        time: { type: 'STRING' },
        content: { type: 'STRING' }
      },
      required: ['content']
    }
  };

  const raw = callGemini_(prompt, schema);
  return sanitizeSuggestions_(raw, 'content', existingContents);
}

function generateItemSuggestions_(title, dates, existingTexts) {
  const prompt = [
    'あなたは旅行の持ち物アドバイザーです。以下の旅行にふさわしい持ち物リストの項目を提案してください。',
    '',
    '旅行タイトル: ' + title,
    '日程: ' + (dates || '未入力'),
    'すでに登録済みの持ち物: ' + (existingTexts.length ? existingTexts.join('、') : 'なし'),
    '',
    '条件:',
    '- 行き先・季節（日程から推測できる場合）・旅行の目的にふさわしい持ち物を8〜12件提案してください。',
    '- 各項目は持ち物の名称のみにしてください（説明文は不要）。',
    '- 登録済みの持ち物と内容が重複するものは含めないでください。'
  ].join('\n');

  const schema = {
    type: 'ARRAY',
    items: {
      type: 'OBJECT',
      properties: {
        text: { type: 'STRING' }
      },
      required: ['text']
    }
  };

  const raw = callGemini_(prompt, schema);
  return sanitizeSuggestions_(raw, 'text', existingTexts);
}

/**
 * Gemini APIを呼び出し、JSON形式のレスポンス本文をパースして返す。
 * responseSchemaにより出力形式を強制しているため、追加のテキスト整形は行わない。
 */
function callGemini_(prompt, schema) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY が設定されていません。スクリプトプロパティに設定してください。');
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: schema
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const bodyText = res.getContentText();
  if (code !== 200) {
    throw new Error('Gemini APIエラー（' + code + '）: ' + bodyText.slice(0, 300));
  }

  const json = JSON.parse(bodyText);
  const text = json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;

  if (!text) {
    throw new Error('Geminiから有効な応答が得られませんでした。');
  }

  return JSON.parse(text);
}

/**
 * Geminiの生の出力を検証・正規化する。
 * - 配列でない/空文字/欠損は除外
 * - 500文字（サーバー・クライアント共通の上限）に切り詰め
 * - 既存項目・レスポンス内での重複を除外（大文字小文字・前後空白を無視して比較）
 * - MAX_SUGGESTIONS件まで
 * - field==='content' の場合、time は HH:MM 形式でなければ空文字にする
 */
function sanitizeSuggestions_(raw, field, existingList) {
  if (!Array.isArray(raw)) return [];

  const seen = new Set(existingList.map(v => v.toLowerCase().trim()));
  const result = [];

  for (let i = 0; i < raw.length && result.length < MAX_SUGGESTIONS; i++) {
    const item = raw[i];
    if (!item) continue;

    let text = (item[field] || '').toString().trim();
    if (!text) continue;
    text = text.slice(0, 500);

    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);

    if (field === 'content') {
      let time = (item.time || '').toString().trim();
      if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) time = '';
      result.push({ time: time, content: text });
    } else {
      result.push({ text: text });
    }
  }

  return result;
}

/* ======================================================================
 * 予定表の読み取り（画像・PDF・Excel/CSV → タイムライン）
 * ====================================================================== */

/**
 * POST body:
 * {
 *   action: 'parseSchedule',
 *   token: 'xxxx',                                   // ログイン必須（AI利用の乱用防止）
 *   title: '出雲旅行',
 *   dates: '2026/11/7-2026/11/9',
 *   files: [{ mimeType: 'image/jpeg', data: '<base64>' }, ...],  // 画像・PDF（最大5件）
 *   text: '# シート：Sheet1\n日付,時刻,内容\n...'        // Excel/CSVを変換したテキスト・貼り付け文字
 * }
 * 返り値: { success: true, entries: [{ date: 'YYYY-MM-DD', time: 'HH:MM', content: '...' }] }
 */
function handleParseSchedule_(body) {
  try {
    const session = getSessionMember_(body.token);
    if (!session) {
      return jsonResponse_({ success: false, error: '予定表の読み込みにはログインが必要です。再ログインしてください。' });
    }

    const files = Array.isArray(body.files) ? body.files.slice(0, SCHEDULE_IMPORT_MAX_FILES) : [];
    const text = (body.text || '').toString().slice(0, SCHEDULE_IMPORT_MAX_TEXT);
    if (files.length === 0 && !text.trim()) {
      return jsonResponse_({ success: false, error: '読み取る内容がありません。' });
    }

    const title = (body.title || '').toString().trim().slice(0, 200);
    const dates = (body.dates || '').toString().trim().slice(0, 100);
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    const prompt = [
      'あなたは旅行の予定表を読み取り、タイムライン形式に変換するアシスタントです。',
      '添付の画像・PDF・表データ（CSV形式に変換済み）から、予定をすべて抜き出してください。',
      '',
      '【旅行情報】',
      '旅行タイトル：' + (title || '（未入力）'),
      '日程：' + (dates || '（未入力）'),
      '今日の日付：' + today,
      '',
      '【ルール】',
      '- 1つの予定を1件とし、時刻順に並べる。',
      '- date は "YYYY-MM-DD"。年が書かれていなければ日程から、それも無ければ今日以降で最も近い年を補う。',
      '- 「1日目」「2日目」「Day1」などは日程の開始日を1日目として日付に変換する。',
      '- 日付が分からない場合は直前の予定と同じ日付、それも無ければ空文字 ""。',
      '- time は "HH:MM"（24時間表記、ゼロ埋め）。「午後2時」→"14:00"。時刻が無ければ ""。',
      '- 「9:00〜10:30」のような時間帯は time に開始時刻を入れ、content の末尾に「（〜10:30）」を付ける。',
      '- content は場所・行動・便名・予約番号・メモなどを簡潔に（60文字程度まで）。',
      '- 手書きやかすれで読めない文字は推測で補い、自信が無い箇所には「(?)」を付ける。',
      '- 予定表と関係のない文字（表の見出し、ページ番号など）は含めない。'
    ].join('\n');

    const parts = [{ text: prompt }];
    for (let i = 0; i < files.length; i++) {
      const f = files[i] || {};
      const mime = (f.mimeType || '').toString().toLowerCase();
      const data = (f.data || '').toString();
      if (SCHEDULE_IMPORT_ALLOWED_MIME.indexOf(mime) === -1) {
        return jsonResponse_({ success: false, error: '対応していないファイル形式です（' + mime + '）。' });
      }
      if (!data || !/^[A-Za-z0-9+/=]+$/.test(data)) {
        return jsonResponse_({ success: false, error: 'ファイルデータが不正です。' });
      }
      parts.push({ inline_data: { mime_type: mime, data: data } });
    }
    if (text.trim()) parts.push({ text: '【予定表データ】\n' + text });

    const schema = {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          date: { type: 'STRING' },
          time: { type: 'STRING' },
          content: { type: 'STRING' }
        },
        required: ['content']
      }
    };

    const raw = callGeminiParts_(parts, schema);
    return jsonResponse_({ success: true, entries: normalizeScheduleEntries_(raw) });
  } catch (err) {
    console.error('parseSchedule error', err);
    return jsonResponse_({ success: false, error: String((err && err.message) || err) });
  }
}

/**
 * 画像・PDFなど複数パーツを含むリクエスト用の Gemini 呼び出し。
 * （テキストのみの callGemini_ と同じ形式でJSONを返す）
 */
function callGeminiParts_(parts, schema) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY が設定されていません。スクリプトプロパティに設定してください。');
  }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);

  const payload = {
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      responseSchema: schema
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const bodyText = res.getContentText();
  if (code === 429) {
    throw new Error('AIが混み合っています。少し時間をおいてお試しください。');
  }
  if (code !== 200) {
    console.error('Gemini error', code, bodyText.slice(0, 1000));
    throw new Error('AIでの読み取りに失敗しました（' + code + '）。');
  }

  const json = JSON.parse(bodyText);
  const outParts = (json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts) || [];
  const text = outParts.map(p => p.text || '').join('').trim();
  if (!text) {
    throw new Error('予定を読み取れませんでした。画像が鮮明か確認してください。');
  }

  try {
    return JSON.parse(text.replace(/```json|```/g, '').trim());
  } catch (e) {
    const m = text.match(/\[[\s\S]*\]/);
    return m ? JSON.parse(m[0]) : [];
  }
}

/**
 * AIの出力を検証・正規化する（date: YYYY-MM-DD / time: HH:MM / content: 500文字まで）。
 */
function normalizeScheduleEntries_(list) {
  if (!Array.isArray(list)) return [];
  const out = [];
  for (let i = 0; i < list.length && out.length < SCHEDULE_IMPORT_MAX_ENTRIES; i++) {
    const e = list[i];
    if (!e) continue;
    const content = (e.content || '').toString().replace(/\s+/g, ' ').trim().slice(0, 500);
    if (!content) continue;
    out.push({
      date: normalizeScheduleDate_(e.date),
      time: normalizeScheduleTime_(e.time),
      content: content
    });
  }
  return out;
}

function normalizeScheduleDate_(v) {
  const m = (v || '').toString().trim().match(/^(\d{4})[-\/.年](\d{1,2})[-\/.月](\d{1,2})/);
  if (!m) return '';
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

function normalizeScheduleTime_(v) {
  const m = (v || '').toString().trim().match(/^(\d{1,2})[:：時](\d{0,2})/);
  if (!m) return '';
  const h = Number(m[1]), mi = Number(m[2] || 0);
  if (h > 23 || mi > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}
