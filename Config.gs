/**
 * 月次の報酬・支払明細（見本）― 設定値・列番号・シート名
 *
 * 想定シナリオ（架空）：
 *   架空の教室「サンプル学習塾」。講師20名（架空）。実績入力・単価マスタ・明細出力の
 *   3ファイルを手作業で突合し、毎月PDFを作ってメールしている業務を自動化する。
 *   金額を扱うため、検証（チェック）を厳しく行い、送信は人が確認してから行う
 *   （このスクリプトが作るのはGmailの下書きまでで、コード内に送信の呼び出しは0件。
 *   下書きは人が1通ずつ確認してから送信する）。
 *
 * 設置・運用の手順は SETUP_GUIDE.md を参照。
 */

// ===== シート名 =====
var SHEET_SETTINGS = '設定';
var SHEET_MASTER_INSTRUCTORS = 'マスタ_講師';
var SHEET_MASTER_ITEMS = 'マスタ_品目';
var SHEET_PREFIX_PERFORMANCE = '実績_'; // + yyyyMM（原データ。script は書き込まない）
var SHEET_PREFIX_CHECK = 'チェック結果_';
var SHEET_PREFIX_CALC = '計算結果_'; // 中間表（講師ごとの合計）。値のみ、script が作り直す
var SHEET_PREFIX_DETAIL = '明細_'; // 中間表（品目別の内訳）。値のみ、script が作り直す
var SHEET_SUMMARY = '合計';
var SHEET_ERROR_LOG = 'エラーログ';
var SHEET_AUDIT_LOG = '監査ログ';
var SHEET_DASHBOARD = 'ダッシュボード';

// ===== 列番号（1始まり）。マスタ・入力シートの列順を変えたらここだけ直す =====
var COL_INSTR = { ID: 1, NAME: 2, RANK: 3, EMAIL: 4, STATUS: 5 };
var COL_ITEM = { ID: 1, NAME: 2, UNIT: 3, RATE_A: 4, RATE_B: 5, RATE_C: 6 };
var COL_PERF = { DATE: 1, INSTRUCTOR_ID: 2, ITEM_ID: 3, QTY: 4, NOTE: 5 };
var COL_CALC = { INSTRUCTOR_ID: 1, NAME: 2, RANK: 3, QTY_TOTAL: 4, SUBTOTAL: 5, TAX: 6, TOTAL: 7, STATUS: 8 };
var COL_DETAIL = { INSTRUCTOR_ID: 1, ITEM_ID: 2, ITEM_NAME: 3, UNIT_PRICE: 4, QTY: 5, AMOUNT: 6 };
var COL_SUMMARY = { MONTH: 1, INSTRUCTOR_ID: 2, NAME: 3, TOTAL: 4, TRANSFER_CHECK: 5, STATUS: 6, UPDATED_AT: 7 };

var SEVERITY_ERROR = 'エラー';
var SEVERITY_WARNING = '注意';

var CHECK_TYPE = {
  UNKNOWN_INSTRUCTOR: '未登録講師',
  INACTIVE_INSTRUCTOR: '無効講師',
  UNKNOWN_ITEM: '未登録品目',
  MISSING_RATE: '単価欠落',
  NEGATIVE_QTY: '負数',
  INVALID_QTY: '数量不正',
  DUPLICATE_ROW: '重複行',
  PREV_MONTH_ANOMALY: '前月比異常'
};

var INSTRUCTOR_STATUS_ACTIVE = '有効';

var ROW_STATUS = {
  NOT_CALCULATED: '未計算',
  CALCULATED: '計算済',
  PDF_DONE: 'PDF作成済',
  DRAFTED: '下書き済',
  FINALIZED: '確定',
  SENT: '送付済' // このステータスは script が自動で付けない。人が送信後に手動で選ぶ（SETUP_GUIDE参照）
};

var MONTH_STATE = {
  NONE: 'NONE',
  CHECKED_OK: 'CHECKED_OK',
  CHECKED_NG: 'CHECKED_NG',
  CALCULATED: 'CALCULATED',
  PDF_DONE: 'PDF_DONE',
  DRAFTED: 'DRAFTED',
  FINALIZED: 'FINALIZED'
};

// PropertiesService キーのプレフィックス
var PROP_STATE_PREFIX = 'STATE_'; // + yyyyMM
var PROP_PROGRESS_PREFIX = 'PROGRESS_'; // + yyyyMM + '_' + STEP
var PROP_PDF_IDS_PREFIX = 'PDF_IDS_'; // + yyyyMM
var PROP_DRAFT_IDS_PREFIX = 'DRAFT_IDS_'; // + yyyyMM
var PROP_TEMPLATE_DOC_ID = 'TEMPLATE_DOC_ID';
var PROP_PDF_ROOT_FOLDER_ID = 'PDF_ROOT_FOLDER_ID';
var PROP_FINALIZED_MONTHS = 'FINALIZED_MONTHS'; // JSON配列
var PROP_DAILY_COUNT_PREFIX = 'DAILY_COUNT_'; // + 'PDF'|'DRAFT' + '_' + yyyy-MM-dd

var DEFAULT_DAILY_PDF_LIMIT = 400; // Drive作成（PDF）の1日あたりの既定上限（控えめな値）
var DEFAULT_DAILY_DRAFT_LIMIT = 400; // Gmail下書きの1日あたりの既定上限（控えめな値）

var CONTINUATION_TRIGGER_HANDLER = 'resumeBatch';
var BATCH_TIME_LIMIT_MS = 4.5 * 60 * 1000; // 6分制限に対して余裕を取る
var LOCK_WAIT_MS = 30 * 1000;

var PDF_ROOT_FOLDER_NAME = '報酬明細PDF';
var BACKUP_SHEET_PREFIX = 'BK_'; // 実行前バックアップのシート名の先頭（同じファイル内・非表示）
var SCHOOL_DISPLAY_NAME = 'Sample School'; // 架空のロゴ枠文言
