require('dotenv').config();
const path = require('path');

function required(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`環境変数 ${name} が設定されていません（.env を確認してください）`);
  }
  return value;
}

function optional(name, fallback) {
  return process.env[name] || fallback;
}

const config = {
  headless: (process.env.HEADLESS || 'false').toLowerCase() === 'true',
  slowMoMs: Number(process.env.SLOWMO_MS || 200),

  airwork: {
    loginUrl: optional('AIRWORK_LOGIN_URL', 'https://ats.rct.airregi.jp/login'),
    id: () => required('AIRWORK_ID'),
    password: () => required('AIRWORK_PASSWORD'),
  },

  doda: {
    loginUrl: optional('DODA_LOGIN_URL', 'https://xtensor.doda.jp/'),
    id: () => required('DODA_ID'),
    password: () => required('DODA_PASSWORD'),
  },

  // ダウンロードフォルダ（通常はブラウザの既定のダウンロード先＝ユーザーのDownloadsフォルダ）
  downloadDir: optional(
    'DOWNLOAD_DIR',
    path.join(require('os').homedir(), 'Downloads')
  ),

  // 手順7/17: 「応募者リスト」フォルダ（仮パス。後で本番パスに差し替え）
  applicantListDir: optional(
    'APPLICANT_LIST_DIR',
    path.join(__dirname, '..', 'work', '応募者リスト')
  ),

  // 手順18: 反映用リスト.xlsx（仮パス。後で本番パスに差し替え）
  reflectExcelPath: optional(
    'REFLECT_EXCEL_PATH',
    path.join(__dirname, '..', 'work', '反映用リスト.xlsx')
  ),

  // 手順21: 応募者数CSVフォルダのルート（仮パス。後で本番パスに差し替え）
  applicantCsvRootDir: optional(
    'APPLICANT_CSV_ROOT_DIR',
    path.join(__dirname, '..', 'work', '応募者数CSV')
  ),

  // 手順22: コピー先の共有ドライブパス
  sharedDriveDestDir: optional(
    'SHARED_DRIVE_DEST_DIR',
    'G:\\共有ドライブ\\RPA運用-採用課連携'
  ),
};

module.exports = config;
