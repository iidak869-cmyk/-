const fs = require('fs');
const fsPromises = fs.promises;
const path = require('path');
const config = require('./config');

const COPY_TIMEOUT_MS = 15000;

function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`timeout: ${label}`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

/**
 * 手順22: 完成した反映用リスト.xlsxをG共有ドライブへコピーする。
 *
 * G:ドライブ（Googleドライブ等のネットワークドライブ想定）が応答しない場合、
 * 同期的なfs呼び出しだとプロセス全体が固まってしまうため、非同期+タイムアウトで
 * 実行し、タイムアウトやエラー時はスクリプト全体を止めずにスキップする。
 */
async function copyReflectExcelToSharedDrive() {
  const src = config.reflectExcelPath;
  if (!fs.existsSync(src)) {
    console.log(`[コピー] コピー元のファイルが見つからないため、共有ドライブへのコピーをスキップしました: ${src}`);
    return null;
  }

  try {
    await withTimeout(
      fsPromises.mkdir(config.sharedDriveDestDir, { recursive: true }),
      COPY_TIMEOUT_MS,
      'mkdir'
    );
    const dest = path.join(config.sharedDriveDestDir, path.basename(src));
    await withTimeout(fsPromises.copyFile(src, dest), COPY_TIMEOUT_MS, 'copyFile');
    console.log(`[コピー] 共有ドライブへコピーしました: ${dest}`);
    return dest;
  } catch (err) {
    console.log(
      `[コピー] 共有ドライブへのコピーに失敗したためスキップしました（${config.sharedDriveDestDir}）: ${err.message}`
    );
    return null;
  }
}

module.exports = { copyReflectExcelToSharedDrive };
