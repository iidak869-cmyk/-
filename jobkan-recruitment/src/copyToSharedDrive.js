const fs = require('fs');
const path = require('path');
const config = require('./config');

/**
 * 手順22: 完成した反映用リスト.xlsxをG共有ドライブへコピーする
 */
function copyReflectExcelToSharedDrive() {
  const src = config.reflectExcelPath;
  if (!fs.existsSync(src)) {
    throw new Error(`コピー元のファイルが見つかりません: ${src}`);
  }

  fs.mkdirSync(config.sharedDriveDestDir, { recursive: true });
  const dest = path.join(config.sharedDriveDestDir, path.basename(src));
  fs.copyFileSync(src, dest);
  console.log(`[コピー] 共有ドライブへコピーしました: ${dest}`);
  return dest;
}

module.exports = { copyReflectExcelToSharedDrive };
