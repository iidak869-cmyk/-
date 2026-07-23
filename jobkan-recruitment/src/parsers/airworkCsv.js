const fs = require('fs');
const { parse } = require('csv-parse/sync');

// AirWORKエクスポートCSVの列インデックス（0始まり）。
// 列の並びが変わった場合はここを実データに合わせて更新すること。
const COL = {
  NAME: 1, // 応募者名
  KANA: 2, // ふりがな
  PHONE: 7, // 電話番号
  EMAIL: 8, // メールアドレス
  GENDER: 9, // 性別（"男性"/"女性"）
  WORK_LOCATION: 29, // 応募勤務地
  EMPLOYMENT_TYPE: 32, // 応募雇用形態
  FREE_Q: [20, 22, 24], // 自由質問文1〜3
  FREE_A: [21, 23, 25], // 自由質問への回答1〜3
  APPLIED_AT: 27, // 応募日時（例: "2026/04/20 19:05:02"）
};

function findAnswerByKeywords(row, questionCols, answerCols, keywords) {
  for (let i = 0; i < questionCols.length; i += 1) {
    const question = row[questionCols[i]] || '';
    if (keywords.every((kw) => question.includes(kw))) {
      return row[answerCols[i]] || '';
    }
  }
  return '';
}

function parseApplicationDate(appliedAt) {
  // "2026/04/20 19:05:02" -> Date(2026,3,20)
  const datePart = (appliedAt || '').split(' ')[0];
  const [y, m, d] = datePart.split('/').map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

/**
 * 応募者リスト(AirWORK).csv を読み込み、正規化した応募者レコードの配列を返す。
 * AirWORKはUTF-8(BOM付き)で出力される。
 */
function parseAirworkCsv(filePath) {
  const raw = fs.readFileSync(filePath, 'utf8').replace(/^﻿/, '');
  const rows = parse(raw, { skip_empty_lines: true });
  const dataRows = rows.slice(1); // 先頭行はヘッダー

  return dataRows.map((row) => ({
    route: 'AirWORK',
    name: row[COL.NAME] || '',
    kana: row[COL.KANA] || '',
    gender: row[COL.GENDER] || '',
    email: row[COL.EMAIL] || '',
    phone: row[COL.PHONE] || '',
    phoneTime: findAnswerByKeywords(row, COL.FREE_Q, COL.FREE_A, ['電話', '時間']),
    workLocation: row[COL.WORK_LOCATION] || '',
    // TODO(要確認): 「入社種類」＝応募雇用形態でよいか要確認
    employmentType: row[COL.EMPLOYMENT_TYPE] || '',
    applicationDate: parseApplicationDate(row[COL.APPLIED_AT]),
  }));
}

module.exports = { parseAirworkCsv };
