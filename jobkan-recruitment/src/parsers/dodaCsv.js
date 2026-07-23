const fs = require('fs');
const iconv = require('iconv-lite');
const { parse } = require('csv-parse/sync');

// dodaエクスポートCSVの列インデックス（0始まり）。
// 列の並びが変わった場合はここを実データに合わせて更新すること。
//
// 注意: このCSVには「応募日」という同名ヘッダーが2箇所ある
// （1列目=YYYY/MM/DD形式のフル日付、133列目=「07月23日」のような年無し表記）。
// 名前ベースでパースすると後者に上書きされてしまうため、必ず列インデックスで読むこと。
const COL = {
  APPLICATION_DATE: 1, // 応募日（フル日付）
  LAST_NAME: 6, // 姓漢字
  FIRST_NAME: 7, // 名漢字
  LAST_KANA: 8, // 姓カナ
  FIRST_KANA: 9, // 名カナ
  GENDER: 10, // 性別（"男"/"女"）
  EMAIL: 13, // メールアドレス
  PHONE: 18, // 電話番号1
  DESIRED_LOCATION_FALLBACK: 99, // 希望勤務地名（未回答の場合のフォールバック）
  // 企業からの質問1〜5 / その回答（doda上で企業側が設定した任意質問。
  // 求人によって何番目に入るか変わるためキーワードで検索する）
  QUESTION: [121, 123, 125, 127, 129],
  ANSWER: [122, 124, 126, 128, 130],
};

const GENDER_MAP = { 男: '男性', 女: '女性' };

function findAnswerByKeywords(row, questionCols, answerCols, keywords) {
  for (let i = 0; i < questionCols.length; i += 1) {
    const question = row[questionCols[i]] || '';
    if (keywords.every((kw) => question.includes(kw))) {
      return row[answerCols[i]] || '';
    }
  }
  return '';
}

function parseApplicationDate(dateStr) {
  // "2026/07/23" -> Date(2026,6,23)
  const [y, m, d] = (dateStr || '').split('/').map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

/**
 * 応募者リスト(doda).csv を読み込み、正規化した応募者レコードの配列を返す。
 * dodaはShift_JIS(CP932)で出力される。
 */
function parseDodaCsv(filePath) {
  const buffer = fs.readFileSync(filePath);
  const raw = iconv.decode(buffer, 'cp932');
  const rows = parse(raw, { skip_empty_lines: true });
  const dataRows = rows.slice(1); // 先頭行はヘッダー

  return dataRows.map((row) => {
    const gender = row[COL.GENDER] || '';
    const location =
      findAnswerByKeywords(row, COL.QUESTION, COL.ANSWER, ['勤務地']) ||
      row[COL.DESIRED_LOCATION_FALLBACK] ||
      '';

    return {
      route: 'doda',
      name: [row[COL.LAST_NAME], row[COL.FIRST_NAME]].filter(Boolean).join(' '),
      kana: [row[COL.LAST_KANA], row[COL.FIRST_KANA]].filter(Boolean).join(' '),
      gender: GENDER_MAP[gender] || gender,
      email: row[COL.EMAIL] || '',
      phone: row[COL.PHONE] || '',
      phoneTime: findAnswerByKeywords(row, COL.QUESTION, COL.ANSWER, ['電話', '時間']),
      workLocation: location,
      // TODO(要確認): dodaのCSVには「入社種類」に対応しそうな固定列が見当たらないため、
      // 現状は空欄にしている。何を入れるべきか要確認。
      employmentType: '',
      applicationDate: parseApplicationDate(row[COL.APPLICATION_DATE]),
    };
  });
}

module.exports = { parseDodaCsv };
