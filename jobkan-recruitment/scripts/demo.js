#!/usr/bin/env node
/**
 * 反映用リスト.xlsxへの転記結果を目で見て確認するためのデモ／テスト用スクリプト。
 *
 * 実際のフォルダ・タスクスケジューラを用意する前に、手元のサンプルファイルだけで
 * 転記ロジックの結果を確認したいときに使う。--reference-date を省略すると
 * 24時間フィルタをかけず全件を転記するので、サンプルデータの日付に関わらず
 * 結果を目視できる。
 *
 * 使い方:
 *   node scripts/demo.js --template <反映用リスト.xlsxのひな形> --out <出力先.xlsx> \
 *     [--airwork <応募者リスト(AirWORK).csv>] \
 *     [--doda <応募者リスト(doda).csv>] \
 *     [--csv-folder <応募者数CSVフォルダ or その中の1サブフォルダ>] \
 *     [--reference-date 2026-07-23T09:00:00]  # 指定時のみ直近24時間でフィルタする
 */
const fs = require('fs');
const path = require('path');
const { parseAirworkCsv } = require('../src/parsers/airworkCsv');
const { parseDodaCsv } = require('../src/parsers/dodaCsv');
const {
  parseApplicantExportFile,
  findLatestDatedSubfolder,
  listApplicantExportFiles,
} = require('../src/parsers/applicantCsvFolder');
const { filterWithinLast24Hours, appendApplicantsToReflectExcel } = require('../src/excel');

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i].startsWith('--')) {
      const key = argv[i].slice(2);
      args[key] = argv[i + 1];
      i += 1;
    }
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.template || !args.out) {
    console.error(
      '使い方: node scripts/demo.js --template <ひな形.xlsx> --out <出力先.xlsx> ' +
        '[--airwork <csv>] [--doda <csv>] [--csv-folder <フォルダ>] [--reference-date <ISO日時>]'
    );
    process.exitCode = 1;
    return;
  }

  fs.copyFileSync(args.template, args.out);
  console.log(`ひな形をコピーしました: ${args.template} -> ${args.out}`);

  const referenceDate = args['reference-date'] ? new Date(args['reference-date']) : null;
  const applyFilter = (records) =>
    referenceDate ? filterWithinLast24Hours(records, referenceDate) : records;

  if (args.airwork) {
    const records = applyFilter(parseAirworkCsv(args.airwork));
    const count = await appendApplicantsToReflectExcel(args.out, records);
    console.log(`[AirWORK] ${count}件を追記しました（${path.basename(args.airwork)}）`);
  }

  if (args.doda) {
    const records = applyFilter(parseDodaCsv(args.doda));
    const count = await appendApplicantsToReflectExcel(args.out, records);
    console.log(`[doda] ${count}件を追記しました（${path.basename(args.doda)}）`);
  }

  if (args['csv-folder']) {
    const targetDir = findLatestDatedSubfolder(args['csv-folder']) || args['csv-folder'];
    const files = listApplicantExportFiles(targetDir);
    for (const filePath of files) {
      const records = applyFilter(await parseApplicantExportFile(filePath));
      const count = await appendApplicantsToReflectExcel(args.out, records);
      console.log(`[応募者数CSV] ${count}件を追記しました（${path.basename(filePath)}）`);
    }
  }

  console.log(`完了しました。結果を確認してください: ${args.out}`);
}

main().catch((err) => {
  console.error('エラーが発生しました:', err);
  process.exitCode = 1;
});
