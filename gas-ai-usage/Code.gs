/**
 * 社内のAI利用状況を可視化するスクリプト。
 *
 * Googleフォームで社内メンバーからAIの利用状況を集め、同じスプレッドシート上に
 * 「誰が・どのツールを・どんな業務に・どれくらい使っているか」を集計したシートと
 * グラフ付きダッシュボードを自動生成する。追加費用はかからない。
 *
 * 事前準備:
 * 1. setup() を一度だけ実行する
 *      - スプレッドシート「AI利用状況ダッシュボード」とGoogleフォーム「AI利用状況アンケート」を新規作成
 *      - フォーム送信時と毎朝8時の自動集計トリガーを登録
 *      - 初回の集計を実行
 * 2. 実行ログに表示されるフォームURLを社内に共有する
 * 3. `設定` シートにメンバー・部署・集計期間を入力し、applySettingsToForm() を実行して
 *    フォームの選択肢に反映する
 *
 * 以降はフォームに回答が入るたび、および毎朝8時に refreshDashboard() が自動で集計し直す。
 */

var SPREADSHEET_NAME = 'AI利用状況ダッシュボード';
var FORM_NAME = 'AI利用状況アンケート';

var SHEET = {
  DASHBOARD: 'ダッシュボード',
  BY_TOOL: '集計_ツール別',
  BY_PURPOSE: '集計_用途別',
  BY_MEMBER: '集計_メンバー別',
  BY_WEEK: '集計_週次推移',
  KNOWLEDGE: '共有ナレッジ',
  SETTINGS: '設定',
  RAW: '回答_生データ'
};

// フォームの質問文。回答シートの見出しと一致するため、変更する場合は既存の回答シートの
// 見出しも合わせて直すこと。
var Q = {
  NAME: '氏名',
  DEPARTMENT: '部署',
  TARGET_DATE: '対象日（いつの利用についての回答ですか）',
  TOOLS: '使ったAIツール（複数選択可）',
  PURPOSES: '主な用途（複数選択可）',
  USAGE_TIME: '1日あたりのAI利用時間',
  SAVED_TIME: '削減できたと感じる作業時間（1日あたり）',
  SKILL: 'AIを使いこなせている実感',
  GOOD: 'うまくいった使い方・共有したいプロンプト（任意）',
  ISSUE: '困っていること・改善してほしいこと（任意）'
};

var EMAIL_HEADER = 'メールアドレス';

var TOOL_CHOICES = [
  'ChatGPT',
  'Claude（チャット）',
  'Claude Code',
  'Gemini',
  'Microsoft Copilot',
  'GitHub Copilot',
  'NotebookLM',
  'Perplexity',
  '画像・動画生成AI',
  '文字起こし・議事録AI'
];

var PURPOSE_CHOICES = [
  '文章作成・添削',
  '資料・提案書づくり',
  'メール・チャットの返信',
  '議事録の要約',
  '情報収集・調査',
  '翻訳',
  'データ集計・分析',
  'プログラミング・業務自動化',
  'アイデア出し・企画',
  '顧客対応'
];

// 選択肢と、集計で使う代表値（時間）の対応。
var USAGE_TIME_OPTIONS = [
  { label: 'ほとんど使わなかった', hours: 0 },
  { label: '15分未満', hours: 0.125 },
  { label: '15〜30分', hours: 0.375 },
  { label: '30分〜1時間', hours: 0.75 },
  { label: '1〜2時間', hours: 1.5 },
  { label: '2〜4時間', hours: 3 },
  { label: '4時間以上', hours: 5 }
];

var SAVED_TIME_OPTIONS = [
  { label: '削減できていない', hours: 0 },
  { label: '15分未満', hours: 0.125 },
  { label: '15〜30分', hours: 0.375 },
  { label: '30分〜1時間', hours: 0.75 },
  { label: '1〜2時間', hours: 1.5 },
  { label: '2時間以上', hours: 3 }
];

var DEFAULT_MEMBERS = ['飯田 圭祐'];
var DEFAULT_DEPARTMENTS = ['営業', '制作', '開発', '管理・バックオフィス'];
var DEFAULT_PERIOD_DAYS = 90;

var TIME_ZONE = 'Asia/Tokyo';

// ---------------------------------------------------------------- セットアップ

function setup() {
  var props = PropertiesService.getScriptProperties();

  var spreadsheetId = props.getProperty('AI_USAGE_SPREADSHEET_ID');
  var ss;
  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log('既存のスプレッドシートを使用します: ' + ss.getUrl());
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    props.setProperty('AI_USAGE_SPREADSHEET_ID', ss.getId());
    Logger.log('スプレッドシートを新規作成しました: ' + ss.getUrl());
  }

  var settings = readSettings(ss);

  var formId = props.getProperty('AI_USAGE_FORM_ID');
  var form;
  if (formId) {
    form = FormApp.openById(formId);
    Logger.log('既存のフォームを使用します: ' + form.getPublishedUrl());
  } else {
    form = createForm(settings);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    props.setProperty('AI_USAGE_FORM_ID', form.getId());
    SpreadsheetApp.flush();
    Logger.log('フォームを新規作成しました（社内に共有するURL）: ' + form.getPublishedUrl());
  }

  applyChoicesToForm(form, settings);
  ensureRawSheet(ss);
  registerTriggers(ss);
  refreshDashboard();
  removeDefaultSheet(ss);

  Logger.log('--- セットアップ完了 ---');
  Logger.log('ダッシュボード: ' + ss.getUrl());
  Logger.log('回答フォーム（社内共有用）: ' + form.getPublishedUrl());
}

function registerTriggers(ss) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshDashboard').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyDays(1).atHour(8).create();
  Logger.log('フォーム送信時と毎日8時の自動集計トリガーを登録しました。');
}

/** `設定` シートの内容をフォームの選択肢（氏名・部署）に反映する。 */
function applySettingsToForm() {
  var ss = openSpreadsheet();
  var form = openForm();
  applyChoicesToForm(form, readSettings(ss));
  Logger.log('フォームの氏名・部署の選択肢を `設定` シートの内容に更新しました。');
}

// ---------------------------------------------------------------- フォーム作成

function createForm(settings) {
  var form = FormApp.create(FORM_NAME);
  form.setTitle(FORM_NAME);
  form.setDescription(
    'ふだんのAI活用状況を教えてください（1〜2分）。\n' +
      '社内でどんな使い方が広がっているかを共有し、うまくいっている使い方を全員に展開するための集計に使います。\n' +
      '個人の評価には使いません。'
  );
  try {
    form.setCollectEmail(true);
  } catch (e) {
    Logger.log('メールアドレスの自動収集は有効化できませんでした（氏名の回答で集計します）: ' + e);
  }

  form
    .addMultipleChoiceItem()
    .setTitle(Q.NAME)
    .setChoiceValues(settings.members)
    .showOtherOption(true)
    .setRequired(true);

  form
    .addMultipleChoiceItem()
    .setTitle(Q.DEPARTMENT)
    .setChoiceValues(settings.departments)
    .showOtherOption(true)
    .setRequired(true);

  form
    .addDateItem()
    .setTitle(Q.TARGET_DATE)
    .setHelpText('直近の代表的な1日を選んでください。')
    .setIncludesYear(true)
    .setRequired(true);

  form
    .addCheckboxItem()
    .setTitle(Q.TOOLS)
    .setChoiceValues(TOOL_CHOICES)
    .showOtherOption(true)
    .setRequired(true);

  form
    .addCheckboxItem()
    .setTitle(Q.PURPOSES)
    .setChoiceValues(PURPOSE_CHOICES)
    .showOtherOption(true)
    .setRequired(true);

  form
    .addMultipleChoiceItem()
    .setTitle(Q.USAGE_TIME)
    .setChoiceValues(optionLabels(USAGE_TIME_OPTIONS))
    .setRequired(true);

  form
    .addMultipleChoiceItem()
    .setTitle(Q.SAVED_TIME)
    .setHelpText('AIを使わなかった場合と比べて、どれくらい作業時間が短くなった感覚がありますか。')
    .setChoiceValues(optionLabels(SAVED_TIME_OPTIONS))
    .setRequired(true);

  form
    .addScaleItem()
    .setTitle(Q.SKILL)
    .setBounds(1, 5)
    .setLabels('うまく使えていない', 'かなり使いこなせている')
    .setRequired(true);

  form.addParagraphTextItem().setTitle(Q.GOOD).setHelpText('他の人にも役立ちそうな使い方があれば教えてください。');
  form.addParagraphTextItem().setTitle(Q.ISSUE);

  return form;
}

function applyChoicesToForm(form, settings) {
  updateMultipleChoice(form, Q.NAME, settings.members);
  updateMultipleChoice(form, Q.DEPARTMENT, settings.departments);
}

function updateMultipleChoice(form, title, values) {
  if (!values || values.length === 0) return;
  var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) {
      items[i].asMultipleChoiceItem().setChoiceValues(values).showOtherOption(true);
      return;
    }
  }
  Logger.log('フォームに「' + title + '」の質問が見つかりませんでした。');
}

// ---------------------------------------------------------------- 集計の実行

function refreshDashboard() {
  var ss = openSpreadsheet();
  var settings = readSettings(ss);
  var rawSheet = ensureRawSheet(ss);

  var allRows = readResponses(rawSheet);
  var range = periodRange(settings.periodDays);
  var rows = allRows.filter(function (r) {
    return r.date >= range.from && r.date <= range.to;
  });

  var stats = aggregate(rows, settings);

  writeToolSheet(ss, stats);
  writePurposeSheet(ss, stats);
  writeMemberSheet(ss, stats);
  writeWeekSheet(ss, stats);
  writeKnowledgeSheet(ss, rows);
  buildDashboard(ss, stats, settings, range, allRows.length);

  Logger.log(
    '集計しました: 対象期間 ' +
      formatDate(range.from) +
      '〜' +
      formatDate(range.to) +
      '（' +
      rows.length +
      '件 / 全' +
      allRows.length +
      '件）'
  );
}

function periodRange(periodDays) {
  var now = new Date();
  var to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  var from = new Date(to.getFullYear(), to.getMonth(), to.getDate() - (periodDays - 1));
  return { from: from, to: to };
}

function aggregate(rows, settings) {
  var byTool = {};
  var byPurpose = {};
  var byMember = {};
  var byWeek = {};
  var respondents = {};

  var totalHours = 0;
  var totalSaved = 0;
  var skillSum = 0;
  var skillCount = 0;

  rows.forEach(function (r) {
    totalHours += r.hours;
    totalSaved += r.savedHours;
    if (r.skill > 0) {
      skillSum += r.skill;
      skillCount++;
    }
    respondents[r.name] = true;

    r.tools.forEach(function (tool) {
      var t = byTool[tool] || (byTool[tool] = { count: 0, hours: 0, members: {} });
      t.count++;
      t.hours += r.hours;
      t.members[r.name] = true;
    });

    r.purposes.forEach(function (purpose) {
      var p = byPurpose[purpose] || (byPurpose[purpose] = { count: 0, savedHours: 0, members: {} });
      p.count++;
      p.savedHours += r.savedHours;
      p.members[r.name] = true;
    });

    var m =
      byMember[r.name] ||
      (byMember[r.name] = {
        department: r.department,
        count: 0,
        hours: 0,
        savedHours: 0,
        skillSum: 0,
        skillCount: 0,
        tools: {},
        purposes: {},
        lastDate: null
      });
    m.count++;
    m.hours += r.hours;
    m.savedHours += r.savedHours;
    if (r.department) m.department = r.department;
    if (r.skill > 0) {
      m.skillSum += r.skill;
      m.skillCount++;
    }
    r.tools.forEach(function (tool) {
      m.tools[tool] = true;
    });
    r.purposes.forEach(function (purpose) {
      m.purposes[purpose] = true;
    });
    if (!m.lastDate || r.date > m.lastDate) m.lastDate = r.date;

    var ws = weekStart(r.date);
    var key = formatDate(ws, 'yyyy-MM-dd');
    var w = byWeek[key] || (byWeek[key] = { start: ws, count: 0, hours: 0, savedHours: 0, members: {} });
    w.count++;
    w.hours += r.hours;
    w.savedHours += r.savedHours;
    w.members[r.name] = true;
  });

  var respondentNames = Object.keys(respondents);
  var notReported = settings.members.filter(function (name) {
    return !respondents[name];
  });
  // カバー率は「登録メンバーのうち申告した人」で計算する。`設定` シートに載っていない人が
  // 回答すると100%を超えてしまうため、登録外の回答者は別のKPIとして出す。
  var registeredRespondents = settings.members.filter(function (name) {
    return respondents[name];
  });
  var unknownRespondents = respondentNames.filter(function (name) {
    return settings.members.indexOf(name) === -1;
  });

  return {
    rowCount: rows.length,
    byTool: byTool,
    byPurpose: byPurpose,
    byMember: byMember,
    byWeek: byWeek,
    respondentCount: respondentNames.length,
    registeredRespondentCount: registeredRespondents.length,
    unknownRespondents: unknownRespondents,
    notReported: notReported,
    totalHours: totalHours,
    totalSavedHours: totalSaved,
    averageSkill: skillCount > 0 ? skillSum / skillCount : 0,
    topTool: topKeyByCount(byTool),
    topPurpose: topKeyByCount(byPurpose)
  };
}

function topKeyByCount(map) {
  var best = '';
  var bestCount = -1;
  Object.keys(map).forEach(function (k) {
    if (map[k].count > bestCount) {
      bestCount = map[k].count;
      best = k;
    }
  });
  return best ? best + '（' + bestCount + '件）' : '—';
}

// ---------------------------------------------------------------- 回答の読み取り

function readResponses(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function (h) {
    return String(h).trim();
  });
  var idx = {};
  header.forEach(function (h, i) {
    if (idx[h] === undefined) idx[h] = i;
  });

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (
      row.join('').toString().trim() === ''
    ) {
      continue;
    }

    var timestamp = toDate(row[0]);
    var date = toDate(cell(row, idx, Q.TARGET_DATE)) || timestamp;
    if (!date) continue;

    var name = String(cell(row, idx, Q.NAME) || '').trim();
    if (!name) name = String(cell(row, idx, EMAIL_HEADER) || '').trim();
    if (!name) name = '（氏名未記入）';

    rows.push({
      timestamp: timestamp,
      date: date,
      name: name,
      department: String(cell(row, idx, Q.DEPARTMENT) || '').trim(),
      tools: splitMulti(cell(row, idx, Q.TOOLS)),
      purposes: splitMulti(cell(row, idx, Q.PURPOSES)),
      usageTimeLabel: String(cell(row, idx, Q.USAGE_TIME) || '').trim(),
      hours: optionHours(USAGE_TIME_OPTIONS, cell(row, idx, Q.USAGE_TIME)),
      savedHours: optionHours(SAVED_TIME_OPTIONS, cell(row, idx, Q.SAVED_TIME)),
      skill: Number(cell(row, idx, Q.SKILL)) || 0,
      good: String(cell(row, idx, Q.GOOD) || '').trim(),
      issue: String(cell(row, idx, Q.ISSUE) || '').trim()
    });
  }
  return rows;
}

function cell(row, idx, header) {
  var i = idx[header];
  return i === undefined ? '' : row[i];
}

function splitMulti(value) {
  if (value === '' || value === null || value === undefined) return [];
  return String(value)
    .split(/\s*,\s*/)
    .map(function (v) {
      return v.trim();
    })
    .filter(function (v) {
      return v !== '';
    });
}

function optionLabels(options) {
  return options.map(function (o) {
    return o.label;
  });
}

function optionHours(options, label) {
  var text = String(label || '').trim();
  for (var i = 0; i < options.length; i++) {
    if (options[i].label === text) return options[i].hours;
  }
  return 0;
}

// ---------------------------------------------------------------- 集計シート

function writeToolSheet(ss, stats) {
  var rows = Object.keys(stats.byTool)
    .map(function (name) {
      var t = stats.byTool[name];
      return [name, t.count, Object.keys(t.members).length, t.hours];
    })
    .sort(function (a, b) {
      return b[1] - a[1] || b[3] - a[3] || String(a[0]).localeCompare(String(b[0]));
    });

  return writeTable(ss, SHEET.BY_TOOL, ['AIツール', '申告件数', '利用人数', '利用時間の合計(h)'], rows, {
    4: '0.0'
  });
}

function writePurposeSheet(ss, stats) {
  var rows = Object.keys(stats.byPurpose)
    .map(function (name) {
      var p = stats.byPurpose[name];
      return [name, p.count, Object.keys(p.members).length, p.savedHours];
    })
    .sort(function (a, b) {
      return b[1] - a[1] || b[3] - a[3] || String(a[0]).localeCompare(String(b[0]));
    });

  return writeTable(ss, SHEET.BY_PURPOSE, ['用途', '申告件数', '利用人数', '削減時間の合計(h)'], rows, {
    4: '0.0'
  });
}

function writeMemberSheet(ss, stats) {
  var rows = Object.keys(stats.byMember)
    .map(function (name) {
      var m = stats.byMember[name];
      return [
        name,
        m.department,
        m.count,
        m.hours,
        m.savedHours,
        m.skillCount > 0 ? m.skillSum / m.skillCount : '',
        Object.keys(m.tools).join('、'),
        Object.keys(m.purposes).join('、'),
        m.lastDate
      ];
    })
    .sort(function (a, b) {
      return b[3] - a[3] || b[2] - a[2] || String(a[0]).localeCompare(String(b[0]));
    });

  return writeTable(
    ss,
    SHEET.BY_MEMBER,
    ['氏名', '部署', '申告件数', '利用時間(h)', '削減時間(h)', '使いこなし度', '使用ツール', '主な用途', '最終申告日'],
    rows,
    { 4: '0.0', 5: '0.0', 6: '0.0', 9: 'yyyy/MM/dd' }
  );
}

function writeWeekSheet(ss, stats) {
  var rows = Object.keys(stats.byWeek)
    .sort()
    .map(function (key) {
      var w = stats.byWeek[key];
      return [formatDate(w.start, 'MM/dd') + '週', w.count, Object.keys(w.members).length, w.hours, w.savedHours];
    });

  return writeTable(
    ss,
    SHEET.BY_WEEK,
    ['週（月曜始まり）', '申告件数', '申告人数', '利用時間(h)', '削減時間(h)'],
    rows,
    { 4: '0.0', 5: '0.0' }
  );
}

function writeKnowledgeSheet(ss, rows) {
  var out = [];
  rows
    .slice()
    .sort(function (a, b) {
      return b.date - a.date;
    })
    .forEach(function (r) {
      if (r.good) out.push([r.date, r.name, r.department, 'うまくいった使い方', r.good]);
      if (r.issue) out.push([r.date, r.name, r.department, '困っていること', r.issue]);
    });

  var sheet = writeTable(ss, SHEET.KNOWLEDGE, ['対象日', '氏名', '部署', '種別', '内容'], out, {
    1: 'yyyy/MM/dd'
  });
  sheet.setColumnWidth(5, 640);
  if (out.length > 0) {
    sheet.getRange(2, 5, out.length, 1).setWrap(true);
  }
  return sheet;
}

function writeTable(ss, name, header, rows, numberFormats) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();

  sheet
    .getRange(1, 1, 1, header.length)
    .setValues([header])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  sheet.setFrozenRows(1);

  if (numberFormats && rows.length > 0) {
    Object.keys(numberFormats).forEach(function (col) {
      sheet.getRange(2, Number(col), rows.length, 1).setNumberFormat(numberFormats[col]);
    });
  }
  sheet.autoResizeColumns(1, header.length);
  return sheet;
}

// ---------------------------------------------------------------- ダッシュボード

function buildDashboard(ss, stats, settings, range, totalRowCount) {
  var sheet = ss.getSheetByName(SHEET.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEET.DASHBOARD, 0);

  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });
  sheet.clear();

  sheet.getRange('A1').setValue('AI利用状況ダッシュボード').setFontSize(18).setFontWeight('bold');
  sheet
    .getRange('A2')
    .setValue(
      '最終更新: ' +
        formatDate(new Date(), 'yyyy/MM/dd HH:mm') +
        '　／　集計対象: 直近' +
        settings.periodDays +
        '日（' +
        formatDate(range.from) +
        '〜' +
        formatDate(range.to) +
        '）　／　全期間の回答数: ' +
        totalRowCount +
        '件'
    )
    .setFontColor('#5f6368');

  var memberCount = settings.members.length;
  var kpis = [
    ['申告件数（対象期間）', stats.rowCount + ' 件'],
    ['申告したメンバー', stats.respondentCount + ' 名'],
    ['登録メンバー', memberCount + ' 名'],
    [
      '申告カバー率（登録メンバー中）',
      memberCount > 0 ? Math.round((stats.registeredRespondentCount / memberCount) * 100) + ' %' : '—'
    ],
    ['AI利用時間の合計（体感）', round1(stats.totalHours) + ' 時間'],
    ['削減できた作業時間の合計（体感）', round1(stats.totalSavedHours) + ' 時間'],
    ['1人あたり平均利用時間', stats.respondentCount > 0 ? round1(stats.totalHours / stats.respondentCount) + ' 時間' : '—'],
    ['平均 使いこなし度（5点満点）', stats.averageSkill > 0 ? round1(stats.averageSkill) : '—'],
    ['最も使われているツール', stats.topTool],
    ['最も多い用途', stats.topPurpose],
    ['未申告のメンバー', stats.notReported.length > 0 ? stats.notReported.join('、') : 'なし（全員が申告済み）'],
    [
      '`設定` シートに未登録の回答者',
      stats.unknownRespondents.length > 0 ? stats.unknownRespondents.join('、') : 'なし'
    ]
  ];

  sheet
    .getRange(4, 1, 1, 2)
    .setValues([['指標', '値']])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  sheet.getRange(5, 1, kpis.length, 2).setValues(kpis);
  sheet.getRange(5, 1, kpis.length, 1).setFontWeight('bold');
  sheet.getRange(5, 2, kpis.length, 1).setWrap(true);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 260);

  if (stats.rowCount === 0) {
    sheet
      .getRange('A' + (5 + kpis.length + 1))
      .setValue('対象期間の回答がまだありません。フォームURLを共有して回答を集めてください。')
      .setFontColor('#b06000');
  }

  var toolRows = countRows(ss, SHEET.BY_TOOL);
  if (toolRows > 0) {
    var toolSheet = ss.getSheetByName(SHEET.BY_TOOL);
    insertChart(
      sheet,
      Charts.ChartType.COLUMN,
      [toolSheet.getRange(1, 1, toolRows + 1, 2), toolSheet.getRange(1, 3, toolRows + 1, 1)],
      'ツール別の利用状況（申告件数・利用人数）',
      4,
      4
    );
  }

  var purposeRows = countRows(ss, SHEET.BY_PURPOSE);
  if (purposeRows > 0) {
    var purposeSheet = ss.getSheetByName(SHEET.BY_PURPOSE);
    insertChart(
      sheet,
      Charts.ChartType.BAR,
      [purposeSheet.getRange(1, 1, purposeRows + 1, 2)],
      '用途別の申告件数（どんな業務に使われているか）',
      4,
      11
    );
  }

  var weekRows = countRows(ss, SHEET.BY_WEEK);
  if (weekRows > 0) {
    var weekSheet = ss.getSheetByName(SHEET.BY_WEEK);
    insertChart(
      sheet,
      Charts.ChartType.LINE,
      [weekSheet.getRange(1, 1, weekRows + 1, 1), weekSheet.getRange(1, 3, weekRows + 1, 3)],
      '週次推移（申告人数・利用時間・削減時間）',
      22,
      4
    );
  }

  var memberRows = countRows(ss, SHEET.BY_MEMBER);
  if (memberRows > 0) {
    var memberSheet = ss.getSheetByName(SHEET.BY_MEMBER);
    insertChart(
      sheet,
      Charts.ChartType.COLUMN,
      [memberSheet.getRange(1, 1, memberRows + 1, 1), memberSheet.getRange(1, 4, memberRows + 1, 2)],
      'メンバー別の利用時間・削減時間(h)',
      22,
      11
    );
  }

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  return sheet;
}

function countRows(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  return Math.max(sheet.getLastRow() - 1, 0);
}

function insertChart(sheet, chartType, ranges, title, anchorRow, anchorCol) {
  var builder = sheet
    .newChart()
    .setChartType(chartType)
    .setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0)
    .setOption('title', title)
    .setOption('width', 560)
    .setOption('height', 340)
    .setOption('legend', { position: 'bottom' });
  ranges.forEach(function (range) {
    builder.addRange(range);
  });
  sheet.insertChart(builder.build());
}

// ---------------------------------------------------------------- 設定シート

function readSettings(ss) {
  var sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) sheet = createSettingsSheet(ss);

  var periodDays = Math.floor(Number(sheet.getRange('B2').getValue()));
  if (!periodDays || periodDays < 1) periodDays = DEFAULT_PERIOD_DAYS;

  var lastRow = Math.max(sheet.getLastRow(), 5);
  var listValues = sheet.getRange(5, 1, lastRow - 4, 2).getValues();
  var members = columnValues(listValues, 0);
  var departments = columnValues(listValues, 1);

  if (members.length === 0) members = DEFAULT_MEMBERS.slice();
  if (departments.length === 0) departments = DEFAULT_DEPARTMENTS.slice();

  return { periodDays: periodDays, members: members, departments: departments };
}

function columnValues(values, colIndex) {
  var out = [];
  values.forEach(function (row) {
    var v = String(row[colIndex] === null || row[colIndex] === undefined ? '' : row[colIndex]).trim();
    if (v !== '' && out.indexOf(v) === -1) out.push(v);
  });
  return out;
}

function createSettingsSheet(ss) {
  var sheet = ss.insertSheet(SHEET.SETTINGS);
  sheet.getRange('A1:B1').setValues([['設定項目', '値']]).setFontWeight('bold').setBackground('#e8eaed');
  sheet.getRange('A2').setValue('集計対象期間（日）');
  sheet.getRange('B2').setValue(DEFAULT_PERIOD_DAYS);
  sheet
    .getRange('A3')
    .setValue('※ メンバー・部署を編集したら applySettingsToForm() を実行するとフォームの選択肢に反映されます')
    .setFontColor('#5f6368');
  sheet
    .getRange('A4:B4')
    .setValues([['メンバー一覧（1行に1名）', '部署一覧（1行に1つ）']])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  sheet.getRange(5, 1, DEFAULT_MEMBERS.length, 1).setValues(
    DEFAULT_MEMBERS.map(function (m) {
      return [m];
    })
  );
  sheet.getRange(5, 2, DEFAULT_DEPARTMENTS.length, 1).setValues(
    DEFAULT_DEPARTMENTS.map(function (d) {
      return [d];
    })
  );
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 260);
  return sheet;
}

// ---------------------------------------------------------------- 共通処理

function openSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('AI_USAGE_SPREADSHEET_ID');
  if (!id) throw new Error('先に setup() を実行してスプレッドシートを作成してください。');
  return SpreadsheetApp.openById(id);
}

function openForm() {
  var id = PropertiesService.getScriptProperties().getProperty('AI_USAGE_FORM_ID');
  if (!id) throw new Error('先に setup() を実行してフォームを作成してください。');
  return FormApp.openById(id);
}

/** フォームの回答シートを見つけて `回答_生データ` に名前をそろえる。 */
function ensureRawSheet(ss) {
  var named = ss.getSheetByName(SHEET.RAW);
  if (named) return named;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getFormUrl()) {
      sheets[i].setName(SHEET.RAW);
      sheets[i].setFrozenRows(1);
      return sheets[i];
    }
  }
  Logger.log('フォームの回答シートがまだ作られていません（最初の回答が入ると作成されます）。');
  return null;
}

/** スプレッドシート新規作成時にできる空の「シート1」を片付ける。 */
function removeDefaultSheet(ss) {
  if (ss.getSheets().length <= 1) return;
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if ((name === 'シート1' || name === 'Sheet1') && sheet.getLastRow() === 0 && !sheet.getFormUrl()) {
      ss.deleteSheet(sheet);
    }
  });
}

function weekStart(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var day = d.getDay(); // 0=日曜
  d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
  return d;
}

function toDate(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (value === '' || value === null || value === undefined) return null;
  var d = new Date(String(value).replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : d;
}

function formatDate(date, pattern) {
  return Utilities.formatDate(date, TIME_ZONE, pattern || 'yyyy/MM/dd');
}

function round1(value) {
  return Math.round(value * 10) / 10;
}
