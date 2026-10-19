/**
 * スクール×インフォマ管理シート｜スクール別タブ自動生成
 *
 * 使い方:
 *   1. 管理シートの 拡張機能 > Apps Script に全文を貼る
 *   2. 保存して再読み込み → メニュー「インフォマ」が出る
 *   3. 「スクール別タブを作成・更新」を実行（初回は権限の承認が出る）
 *
 * 前提（ここを崩すと集計がずれる）:
 *   - ⑥素材マスタ : 4行目から1素材＝1行。A=スクール, B=素材タイトル, C=種類, D=尺, E=ステータス
 *   - ⑤インフォマ効果分析 : 3行目から1配信＝1行。A=配信日, B=スクール, C=番組, D=動画タイトル,
 *                           E=動画URL, F=使用素材(⑥の素材タイトル), K=LINE登録数
 *   - 1配信につき使う素材は1つ（使用回数＝⑤でその素材が入っている行数）
 */

const SRC_MASTER = '⑥素材マスタ';
const SRC_EFFECT = '⑤インフォマ効果分析';

// tab  : 作るタブ名
// key  : ⑤⑥の「スクール」列に入っている値（完全一致。全角/半角の「＋」に注意）
const SCHOOLS = [
  { tab: "HERO'ZZ", key: "HERO'ZZ" },
  { tab: "CREATOR'ZZ", key: "CREATOR'ZZ" },
  { tab: 'RVA', key: 'RVA' },
  { tab: 'AI+', key: 'AI+' },
];

const LIST_ROW = 9;     // 素材一覧の1行目
const LIST_ROWS = 30;   // 素材一覧の最大行数（1スクール30素材まで）
const HIST_ROWS = 500;  // 使用履歴の最大行数（数式の範囲）
const HIST_FRAME_ROWS = 100; // 使用履歴で罫線を引く行数

const COLOR_HEADER = '#1f3a5f';
const COLOR_UNUSED = '#fde7e7';
const COLOR_STALE = '#fff0d6';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('インフォマ')
    .addItem('スクール別タブを作成・更新', 'buildSchoolTabs')
    .addItem('⑤「使用素材」にドロップダウンを設定', 'setupMaterialDropdown')
    .addItem('⑥の使用回数・最終使用日・累計LINEを自動化', 'fillMasterUsage')
    .addToUi();
}

function buildSchoolTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SRC_MASTER, SRC_EFFECT].forEach(function (n) {
    if (!ss.getSheetByName(n)) throw new Error('シートが見つかりません: ' + n);
  });
  SCHOOLS.forEach(function (s) { buildTab_(ss, s); });
  notify_('スクール別タブを作成・更新しました: ' + SCHOOLS.map(function (s) { return s.tab; }).join(' / '));
}

function buildTab_(ss, school) {
  let sh = ss.getSheetByName(school.tab);
  if (!sh) {
    sh = ss.insertSheet(school.tab);
  } else {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clear();
    sh.clearConditionalFormatRules();
  }

  const K = JSON.stringify(school.key);          // 数式に埋め込むスクール名
  const M = "'" + SRC_MASTER + "'!";
  const E = "'" + SRC_EFFECT + "'!";
  const lastRow = LIST_ROW + LIST_ROWS - 1;

  // ---- タイトル ----
  sh.getRange('A1').setValue(school.tab + '｜インフォマ素材 使用状況')
    .setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('※ ⑥素材マスタ・⑤インフォマ効果分析から自動集計。数式が入っているので、この中は触らない')
    .setFontColor('#777777').setFontSize(9);

  // ---- サマリー（KPI 4つ）----
  const kpis = [
    { label: '素材数', range: 'A4:A5', formula: '=COUNTIFS(' + M + '$A$4:$A,' + K + ',' + M + '$B$4:$B,"<>")' },
    { label: '延べ使用回数', range: 'B4:C5', formula: '=COUNTIFS(' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,"<>")' },
    {
      label: '最終使用日', range: 'D4:E5',
      formula: '=IFERROR(IF(MAXIFS(' + E + '$A$3:$A,' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,"<>")=0,"-",MAXIFS(' +
        E + '$A$3:$A,' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,"<>")),"-")',
    },
    { label: '未使用の素材数', range: 'F4:G5', formula: '=COUNTIFS($A$' + LIST_ROW + ':$A$' + lastRow + ',"<>",$E$' + LIST_ROW + ':$E$' + lastRow + ',0)' },
  ];
  kpis.forEach(function (k) {
    const r = sh.getRange(k.range);
    const top = r.offset(0, 0, 1, r.getNumColumns());
    const bottom = r.offset(1, 0, 1, r.getNumColumns());
    if (r.getNumColumns() > 1) { top.merge(); bottom.merge(); }
    top.getCell(1, 1).setValue(k.label);
    bottom.getCell(1, 1).setFormula(k.formula);
    top.setFontSize(9).setFontColor('#666666').setHorizontalAlignment('center').setBackground('#f3f6fa');
    bottom.setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f3f6fa');
    r.setBorder(true, true, true, true, false, false, '#c9d3e0', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange('D5').setNumberFormat('yyyy/mm/dd');
  sh.setRowHeight(5, 40);

  // ---- 見出し ----
  sh.getRange('A7').setValue('■ 素材別の使用状況').setFontWeight('bold');
  sh.getRange('K7').setValue('■ 使用履歴（新しい順）').setFontWeight('bold');

  const listHeaders = ['素材タイトル', '種類', '尺(秒)', 'ステータス', '使用回数', '最終使用日', '最終使用からの日数', '累計LINE登録数'];
  const histHeaders = ['配信日', '番組', '動画タイトル', '使用素材', '動画URL'];
  styleHeader_(sh.getRange(8, 1, 1, listHeaders.length).setValues([listHeaders]));
  styleHeader_(sh.getRange(8, 11, 1, histHeaders.length).setValues([histHeaders]));

  // ---- 素材一覧（A:D は⑥から取得、E:H は行ごとの数式）----
  sh.getRange(LIST_ROW, 1).setFormula(
    '=IFERROR(FILTER(' + M + '$B$4:$E,' + M + '$A$4:$A=' + K + ',' + M + '$B$4:$B<>""),"")');

  const rows = [];
  for (let i = 0; i < LIST_ROWS; i++) {
    const r = LIST_ROW + i;
    rows.push([
      '=IF($A' + r + '="","",COUNTIFS(' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,$A' + r + '))',
      '=IF($A' + r + '="","",IF($E' + r + '=0,"未使用",MAXIFS(' + E + '$A$3:$A,' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,$A' + r + ')))',
      '=IF(ISNUMBER($F' + r + '),TODAY()-$F' + r + ',"")',
      '=IF($A' + r + '="","",SUMIFS(' + E + '$K$3:$K,' + E + '$B$3:$B,' + K + ',' + E + '$F$3:$F,$A' + r + '))',
    ]);
  }
  sh.getRange(LIST_ROW, 5, LIST_ROWS, 4).setFormulas(rows);

  sh.getRange(LIST_ROW, 3, LIST_ROWS, 1).setHorizontalAlignment('right');
  sh.getRange(LIST_ROW, 5, LIST_ROWS, 1).setNumberFormat('0').setHorizontalAlignment('right');
  sh.getRange(LIST_ROW, 6, LIST_ROWS, 1).setNumberFormat('yyyy/mm/dd').setHorizontalAlignment('right');
  sh.getRange(LIST_ROW, 7, LIST_ROWS, 1).setNumberFormat('0"日"').setHorizontalAlignment('right');
  sh.getRange(LIST_ROW, 8, LIST_ROWS, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');

  // ---- 使用履歴（⑤から、使用素材が入っている行だけ、新しい順）----
  sh.getRange(LIST_ROW, 11).setFormula(
    '=IFERROR(SORT(FILTER({' + E + '$A$3:$A,' + E + '$C$3:$C,' + E + '$D$3:$D,' + E + '$F$3:$F,' + E + '$E$3:$E},' +
    E + '$B$3:$B=' + K + ',' + E + '$F$3:$F<>""),1,FALSE),"")');
  sh.getRange(LIST_ROW, 11, HIST_ROWS, 1).setNumberFormat('yyyy/mm/dd').setHorizontalAlignment('left');

  // ---- 罫線（入力・表示する範囲を区切る）----
  const BORDER = '#b7c0cc';
  [sh.getRange(8, 1, LIST_ROWS + 1, 8), sh.getRange(8, 11, HIST_FRAME_ROWS + 1, 5)].forEach(function (r) {
    r.setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  // ---- 条件付き書式 ----
  const listRange = sh.getRange(LIST_ROW, 1, LIST_ROWS, 8);
  const statusRange = sh.getRange(LIST_ROW, 4, LIST_ROWS, 1);
  const staleRange = sh.getRange(LIST_ROW, 7, LIST_ROWS, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A' + LIST_ROW + '<>"",$E' + LIST_ROW + '=0)')
      .setBackground(COLOR_UNUSED).setRanges([listRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G' + LIST_ROW + '),$G' + LIST_ROW + '>=30)')
      .setBackground(COLOR_STALE).setRanges([staleRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('完成').setBackground('#d9f0de').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('制作中').setBackground('#fff4c2').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('未着手').setBackground('#e6e6e6').setRanges([statusRange]).build(),
  ];
  sh.setConditionalFormatRules(rules);

  // ---- 体裁 ----
  const widths = { 1: 240, 2: 70, 3: 60, 4: 90, 5: 80, 6: 100, 7: 120, 8: 120, 9: 20, 10: 20, 11: 100, 12: 120, 13: 320, 14: 200, 15: 260 };
  Object.keys(widths).forEach(function (c) { sh.setColumnWidth(Number(c), widths[c]); });
  sh.setFrozenRows(8);
  sh.setHiddenGridlines(true);
  sh.getRange(8, 1, 1, 8).setWrap(true);
}

/** 完了メッセージ。エディタ実行など alert が使えない場面ではシート右下のトーストで出す。 */
function notify_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'インフォマ', 8);
  }
}

function styleHeader_(range) {
  range.setBackground(COLOR_HEADER).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/** ⑤の「使用素材」列(F)に、⑥の素材タイトルのドロップダウンを付ける（表記ゆれ防止）。 */
function setupMaterialDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const effect = ss.getSheetByName(SRC_EFFECT);
  const master = ss.getSheetByName(SRC_MASTER);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(master.getRange('B4:B300'), true)
    .setAllowInvalid(true) // 入力は止めず、警告表示のみ
    .build();
  effect.getRange('F3:F' + effect.getMaxRows()).setDataValidation(rule);
  notify_('⑤の「使用素材」列にドロップダウンを設定しました。');
}

/** ⑥の 使用回数(I)・最終使用日(J)・累計LINE登録数(K) を数式で埋める。 */
function fillMasterUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(SRC_MASTER);
  const E = "'" + SRC_EFFECT + "'!";
  const first = 4, last = 300;

  const existing = master.getRange(first, 9, last - first + 1, 3).getValues()
    .some(function (row) { return row.some(function (v) { return v !== ''; }); });
  if (existing) {
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('⑥のI〜K列に既に値があります。数式で上書きしますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }

  const rows = [];
  for (let r = first; r <= last; r++) {
    rows.push([
      '=IF($B' + r + '="","",COUNTIFS(' + E + '$B$3:$B,$A' + r + ',' + E + '$F$3:$F,$B' + r + '))',
      '=IF($B' + r + '="","",IF($I' + r + '=0,"",MAXIFS(' + E + '$A$3:$A,' + E + '$B$3:$B,$A' + r + ',' + E + '$F$3:$F,$B' + r + ')))',
      '=IF($B' + r + '="","",SUMIFS(' + E + '$K$3:$K,' + E + '$B$3:$B,$A' + r + ',' + E + '$F$3:$F,$B' + r + '))',
    ]);
  }
  master.getRange(first, 9, rows.length, 3).setFormulas(rows);
  master.getRange(first, 10, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  notify_('⑥の使用回数・最終使用日・累計LINE登録数を数式化しました。');
}
