/**
 * 企画進行管理シート セットアップ
 *
 * 対象: スクール×インフォマ管理シート の「シート6」
 * やること:
 *   1. シート6 を「シート6_元データ」として複製（バックアップ）
 *   2. シート6 を「企画進行管理」にリネームし、④台本進行管理 の直前に移動
 *   3. ④台本進行管理 と同じ作法（2行目＝逆算日数、3行目＝見出し、4行目以降＝ARRAYFORMULA）で作り直す
 *   4. 9/23〜10/31 の放送枠を、シート6 にあった番組・スクール・メモを引き継いで流し込む
 *
 * 使い方: スプレッドシートの「拡張機能 > Apps Script」に貼り付けて setupPlanningSheet を実行。
 *         12月末までの撮影スケジュールを ④台本進行管理 に入れるときは addShootingSchedule を実行。
 * 再実行しても元データのバックアップは上書きされない（放送枠は初期値で入れ直される）。
 */

const SRC_NAME = 'シート6';
const NEW_NAME = '企画進行管理';
const BACKUP_NAME = 'シート6_元データ';
const SCRIPT_SHEET = '④台本進行管理';
const LAST_ROW = 500;

// 放送日の何日前か（2行目に置くので、シート上で数字を変えれば全行の期限が変わる）
// 台本依頼日 = 撮影日目安(放送-10) - 21日 で ④台本進行管理 の依頼日と揃えている
const LEAD = {
  K: 45, // 企画着手日
  L: 40, // 企画案提出期限
  M: 38, // 一次FB期限【事業部・松崎】
  N: 36, // 修正案提出期限
  O: 34, // 二次FB期限【千葉】
  P: 32, // 企画FIX期限【小澤】
  T: 31, // 台本依頼日（④へ転記）
  U: 10, // 撮影日目安
};

const HEADERS = [
  '放送日', '曜日', '番組', 'スクール', '素材区分', 'ステータス', 'アラート', '担当',
  '企画テーマ・概要', '企画書URL',
  '企画着手日', '企画案提出期限', '一次FB期限\n【事業部・松崎】', '修正案提出期限',
  '二次FB期限\n【千葉】', '企画FIX期限\n【小澤】',
  '次のアクション', '次の期限', '残日数', '台本依頼日\n（④へ転記）', '撮影日目安', '撮影日\n（確定）',
  '④台本進行管理に\n転記済', 'LINEリンク発行済\n（⑧概要欄リンク）', '差し込み素材URL', '配信後リンク', 'メモ',
];

const PROGRAMS = ['LASTCALL', 'LASTCALLサブ', 'HOSTCALL', 'HOSTCALLサブ', 'REAL INFLUENCER',
  'NoBorder X File', 'NoBorder', 'NoBorder News', 'BreakingDown', 'REAL VALUE', 'REAL FOOD', '令和の龍',
  '【AI・撮影不要】汎用インフォマ', '汎用インフォマ'];
const SCHOOLS = ["HERO'ZZ", "CREATOR'ZZ", 'RVA', 'AI＋', 'Lアカデミア'];
const OWNERS = ['森本', '鈴木', '福谷', '松崎', '鳥居', '杉山', '武田'];
const MATERIALS = ['新素材', '新素材（再放送）', '撮影済素材', '既存素材', 'AI動画', '配信なし', '未定'];
const STATUSES = ['企画待ち', '企画案提出済', '一次FB済', '修正案提出済', '二次FB済',
  '企画FIX（小澤さん承認済）', '台本進行へ移行', '既存素材で対応', '配信なし'];

// シート6 の内容を引き継いだ初期データ
// [放送日, 番組, スクール, 素材区分, ステータス, 撮影日(確定), メモ]
// 「※④から推定」は ④台本進行管理 の撮影日と番組・スクールが一致する行から当てたもの
const ROWS = [
  ['2026/09/23', 'REAL VALUE', 'RVA', '既存素材', '既存素材で対応', '', '新素材無し／過去素材をつなぎ合わせてうまく1本の新素材に見せる'],
  ['2026/09/24', 'HOSTCALL', "CREATOR'ZZ", '配信なし', '配信なし', '', '新素材無し／放送回数、LINE流入数、面談予約、契約数、視聴維持率のデータ確認してまとめたうえで放送できるか、放送できないか、AIで作るかを小澤さんに共有'],
  ['2026/09/25', 'NoBorder X File', 'AI＋', '配信なし', '配信なし', '', '新素材無し／AI＋のNoBorderとXFileの実績値をヒヤリング'],
  ['2026/09/26', 'NoBorder', 'RVA', '新素材', '台本進行へ移行', '2026/09/12', '新素材 1回目（撮影日は※④から推定）'],
  ['2026/09/27', 'LASTCALL', "HERO'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/24', '24日撮影素材／編集間に合うか、間に合わない場合は別スタジオで撮影／溝口さんとクイーン、ゲストいけるか確認'],
  ['2026/09/29', '令和の龍', 'RVA', '配信なし', '配信なし', '', '編集間に合うか'],
  ['2026/09/30', 'REAL VALUE', 'RVA', '未定', '企画待ち', '', '新素材無し／27日撮影できるのかは水曜判断。それまでに承認取れる台本を用意。お知らせやタイムリーな情報だとよい。内山さんチェックも入れる'],
  ['2026/10/01', 'HOSTCALL', "HERO'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/25', '25日撮影素材／編集間に合うか'],
  ['2026/10/02', 'NoBorder X File', 'AI＋', '既存素材', '既存素材で対応', '', '新素材無し'],
  ['2026/10/03', 'NoBorder', 'AI＋', '既存素材', '既存素材で対応', '', '新素材無し'],
  ['2026/10/04', 'LASTCALL', "CREATOR'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/24', '9/24撮影素材'],
  ['2026/10/06', '令和の龍', 'RVA', '新素材', '台本進行へ移行', '2026/09/16', '新素材 1回目（撮影日は※④から推定）'],
  ['2026/10/07', 'REAL VALUE', 'RVA', '未定', '企画待ち', '', '卒業式の動画をしっかりとってインフォマにできないか、RVAメンバーでインフォマ素材を企画できる方がいないか福谷さんに確認。カメラマンでいい人いないか、はたけんさん国木田さんに確認'],
  ['2026/10/08', 'HOSTCALL', "CREATOR'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/25', '9/25撮影素材'],
  ['2026/10/09', 'NoBorder X File', 'RVA', '新素材', '台本進行へ移行', '2026/09/26', '新素材 1回目（撮影日は※④から推定）'],
  ['2026/10/10', 'NoBorder', 'RVA', '新素材（再放送）', '既存素材で対応', '2026/09/12', '新素材 2回目（9/26放送分の再放送）'],
  ['2026/10/11', 'LASTCALL', "HERO'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/24', '9/24撮影素材'],
  ['2026/10/13', '令和の龍', 'RVA', '新素材（再放送）', '既存素材で対応', '2026/09/16', '新素材 2回目（10/6放送分の再放送）'],
  ['2026/10/14', 'REAL VALUE', 'RVA', '未定', '企画待ち', '', ''],
  ['2026/10/15', 'HOSTCALL', "HERO'ZZ", '撮影済素材', '台本進行へ移行', '2026/09/25', '9/25撮影素材'],
  ['2026/10/16', 'NoBorder X File', 'RVA', '新素材（再放送）', '既存素材で対応', '2026/09/26', '新素材 2回目（10/9放送分の再放送）'],
  ['2026/10/17', 'NoBorder', 'AI＋', '未定', '企画待ち', '', ''],
  ['2026/10/18', 'LASTCALL', "CREATOR'ZZ", '未定', '企画待ち', '', ''],
  ['2026/10/20', '令和の龍', 'RVA', '未定', '企画待ち', '', ''],
  ['2026/10/21', 'REAL VALUE', 'RVA', '未定', '企画待ち', '', ''],
  ['2026/10/22', 'HOSTCALL', "CREATOR'ZZ", '未定', '企画待ち', '', ''],
  ['2026/10/23', 'NoBorder X File', 'AI＋', '未定', '企画待ち', '', ''],
  ['2026/10/24', 'NoBorder', 'RVA', '未定', '企画待ち', '', ''],
  ['2026/10/25', 'LASTCALL', "HERO'ZZ", '未定', '企画待ち', '', ''],
  ['2026/10/27', '令和の龍', 'RVA', '新素材', '企画待ち', '2026/10/14', '10/14収録で差し込み撮影（③収録スケジュール）'],
  ['2026/10/28', 'REAL VALUE', 'RVA', '未定', '企画待ち', '', ''],
  ['2026/10/29', 'HOSTCALL', "HERO'ZZ", '未定', '企画待ち', '', ''],
  ['2026/10/30', 'NoBorder X File', 'AI＋', '新素材', '企画待ち', '2026/10/17', '10/17収録で差し込み撮影（③収録スケジュール）。撮影するスクールは要決定'],
  ['2026/10/31', 'NoBorder', 'AI＋', '未定', '企画待ち', '', ''],
];

function setupPlanningSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC_NAME) || ss.getSheetByName(NEW_NAME);
  if (!sh) throw new Error(`「${SRC_NAME}」も「${NEW_NAME}」も見つかりません`);

  if (!ss.getSheetByName(BACKUP_NAME)) {
    sh.copyTo(ss).setName(BACKUP_NAME);
  }

  sh.setName(NEW_NAME);
  const scriptSheet = ss.getSheetByName(SCRIPT_SHEET);
  if (scriptSheet) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(scriptSheet.getIndex());
  }

  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  if (sh.getMaxColumns() < HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < LAST_ROW) {
    sh.insertRowsAfter(sh.getMaxRows(), LAST_ROW - sh.getMaxRows());
  }

  // 1行目: 説明
  sh.getRange('A1').setValue('企画進行管理（放送日から逆算）｜白セル＝入力／グレー列＝自動計算（触らない）／2行目の日数を変えると全行の期限が変わる')
    .setFontWeight('bold');

  // 2行目: 逆算日数
  sh.getRange('A2').setValue('▼ 逆算日数（放送日の何日前）').setFontColor('#666666');
  Object.keys(LEAD).forEach(col => {
    sh.getRange(`${col}2`).setValue(LEAD[col]).setHorizontalAlignment('center').setBackground('#FFF2CC');
  });

  // 3行目: 見出し
  sh.getRange(3, 1, 1, HEADERS.length).setValues([HEADERS])
    .setBackground('#D6DCE4').setFontWeight('bold').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 4行目以降: 入力データ
  const n = ROWS.length;
  sh.getRange(4, 1, n, 1).setValues(ROWS.map(r => [r[0]]));
  sh.getRange(4, 3, n, 4).setValues(ROWS.map(r => [r[1], r[2], r[3], r[4]]));
  sh.getRange(4, 22, n, 1).setValues(ROWS.map(r => [r[5]]));
  sh.getRange(4, 27, n, 1).setValues(ROWS.map(r => [r[6]]));

  // 自動計算列
  const col = c => `$${c}$4:$${c}$${LAST_ROW}`;
  const A = col('A');
  const F = col('F');
  const DONE = '台本進行へ移行|既存素材で対応|配信なし';

  sh.getRange('B4').setFormula(`=ARRAYFORMULA(IF(${A}="","",MID("日月火水木金土",WEEKDAY(${A}),1)))`);
  Object.keys(LEAD).forEach(c => {
    sh.getRange(`${c}4`).setFormula(`=ARRAYFORMULA(IF(${A}="","",${A}-${c}$2))`);
  });
  sh.getRange('Q4').setFormula(
    `=ARRAYFORMULA(IF(${A}="","",IF(REGEXMATCH(${F},"${DONE}"),"—",` +
    `IF(${F}="企画案提出済","一次FB（松崎）",IF(${F}="一次FB済","修正案提出",IF(${F}="修正案提出済","二次FB（千葉）",` +
    `IF(${F}="二次FB済","企画FIX（小澤さん承認）",IF(REGEXMATCH(${F},"企画FIX"),"台本依頼・④へ転記","企画案提出"))))))))`);
  sh.getRange('R4').setFormula(
    `=ARRAYFORMULA(IF(${A}="","",IF(REGEXMATCH(${F},"${DONE}"),"",` +
    `IF(${F}="企画案提出済",${col('M')},IF(${F}="一次FB済",${col('N')},IF(${F}="修正案提出済",${col('O')},` +
    `IF(${F}="二次FB済",${col('P')},IF(REGEXMATCH(${F},"企画FIX"),${col('T')},${col('L')}))))))))`);
  sh.getRange('S4').setFormula(`=ARRAYFORMULA(IF(${col('R')}="","",${col('R')}-TODAY()))`);
  sh.getRange('G4').setFormula(
    `=ARRAYFORMULA(IF(${A}="","",IF(REGEXMATCH(${F},"${DONE}"),"完了",IF(${col('R')}="","",` +
    `IF(${col('R')}<TODAY(),"超過",IF(${col('R')}<=TODAY()+1,"直前","OK"))))))`);

  // 書式
  const rows = LAST_ROW - 3;
  ['A', 'K', 'L', 'M', 'N', 'O', 'P', 'R', 'T', 'U', 'V'].forEach(c => {
    sh.getRange(`${c}4:${c}${LAST_ROW}`).setNumberFormat('yyyy/mm/dd');
  });
  sh.getRange(`S4:S${LAST_ROW}`).setNumberFormat('0');
  ['B', 'G', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U'].forEach(c => {
    sh.getRange(`${c}4:${c}${LAST_ROW}`).setBackground('#F3F3F3');
  });
  sh.getRange(4, 1, rows, HEADERS.length).setVerticalAlignment('middle');
  sh.getRange(`B4:B${LAST_ROW}`).setHorizontalAlignment('center');
  sh.getRange(`G4:G${LAST_ROW}`).setHorizontalAlignment('center');
  sh.getRange(`AA4:AA${LAST_ROW}`).setWrap(true);

  // 入力規則
  const list = (values) => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
  sh.getRange(`C4:C${LAST_ROW}`).setDataValidation(list(PROGRAMS));
  sh.getRange(`D4:D${LAST_ROW}`).setDataValidation(list(SCHOOLS));
  sh.getRange(`E4:E${LAST_ROW}`).setDataValidation(list(MATERIALS));
  sh.getRange(`F4:F${LAST_ROW}`).setDataValidation(list(STATUSES));
  sh.getRange(`H4:H${LAST_ROW}`).setDataValidation(list(OWNERS));
  sh.getRange(`W4:X${LAST_ROW}`).insertCheckboxes();

  // 条件付き書式（④台本進行管理 と同じ色）
  const alert = sh.getRange(`G4:G${LAST_ROW}`);
  const cf = (text, color) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(color).setRanges([alert]).build();
  const muted = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($F4,"配信なし")')
    .setFontColor('#999999').setRanges([sh.getRange(`A4:F${LAST_ROW}`)]).build();
  sh.setConditionalFormatRules([
    cf('超過', '#F8696B'), cf('直前', '#FFEB84'), cf('OK', '#63BE7B'), cf('完了', '#A6A6A6'), muted,
  ]);

  // 列幅・固定
  const widths = {
    A: 95, B: 40, C: 180, D: 100, E: 110, F: 170, G: 60, H: 70, I: 240, J: 110,
    K: 95, L: 100, M: 110, N: 100, O: 100, P: 100, Q: 160, R: 95, S: 55, T: 100, U: 95, V: 95,
    W: 110, X: 120, Y: 140, Z: 140, AA: 360,
  };
  Object.keys(widths).forEach(c => sh.setColumnWidth(sh.getRange(`${c}1`).getColumn(), widths[c]));
  sh.setRowHeight(3, 48);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(4);

  SpreadsheetApp.flush();
}

// ③放送カレンダー【収録スケジュール】の 10/14〜12月末の収録日（+ 番組側から聞いた令和の龍 1/13）
// [番組, スクール, 撮影/収録日, メモ]
// 初回放送目安 = 収録日+10日（編集期間）以降で最初の同番組の放送枠
// ・LASTCALL/HOSTCALL は 9/24・9/25 と同じく1回の収録で CREATOR'ZZ と HERO'ZZ の2本を撮る
// ・NoBorder / NoBorder X File は AI＋ と RVA を交互に入れているため、スクールは空欄（要決定）
// ・BreakingDown と、11月以降の LASTCALL / HOSTCALL は収録日未定
const SHOOTS = [
  ['令和の龍', 'RVA', '2026/10/14', '初回放送目安 10/27(火)'],
  ['NoBorder X File', '', '2026/10/17', '初回放送目安 10/30(金)／スクール要決定（AI＋ or RVA）'],
  ['LASTCALL', "CREATOR'ZZ", '2026/10/22', '初回放送目安 11/1(日)'],
  ['LASTCALL', "HERO'ZZ", '2026/10/22', '初回放送目安 11/1(日)'],
  ['HOSTCALL', "CREATOR'ZZ", '2026/10/23', '初回放送目安 11/5(木)'],
  ['HOSTCALL', "HERO'ZZ", '2026/10/23', '初回放送目安 11/5(木)'],
  ['NoBorder', '', '2026/10/24', '初回放送目安 11/7(土)／スクール要決定（AI＋ or RVA）'],
  ['REAL VALUE', 'RVA', '2026/10/25', '初回放送目安 11/4(水)'],
  ['令和の龍', 'RVA', '2026/11/11', '初回放送目安 11/24(火)'],
  ['NoBorder X File', '', '2026/11/21', '初回放送目安 12/4(金)／スクール要決定（AI＋ or RVA）'],
  ['REAL VALUE', 'RVA', '2026/11/22', '初回放送目安 12/2(水)'],
  ['NoBorder', '', '2026/11/28', '初回放送目安 12/12(土)／スクール要決定（AI＋ or RVA）'],
  ['REAL VALUE', 'RVA', '2026/12/06', '初回放送目安 12/16(水)／12/6〜10は連日収録：差し込み撮影する日を決めて不要な行は削除'],
  ['REAL FOOD', 'RVA', '2026/12/06', '定期放送枠なし：使う枠を要決定'],
  ['REAL VALUE', 'RVA', '2026/12/07', '初回放送目安 12/23(水)／12/6〜10連日収録'],
  ['REAL VALUE', 'RVA', '2026/12/08', '初回放送目安 12/23(水)／12/6〜10連日収録'],
  ['REAL VALUE', 'RVA', '2026/12/09', '初回放送目安 12/23(水)／12/6〜10連日収録'],
  ['REAL VALUE', 'RVA', '2026/12/10', '初回放送目安 12/23(水)／12/6〜10連日収録'],
  ['令和の龍', 'RVA', '2026/12/16', '初回放送目安 12/29(火)'],
  ['NoBorder', '', '2026/12/19', '初回放送目安 1/2(土)／スクール要決定（AI＋ or RVA）'],
  ['令和の龍', 'RVA', '2027/01/13', '初回放送目安 1/26(火)／③の一覧には未記載（番組側から確認済）'],
];

function addShootingSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SCRIPT_SHEET);
  if (!sh) throw new Error(`「${SCRIPT_SHEET}」が見つかりません`);
  const tz = ss.getSpreadsheetTimeZone();
  const fmt = v => (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy/MM/dd') : String(v));

  // 既存行（番組・スクール・撮影日が同じもの）は追加しない
  const data = sh.getRange(4, 1, LAST_ROW - 3, 3).getValues();
  const existing = new Set(data.filter(r => r[0] !== '').map(r => `${r[0]}|${r[1]}|${fmt(r[2])}`));
  let lastFilled = 3;
  data.forEach((r, i) => { if (r[0] !== '') lastFilled = i + 4; });

  const add = SHOOTS.filter(r => !existing.has(`${r[0]}|${r[1]}|${r[2]}`));
  if (add.length === 0) return;
  const start = lastFilled + 1;
  const n = add.length;

  // I〜Q列は ④ の ARRAYFORMULA が撮影日から自動で埋める
  sh.getRange(start, 1, n, 5).setValues(add.map(r => [r[0], r[1], r[2], '公式撮影', '一次提出待ち']));
  sh.getRange(start, 3, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(start, 21, n, 1).setValues(add.map(r => [r[3]]));
  sh.getRange(start, 18, n, 3).insertCheckboxes();

  // ④ の入力規則は40行目前後までしか付いていないので、追加行にも同じプルダウンを付ける
  const list = (values) => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
  sh.getRange(start, 1, n, 1).setDataValidation(list(PROGRAMS));
  sh.getRange(start, 2, n, 1).setDataValidation(list(SCHOOLS));
  sh.getRange(start, 4, n, 1).setDataValidation(list(['公式撮影', '非公式撮影']));
  sh.getRange(start, 5, n, 1).setDataValidation(list(['一次提出待ち', '二次提出待ち', '三次提出待ち', '小澤さん承認済', '撮影済']));
  sh.getRange(start, 8, n, 1).setDataValidation(list(OWNERS));

  SpreadsheetApp.flush();
}
