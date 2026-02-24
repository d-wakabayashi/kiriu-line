/**
 * KIRIU ライン負荷最適化システム - Google Apps Script
 */

// ========================================
// 設定
// ========================================
const API_URL = 'https://kiriu-line-optimizer-395896719333.asia-northeast1.run.app';

const DISC_LINES = ['4915', '4919', '4927', '4928', '4934', '4935', '4945', '4G01', '4J01'];
const MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月'];

// ========================================
// ログ出力ユーティリティ
// ========================================
function log(message, data = null) {
  const timestamp = new Date().toISOString();
  if (data) {
    console.log(`[${timestamp}] ${message}`, JSON.stringify(data).substring(0, 500));
  } else {
    console.log(`[${timestamp}] ${message}`);
  }
}

// ========================================
// メニュー追加
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ライン最適化')
    .addItem('最適化を実行（100%）', 'runOptimization')
    .addItem('パターン比較（100/90/80%）', 'runComparisonOptimization')
    .addItem('勤務体制パターン比較', 'runWorkPatternComparison')
    .addSeparator()
    .addItem('テンプレートを作成', 'createTemplateSheets')
    .addItem('サンプルデータを入力', 'insertSampleData')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

// ========================================
// テンプレートシート作成
// ========================================
function createTemplateSheets() {
  log('テンプレート作成開始');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 入力シート
  let inputSheet = ss.getSheetByName('入力');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('入力');
  }
  inputSheet.clear();

  // ヘッダー
  const inputHeaders = ['部品番号', 'メインライン', 'サブ1', 'サブ2'].concat(MONTHS);
  inputSheet.getRange(1, 1, 1, inputHeaders.length).setValues([inputHeaders]);
  inputSheet.getRange(1, 1, 1, inputHeaders.length)
    .setBackground('#4472C4')
    .setFontColor('white')
    .setFontWeight('bold');

  // ライン選択用のドロップダウン
  const lineRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DISC_LINES, true)
    .build();
  inputSheet.getRange('B2:D100').setDataValidation(lineRule);

  // ライン能力シート（月別対応）
  let capSheet = ss.getSheetByName('ライン能力');
  if (!capSheet) {
    capSheet = ss.insertSheet('ライン能力');
  }
  capSheet.clear();

  // 月別ヘッダー
  const capHeaders = ['ライン'].concat(MONTHS);
  capSheet.getRange(1, 1, 1, capHeaders.length).setValues([capHeaders]);
  capSheet.getRange(1, 1, 1, capHeaders.length)
    .setBackground('#4472C4')
    .setFontColor('white')
    .setFontWeight('bold');

  // デフォルト能力値（月別）
  const defaultCaps = {
    '4915': 70000, '4919': 80000, '4927': 40000, '4928': 40000,
    '4934': 50000, '4935': 85000, '4945': 50000, '4G01': 50000, '4J01': 10000
  };
  // 各ラインに対して12ヶ月分のデフォルト値を設定
  const capData = DISC_LINES.map(line => {
    const cap = defaultCaps[line] || 50000;
    return [line].concat(Array(12).fill(cap));
  });
  capSheet.getRange(2, 1, capData.length, 13).setValues(capData);
  capSheet.getRange(2, 2, capData.length, 12).setNumberFormat('#,##0');

  // 入力セルを黄色でハイライト
  capSheet.getRange(2, 2, capData.length, 12).setBackground('#FFFFCC');

  // 結果シート（ライン負荷）
  let resultLineSheet = ss.getSheetByName('結果_ライン負荷');
  if (!resultLineSheet) {
    resultLineSheet = ss.insertSheet('結果_ライン負荷');
  }
  resultLineSheet.clear();

  // 結果シート（部品割当）
  let resultAllocSheet = ss.getSheetByName('結果_部品割当');
  if (!resultAllocSheet) {
    resultAllocSheet = ss.insertSheet('結果_部品割当');
  }
  resultAllocSheet.clear();

  // 結果シート（未割当）
  let resultUnmetSheet = ss.getSheetByName('結果_未割当');
  if (!resultUnmetSheet) {
    resultUnmetSheet = ss.insertSheet('結果_未割当');
  }
  resultUnmetSheet.clear();

  // 結果シート（月別能力）
  let resultCapSheet = ss.getSheetByName('結果_月別能力');
  if (!resultCapSheet) {
    resultCapSheet = ss.insertSheet('結果_月別能力');
  }
  resultCapSheet.clear();

  // --- 勤務体制パターン関連シート ---

  // 負荷率計算シート
  let wpSheet = ss.getSheetByName('負荷率計算');
  if (!wpSheet) {
    wpSheet = ss.insertSheet('負荷率計算');
  }
  wpSheet.clear();

  const wpHeaders = ['勤務体制', '月稼働時間計算式', '月除外時間'];
  wpSheet.getRange(1, 1, 1, wpHeaders.length).setValues([wpHeaders]);
  wpSheet.getRange(1, 1, 1, wpHeaders.length)
    .setBackground('#4472C4')
    .setFontColor('white')
    .setFontWeight('bold');

  const wpData = [
    ['2直2交替', '{月間稼働日数} * 7.5 * 2 - {月除外時間}', 5],
    ['3直3交替', '{月間稼働日数} * 7.5 * 3 - {月除外時間}', 8],
  ];
  wpSheet.getRange(2, 1, wpData.length, wpData[0].length).setValues(wpData);
  wpSheet.getRange(2, 1, wpData.length, wpData[0].length).setBackground('#FFFFCC');

  // ライン製造能力シート（JPH）
  let jphSheet = ss.getSheetByName('ライン製造能力');
  if (!jphSheet) {
    jphSheet = ss.insertSheet('ライン製造能力');
  }
  jphSheet.clear();

  const jphHeaders = ['ライン', 'JPH'];
  jphSheet.getRange(1, 1, 1, jphHeaders.length).setValues([jphHeaders]);
  jphSheet.getRange(1, 1, 1, jphHeaders.length)
    .setBackground('#4472C4')
    .setFontColor('white')
    .setFontWeight('bold');

  const defaultJph = {
    '4915': 350, '4919': 400, '4927': 200, '4928': 200,
    '4934': 250, '4935': 425, '4945': 250, '4G01': 250, '4J01': 50
  };
  const jphData = DISC_LINES.map(line => [line, defaultJph[line] || 0]);
  jphSheet.getRange(2, 1, jphData.length, 2).setValues(jphData);
  jphSheet.getRange(2, 2, jphData.length, 1).setNumberFormat('#,##0');
  jphSheet.getRange(2, 2, jphData.length, 1).setBackground('#FFFFCC');

  // 月間稼働日数シート
  let daysSheet = ss.getSheetByName('月間稼働日数');
  if (!daysSheet) {
    daysSheet = ss.insertSheet('月間稼働日数');
  }
  daysSheet.clear();

  daysSheet.getRange(1, 1, 1, MONTHS.length).setValues([MONTHS]);
  daysSheet.getRange(1, 1, 1, MONTHS.length)
    .setBackground('#4472C4')
    .setFontColor('white')
    .setFontWeight('bold');

  const defaultDays = [[20, 19, 21, 22, 21, 20, 22, 19, 21, 20, 18, 21]];
  daysSheet.getRange(2, 1, 1, 12).setValues(defaultDays);
  daysSheet.getRange(2, 1, 1, 12).setBackground('#FFFFCC');

  log('テンプレート作成完了');
  SpreadsheetApp.getUi().alert('テンプレートシートを作成しました');
}

// ========================================
// サンプルデータ入力
// ========================================
function insertSampleData() {
  log('サンプルデータ入力開始');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('入力');

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('先にテンプレートを作成してください');
    return;
  }

  // サンプルデータ
  const sampleData = [
    ['PART001', '4915', '4919', '', 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000],
    ['PART002', '4919', '', '', 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000],
    ['PART003', '4927', '4928', '', 3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000, 3000],
    ['PART004', '4935', '', '', 10000, 10000, 10000, 10000, 10000, 10000, 10000, 10000, 10000, 10000, 10000, 10000],
    ['PART005', '4G01', '4945', '', 7000, 7000, 7000, 7000, 7000, 7000, 7000, 7000, 7000, 7000, 7000, 7000],
  ];

  inputSheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  inputSheet.getRange(2, 5, sampleData.length, 12).setNumberFormat('#,##0');

  log('サンプルデータ入力完了', { rows: sampleData.length });
  SpreadsheetApp.getUi().alert('サンプルデータを入力しました');
}

// ========================================
// 最適化実行
// ========================================
function runOptimization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  log('最適化実行開始');

  // 入力データ取得
  const inputSheet = ss.getSheetByName('入力');
  if (!inputSheet) {
    log('エラー: 入力シートが見つかりません');
    ui.alert('エラー', '「入力」シートが見つかりません', ui.ButtonSet.OK);
    return;
  }

  const inputData = inputSheet.getDataRange().getValues();
  log('入力データ取得', { rows: inputData.length, cols: inputData[0]?.length });

  if (inputData.length < 2) {
    log('エラー: 入力データがありません');
    ui.alert('エラー', '入力データがありません', ui.ButtonSet.OK);
    return;
  }

  // 能力データ取得
  const capSheet = ss.getSheetByName('ライン能力');
  let capacitiesData = null;
  if (capSheet) {
    capacitiesData = capSheet.getDataRange().getValues().slice(1);  // ヘッダー除外
    log('能力データ取得', { rows: capacitiesData.length });
  }

  try {
    log('API呼び出し開始', { url: API_URL });

    // API呼び出し
    const response = callOptimizeApi(inputData, capacitiesData);

    log('API呼び出し完了', {
      success: response.success,
      status: response.status,
      parts_count: response.parts_count
    });

    if (!response.success) {
      ui.alert('最適化失敗', `ステータス: ${response.status}`, ui.ButtonSet.OK);
      return;
    }

    // 結果を書き込み
    log('結果書き込み開始');
    writeResults(ss, response);
    log('結果書き込み完了');

    const unmetMsg = response.total_unmet > 0
      ? `\n未割当合計: ${response.total_unmet.toLocaleString()}（※能力超過分）`
      : '';

    ui.alert('完了',
      `最適化が完了しました\n\n` +
      `ステータス: ${response.status}\n` +
      `部品数: ${response.parts_count}\n` +
      `年間総需要: ${response.total_demand.toLocaleString()}${unmetMsg}\n` +
      `実行時間: ${response.solve_time.toFixed(2)}秒`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    log('エラー発生', { message: error.message, stack: error.stack });
    ui.alert('エラー', `API呼び出しに失敗しました:\n${error.message}`, ui.ButtonSet.OK);
  }
}

// ========================================
// API呼び出し
// ========================================
function callOptimizeApi(partsData, capacitiesData) {
  const payload = {
    parts_data: partsData,
    capacities_data: capacitiesData,
    time_limit: 60
  };

  log('APIリクエスト送信', { endpoint: '/optimize/simple' });

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(`${API_URL}/optimize/simple`, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  log('APIレスポンス受信', { code: responseCode, length: responseText.length });

  if (responseCode !== 200) {
    log('APIエラー', { code: responseCode, body: responseText.substring(0, 500) });
    throw new Error(`HTTPエラー: ${responseCode}\n${responseText.substring(0, 200)}`);
  }

  const result = JSON.parse(responseText);
  log('APIレスポンスパース完了', { success: result.success });

  return result;
}

// ========================================
// 結果書き込み
// ========================================
function writeResults(ss, response) {
  // ライン負荷結果
  const lineSheet = ss.getSheetByName('結果_ライン負荷');
  if (lineSheet && response.line_loads) {
    log('ライン負荷シート書き込み開始', { rows: response.line_loads.length });

    lineSheet.clear();
    const lineData = response.line_loads;
    lineSheet.getRange(1, 1, lineData.length, lineData[0].length).setValues(lineData);

    // ヘッダースタイル
    lineSheet.getRange(1, 1, 1, lineData[0].length)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold');

    // 数値フォーマット（3列目から最後から2列目まで）
    if (lineData.length > 1) {
      lineSheet.getRange(2, 3, lineData.length - 1, 12)
        .setNumberFormat('#,##0');
    }

    // 100%超えを赤色表示
    const lastCol = lineData[0].length;
    for (let i = 2; i <= lineData.length; i++) {
      try {
        const rateValue = lineSheet.getRange(i, lastCol).getValue();
        // 文字列か数値かを判定して処理
        let rate = 0;
        if (typeof rateValue === 'string') {
          rate = parseFloat(rateValue.replace('%', ''));
        } else if (typeof rateValue === 'number') {
          rate = rateValue * 100;  // 小数の場合
        }

        log(`行${i}の負荷率`, { value: rateValue, rate: rate });

        if (rate > 100) {
          lineSheet.getRange(i, 1, 1, lastCol).setBackground('#FFC7CE');
        }
      } catch (e) {
        log(`行${i}の負荷率処理でエラー`, { error: e.message });
      }
    }

    log('ライン負荷シート書き込み完了');
  }

  // 部品割当結果
  const allocSheet = ss.getSheetByName('結果_部品割当');
  if (allocSheet && response.allocations) {
    log('部品割当シート書き込み開始', { rows: response.allocations.length });

    allocSheet.clear();
    const allocData = response.allocations;
    allocSheet.getRange(1, 1, allocData.length, allocData[0].length).setValues(allocData);

    // ヘッダースタイル
    allocSheet.getRange(1, 1, 1, allocData[0].length)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold');

    // 数値フォーマット（3列目以降）
    if (allocData.length > 1) {
      allocSheet.getRange(2, 3, allocData.length - 1, allocData[0].length - 2)
        .setNumberFormat('#,##0');
    }

    log('部品割当シート書き込み完了');
  }

  // 未割当結果
  const unmetSheet = ss.getSheetByName('結果_未割当');
  if (unmetSheet && response.unmet_demands) {
    log('未割当シート書き込み開始', { rows: response.unmet_demands.length });

    unmetSheet.clear();
    const unmetData = response.unmet_demands;
    unmetSheet.getRange(1, 1, unmetData.length, unmetData[0].length).setValues(unmetData);

    // ヘッダースタイル
    unmetSheet.getRange(1, 1, 1, unmetData[0].length)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold');

    // 数値フォーマット（2列目以降）
    if (unmetData.length > 1) {
      unmetSheet.getRange(2, 2, unmetData.length - 1, unmetData[0].length - 1)
        .setNumberFormat('#,##0');

      // 未割当セルを警告色で表示
      for (let i = 2; i <= unmetData.length; i++) {
        unmetSheet.getRange(i, 1, 1, unmetData[0].length).setBackground('#FFC7CE');
      }
    }

    log('未割当シート書き込み完了');
  }

  // 月別能力結果
  const capResultSheet = ss.getSheetByName('結果_月別能力');
  if (capResultSheet && response.capacities) {
    log('月別能力シート書き込み開始', { rows: response.capacities.length });

    capResultSheet.clear();
    const capData = response.capacities;
    capResultSheet.getRange(1, 1, capData.length, capData[0].length).setValues(capData);

    // ヘッダースタイル
    capResultSheet.getRange(1, 1, 1, capData[0].length)
      .setBackground('#4472C4')
      .setFontColor('white')
      .setFontWeight('bold');

    // 数値フォーマット（2列目以降）
    if (capData.length > 1) {
      capResultSheet.getRange(2, 2, capData.length - 1, capData[0].length - 1)
        .setNumberFormat('#,##0');
    }

    log('月別能力シート書き込み完了');
  }
}

// ========================================
// パターン比較最適化（100%/90%/80%）
// ========================================
function runComparisonOptimization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  log('パターン比較最適化実行開始');

  // 入力データ取得
  const inputSheet = ss.getSheetByName('入力');
  if (!inputSheet) {
    ui.alert('エラー', '「入力」シートが見つかりません', ui.ButtonSet.OK);
    return;
  }

  const inputData = inputSheet.getDataRange().getValues();
  if (inputData.length < 2) {
    ui.alert('エラー', '入力データがありません', ui.ButtonSet.OK);
    return;
  }

  // 能力データ取得
  const capSheet = ss.getSheetByName('ライン能力');
  let capacitiesData = null;
  if (capSheet) {
    capacitiesData = capSheet.getDataRange().getValues().slice(1);
  }

  try {
    log('比較API呼び出し開始');

    const response = callCompareApi(inputData, capacitiesData);

    log('比較API呼び出し完了', { success: response.success, patterns: response.patterns });

    if (!response.success) {
      ui.alert('最適化失敗', '全パターンで最適化に失敗しました', ui.ButtonSet.OK);
      return;
    }

    // 比較結果を書き込み
    writeComparisonResults(ss, response);

    ui.alert('完了',
      `パターン比較最適化が完了しました\n\n` +
      `パターン: ${response.patterns.map(p => p + '%').join(', ')}\n` +
      `部品数: ${response.parts_count}\n` +
      `年間総需要: ${response.total_demand.toLocaleString()}\n\n` +
      `結果シートを確認してください。`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    log('エラー発生', { message: error.message, stack: error.stack });
    ui.alert('エラー', `API呼び出しに失敗しました:\n${error.message}`, ui.ButtonSet.OK);
  }
}

// ========================================
// 勤務体制パターン比較最適化
// ========================================
function runWorkPatternComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  log('勤務体制パターン比較実行開始');

  // 入力データ取得
  const inputSheet = ss.getSheetByName('入力');
  if (!inputSheet) {
    ui.alert('エラー', '「入力」シートが見つかりません', ui.ButtonSet.OK);
    return;
  }

  const inputData = inputSheet.getDataRange().getValues();
  if (inputData.length < 2) {
    ui.alert('エラー', '入力データがありません', ui.ButtonSet.OK);
    return;
  }

  // 勤務体制パターン関連シートの読み込み
  const wpSheet = ss.getSheetByName('負荷率計算');
  const jphSheet = ss.getSheetByName('ライン製造能力');
  const daysSheet = ss.getSheetByName('月間稼働日数');

  if (!wpSheet || !jphSheet || !daysSheet) {
    ui.alert('エラー',
      '勤務体制パターン関連シートが見つかりません。\n' +
      'テンプレートを作成してデータを入力してください。\n\n' +
      '必要なシート: 負荷率計算、ライン製造能力、月間稼働日数',
      ui.ButtonSet.OK);
    return;
  }

  // 勤務体制パターン読み込み
  const wpData = wpSheet.getDataRange().getValues();
  const workPatterns = [];
  for (let i = 1; i < wpData.length; i++) {
    const row = wpData[i];
    if (row[0] && row[1]) {
      workPatterns.push({
        name: String(row[0]).trim(),
        formula: String(row[1]).trim(),
        exclusion_hours: parseFloat(row[2]) || 0
      });
    }
  }

  if (workPatterns.length === 0) {
    ui.alert('エラー', '負荷率計算シートに有効なパターンがありません', ui.ButtonSet.OK);
    return;
  }

  // JPHデータ読み込み
  const jphData = jphSheet.getDataRange().getValues();

  // 月間稼働日数読み込み
  const daysData = daysSheet.getDataRange().getValues();
  let monthlyDays = [];
  if (daysData.length >= 2) {
    monthlyDays = daysData[1].slice(0, 12).map(v => parseFloat(v) || 20);
  }

  try {
    log('勤務体制パターンAPI呼び出し開始', { patterns: workPatterns.length });

    const response = callWorkPatternApi(inputData, jphData, workPatterns, monthlyDays);

    log('勤務体制パターンAPI呼び出し完了', {
      success: response.success,
      pattern_names: response.pattern_names
    });

    if (!response.success) {
      ui.alert('最適化失敗', '全パターンで最適化に失敗しました', ui.ButtonSet.OK);
      return;
    }

    // 結果を書き込み
    writeWorkPatternResults(ss, response);

    ui.alert('完了',
      `勤務体制パターン比較最適化が完了しました\n\n` +
      `パターン: ${response.pattern_names.join(', ')}\n` +
      `部品数: ${response.parts_count}\n` +
      `年間総需要: ${response.total_demand.toLocaleString()}\n\n` +
      `結果シートを確認してください。`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    log('エラー発生', { message: error.message, stack: error.stack });
    ui.alert('エラー', `API呼び出しに失敗しました:\n${error.message}`, ui.ButtonSet.OK);
  }
}

// ========================================
// 勤務体制パターンAPI呼び出し
// ========================================
function callWorkPatternApi(partsData, jphData, workPatterns, monthlyDays) {
  const payload = {
    parts_data: partsData,
    jph_data: jphData,
    work_patterns: workPatterns,
    monthly_working_days: monthlyDays,
    time_limit: 60
  };

  log('勤務体制パターンAPIリクエスト送信', { endpoint: '/optimize/simple/compare-patterns' });

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(`${API_URL}/optimize/simple/compare-patterns`, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  log('勤務体制パターンAPIレスポンス受信', { code: responseCode, length: responseText.length });

  if (responseCode !== 200) {
    throw new Error(`HTTPエラー: ${responseCode}\n${responseText.substring(0, 200)}`);
  }

  return JSON.parse(responseText);
}

// ========================================
// 勤務体制パターン結果書き込み
// ========================================
function writeWorkPatternResults(ss, response) {
  const patternNames = response.pattern_names;

  // --- パターン比較サマリーシート ---
  const summarySheet = writeSheetData(ss, '結果_勤務体制比較', response.comparison_summary, {
    headerBg: '#4472C4',
    numberCols: [3, 4, 6],
  });
  if (summarySheet) {
    formatComparisonSummary(summarySheet, response.comparison_summary);
  }

  // --- ライン別負荷率比較シート ---
  writeSheetData(ss, '結果_勤務体制負荷率比較', response.line_comparison, {
    headerBg: '#4472C4',
  });

  // --- 未割当比較シート ---
  if (response.unmet_comparison && response.unmet_comparison.length > 1) {
    writeSheetData(ss, '結果_勤務体制未割当比較', response.unmet_comparison, {
      headerBg: '#4472C4',
      warnRows: true,
    });
  }

  // パターンヘッダー色
  const patternColors = ['#4472C4', '#ED7D31', '#70AD47', '#FFC000'];

  // --- パターン別ライン負荷シート ---
  for (let i = 0; i < patternNames.length; i++) {
    const name = patternNames[i];
    const sheetName = `結果_負荷_${name}`;
    const data = response.patterns_line_loads[name];
    if (data && data.length > 0) {
      const sheet = writeSheetData(ss, sheetName, data, {
        headerBg: patternColors[i % patternColors.length],
        numberStartCol: 2,
        numberEndCol: 13,
      });
      if (sheet) {
        formatLineLoads(sheet, data);
      }
    }
  }

  // --- パターン別部品割当シート ---
  for (let i = 0; i < patternNames.length; i++) {
    const name = patternNames[i];
    const sheetName = `結果_割当_${name}`;
    const data = response.patterns_allocations[name];
    if (data && data.length > 0) {
      const sheet = writeSheetData(ss, sheetName, data, {
        headerBg: patternColors[i % patternColors.length],
        numberStartCol: 3,
        numberEndCol: 15,
      });
      if (sheet) {
        formatAllocations(sheet, data);
      }
    }
  }

  // --- パターン別未割当シート ---
  for (let i = 0; i < patternNames.length; i++) {
    const name = patternNames[i];
    const sheetName = `結果_未割当_${name}`;
    const data = response.patterns_unmet[name];
    if (data && data.length > 1) {
      writeSheetData(ss, sheetName, data, {
        headerBg: patternColors[i % patternColors.length],
        numberStartCol: 2,
        numberEndCol: 14,
        warnRows: true,
      });
    }
  }

  // --- パターン別キャパシティシート ---
  for (let i = 0; i < patternNames.length; i++) {
    const name = patternNames[i];
    const sheetName = `結果_能力_${name}`;
    const data = response.patterns_capacities[name];
    if (data && data.length > 0) {
      const sheet = writeSheetData(ss, sheetName, data, {
        headerBg: patternColors[i % patternColors.length],
        numberStartCol: 2,
        numberEndCol: 13,
      });
      if (sheet) {
        formatCapacities(sheet, data);
      }
    }
  }

  log('勤務体制パターン結果書き込み完了');
}

// ========================================
// 比較API呼び出し
// ========================================
function callCompareApi(partsData, capacitiesData) {
  const payload = {
    parts_data: partsData,
    capacities_data: capacitiesData,
    time_limit: 60
  };

  log('比較APIリクエスト送信', { endpoint: '/optimize/simple/compare' });

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(`${API_URL}/optimize/simple/compare`, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  log('比較APIレスポンス受信', { code: responseCode, length: responseText.length });

  if (responseCode !== 200) {
    throw new Error(`HTTPエラー: ${responseCode}\n${responseText.substring(0, 200)}`);
  }

  return JSON.parse(responseText);
}

// ========================================
// 比較結果書き込み
// ========================================
function writeComparisonResults(ss, response) {
  const patterns = response.patterns; // [100, 90, 80]

  // --- パターン比較サマリーシート ---
  const summarySheet = writeSheetData(ss, '結果_パターン比較', response.comparison_summary, {
    headerBg: '#4472C4',
    numberCols: [3, 4, 6],  // 目的関数値、実行時間、未割当合計
  });
  if (summarySheet) {
    formatComparisonSummary(summarySheet, response.comparison_summary);
  }

  // --- ライン別負荷率比較シート ---
  const lineCompSheet = writeSheetData(ss, '結果_負荷率比較', response.line_comparison, {
    headerBg: '#4472C4',
    numberCols: [2, 3, 5, 7],  // 平均能力 + 平均負荷列
  });
  if (lineCompSheet) {
    formatLineComparison(lineCompSheet, response.line_comparison, patterns);
  }

  // --- 未割当比較シート ---
  if (response.unmet_comparison && response.unmet_comparison.length > 1) {
    writeSheetData(ss, '結果_未割当比較', response.unmet_comparison, {
      headerBg: '#4472C4',
      numberCols: Array.from({length: patterns.length}, (_, i) => i + 2),
      warnRows: true,
    });
  }

  // --- パターン別ライン負荷シート ---
  for (const pct of patterns) {
    const key = `${pct}pct`;
    const sheetName = `結果_負荷_${pct}%`;
    const data = response.patterns_line_loads[key];
    if (data && data.length > 0) {
      const sheet = writeSheetData(ss, sheetName, data, {
        headerBg: getPatternColor(pct),
        numberStartCol: 2,
        numberEndCol: 13,
      });
      if (sheet) {
        formatLineLoads(sheet, data);
      }
    }
  }

  // --- パターン別部品割当シート ---
  for (const pct of patterns) {
    const key = `${pct}pct`;
    const sheetName = `結果_割当_${pct}%`;
    const data = response.patterns_allocations[key];
    if (data && data.length > 0) {
      const sheet = writeSheetData(ss, sheetName, data, {
        headerBg: getPatternColor(pct),
        numberStartCol: 3,
        numberEndCol: 15,
      });
      if (sheet) {
        formatAllocations(sheet, data);
      }
    }
  }

  // --- パターン別未割当シート ---
  for (const pct of patterns) {
    const key = `${pct}pct`;
    const sheetName = `結果_未割当_${pct}%`;
    const data = response.patterns_unmet[key];
    if (data && data.length > 1) {
      writeSheetData(ss, sheetName, data, {
        headerBg: getPatternColor(pct),
        numberStartCol: 2,
        numberEndCol: 14,
        warnRows: true,
      });
    }
  }

  // --- キャパシティシート（共通） ---
  const capSheet = writeSheetData(ss, '結果_月別能力', response.capacities, {
    headerBg: '#4472C4',
    numberStartCol: 2,
    numberEndCol: 13,
  });
  if (capSheet && response.capacities && response.capacities.length > 1) {
    formatCapacities(capSheet, response.capacities);
  }

  log('比較結果書き込み完了');
}

/**
 * シートにデータを書き込む汎用関数
 */
function writeSheetData(ss, sheetName, data, options) {
  if (!data || data.length === 0) return;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();

  // データ書き込み
  const numCols = Math.max(...data.map(row => row.length));
  // 行の列数を揃える
  const normalizedData = data.map(row => {
    const padded = [...row];
    while (padded.length < numCols) padded.push('');
    return padded;
  });

  sheet.getRange(1, 1, normalizedData.length, numCols).setValues(normalizedData);

  // ヘッダースタイル
  const headerBg = options.headerBg || '#4472C4';
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(headerBg)
    .setFontColor('white')
    .setFontWeight('bold');

  // 数値フォーマット（範囲指定）
  if (options.numberStartCol && options.numberEndCol && normalizedData.length > 1) {
    const startCol = options.numberStartCol;
    const endCol = Math.min(options.numberEndCol, numCols);
    sheet.getRange(2, startCol, normalizedData.length - 1, endCol - startCol + 1)
      .setNumberFormat('#,##0');
  }

  // 特定列の数値フォーマット
  if (options.numberCols && normalizedData.length > 1) {
    for (const col of options.numberCols) {
      if (col <= numCols) {
        sheet.getRange(2, col, normalizedData.length - 1, 1)
          .setNumberFormat('#,##0');
      }
    }
  }

  // 未割当行の警告色
  if (options.warnRows && normalizedData.length > 1) {
    for (let i = 2; i <= normalizedData.length; i++) {
      // データ行に値がある場合は警告色
      const rowData = normalizedData[i - 1];
      const hasValue = rowData.slice(1).some(v => v !== '' && v !== 0 && v > 0);
      if (hasValue) {
        sheet.getRange(i, 1, 1, numCols).setBackground('#FFC7CE');
      }
    }
  }

  log(`${sheetName} 書き込み完了`, { rows: normalizedData.length });
  return sheet;
}

/**
 * パターンごとのヘッダー色を取得
 */
function getPatternColor(pct) {
  const colors = {
    100: '#4472C4',  // 青
    90: '#ED7D31',   // オレンジ
    80: '#70AD47',   // 緑
  };
  return colors[pct] || '#4472C4';
}

// ========================================
// 色設定ヘルパー関数
// ========================================

/**
 * パーセント文字列を数値に変換（例: "7.8%" → 7.8）
 */
function parseRatePercent(value) {
  if (typeof value === 'number') return value >= 1 ? value : value * 100;
  if (typeof value === 'string') {
    const num = parseFloat(value.replace('%', ''));
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * 負荷率に基づく背景色を取得
 */
function getLoadRateColor(rate) {
  if (rate > 100) return '#FFC7CE';  // 赤（超過）
  if (rate > 90)  return '#FFEB9C';  // 黄（警戒）
  if (rate > 80)  return '#C6EFCE';  // 薄緑（やや高い）
  return null;
}

/**
 * 負荷率に基づくフォント色を取得
 */
function getLoadRateFontColor(rate) {
  if (rate > 100) return '#9C0006';  // 暗赤
  if (rate > 90)  return '#9C6500';  // 暗黄
  if (rate > 80)  return '#006100';  // 暗緑
  return null;
}

/**
 * ステータスに基づく背景色を取得
 */
function getStatusColor(status) {
  const colors = {
    'OPTIMAL': '#C6EFCE',
    'FEASIBLE': '#FFEB9C',
    'INFEASIBLE': '#FFC7CE',
    'ERROR': '#FFC7CE',
  };
  return colors[status] || null;
}

/**
 * ステータスに基づくフォント色を取得
 */
function getStatusFontColor(status) {
  const colors = {
    'OPTIMAL': '#006100',
    'FEASIBLE': '#9C6500',
    'INFEASIBLE': '#9C0006',
    'ERROR': '#9C0006',
  };
  return colors[status] || null;
}

// ========================================
// 比較結果フォーマット関数
// ========================================

/**
 * パターン比較サマリーシートのフォーマット
 * - ステータス列: OPTIMAL=緑, FEASIBLE=黄, ERROR=赤
 * - 平均負荷率列: 負荷率に応じた色分け
 * - 未割当合計列: 値がある場合は警告色
 */
function formatComparisonSummary(sheet, data) {
  if (!data || data.length <= 1) return;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    // ステータス列（2列目）の色分け
    const status = row[1];
    const statusBg = getStatusColor(status);
    const statusFont = getStatusFontColor(status);
    if (statusBg) {
      sheet.getRange(rowNum, 2).setBackground(statusBg);
    }
    if (statusFont) {
      sheet.getRange(rowNum, 2).setFontColor(statusFont).setFontWeight('bold');
    }

    // 平均負荷率列（5列目）の色分け
    const avgRate = parseRatePercent(row[4]);
    const rateBg = getLoadRateColor(avgRate);
    const rateFont = getLoadRateFontColor(avgRate);
    if (rateBg) {
      sheet.getRange(rowNum, 5).setBackground(rateBg);
    }
    if (rateFont) {
      sheet.getRange(rowNum, 5).setFontColor(rateFont);
    }

    // 未割当合計列（6列目）の警告色
    const unmet = typeof row[5] === 'number' ? row[5] : parseFloat(row[5]) || 0;
    if (unmet > 0) {
      sheet.getRange(rowNum, 6)
        .setBackground('#FFC7CE')
        .setFontColor('#9C0006')
        .setFontWeight('bold');
    }
  }
}

/**
 * ライン別負荷率比較シートのフォーマット
 * - 各パターンの負荷率列を条件付き色分け
 */
function formatLineComparison(sheet, data, patterns) {
  if (!data || data.length <= 1) return;

  // 負荷率列: ライン, 平均能力, 平均負荷(100%), 負荷率(100%), 平均負荷(90%), 負荷率(90%), ...
  // 負荷率列は 4, 6, 8 (1-based)
  const rateCols = patterns.map((_, idx) => 4 + idx * 2);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    for (const col of rateCols) {
      if (col - 1 < row.length) {
        const rate = parseRatePercent(row[col - 1]);
        const bg = getLoadRateColor(rate);
        const font = getLoadRateFontColor(rate);
        if (bg) {
          sheet.getRange(rowNum, col).setBackground(bg);
        }
        if (font) {
          sheet.getRange(rowNum, col).setFontColor(font).setFontWeight('bold');
        }
      }
    }
  }
}

/**
 * パターン別ライン負荷シートのフォーマット
 * - 負荷率列（最終列）の条件付き色分け
 * - 平均能力・平均負荷の数値フォーマット
 * - 100%超えの行全体を警告色
 */
function formatLineLoads(sheet, data) {
  if (!data || data.length <= 1) return;

  const numCols = data[0].length;
  const rateCol = numCols;        // 負荷率（最終列）
  const avgCapCol = numCols - 2;  // 平均能力

  // 平均能力・平均負荷列の数値フォーマット
  sheet.getRange(2, avgCapCol, data.length - 1, 2).setNumberFormat('#,##0');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    const rate = parseRatePercent(row[rateCol - 1]);
    const bg = getLoadRateColor(rate);
    const font = getLoadRateFontColor(rate);

    if (rate > 100) {
      // 100%超え: 行全体を警告色
      sheet.getRange(rowNum, 1, 1, numCols).setBackground('#FFC7CE');
      sheet.getRange(rowNum, rateCol).setFontColor('#9C0006').setFontWeight('bold');
    } else {
      // 負荷率セルのみ色分け
      if (bg) {
        sheet.getRange(rowNum, rateCol).setBackground(bg);
      }
      if (font) {
        sheet.getRange(rowNum, rateCol).setFontColor(font).setFontWeight('bold');
      }
    }
  }
}

/**
 * パターン別部品割当シートのフォーマット
 * - 交互背景色で読みやすさ向上
 */
function formatAllocations(sheet, data) {
  if (!data || data.length <= 1) return;

  const numCols = Math.max(...data.map(row => row.length));

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    if (i % 2 === 0) {
      sheet.getRange(rowNum, 1, 1, numCols).setBackground('#F2F2F2');
    }
  }
}

/**
 * 月別能力シートのフォーマット
 * - 交互背景色
 */
function formatCapacities(sheet, data) {
  if (!data || data.length <= 1) return;

  const numCols = Math.max(...data.map(row => row.length));

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    if (i % 2 === 0) {
      sheet.getRange(rowNum, 1, 1, numCols).setBackground('#F2F2F2');
    }
  }
}

// ========================================
// ヘルプ
// ========================================
function showHelp() {
  const help = `
KIRIU ライン負荷最適化システム

【使い方】
1. メニュー「ライン最適化」→「テンプレートを作成」
2. 「入力」シートに部品データを入力
   - 部品番号、メインライン、サブライン、月別需要
3. 能力設定（以下いずれか）:
   a) 「ライン能力」シートで月別の能力を直接設定
   b) 勤務体制パターン方式:
      - 「負荷率計算」シートで勤務体制と計算式を設定
      - 「ライン製造能力」シートでJPH（時間あたり生産数）を設定
      - 「月間稼働日数」シートで月ごとの稼働日数を設定
4. メニュー「ライン最適化」から実行方法を選択:
   - 「最適化を実行（100%）」: 従来の単一パターン実行
   - 「パターン比較（100/90/80%）」: 負荷率3パターン比較
   - 「勤務体制パターン比較」: 勤務体制ごとに能力を計算して比較
5. 結果シートで確認

【シート構成 - 勤務体制パターン比較時】
- 結果_勤務体制比較: パターンの概要比較
- 結果_勤務体制負荷率比較: ライン別負荷率比較
- 結果_負荷_○○: 各パターンのライン負荷詳細
- 結果_割当_○○: 各パターンの部品割当詳細
- 結果_能力_○○: 各パターンの月別能力

【能力計算式】
月間能力 = JPH × 月稼働時間
月稼働時間 = 計算式（例: 月間稼働日数 × 7.5 × 直数 - 月除外時間）

【重要】
- 負荷率は能力上限を超えません（ハード制約）
- 上限を下げると未割当が増える場合があります

【ログ確認方法】
Apps Script エディタ → 表示 → 実行ログ
  `;

  SpreadsheetApp.getUi().alert('ヘルプ', help, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// テスト用（API無しで動作確認）
// ========================================
function testWithoutApi() {
  log('テスト実行開始');

  const dummyResponse = {
    success: true,
    status: 'OPTIMAL',
    solve_time: 0.5,
    parts_count: 5,
    total_demand: 396000,
    line_loads: [
      ['ライン', '能力', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月', '平均', '負荷率'],
      ['4915', 70000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, '7.1%'],
      ['4919', 80000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, 8000, '10.0%'],
    ],
    allocations: [
      ['部品番号', '割当ライン', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月', '年間計'],
      ['PART001', '4915', 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 60000],
    ]
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeResults(ss, dummyResponse);

  log('テスト実行完了');
  SpreadsheetApp.getUi().alert('テストデータを書き込みました');
}
