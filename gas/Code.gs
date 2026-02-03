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
    .addItem('最適化を実行', 'runOptimization')
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
// ヘルプ
// ========================================
function showHelp() {
  const help = `
KIRIU ライン負荷最適化システム

【使い方】
1. メニュー「ライン最適化」→「テンプレートを作成」
2. 「入力」シートに部品データを入力
   - 部品番号、メインライン、サブライン、月別需要
3. 「ライン能力」シートで月別の能力を調整
   - 各月ごとに異なる能力を設定可能
4. メニュー「ライン最適化」→「最適化を実行」
5. 結果シートで確認

【シート構成】
- 入力: 部品データ入力
- ライン能力: 月別生産能力の設定（黄色セルを編集）
- 結果_ライン負荷: ライン別月別負荷
- 結果_部品割当: 部品別生産割当
- 結果_未割当: 能力超過で生産できなかった数量
- 結果_月別能力: 最適化に使用した月別能力

【重要】
- 負荷率は100%を超えません（ハード制約）
- 超過分は「結果_未割当」シートに表示されます

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
