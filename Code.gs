/**
 * AI経費精算システム - サーバーサイドスクリプト
 * v2.5: 権限修正、スプシURL取得、CSV文字化け(Shift-JIS)対応
 */

// 定数定義：シート名
const SHEETS = {
  RECEIPTS: '領収書',
  DETAILS: '明細',
  ACCOUNTS: '勘定科目マスタ',
  USERS: '使用者マスタ',
  LEARNING: '学習データ',
  RULES: '承認ルール設定',
  CSV_CONFIG: 'CSV出力設定'
};

// APIキー設定
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
// 領収書保存先フォルダID
const RECEIPT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('RECEIPT_FOLDER_ID') || '';

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('AI経費精算システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 初期セットアップ用関数
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const headers = {
    [SHEETS.RECEIPTS]: ['登録ID', '登録日時', '利用日', '使用者', '支払先', '合計金額', 'メモ', 'ファイル名', 'ファイルID', 'インボイス有無', 'ステータス', '現在の承認者'],
    [SHEETS.DETAILS]: ['明細ID', '親登録ID', '勘定科目', '項目名', '取引先', '参加人数', '金額(税込)', '税抜', '消費税', 'メモ'],
    [SHEETS.ACCOUNTS]: ['科目名', 'メモ'],
    [SHEETS.USERS]: ['氏名', 'メールアドレス', '権限', '銀行コード'],
    [SHEETS.LEARNING]: ['品目キーワード', '正解勘定科目'],
    [SHEETS.RULES]: ['ルールID', '優先順位', '対象科目', '金額条件', 'キーワード条件', '必須承認者ルート'],
    [SHEETS.CSV_CONFIG]: ['出力種別', '列順', 'ヘッダー名', '参照元', 'カラム名', 'フォーマット指定']
  };

  for (const [sheetName, headerRow] of Object.entries(headers)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headerRow);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold');
    }
  }
}

/**
 * ログインユーザー情報の取得
 */
function getCurrentUser(manualEmail = null) {
  const email = manualEmail || Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  
  console.log(`Login Attempt: ${email}`);
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      return {
        name: data[i][0],
        email: data[i][1],
        role: data[i][2],
        bankCode: data[i][3]
      };
    }
  }
  
  console.warn(`User not found in master: ${email}`);
  return {
    name: '未登録ユーザー',
    email: email,
    role: 'UNREGISTERED',
    bankCode: ''
  };
}

/**
 * 現在のスプレッドシートのURLを取得する
 */
function getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

/**
 * ダッシュボードデータ取得
 */
function getDashboardData(userEmail) {
  const user = getCurrentUser(userEmail);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEETS.RECEIPTS);
  
  if (user.role === 'UNREGISTERED') {
    return { user: user, myApplications: [], approvalTasks: [] };
  }

  if (rSheet.getLastRow() <= 1) {
    return { user: user, myApplications: [], approvalTasks: [] };
  }

  const data = rSheet.getDataRange().getValues();
  const rows = data.slice(1);

  const myApplications = [];
  const approvalTasks = [];
  
  const currentUserName = String(user.name).trim();
  console.log(`Dashboard Fetch for: ${currentUserName} (Role: ${user.role})`);

  rows.forEach((row, index) => {
    const formatDateStr = (d) => {
      if (d instanceof Date) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }
      return String(d);
    };

    const item = {
      id: row[0],
      date: formatDateStr(row[1]),
      useDate: formatDateStr(row[2]),
      user: row[3],
      store: row[4],
      amount: row[5],
      category: '詳細', 
      status: row[10],
      approver: row[11]
    };

    const rowUserName = String(item.user).trim();

    // 自分の申請
    if (rowUserName === currentUserName) {
      myApplications.push(item);
    } 

    // 承認タスク: 自分の役職 と 現在の承認者 が一致 かつ ステータスが承認待ち
    if (String(item.approver).trim() === String(user.role).trim() && item.status === '承認待ち') {
      approvalTasks.push(item);
    }
  });

  return {
    user: user,
    myApplications: myApplications.reverse(),
    approvalTasks: approvalTasks
  };
}

/**
 * Gemini APIを使用して画像を解析する
 */
function analyzeReceiptImage(base64Image, mimeType) {
  if (!GEMINI_API_KEY) {
    throw new Error('Gemini APIキーが設定されていません。');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const prompt = `
    あなたは、経費精算システムにデータ入力を行う、非常に精度が高いAIアシスタントです。
    提供された画像の領収書を分析し、以下のJSON形式で出力してください。
    
    抽出項目:
    use_date: YYYY/MM/DD形式
    store_name: 支払先名称
    has_invoice: インボイス登録番号(T+13桁)があればtrue, なければfalse
    client: 取引先・顧客名(あれば)
    participants: 参加人数(数値, あれば)
    category: 勘定科目(旅費交通費, 交際費, 消耗品, 会議費, 研修費 から推測)
    total_amount: 税込合計金額(数値)
    items: 明細リスト [{name, subtotal, tax, total_price}]

    JSON以外の説明文は不要です。配列形式で返してください。
  `;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64Image } }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(json.error.message);
    }

    const text = json.candidates[0].content.parts[0].text;
    const jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(jsonStr);

  } catch (e) {
    throw new Error('AI解析に失敗しました: ' + e.message);
  }
}

/**
 * 承認ルート取得ロジック
 */
function getApprovalRoute(category, amount, items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RULES);
  const data = sheet.getDataRange().getValues();
  const rules = data.slice(1).sort((a, b) => a[1] - b[1]);

  for (const rule of rules) {
    const targetCategory = rule[2];
    const amountLimit = rule[3];
    const keyword = rule[4];
    const route = rule[5];

    let match = true;
    if (targetCategory && targetCategory !== '*' && targetCategory !== category) match = false;
    if (amountLimit && amount < amountLimit) match = false;
    if (keyword) {
      const allText = items.map(i => i.name).join(' ');
      if (!allText.includes(keyword)) match = false;
    }

    if (match) {
      return route.split(',').map(s => s.trim()).filter(s => s); 
    }
  }
  return [];
}

/**
 * 申請データを保存する
 */
function saveApplication(formData, itemsData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEETS.RECEIPTS);
  const dSheet = ss.getSheetByName(SHEETS.DETAILS);
  
  const regId = Utilities.getUuid();
  const now = new Date();
  
  // 1. ファイル保存処理
  let fileId = '';
  if (formData.fileBase64 && formData.mimeType && formData.fileName) {
    try {
      let folder;
      if (RECEIPT_FOLDER_ID) {
        folder = DriveApp.getFolderById(RECEIPT_FOLDER_ID);
      } else {
        folder = DriveApp.getRootFolder();
      }
      const decoded = Utilities.base64Decode(formData.fileBase64);
      const blob = Utilities.newBlob(decoded, formData.mimeType, formData.fileName);
      const file = folder.createFile(blob);
      fileId = file.getId(); 
    } catch (e) {
      console.error('File Save Error:', e.message);
      fileId = 'SAVE_ERROR'; 
    }
  } else {
    fileId = 'NO_FILE'; 
  }

  // 2. 承認ルート判定
  const safeItems = itemsData || [];
  const route = getApprovalRoute(formData.category, formData.totalAmount, safeItems);
  
  const firstApprover = route.length > 0 ? route[0] : '承認済';
  const initialStatus = (firstApprover === '承認済') ? '承認済' : '承認待ち';

  // 3. 領収書データ保存
  rSheet.appendRow([
    regId,
    now,
    formData.useDate,
    formData.userName,
    formData.storeName,
    formData.totalAmount,
    formData.memo,
    formData.fileName || '領収書なし',
    fileId,
    formData.hasInvoice,
    initialStatus,
    firstApprover
  ]);

  // 4. 明細データ保存
  safeItems.forEach(item => {
    dSheet.appendRow([
      Utilities.getUuid(),
      regId,
      formData.category, 
      item.name,
      formData.client || '',
      formData.participants || '',
      item.total_price || 0,
      item.subtotal || 0,
      item.tax || 0,
      ''
    ]);
  });
  
  SpreadsheetApp.flush();

  return { success: true, message: '申請が完了しました', id: regId };
}

/**
 * 承認実行処理
 */
function approveApplication(id, userEmail) {
  const user = getCurrentUser(userEmail); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName(SHEETS.RECEIPTS);
  const rData = rSheet.getDataRange().getValues();

  for (let i = 1; i < rData.length; i++) {
    if (rData[i][0] == id) {
      const currentRole = rData[i][11];
      const totalAmount = rData[i][5];
      
      const dSheet = ss.getSheetByName(SHEETS.DETAILS);
      const dData = dSheet.getDataRange().getValues();
      const relatedItems = dData.slice(1).filter(d => d[1] == id);
      
      if (relatedItems.length === 0) {
         var category = 'その他'; 
         var itemsForCheck = [];
      } else {
         var category = relatedItems[0][2];
         var itemsForCheck = relatedItems.map(r => ({ name: r[3] }));
      }

      const route = getApprovalRoute(category, totalAmount, itemsForCheck);
      const currentIndex = route.indexOf(currentRole);
      
      if (currentIndex !== -1 && currentIndex < route.length - 1) {
        const nextApprover = route[currentIndex + 1];
        rSheet.getRange(i + 1, 12).setValue(nextApprover);
      } else {
        rSheet.getRange(i + 1, 11).setValue('承認済');
        rSheet.getRange(i + 1, 12).setValue('承認済');
      }
      
      return { success: true };
    }
  }
  return { success: false, message: 'IDが見つかりません' };
}

function rejectApplication(id, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RECEIPTS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 11).setValue('差戻し');
      const currentMemo = data[i][6];
      sheet.getRange(i + 1, 7).setValue(currentMemo + `\n[差戻し理由] ${comment}`);
      return { success: true };
    }
  }
}

function saveLearningData(keyword, correctCategory) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LEARNING);
  sheet.appendRow([keyword, correctCategory]);
  return { success: true };
}

function getMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName(SHEETS.ACCOUNTS);
  const accounts = accSheet.getDataRange().getValues().slice(1).map(r => r[0]);
  return { accounts: accounts };
}

/**
 * CSV出力処理 (Shift-JIS対応版)
 * 文字列ではなく、Base64エンコードされたShift-JISデータを返す
 */
function generateCSV(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CSV_CONFIG);
  const configs = configSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] === type)
    .sort((a, b) => a[1] - b[1]);

  if (configs.length === 0) throw new Error('CSV設定が見つかりません: ' + type);

  const rSheet = ss.getSheetByName(SHEETS.RECEIPTS);
  const rData = rSheet.getDataRange().getValues();
  const rHeaders = rData[0];
  const rRows = rData.slice(1).filter(r => r[10] === '承認済');

  let csvContent = configs.map(c => c[2]).join(',') + '\r\n'; // Windows向け改行

  rRows.forEach(row => {
    const line = configs.map(config => {
      const source = config[3];
      const colName = config[4];
      const format = config[5];
      let value = '';

      if (source === '領収書') {
        const colIndex = rHeaders.indexOf(colName);
        if (colIndex > -1) value = row[colIndex];
      } else if (source === '固定値') {
        value = colName;
      }

      if (format === 'YYYYMMDD' && value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyyMMdd');
      }

      return `"${value}"`;
    }).join(',');
    csvContent += line + '\r\n';
  });

  // Shift-JISに変換してBlobを作成
  const blob = Utilities.newBlob('', 'text/csv', type + '.csv').setDataFromString(csvContent, 'Shift_JIS');
  
  // Base64にエンコードして返す (フロントエンドでデコードする)
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: `${type}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}.csv`
  };
}
