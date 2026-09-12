/**
 * BAKERU 商品・在庫管理 v2.7.1 - Google Apps Script backend
 * Bound Script（現在のBAKERUスプレッドシートに紐づけて使用）
 *
 * Script Properties:
 *   GEMINI_API_KEY   必須（Gemini利用時）
 *   GEMINI_MODEL     任意（既定: gemini-3.6-flash）
 *   BAKERU_SYNC_TOKEN setupBAKERU() が自動生成
 *   SPREADSHEET_ID   setupBAKERU() が自動保存
 */

const BAKERU_VERSION = '2.7.0-qr-inventory';

const BAKERU_SHEETS = {
  settings: '00_設定',
  products: '01_商品マスタ',
  movements: '02_入出庫履歴',
  receipts: '03_領収書OCR',
  stocktake: '04_棚卸履歴',
  alerts: '05_差異アラート',
  summary: '06_在庫集計',
  logs: '07_同期ログ'
};

const SCHEMAS = {};
SCHEMAS[BAKERU_SHEETS.settings] = ['項目','値','説明'];
SCHEMAS[BAKERU_SHEETS.products] = ['商品ID','SKU','バーコード','商品名','カテゴリ','単位','仕入単価','販売単価','在庫下限','現在庫','有効','別名','更新日時','仕入単位','換算量','注意在庫','通知Gmail','内容量','内容量単位','商品特徴','棚・保管場所'];
SCHEMAS[BAKERU_SHEETS.movements] = ['取引ID','日時','種別','商品ID','SKU','バーコード','商品名','数量','在庫増減','単価','金額','入力元','領収書ID','備考','担当'];
SCHEMAS[BAKERU_SHEETS.receipts] = ['領収書ID','読取日時','仕入先','領収書日付','領収書合計','行番号','OCR商品名','商品ID','SKU','バーコード','商品名','領収書数量','検品数量','差異','単価','金額','検品状態','備考'];
SCHEMAS[BAKERU_SHEETS.stocktake] = ['棚卸ID','棚卸日時','商品ID','SKU','バーコード','商品名','帳簿在庫','実地在庫','差異','仕入単価','差異金額','原因','対応','担当'];
SCHEMAS[BAKERU_SHEETS.alerts] = ['アラートID','発生日時','種類','重要度','商品ID','商品名','期待値','実績値','差異','領収書ID','状態','詳細'];
SCHEMAS[BAKERU_SHEETS.summary] = ['商品ID','SKU','バーコード','商品名','カテゴリ','現在庫','在庫下限','在庫状態','仕入単価','在庫原価','販売単価','更新日時'];
SCHEMAS[BAKERU_SHEETS.logs] = ['日時','処理','結果','詳細'];


// --- テストモード用シート（本番データとは完全分離） ---
const BAKERU_TEST_SHEETS = {
  settings: 'T00_設定',
  products: 'T01_商品マスタ',
  movements: 'T02_入出庫履歴',
  receipts: 'T03_領収書OCR',
  stocktake: 'T04_棚卸履歴',
  alerts: 'T05_差異アラート',
  summary: 'T06_在庫集計',
  logs: 'T07_同期ログ'
};

const TEST_SCHEMAS = {};
TEST_SCHEMAS[BAKERU_TEST_SHEETS.settings] = SCHEMAS[BAKERU_SHEETS.settings];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.products] = SCHEMAS[BAKERU_SHEETS.products];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.movements] = SCHEMAS[BAKERU_SHEETS.movements];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.receipts] = SCHEMAS[BAKERU_SHEETS.receipts];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.stocktake] = SCHEMAS[BAKERU_SHEETS.stocktake];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.alerts] = SCHEMAS[BAKERU_SHEETS.alerts];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.summary] = SCHEMAS[BAKERU_SHEETS.summary];
TEST_SCHEMAS[BAKERU_TEST_SHEETS.logs] = SCHEMAS[BAKERU_SHEETS.logs];

/**
 * STEP 1: テスト専用シートを作成するだけの安全な初期化。
 * この段階では本番APIの保存先はまだ変更しない。
 * 本番商品マスタから最大4商品をコピーし、テスト側だけで入出庫を試せる土台を作る。
 */
function setupBAKERUTestMode() {
  const ss = getSs_();
  Object.keys(TEST_SCHEMAS).forEach(name => ensureSheet_(ss, name, TEST_SCHEMAS[name]));

  const src = ss.getSheetByName(BAKERU_SHEETS.products);
  const dst = ss.getSheetByName(BAKERU_TEST_SHEETS.products);
  if (src && dst) {
    // 既存のテスト商品がなければ、現在の本番商品から最大4件だけコピー。
    const hasTestProducts = dst.getLastRow() >= 2;
    if (!hasTestProducts && src.getLastRow() >= 2) {
      const rows = src.getRange(2, 1, src.getLastRow() - 1, SCHEMAS[BAKERU_SHEETS.products].length).getValues()
        .filter(r => r.some(v => v !== '' && v !== null) && String(r[10]).toUpperCase() !== 'FALSE')
        .slice(0, 4)
        .map(r => {
          const copy = r.slice();
          copy[0] = 'TEST_' + String(copy[0] || Utilities.getUuid());
          copy[12] = new Date();
          return copy;
        });
      if (rows.length) dst.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
  }

  const testSettings = ss.getSheetByName(BAKERU_TEST_SHEETS.settings);
  if (testSettings) {
    testSettings.clearContents();
    testSettings.getRange(1,1,1,TEST_SCHEMAS[BAKERU_TEST_SHEETS.settings].length)
      .setValues([TEST_SCHEMAS[BAKERU_TEST_SHEETS.settings]]);
    testSettings.getRange(2,1,3,3).setValues([
      ['MODE', 'TEST', 'このT00〜T07はテスト専用。通常シートとは分離'],
      ['注意', '本番データではありません', 'テスト完了後に削除しても本番には影響しません'],
      ['次の手順', '画面にテスト/本番切替を追加', 'STEP 2で実装']
    ]);
  }

  Object.keys(TEST_SCHEMAS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const cols = TEST_SCHEMAS[name].length;
    sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  const msg = 'テスト用シート T00〜T07 を作成しました。次は画面のテストモード切替を実装します。';
  try { ss.toast(msg, 'BAKERU テスト準備', 8); } catch (_) {}
  console.log(msg);
  return { ok:true, created:Object.values(BAKERU_TEST_SHEETS) };
}

function setupBAKERU() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('このスクリプトをBAKERUのスプレッドシートに紐づけて実行してください。');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('BAKERU_SYNC_TOKEN')) props.setProperty('BAKERU_SYNC_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  const currentModel = props.getProperty('GEMINI_MODEL');
  if (!currentModel || currentModel === 'gemini-3.8-flash') props.setProperty('GEMINI_MODEL', 'gemini-3.6-flash');

  Object.keys(SCHEMAS).forEach(name => ensureSheet_(ss, name, SCHEMAS[name]));
  // 既にテスト用シートがある場合は、v2.7.1の追加列だけ安全に拡張する。
  Object.keys(TEST_SCHEMAS).forEach(name => { if (ss.getSheetByName(name)) ensureSheet_(ss, name, TEST_SCHEMAS[name]); });
  writeSettings_();
  rebuildSummary_();
  formatAllSheets_();
  logSync_('setupBAKERU', 'OK', '初期シート作成・設定完了');

  // Apps Scriptエディタから直接実行しても待機しないよう、モーダルalertは使わない。
  const message = '初期設定完了。GEMINI_API_KEYをスクリプトプロパティに設定してください。';
  try { ss.toast(message, 'BAKERU v2.7.1', 8); } catch (_) {}
  console.log(message);
  console.log('GEMINI_MODEL=' + props.getProperty('GEMINI_MODEL'));
  console.log('SPREADSHEET_ID=' + ss.getId());
  return {
    ok: true,
    version: BAKERU_VERSION,
    spreadsheetId: ss.getId(),
    geminiModel: props.getProperty('GEMINI_MODEL'),
    syncTokenConfigured: !!props.getProperty('BAKERU_SYNC_TOKEN')
  };
}

function setGeminiModel36Flash() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('GEMINI_MODEL', 'gemini-3.6-flash');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { if (ss) ss.toast('Geminiモデルを gemini-3.6-flash に設定しました。', 'BAKERU', 6); } catch (_) {}
  console.log('GEMINI_MODEL=gemini-3.6-flash');
  return 'gemini-3.6-flash';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('BAKERU管理')
    .addItem('初期セットアップ', 'setupBAKERU')
    .addItem('接続情報を表示', 'showConnectionInfo')
    .addItem('テスト用シートを作成', 'setupBAKERUTestMode')
    .addItem('Gemini 3.6 Flashに設定', 'setGeminiModel36Flash')
    .addItem('旧4_棚卸表から移行', 'migrateLegacyStocktakeMenu')
    .addSeparator()
    .addItem('在庫集計を再構築', 'rebuildSummaryMenu')
    .addToUi();
}

function showConnectionInfo() {
  const props = PropertiesService.getScriptProperties();
  const url = ScriptApp.getService().getUrl() || 'Webアプリを一度デプロイするとURLが表示されます';
  const token = props.getProperty('BAKERU_SYNC_TOKEN') || 'setupBAKERU()を実行してください';
  SpreadsheetApp.getUi().alert('BAKERU 接続情報', `GAS WebアプリURL:\n${url}\n\nBAKERU同期トークン:\n${token}\n\nこの2つをBAKERU管理ツールの「設定・連携」に入力します。`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function migrateLegacyStocktakeMenu() {
  const r = migrateLegacyStocktake_();
  SpreadsheetApp.getUi().alert(`移行完了: ${r.imported}件 / スキップ: ${r.skipped}件`);
}

function rebuildSummaryMenu() {
  rebuildSummary_();
  SpreadsheetApp.getUi().alert('在庫集計を再構築しました。');
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(req.action || '');
    verifyToken_(req.token);
    const p = req.payload || {};
    let data;

    switch (action) {
      case 'ping': data = ping_(); break;
      case 'bootstrap': data = bootstrap_(); break;
      case 'testBootstrap': data = testBootstrap_(); break;
      case 'upsertProduct': data = { product: upsertProduct_(p) }; break;
      case 'testUpsertProduct': data = { product: testUpsertProduct_(p) }; break;
      case 'deleteProduct': data = deleteProduct_(p.id); break;
      case 'recordMovement': data = recordMovement_(p); break;
      case 'testRecordMovement': data = testRecordMovement_(p); break;
      case 'geminiReceiptOcr': data = geminiReceiptOcr_(p.image); break;
      case 'geminiIdentifyProduct': data = geminiIdentifyProduct_(p.image); break;
      case 'commitReceipt': data = commitReceipt_(p); break;
      case 'saveStocktake': data = saveStocktake_(p); break;
      case 'resolveAlert': data = resolveAlert_(p.id); break;
      case 'geminiAnalysis': data = { text: geminiAnalysis_(p) }; break;
      case 'migrateLegacyStocktake': data = migrateLegacyStocktake_(); break;
      case 'nekkoExport': data = nekkoExport_(); break;
      default: throw new Error('未対応のactionです: ' + action);
    }
    if (action.indexOf('test') === 0) logTestSync_(action, 'OK', summarizePayload_(p));
    else logSync_(action, 'OK', summarizePayload_(p));
    return json_({ ok: true, version: BAKERU_VERSION, data: data });
  } catch (err) {
    try { logSync_('doPost', 'ERROR', String(err && err.stack || err)); } catch (_) {}
    return json_({ ok: false, version: BAKERU_VERSION, error: String(err && err.message || err) });
  }
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'ping');
    verifyToken_((e && e.parameter && e.parameter.token) || '');
    let data;
    if (action === 'ping') data = ping_();
    else if (action === 'nekkoExport') data = nekkoExport_();
    else if (action === 'bootstrap') data = bootstrap_();
    else throw new Error('未対応のactionです: ' + action);
    return json_({ ok: true, version: BAKERU_VERSION, data: data });
  } catch (err) {
    return json_({ ok: false, version: BAKERU_VERSION, error: String(err && err.message || err) });
  }
}

function ping_() {
  const ss = getSs_();
  return {
    version: BAKERU_VERSION,
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    time: new Date().toISOString(),
    geminiConfigured: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
  };
}

function bootstrap_() {
  return {
    inventory: getProducts_(),
    transactions: getMovements_(500),
    alerts: getAlerts_(300)
  };
}

function getSs_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('SPREADSHEET_IDが未設定です。setupBAKERU()を実行してください。');
  return active;
}

function verifyToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('BAKERU_SYNC_TOKEN');
  if (!expected) throw new Error('BAKERU_SYNC_TOKENが未設定です。setupBAKERU()を実行してください。');
  if (!token || String(token) !== String(expected)) throw new Error('同期トークンが一致しません。');
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = sh.getLastColumn() > 0 ? sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn())).getValues()[0] : [];
  const needsHeader = headers.some((h, i) => current[i] !== h);
  if (needsHeader) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function formatAllSheets_() {
  const ss = getSs_();
  Object.keys(SCHEMAS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const cols = SCHEMAS[name].length;
    sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, cols); } catch (_) {}
    for (let c=1;c<=cols;c++) if (sh.getColumnWidth(c) > 320) sh.setColumnWidth(c, 320);
  });
  const alertSh = ss.getSheetByName(BAKERU_SHEETS.alerts);
  if (alertSh) alertSh.getRange('A:L').setWrap(true);
}

function writeSettings_() {
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.settings);
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  const values = [
    ['BAKERU_VERSION', BAKERU_VERSION, '管理ツールのデータ構造バージョン'],
    ['GEMINI_MODEL', model, 'Geminiモデル名。実際のAPIキーはScript Propertiesに保存'],
    ['在庫の正本', '01_商品マスタ 現在庫', 'すべての入出庫・棚卸調整で自動更新'],
    ['入出庫履歴', '02_入出庫履歴', '仕入・販売・OCR入庫・棚卸調整を1本の台帳で管理'],
    ['旧棚卸表', '4_棚卸表', '削除せず初回移行元として保持']
  ];
  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([SCHEMAS[BAKERU_SHEETS.settings]]);
  sh.getRange(2,1,values.length,3).setValues(values);
}

function rowsAsObjects_(sheetName) {
  const sh = getSs_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '' && v !== null)).map((r, idx) => {
    const o = { _row: idx + 2 };
    headers.forEach((h,i) => o[h] = r[i]);
    return o;
  });
}

function getProducts_() {
  return rowsAsObjects_(BAKERU_SHEETS.products)
    .filter(r => String(r['有効']).toUpperCase() !== 'FALSE' && r['商品ID'])
    .map(r => ({
      id: String(r['商品ID']), sku: String(r['SKU'] || ''), barcode: String(r['バーコード'] || ''),
      name: String(r['商品名'] || ''), category: String(r['カテゴリ'] || ''), unit: String(r['単位'] || '個'),
      cost: num_(r['仕入単価']), price: num_(r['販売単価']), minStock: num_(r['在庫下限']),
      stock: num_(r['現在庫']), active: true, aliases: String(r['別名'] || ''),
      purchaseUnit: String(r['仕入単位'] || r['単位'] || '個'), purchaseFactor: Math.max(0.000001, num_(r['換算量']) || 1),
      warningStock: num_(r['注意在庫']) > num_(r['在庫下限']) ? num_(r['注意在庫']) : Math.max(num_(r['在庫下限']) * 2, num_(r['在庫下限']) + 1),
      notifyEmail: String(r['通知Gmail'] || ''), contentAmount: num_(r['内容量']), contentUnit: String(r['内容量単位'] || 'g'),
      details: String(r['商品特徴'] || ''), shelf: String(r['棚・保管場所'] || ''),
      updatedAt: dateMs_(r['更新日時']) || Date.now()
    }));
}

function getMovements_(limit) {
  const rows = rowsAsObjects_(BAKERU_SHEETS.movements);
  return rows.slice(-Math.max(1, limit || 500)).reverse().map(r => ({
    id: String(r['取引ID'] || ''), timestamp: dateMs_(r['日時']) || Date.now(),
    type: String(r['種別'] || '') === '販売' ? 'sale' : String(r['種別'] || '') === '使用' ? 'use' : num_(r['在庫増減']) < 0 ? 'out' : 'in', productId: String(r['商品ID'] || ''),
    productName: String(r['商品名'] || ''), barcode: String(r['バーコード'] || ''),
    quantity: Math.abs(num_(r['数量'])), unitPrice: num_(r['単価']), source: String(r['入力元'] || ''), note: String(r['備考'] || '')
  }));
}

function getAlerts_(limit) {
  const rows = rowsAsObjects_(BAKERU_SHEETS.alerts);
  return rows.slice(-Math.max(1, limit || 300)).reverse().map(r => ({
    id: String(r['アラートID'] || ''), alertId: String(r['アラートID'] || ''), timestamp: dateMs_(r['発生日時']) || Date.now(),
    type: String(r['種類'] || ''), severity: String(r['重要度'] || ''), productId: String(r['商品ID'] || ''),
    productName: String(r['商品名'] || ''), expected: r['期待値'], actual: r['実績値'], difference: num_(r['差異']),
    receiptId: String(r['領収書ID'] || ''), status: String(r['状態'] || '未確認'), detail: String(r['詳細'] || ''), message: String(r['詳細'] || '')
  }));
}

function upsertProduct_(p) {
  if (!p || !String(p.name || '').trim()) throw new Error('商品名は必須です。');
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.products);
  const rows = rowsAsObjects_(BAKERU_SHEETS.products);
  const id = String(p.id || Utilities.getUuid());
  const existing = rows.find(r => String(r['商品ID']) === id);
  const barcode = String(p.barcode || '').trim();
  if (barcode) {
    const dup = rows.find(r => String(r['バーコード'] || '').trim() === barcode && String(r['商品ID']) !== id && String(r['有効']).toUpperCase() !== 'FALSE');
    if (dup) throw new Error(`このバーコードは既に「${dup['商品名']}」に登録されています。`);
  }
  const sku = String(p.sku || existing && existing['SKU'] || nextSku_(rows));
  const initialStock = existing ? num_(existing['現在庫']) : Math.max(0, num_(p.stock));
  const row = [id, sku, barcode, String(p.name).trim(), String(p.category || ''), String(p.unit || '個'), num_(p.cost), num_(p.price), Math.max(0,num_(p.minStock)), existing ? initialStock : 0, true, String(p.aliases || ''), new Date(), String(p.purchaseUnit || p.unit || '個'), Math.max(0.000001, num_(p.purchaseFactor) || 1), Math.max(Math.max(0,num_(p.minStock)) + 0.001, num_(p.warningStock) || Math.max(num_(p.minStock) * 2, num_(p.minStock) + 1)), String(p.notifyEmail || ''), Math.max(0,num_(p.contentAmount)), String(p.contentUnit || 'g'), String(p.details || ''), String(p.shelf || '')];
  if (existing) sh.getRange(existing._row,1,1,row.length).setValues([row]);
  else sh.appendRow(row);

  if (!existing && initialStock > 0) {
    recordMovementInternal_({ id:'init_'+Utilities.getUuid(), productId:id, type:'in', quantity:initialStock, unitPrice:num_(p.cost), source:'初期登録', note:'新規商品登録時の初期在庫' }, false);
  }
  rebuildSummary_();
  return getProducts_().find(x => x.id === id);
}

function deleteProduct_(id) {
  if (!id) throw new Error('商品IDがありません。');
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.products);
  const row = rowsAsObjects_(BAKERU_SHEETS.products).find(r => String(r['商品ID']) === String(id));
  if (!row) throw new Error('商品が見つかりません。');
  sh.getRange(row._row, 11).setValue(false);
  sh.getRange(row._row, 13).setValue(new Date());
  rebuildSummary_();
  return { deleted: true, inventory: getProducts_() };
}

function recordMovement_(p) {
  recordMovementInternal_(p, true);
  return { inventory: getProducts_(), transactions: getMovements_(500), alerts: getAlerts_(300) };
}

function recordMovementInternal_(p, doSummary) {
  const product = findProductRow_(p.productId);
  if (!product) throw new Error('商品が見つかりません。');
  const qty = Math.abs(num_(p.quantity));
  if (qty <= 0) throw new Error('数量は1以上で入力してください。');
  let signed;
  let kind;
  if (p.type === 'sale' || p.type === '販売') { signed = -qty; kind = '販売'; }
  else if (p.type === 'use' || p.type === '使用') { signed = -qty; kind = '使用'; }
  else if (p.type === 'out' || p.type === '出庫') { signed = -qty; kind = '出庫'; }
  else if (p.type === 'adjustment') { signed = num_(p.signedQuantity); kind = '棚卸調整'; }
  else { signed = qty; kind = String(p.type || '') === 'initial' ? '初期在庫' : '入庫'; }

  const before = num_(product['現在庫']);
  const after = before + signed;
  if (after < 0) throw new Error(`${product['商品名']}の在庫が不足しています（現在庫 ${before}）。`);

  const defaultUnitPrice = kind === '販売' ? num_(product['販売単価']) : num_(product['仕入単価']);
  const unitPrice = (p.unitPrice !== undefined && p.unitPrice !== null && num_(p.unitPrice) > 0) ? num_(p.unitPrice) : defaultUnitPrice;
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.movements);
  sh.appendRow([
    String(p.id || 'tx_'+Utilities.getUuid()), new Date(p.timestamp || Date.now()), kind, String(product['商品ID']), String(product['SKU'] || ''),
    String(product['バーコード'] || ''), String(product['商品名'] || ''), qty, signed, unitPrice, Math.abs(signed) * unitPrice,
    String(p.source || '手入力'), String(p.receiptId || ''), String(p.note || ''), String(p.operator || '')
  ]);
  const productSh = getSs_().getSheetByName(BAKERU_SHEETS.products);
  productSh.getRange(product._row, 10).setValue(after);
  productSh.getRange(product._row, 13).setValue(new Date());

  syncLowStockAlerts_(product, after);
  if (doSummary !== false) rebuildSummary_();
  return after;
}

function syncLowStockAlerts_(product, stock) {
  const critical = Math.max(0, num_(product['在庫下限']));
  const warningRaw = num_(product['注意在庫']);
  const warning = warningRaw > critical ? warningRaw : Math.max(critical * 2, critical + 1);
  const productId = String(product['商品ID']);
  const name = String(product['商品名']);
  const rows = rowsAsObjects_(BAKERU_SHEETS.alerts);
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.alerts);
  const active = rows.filter(r => String(r['商品ID']) === productId && ['在庫注意','在庫レッド','在庫不足'].includes(String(r['種類'])) && String(r['状態']) !== '解決済み');

  const resolveTypes = (types) => active.filter(r => types.includes(String(r['種類']))).forEach(r => sh.getRange(r._row, 11).setValue('解決済み'));
  const hasType = (type) => active.some(r => String(r['種類']) === type);

  if (stock <= critical) {
    resolveTypes(['在庫注意','在庫不足']);
    if (!hasType('在庫レッド')) {
      const detail = `🔴 現在庫 ${stock} がレッドライン ${critical} 以下です。早急に発注・補充を確認してください。`;
      addAlert_({ type:'在庫レッド', severity:'重要', productId, productName:name, expected:critical, actual:stock, difference:stock-critical, detail });
      notifyLowStockEmail_(product, '🔴 レッド在庫', detail);
    }
    return;
  }
  if (stock <= warning) {
    resolveTypes(['在庫レッド','在庫不足']);
    if (!hasType('在庫注意')) {
      const detail = `🟡 現在庫 ${stock} が注意ライン ${warning} 以下です。発注準備を確認してください。`;
      addAlert_({ type:'在庫注意', severity:'注意', productId, productName:name, expected:warning, actual:stock, difference:stock-warning, detail });
      notifyLowStockEmail_(product, '🟡 注意在庫', detail);
    }
    return;
  }
  resolveTypes(['在庫注意','在庫レッド','在庫不足']);
}

function notifyLowStockEmail_(product, level, detail) {
  const productEmail = String(product['通知Gmail'] || '').trim();
  const commonEmail = String(PropertiesService.getScriptProperties().getProperty('BAKERU_ALERT_EMAIL') || '').trim();
  const to = productEmail || commonEmail;
  if (!to) return;
  try {
    MailApp.sendEmail({
      to,
      subject: `[BAKERU] ${level}：${String(product['商品名'] || '')}`,
      body: `${detail}\n\n商品ID: ${String(product['商品ID'] || '')}\nSKU: ${String(product['SKU'] || '')}\n棚・保管場所: ${String(product['棚・保管場所'] || '未設定')}\n\nBAKERU 商品・在庫管理`
    });
  } catch (err) {
    console.warn('BAKERU low stock email failed', err);
  }
}

function resolveLowStockAlert_(productId) {
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.alerts);
  rowsAsObjects_(BAKERU_SHEETS.alerts)
    .filter(r => String(r['商品ID']) === String(productId) && ['在庫注意','在庫レッド','在庫不足'].includes(String(r['種類'])) && String(r['状態']) !== '解決済み')
    .forEach(r => sh.getRange(r._row, 11).setValue('解決済み'));
}

function resolveAlert_(id) {
  if (!id) throw new Error('アラートIDがありません。');
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.alerts);
  const row = rowsAsObjects_(BAKERU_SHEETS.alerts).find(r => String(r['アラートID']) === String(id));
  if (!row) throw new Error('アラートが見つかりません。');
  sh.getRange(row._row, 11).setValue('解決済み');
  return { alerts:getAlerts_(300) };
}

function addAlert_(a) {
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.alerts);
  sh.appendRow([
    String(a.id || 'alert_'+Utilities.getUuid()), new Date(), String(a.type || '差異'), String(a.severity || '要確認'), String(a.productId || ''), String(a.productName || ''),
    a.expected === undefined ? '' : a.expected, a.actual === undefined ? '' : a.actual, a.difference === undefined ? '' : a.difference,
    String(a.receiptId || ''), String(a.status || '未確認'), String(a.detail || '')
  ]);
}

function geminiReceiptOcr_(imageDataUrl) {
  if (!imageDataUrl) throw new Error('領収書画像がありません。');
  const schema = {
    type:'OBJECT', properties:{
      receiptId:{type:'STRING'}, vendor:{type:'STRING'}, date:{type:'STRING'}, total:{type:'NUMBER'},
      items:{type:'ARRAY',items:{type:'OBJECT',properties:{name:{type:'STRING'},quantity:{type:'NUMBER'},unitPrice:{type:'NUMBER'},amount:{type:'NUMBER'}}}}
    }
  };
  const prompt = [
    '日本語の領収書・レシート・納品書画像を在庫入庫用に正確に読み取ってください。',
    '商品明細ごとに商品名、数量、単価、金額を抽出してください。数量が省略されている商品は1としてください。',
    '値引き・小計・消費税・合計は商品明細として扱わないでください。',
    'receiptIdは空文字で構いません。dateは可能ならYYYY-MM-DD。totalは支払合計。',
    '判読できない値を勝手に作らず、商品名は画像上の表記をできるだけ保ってください。'
  ].join('\n');
  const result = callGeminiJson_(imageDataUrl, prompt, schema);
  result.receiptId = result.receiptId || 'receipt_'+Utilities.getUuid();
  result.items = Array.isArray(result.items) ? result.items : [];
  return result;
}

function geminiIdentifyProduct_(imageDataUrl) {
  if (!imageDataUrl) throw new Error('商品画像がありません。');
  const schema = { type:'OBJECT', properties:{ barcode:{type:'STRING'}, productName:{type:'STRING'}, category:{type:'STRING'}, confidence:{type:'NUMBER'} } };
  const prompt = [
    '商品パッケージ画像を解析してください。',
    '1次元バーコードの数字が画像から明確に読める場合だけbarcodeに入れてください。推測で数字を作らないでください。',
    'productNameは商品ラベルから特定し、categoryも短く日本語で返してください。',
    'バーコードが不鮮明ならbarcodeは空文字にしてください。'
  ].join('\n');
  return callGeminiJson_(imageDataUrl, prompt, schema);
}

function callGeminiJson_(imageDataUrl, prompt, schema) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEYがScript Propertiesに設定されていません。');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  const image = parseDataUrl_(imageDataUrl);
  const body = {
    contents:[{role:'user',parts:[{text:prompt},{inlineData:{mimeType:image.mimeType,data:image.base64}}]}],
    generationConfig:{responseMimeType:'application/json',responseSchema:schema,temperature:0.1}
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify(body),muteHttpExceptions:true});
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`Gemini APIエラー ${code}: ${txt.slice(0,400)}`);
  const obj = JSON.parse(txt);
  const out = obj && obj.candidates && obj.candidates[0] && obj.candidates[0].content && obj.candidates[0].content.parts && obj.candidates[0].content.parts[0] && obj.candidates[0].content.parts[0].text;
  if (!out) throw new Error('Geminiの解析結果が空です。');
  try { return JSON.parse(out); } catch (_) { throw new Error('GeminiのJSON解析に失敗しました: ' + out.slice(0,250)); }
}

function geminiAnalysis_(p) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEYがScript Propertiesに設定されていません。');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  const prompt = `あなたはBAKERUの商品・在庫管理アナリストです。次のデータを日本語で分析してください。\n\n必ず、(1)発注優先商品 (2)在庫過多・滞留リスク (3)最近の入出庫傾向 (4)在庫差異を減らす運用提案 の順に、具体的かつ簡潔に書いてください。データにないことは断定しないでください。\n\n${JSON.stringify(p || {})}`;
  const body = { contents:[{role:'user',parts:[{text:prompt}]}], generationConfig:{temperature:0.2} };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify(body),muteHttpExceptions:true});
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error('Gemini分析に失敗しました: ' + res.getContentText().slice(0,400));
  const obj = JSON.parse(res.getContentText());
  return obj.candidates[0].content.parts[0].text || '';
}

function commitReceipt_(p) {
  if (!p || !Array.isArray(p.items) || !p.items.length) throw new Error('領収書明細がありません。');
  const receiptId = String(p.receiptId || 'receipt_'+Utilities.getUuid());
  const inspected = p.inspectedCounts || {};
  const skipped = !!p.inspectionSkipped;
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.receipts);
  const now = new Date();

  const expectedBy = {};
  p.items.forEach(it => {
    if (!it.productId) throw new Error(`商品マスタ未紐付け: ${it.name || '名称不明'}`);
    expectedBy[it.productId] = (expectedBy[it.productId] || 0) + Math.max(0, num_(it.quantity) || 1);
  });

  const actualBy = {};
  Object.keys(expectedBy).forEach(productId => {
    actualBy[productId] = skipped ? expectedBy[productId] : Math.max(0, num_(inspected[productId]));
  });

  p.items.forEach((it, idx) => {
    const product = findProductRow_(it.productId);
    if (!product) throw new Error(`商品IDが見つかりません: ${it.productId}`);
    const expected = Math.max(0, num_(it.quantity) || 1);
    // 同じ商品が複数行ある場合の行別検品数量は参考値。差異判定は商品合算で行う。
    const totalExpected = expectedBy[it.productId];
    const totalActual = actualBy[it.productId];
    const rowActual = totalExpected > 0 ? Math.round(totalActual * expected / totalExpected * 1000) / 1000 : 0;
    const diff = rowActual - expected;
    sh.appendRow([
      receiptId, now, String(p.vendor || ''), String(p.receiptDate || ''), num_(p.total), idx+1, String(it.name || ''), String(product['商品ID']), String(product['SKU'] || ''),
      String(product['バーコード'] || ''), String(product['商品名'] || ''), expected, rowActual, diff, num_(it.unitPrice !== undefined ? it.unitPrice : it.price), num_(it.amount || (num_(it.unitPrice !== undefined ? it.unitPrice : it.price) * expected)),
      skipped ? '検品省略' : (diff === 0 ? '一致' : '差異あり'), ''
    ]);
  });

  Object.keys(expectedBy).forEach(productId => {
    const product = findProductRow_(productId);
    const expected = expectedBy[productId];
    const actual = actualBy[productId];
    const diff = actual - expected;
    if (!skipped && diff !== 0) {
      addAlert_({ type:'入荷差異', severity: diff < 0 ? '重要' : '要確認', productId, productName:String(product['商品名']), expected, actual, difference:diff, receiptId, detail:`領収書数量 ${expected} に対し実物検品 ${actual}。差異 ${diff > 0 ? '+' : ''}${diff}。` });
    }
    if (actual > 0) {
      recordMovementInternal_({ id:'receipt_tx_'+Utilities.getUuid(), productId, type:'in', quantity:actual, unitPrice:num_(product['仕入単価']), source:'領収書OCR・検品', receiptId, note: skipped ? 'バーコード検品省略' : 'バーコード検品済み' }, false);
    }
  });
  rebuildSummary_();
  return bootstrap_();
}

function saveStocktake_(p) {
  if (!p || !Array.isArray(p.items) || !p.items.length) throw new Error('棚卸データがありません。');
  const countId = String(p.countId || 'count_'+Utilities.getUuid());
  const sh = getSs_().getSheetByName(BAKERU_SHEETS.stocktake);
  p.items.forEach(it => {
    const product = findProductRow_(it.productId);
    if (!product) return;
    const theoretical = num_(product['現在庫']);
    const actual = Math.max(0, num_(it.actual));
    const diff = actual - theoretical;
    const cost = num_(product['仕入単価']);
    sh.appendRow([countId,new Date(),String(product['商品ID']),String(product['SKU']||''),String(product['バーコード']||''),String(product['商品名']||''),theoretical,actual,diff,cost,diff*cost,String(it.reason||''),String(it.action||''),String(p.operator||'')]);
    if (diff !== 0) {
      addAlert_({ type:'棚卸差異', severity:Math.abs(diff) >= 5 ? '重要' : '要確認', productId:String(product['商品ID']), productName:String(product['商品名']), expected:theoretical, actual, difference:diff, detail:`棚卸: 帳簿 ${theoretical} / 実地 ${actual} / 差異 ${diff > 0 ? '+' : ''}${diff}` });
      recordMovementInternal_({ id:'stocktake_'+Utilities.getUuid(), productId:String(product['商品ID']), type:'adjustment', quantity:Math.abs(diff), signedQuantity:diff, unitPrice:cost, source:'棚卸', note:'棚卸差異を実地在庫へ調整' }, false);
    }
  });
  rebuildSummary_();
  return { inventory:getProducts_(), alerts:getAlerts_(300) };
}

function rebuildSummary_() {
  const ss = getSs_();
  const sh = ss.getSheetByName(BAKERU_SHEETS.summary) || ensureSheet_(ss, BAKERU_SHEETS.summary, SCHEMAS[BAKERU_SHEETS.summary]);
  const products = getProducts_();
  sh.clearContents();
  sh.getRange(1,1,1,SCHEMAS[BAKERU_SHEETS.summary].length).setValues([SCHEMAS[BAKERU_SHEETS.summary]]);
  if (products.length) {
    const rows = products.map(p => [p.id,p.sku,p.barcode,p.name,p.category,p.stock,p.minStock,p.stock<=p.minStock?'🔴要発注':p.stock<=p.warningStock?'🟡注意':'適正',p.cost,p.stock*p.cost,p.price,new Date(p.updatedAt)]);
    sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,SCHEMAS[BAKERU_SHEETS.summary].length).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
}

function migrateLegacyStocktake_() {
  const ss = getSs_();
  const sh = ss.getSheetByName('4_棚卸表');
  if (!sh) throw new Error('「4_棚卸表」が見つかりません。');
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return { imported:0, skipped:0 };
  const h = values[0].map(String);
  const idx = name => h.indexOf(name);
  const nameI = idx('品名'), skuI = idx('SKU'), stockI = idx('帳簿在庫'), costI = idx('単価');
  if (nameI < 0) throw new Error('旧棚卸表に「品名」列がありません。');
  let imported=0, skipped=0;
  const existing = getProducts_();
  for (let i=1;i<values.length;i++) {
    const name = String(values[i][nameI] || '').trim();
    if (!name) { skipped++; continue; }
    const sku = skuI >= 0 ? String(values[i][skuI] || '').trim() : '';
    const duplicate = existing.find(p => (sku && p.sku === sku) || normalize_(p.name) === normalize_(name));
    if (duplicate) { skipped++; continue; }
    const stock = stockI >= 0 ? parseMoney_(values[i][stockI]) : 0;
    const cost = costI >= 0 ? parseMoney_(values[i][costI]) : 0;
    upsertProduct_({ sku, name, stock, cost, price:0, minStock:5, unit:'個', barcode:'', category:'', aliases:'' });
    imported++;
  }
  rebuildSummary_();
  return { imported, skipped };
}

function nekkoExport_() {
  const inventory = getProducts_();
  const alerts = getAlerts_(300).filter(a => a.status !== '解決済み');
  const tx = getMovements_(500);
  return {
    source:'BAKERU', version:BAKERU_VERSION, exportedAt:new Date().toISOString(),
    summary:{ productCount:inventory.length, totalUnits:inventory.reduce((s,x)=>s+num_(x.stock),0), stockCost:inventory.reduce((s,x)=>s+num_(x.stock)*num_(x.cost),0), activeAlertCount:alerts.length },
    products:inventory, alerts:alerts, recentMovements:tx
  };
}

function findProductRow_(id) {
  return rowsAsObjects_(BAKERU_SHEETS.products).find(r => String(r['商品ID']) === String(id) && String(r['有効']).toUpperCase() !== 'FALSE');
}

function nextSku_(rows) {
  let max = 0;
  rows.forEach(r => {
    const m = String(r['SKU'] || '').match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'SKU-' + String(max + 1).padStart(4,'0');
}

function parseDataUrl_(s) {
  const m = String(s).match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('画像データ形式が不正です。');
  return { mimeType:m[1], base64:m[2] };
}

function parseMoney_(v) {
  return num_(String(v || '').replace(/[¥￥,\s]/g,''));
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function dateMs_(v) {
  if (v instanceof Date && !isNaN(v)) return v.getTime();
  const t = new Date(v).getTime();
  return isFinite(t) ? t : 0;
}

function normalize_(s) {
  return String(s || '').normalize('NFKC').toLowerCase().replace(/[\s　()（）・\-_/]/g,'');
}

function summarizePayload_(p) {
  try {
    const s = JSON.stringify(p || {});
    return s.length > 500 ? s.slice(0,500) + '…' : s;
  } catch (_) { return ''; }
}

function logSync_(action, result, detail) {
  const ss = getSs_();
  let sh = ss.getSheetByName(BAKERU_SHEETS.logs);
  if (!sh) sh = ensureSheet_(ss, BAKERU_SHEETS.logs, SCHEMAS[BAKERU_SHEETS.logs]);
  sh.appendRow([new Date(), String(action || ''), String(result || ''), String(detail || '')]);
}


// ===== STEP2: テストモード専用の最小API =====
// 本番シートには一切書き込まず、T01/T02/T05/T06/T07 のみを使用する。
function testBootstrap_() {
  return {
    inventory: getProductsFrom_(BAKERU_TEST_SHEETS.products),
    transactions: getMovementsFrom_(BAKERU_TEST_SHEETS.movements, 500),
    alerts: getAlertsFrom_(BAKERU_TEST_SHEETS.alerts, 300)
  };
}

function getProductsFrom_(sheetName) {
  return rowsAsObjects_(sheetName)
    .filter(r => String(r['有効']).toUpperCase() !== 'FALSE' && r['商品ID'])
    .map(r => ({
      id: String(r['商品ID']), sku: String(r['SKU'] || ''), barcode: String(r['バーコード'] || ''),
      name: String(r['商品名'] || ''), category: String(r['カテゴリ'] || ''), unit: String(r['単位'] || '個'),
      cost: num_(r['仕入単価']), price: num_(r['販売単価']), minStock: num_(r['在庫下限']),
      stock: num_(r['現在庫']), active: true, aliases: String(r['別名'] || ''),
      purchaseUnit: String(r['仕入単位'] || r['単位'] || '個'), purchaseFactor: Math.max(0.000001, num_(r['換算量']) || 1),
      warningStock: num_(r['注意在庫']) > num_(r['在庫下限']) ? num_(r['注意在庫']) : Math.max(num_(r['在庫下限']) * 2, num_(r['在庫下限']) + 1),
      notifyEmail: String(r['通知Gmail'] || ''), contentAmount: num_(r['内容量']), contentUnit: String(r['内容量単位'] || 'g'),
      details: String(r['商品特徴'] || ''), shelf: String(r['棚・保管場所'] || ''),
      updatedAt: dateMs_(r['更新日時']) || Date.now()
    }));
}

function getMovementsFrom_(sheetName, limit) {
  const rows = rowsAsObjects_(sheetName);
  return rows.slice(-Math.max(1, limit || 500)).reverse().map(r => ({
    id: String(r['取引ID'] || ''), timestamp: dateMs_(r['日時']) || Date.now(),
    type: String(r['種別'] || '') === '販売' ? 'sale' : String(r['種別'] || '') === '使用' ? 'use' : num_(r['在庫増減']) < 0 ? 'out' : 'in', productId: String(r['商品ID'] || ''),
    productName: String(r['商品名'] || ''), barcode: String(r['バーコード'] || ''),
    quantity: Math.abs(num_(r['数量'])), unitPrice: num_(r['単価']), source: String(r['入力元'] || ''), note: String(r['備考'] || '')
  }));
}

function getAlertsFrom_(sheetName, limit) {
  const rows = rowsAsObjects_(sheetName);
  return rows.slice(-Math.max(1, limit || 300)).reverse().map(r => ({
    id: String(r['アラートID'] || ''), alertId: String(r['アラートID'] || ''), timestamp: dateMs_(r['発生日時']) || Date.now(),
    type: String(r['種類'] || ''), severity: String(r['重要度'] || ''), productId: String(r['商品ID'] || ''),
    productName: String(r['商品名'] || ''), expected: r['期待値'], actual: r['実績値'], difference: num_(r['差異']),
    receiptId: String(r['領収書ID'] || ''), status: String(r['状態'] || '未確認'), detail: String(r['詳細'] || ''), message: String(r['詳細'] || '')
  }));
}

function testUpsertProduct_(p) {
  if (!p || !String(p.name || '').trim()) throw new Error('商品名は必須です。');
  const sh = getSs_().getSheetByName(BAKERU_TEST_SHEETS.products);
  if (!sh) throw new Error('T01_商品マスタがありません。setupBAKERUTestMode()を先に実行してください。');
  const rows = rowsAsObjects_(BAKERU_TEST_SHEETS.products);
  const id = String(p.id || ('TEST_' + Utilities.getUuid()));
  const existing = rows.find(r => String(r['商品ID']) === id);
  const barcode = String(p.barcode || '').trim();
  if (barcode) {
    const dup = rows.find(r => String(r['バーコード'] || '').trim() === barcode && String(r['商品ID']) !== id && String(r['有効']).toUpperCase() !== 'FALSE');
    if (dup) throw new Error(`このバーコードはテスト商品「${dup['商品名']}」に登録済みです。`);
  }
  const sku = String(p.sku || (existing && existing['SKU']) || ('T-' + nextSku_(rows)));
  const initialStock = existing ? num_(existing['現在庫']) : Math.max(0, num_(p.stock));
  const row = [id, sku, barcode, String(p.name).trim(), String(p.category || ''), String(p.unit || '個'), num_(p.cost), num_(p.price), Math.max(0,num_(p.minStock)), existing ? initialStock : 0, true, String(p.aliases || ''), new Date(), String(p.purchaseUnit || p.unit || '個'), Math.max(0.000001, num_(p.purchaseFactor) || 1), Math.max(Math.max(0,num_(p.minStock)) + 0.001, num_(p.warningStock) || Math.max(num_(p.minStock) * 2, num_(p.minStock) + 1)), String(p.notifyEmail || ''), Math.max(0,num_(p.contentAmount)), String(p.contentUnit || 'g'), String(p.details || ''), String(p.shelf || '')];
  if (existing) sh.getRange(existing._row,1,1,row.length).setValues([row]);
  else sh.appendRow(row);

  if (!existing && initialStock > 0) {
    testRecordMovementInternal_({ id:'test_init_'+Utilities.getUuid(), productId:id, type:'in', quantity:initialStock, unitPrice:num_(p.cost), source:'テスト初期登録', note:'テスト商品登録時の初期在庫' });
  }
  rebuildTestSummary_();
  return getProductsFrom_(BAKERU_TEST_SHEETS.products).find(x => x.id === id);
}

function testRecordMovement_(p) {
  testRecordMovementInternal_(p);
  return testBootstrap_();
}

function testRecordMovementInternal_(p) {
  const product = rowsAsObjects_(BAKERU_TEST_SHEETS.products).find(r => String(r['商品ID']) === String(p.productId) && String(r['有効']).toUpperCase() !== 'FALSE');
  if (!product) throw new Error('テスト商品が見つかりません。');
  const qty = Math.abs(num_(p.quantity));
  if (qty <= 0) throw new Error('数量は0より大きい値で入力してください。');
  let signed, kind;
  if (p.type === 'sale' || p.type === '販売') { signed = -qty; kind = '販売'; }
  else if (p.type === 'use' || p.type === '使用') { signed = -qty; kind = '使用'; }
  else if (p.type === 'out' || p.type === '出庫') { signed = -qty; kind = '出庫'; }
  else { signed = qty; kind = '入庫'; }

  const before = num_(product['現在庫']);
  const after = before + signed;
  if (after < 0) throw new Error(`${product['商品名']}のテスト在庫が不足しています（現在庫 ${before}）。`);

  const defaultUnitPrice = kind === '販売' ? num_(product['販売単価']) : num_(product['仕入単価']);
  const unitPrice = (p.unitPrice !== undefined && p.unitPrice !== null && num_(p.unitPrice) > 0) ? num_(p.unitPrice) : defaultUnitPrice;
  const sh = getSs_().getSheetByName(BAKERU_TEST_SHEETS.movements);
  sh.appendRow([
    String(p.id || 'test_tx_'+Utilities.getUuid()), new Date(p.timestamp || Date.now()), kind, String(product['商品ID']), String(product['SKU'] || ''),
    String(product['バーコード'] || ''), String(product['商品名'] || ''), qty, signed, unitPrice, Math.abs(signed) * unitPrice,
    String(p.source || 'テスト手入力'), String(p.receiptId || ''), String(p.note || ''), String(p.operator || '')
  ]);
  const productSh = getSs_().getSheetByName(BAKERU_TEST_SHEETS.products);
  productSh.getRange(product._row, 10).setValue(after);
  productSh.getRange(product._row, 13).setValue(new Date());
  syncTestLowStockAlerts_(product, after);
  rebuildTestSummary_();
  return after;
}

function syncTestLowStockAlerts_(product, stock) {
  const critical = Math.max(0, num_(product['在庫下限']));
  const warningRaw = num_(product['注意在庫']);
  const warning = warningRaw > critical ? warningRaw : Math.max(critical * 2, critical + 1);
  const productId = String(product['商品ID']);
  const name = String(product['商品名']);
  const rows = rowsAsObjects_(BAKERU_TEST_SHEETS.alerts);
  const sh = getSs_().getSheetByName(BAKERU_TEST_SHEETS.alerts);
  const active = rows.filter(r => String(r['商品ID']) === productId && ['在庫注意','在庫レッド','在庫不足'].includes(String(r['種類'])) && String(r['状態']) !== '解決済み');
  const resolveTypes = (types) => active.filter(r => types.includes(String(r['種類']))).forEach(r => sh.getRange(r._row, 11).setValue('解決済み'));
  const hasType = (type) => active.some(r => String(r['種類']) === type);
  const add = (a) => sh.appendRow([
    String(a.id || 'test_alert_'+Utilities.getUuid()), new Date(), String(a.type), String(a.severity), productId, name,
    a.expected, a.actual, a.difference, '', '未確認', String(a.detail || '')
  ]);
  if (stock <= critical) {
    resolveTypes(['在庫注意','在庫不足']);
    if (!hasType('在庫レッド')) add({type:'在庫レッド',severity:'重要',expected:critical,actual:stock,difference:stock-critical,detail:`🔴 現在庫 ${stock} がレッドライン ${critical} 以下です。早急に発注・補充を確認してください。`});
    return;
  }
  if (stock <= warning) {
    resolveTypes(['在庫レッド','在庫不足']);
    if (!hasType('在庫注意')) add({type:'在庫注意',severity:'注意',expected:warning,actual:stock,difference:stock-warning,detail:`🟡 現在庫 ${stock} が注意ライン ${warning} 以下です。発注準備を確認してください。`});
    return;
  }
  resolveTypes(['在庫注意','在庫レッド','在庫不足']);
}

function rebuildTestSummary_() {
  const ss = getSs_();
  const headers = TEST_SCHEMAS[BAKERU_TEST_SHEETS.summary];
  const sh = ss.getSheetByName(BAKERU_TEST_SHEETS.summary) || ensureSheet_(ss, BAKERU_TEST_SHEETS.summary, headers);
  const products = getProductsFrom_(BAKERU_TEST_SHEETS.products);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (products.length) {
    const rows = products.map(p => [p.id,p.sku,p.barcode,p.name,p.category,p.stock,p.minStock,p.stock<=p.minStock?'🔴要発注':p.stock<=p.warningStock?'🟡注意':'適正',p.cost,p.stock*p.cost,p.price,new Date(p.updatedAt)]);
    sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#7c2d12').setFontColor('#ffffff');
}

function logTestSync_(action, result, detail) {
  const ss = getSs_();
  let sh = ss.getSheetByName(BAKERU_TEST_SHEETS.logs);
  if (!sh) sh = ensureSheet_(ss, BAKERU_TEST_SHEETS.logs, TEST_SCHEMAS[BAKERU_TEST_SHEETS.logs]);
  sh.appendRow([new Date(), String(action || ''), String(result || ''), String(detail || '')]);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
