/**
 * CHAEN POS → Google Sheets 連携 受け口スクリプト v10
 * 設置場所: 管理スプレッドシートの [拡張機能] > [Apps Script]
 *
 * v2: POSでの注文「編集・削除」をシートに反映(R列のPOS Order IDで照合)。
 * v3: 経費の「削除」に対応。 v4: 経費の「編集」に対応(削除→再追記方式)。
 * v5: doGet に action=today を追加。同じSTOREの「本日分の注文」を
 *     orderIdごとにまとめて返す(複数端末間で当日売上を突き合わせるため)。
 * v6: 同じレジをシフト交代で共有する運用に対応。
 *     - 注文に「時刻」(S列)、経費に「時刻」(O列。J〜N列は経費サマリー表が
 *       使っているため空いているO列を使用)を追加記録。
 *     - doGet(action=today) が経費(Expense Log)も一緒に返すよう拡張。
 * v7: 注文にスタッフ名(T列)を追加記録。どのスタッフが対応した注文かを
 *     シート側でも把握できるようにした。
 * v8: 経費にもスタッフ名(P列)を追加記録。どのスタッフが記録した経費かを
 *     把握できるようにした。
 * v10: 在庫管理に対応。
 *      - doGet に action=expensepresets（経費プリセット配信）と
 *        action=inventory（残在庫配信）を追加。
 *      - 経費に 材料ID / 入庫数量 / 個数 / 仕入単位 を記録（Q〜T列）。
 *        これで Inventory シートの「入庫」が自動で積み上がる。
 *      - 未登録の品目をPOSから登録すると Expense Preset に
 *        下書き行（有効=FALSE / 要確認）が自動で追加される。
 *      - doPost の売上まわりには手を入れていない。
 * v9: doGet に action=products を追加。Recipe Master の先頭行(商品名がある行)だけを
 *     読んで商品リストをJSONで返す。POSがこれを取得してボタンを生成する。
 *     doPost(売上・経費の書き込み)には一切手を入れていない。
 *Expense LogのI列にPOS Expense IDを記録し、
 *  削除時は該当行のA〜I列をクリアする。
 *  ※行ごと削除しない理由: 同じ行の右側(J〜N列)に経費サマリー表が
 *    あるため、deleteRowを使うとサマリーが崩れる。
 */

const TOKEN = 'CHAENPOS-2026';   // ← POS側(sheet-sync.js)と同じ値
const SHEET_SALES   = 'Daily Product Sales';
const SHEET_EXPENSE = 'Expense Log';
const COL_ORDER_ID    = 18; // Daily Product Sales R列
const COL_ORDER_TIME  = 19; // Daily Product Sales S列(v6で追加)
const COL_ORDER_STAFF = 20; // Daily Product Sales T列(v7で追加。対応したスタッフ名)
const COL_EXPENSE_ID   = 9;  // Expense Log I列
const COL_EXPENSE_TIME  = 15; // Expense Log O列(v6で追加。J〜N列は経費サマリー表のため使用不可)
const COL_EXPENSE_STAFF = 16; // Expense Log P列(v8で追加。記録したスタッフ名)
const SHEET_RECIPE     = 'Recipe Master';
const RECIPE_FIRST_ROW = 5;  // Recipe Master のデータ開始行(1〜4はヘッダー)
const SHEET_PRESET     = 'Expense Preset';
const PRESET_FIRST_ROW = 5;
const SHEET_INV        = 'Inventory';
const INV_FIRST_ROW    = 6;
// Expense Log の追加列(J〜Nは経費サマリー表、O=時刻、P=スタッフで使用済み)
const COL_EXP_INGID = 17; // Q 材料ID
const COL_EXP_QTY   = 18; // R 入庫数量(基本単位)
const COL_EXP_PACKS = 19; // S 個数
const COL_EXP_UNIT  = 20; // T 仕入単位

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return jsonOut({ ok: false, error: 'auth' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // 複数端末の同時送信対策
    try {
      switch (data.type) {
        case 'order_add':    writeOrder(ss, data); break;
        case 'order_edit':   deleteOrderRows(ss, data.orderId); writeOrder(ss, data); break;
        case 'order_delete': deleteOrderRows(ss, data.orderId); break;
        case 'expense':        writeExpense(ss, data); break;
        case 'expense_edit':   deleteExpenseRows(ss, data.expenseId); writeExpense(ss, data); break;
        case 'expense_delete': deleteExpenseRows(ss, data.expenseId); break;
        default: return jsonOut({ ok: false, error: 'unknown type' });
      }
    } finally {
      lock.releaseLock();
    }
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/**
 * GET:
 *  - パラメータなし → 動作確認用の生存確認
 *  - action=today&token=...&store=... → 同じ店舗の「本日分」の注文と経費を
 *    返す(シフト交代で端末が変わっても、その日のデータを引き継ぐため)
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'today') {
    if (p.token !== TOKEN) return jsonOut({ ok: false, error: 'auth' });
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const tz = ss.getSpreadsheetTimeZone();
      const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

      // 注文(商品1行ずつ→orderIdでグループ化)
      const shSales = ss.getSheetByName(SHEET_SALES);
      const lastSales = shSales.getLastRow();
      const byOrder = {};
      if (lastSales >= 6) {
        const rows = shSales.getRange(6, 1, lastSales - 5, COL_ORDER_STAFF).getValues();
        rows.forEach(function (r) {
          const rowDate = normDate(r[0], tz);
          if (rowDate !== today) return;
          const store = r[1], recipeId = r[2], qty = r[3], promo = r[4], pay = r[5];
          const orderId = r[COL_ORDER_ID - 1];
          const time = r[COL_ORDER_TIME - 1];
          const staff = r[COL_ORDER_STAFF - 1];
          if (!orderId) return;
          if (p.store && String(store) !== String(p.store)) return;
          if (!byOrder[orderId]) {
            byOrder[orderId] = { orderId: orderId, date: rowDate, time: time || '', staff: staff || '', store: store, pay: pay, items: [] };
          }
          byOrder[orderId].items.push({ recipeId: recipeId, qty: qty, promo: promo });
        });
      }

      // 経費(1行=1件)
      const shExp = ss.getSheetByName(SHEET_EXPENSE);
      const lastExp = shExp.getLastRow();
      const expenseList = [];
      if (lastExp >= 6) {
        const rows = shExp.getRange(6, 1, lastExp - 5, COL_EXPENSE_STAFF).getValues();
        rows.forEach(function (r) {
          const expenseId = r[COL_EXPENSE_ID - 1];
          if (!expenseId) return; // 削除済み(クリア済み)行は対象外
          const rowDate = normDate(r[0], tz);
          if (rowDate !== today) return;
          const store = r[1];
          if (p.store && String(store) !== String(p.store)) return;
          expenseList.push({
            expenseId: expenseId, date: rowDate, time: r[COL_EXPENSE_TIME - 1] || '',
            staff: r[COL_EXPENSE_STAFF - 1] || '',
            store: store, category: r[3], amount: r[4], pay: r[5], notes: r[6]
          });
        });
      }

      return jsonOut({
        ok: true,
        orders: Object.keys(byOrder).map(function (k) { return byOrder[k]; }),
        expenses: expenseList
      });
    } catch (err) {
      return jsonOut({ ok: false, error: String(err) });
    }
  }
  if (p.action === 'products') {
    // 商品マスターは Recipe Master が正。POSは起動時にここから取得してボタンを生成する。
    // tokenは任意(商品リストは機微情報ではない)。渡された場合だけ検証する。
    if (p.token && p.token !== TOKEN) return jsonOut({ ok: false, error: 'auth' });
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_RECIPE);
      const last = sh.getLastRow();
      const products = [];
      if (last >= RECIPE_FIRST_ROW) {
        const rows = sh.getRange(RECIPE_FIRST_ROW, 1, last - RECIPE_FIRST_ROW + 1, 5).getValues();
        rows.forEach(function (r) {
          const id       = String(r[0] == null ? '' : r[0]).trim();
          const name     = String(r[1] == null ? '' : r[1]).trim();
          const size     = String(r[2] == null ? '' : r[2]).trim();
          const category = String(r[3] == null ? '' : r[3]).trim();
          const price    = r[4];
          if (!id || !name) return;                         // 材料の継続行(商品名が空)はスキップ
          if (price === '' || price === null) return;        // 売価が空の行はスキップ
          const n = Number(price);
          if (!isFinite(n) || n <= 0) return;
          products.push({ id: id, name: name, size: size, category: category || 'Other', price: n });
        });
      }
      return jsonOut({ ok: true, updatedAt: new Date().toISOString(), products: products });
    } catch (err) {
      return jsonOut({ ok: false, error: String(err) });
    }
  }

  if (p.action === 'expensepresets') {
    // POSの EXPENSE 画面に出すボタンの一覧。有効=TRUE の行だけ返す。
    if (p.token && p.token !== TOKEN) return jsonOut({ ok: false, error: 'auth' });
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRESET);
      if (!sh) return jsonOut({ ok: true, updatedAt: new Date().toISOString(), presets: [] });
      const last = sh.getLastRow();
      const out = [];
      if (last >= PRESET_FIRST_ROW) {
        const rows = sh.getRange(PRESET_FIRST_ROW, 1, last - PRESET_FIRST_ROW + 1, 11).getValues();
        rows.forEach(function (r) {
          if (r[0] !== true && String(r[0]).toUpperCase() !== 'TRUE') return; // 有効のみ
          const name = String(r[3] == null ? '' : r[3]).trim();
          const cat  = String(r[2] == null ? '' : r[2]).trim();
          if (!name || !cat) return;
          out.push({
            order:    Number(r[1]) || 0,
            category: cat,
            name:     name,
            buyUnit:  String(r[4] == null ? '' : r[4]).trim(),
            qtyPerUnit: (r[5] === '' || r[5] == null) ? null : Number(r[5]),
            baseUnit: String(r[6] == null ? '' : r[6]).trim(),
            price:    Number(r[7]) || 0,
            pay:      String(r[8] == null ? '' : r[8]).trim() || 'CASH',
            ingId:    String(r[9] == null ? '' : r[9]).trim()
          });
        });
      }
      out.sort(function (a, b) { return a.order - b.order; });
      return jsonOut({ ok: true, updatedAt: new Date().toISOString(), presets: out });
    } catch (err) { return jsonOut({ ok: false, error: String(err) }); }
  }

  if (p.action === 'inventory') {
    // Inventory シートの計算結果をそのまま返す。POSは表示するだけ。
    if (p.token && p.token !== TOKEN) return jsonOut({ ok: false, error: 'auth' });
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INV);
      if (!sh) return jsonOut({ ok: true, updatedAt: new Date().toISOString(), items: [] });
      const last = sh.getLastRow();
      const out = [];
      if (last >= INV_FIRST_ROW) {
        const rows = sh.getRange(INV_FIRST_ROW, 1, last - INV_FIRST_ROW + 1, 13).getValues();
        rows.forEach(function (r) {
          const id = String(r[0] == null ? '' : r[0]).trim();
          const name = String(r[1] == null ? '' : r[1]).trim();
          if (!id || !name) return;
          out.push({
            id: id, name: name,
            category: String(r[2] == null ? '' : r[2]).trim(),
            unit:     String(r[3] == null ? '' : r[3]).trim(),
            stock:    Number(r[7]) || 0,
            perDay:   Number(r[8]) || 0,
            days:     (r[9] === '' || r[9] == null) ? null : Number(r[9]),
            reorder:  Number(r[10]) || 0,
            status:   String(r[11] == null ? '' : r[11]).trim()
          });
        });
      }
      return jsonOut({ ok: true, updatedAt: new Date().toISOString(), items: out });
    } catch (err) { return jsonOut({ ok: false, error: String(err) }); }
  }

  return jsonOut({ ok: true, message: 'CHAEN POS endpoint alive (v10)' });
}

/** シートの日付セル(Date型 or 文字列)を 'yyyy-MM-dd' に統一する */
function normDate(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v);
}

/** 注文 → Daily Product Sales(商品1つにつき1行、R列に注文ID、S列に時刻、T列に担当スタッフ) */
function writeOrder(ss, d) {
  const sh = ss.getSheetByName(SHEET_SALES);
  let row = nextEmptyRow(sh, 6);
  d.items.forEach(function (it) {
    // A〜F: 日付 / 店舗ID / 商品ID / 数量 / 割引枚数 / 支払方法
    sh.getRange(row, 1, 1, 6).setValues([[
      d.date, d.store, it.recipeId, it.qty, it.promo || '', d.pay
    ]]);
    // G〜Q の数式(単価・売上・原価・粗利など)を直前の行からコピー
    sh.getRange(row - 1, 7, 1, 11).copyTo(sh.getRange(row, 7, 1, 11));
    // R: POSの注文ID(編集・削除の照合キー)
    sh.getRange(row, COL_ORDER_ID).setValue(d.orderId || '');
    // S: 注文時刻(v6)
    sh.getRange(row, COL_ORDER_TIME).setValue(d.time || '');
    // T: 対応スタッフ名(v7)
    sh.getRange(row, COL_ORDER_STAFF).setValue(d.staff || '');
    row++;
  });
}

/** R列の注文IDが一致する行をすべて削除(下から順に) */
function deleteOrderRows(ss, orderId) {
  if (!orderId) return;
  const sh = ss.getSheetByName(SHEET_SALES);
  const last = sh.getLastRow();
  if (last < 6) return;
  const ids = sh.getRange(6, COL_ORDER_ID, last - 5, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === orderId) sh.deleteRow(6 + i);
  }
}

/** 経費 → Expense Log */
function writeExpense(ss, d) {
  const sh = ss.getSheetByName(SHEET_EXPENSE);
  const row = nextEmptyRow(sh, 6);
  // A〜G: 日付 / 店舗ID / カテゴリー / 品目 / 金額 / 支払元 / 備考
  sh.getRange(row, 1, 1, 7).setValues([[
    d.date, d.store, d.category, d.item, d.amount, d.pay, d.notes || ''
  ]]);
  // H列 月(自動) の数式をコピー
  sh.getRange(row - 1, 8, 1, 1).copyTo(sh.getRange(row, 8, 1, 1));
  // I: POSの経費ID(削除の照合キー)
  sh.getRange(row, COL_EXPENSE_ID).setValue(d.expenseId || '');
  // O: 経費の時刻(v6。J〜N列は経費サマリー表のため使用不可)
  sh.getRange(row, COL_EXPENSE_TIME).setValue(d.time || '');
  // P: 記録したスタッフ名(v8)
  sh.getRange(row, COL_EXPENSE_STAFF).setValue(d.staff || '');
  // Q〜T: 在庫連動用(v10)。材料IDが無い経費(人件費など)は空のまま。
  sh.getRange(row, COL_EXP_INGID).setValue(d.ingId || '');
  sh.getRange(row, COL_EXP_QTY).setValue(d.ingId ? (Number(d.qty) || '') : '');
  sh.getRange(row, COL_EXP_PACKS).setValue(d.packs === undefined ? '' : d.packs);
  sh.getRange(row, COL_EXP_UNIT).setValue(d.buyUnit || '');

  // 未登録の品目なら Expense Preset に下書きを積む(v10)
  if (d.draft) addPresetDraft_(ss, d);
}

/** POSから登録された未登録品目を、Expense Preset に「要確認」の下書きとして追加する。
 *  すでに同じカテゴリー・同じ表示名の行があれば何もしない。 */
function addPresetDraft_(ss, d) {
  try {
    const sh = ss.getSheetByName(SHEET_PRESET);
    if (!sh) return;
    const name = String(d.item || '').trim();
    if (!name) return;
    const last = sh.getLastRow();
    let maxOrder = 0;
    if (last >= PRESET_FIRST_ROW) {
      const rows = sh.getRange(PRESET_FIRST_ROW, 1, last - PRESET_FIRST_ROW + 1, 11).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][3]).trim() === name && String(rows[i][2]).trim() === String(d.category).trim()) return;
        maxOrder = Math.max(maxOrder, Number(rows[i][1]) || 0);
      }
    }
    sh.getRange(last + 1, 1, 1, 11).setValues([[
      false, maxOrder + 10, d.category || '', name, d.buyUnit || '',
      '', '', Number(d.unitPrice) || Number(d.amount) || 0, d.pay || 'CASH', '',
      'POSから登録 要確認 (' + (d.date || '') + ')'
    ]]);
    sh.getRange(last + 1, 1, 1, 11).setBackground('#ffe9e0');
  } catch (e) { /* 下書きに失敗しても経費の記録は成立させる */ }
}

/** I列の経費IDが一致する行のA〜I列・O列(時刻)・P列(スタッフ)をクリア(J〜N列のサマリー表は残す) */
function deleteExpenseRows(ss, expenseId) {
  if (!expenseId) return;
  const sh = ss.getSheetByName(SHEET_EXPENSE);
  const last = sh.getLastRow();
  if (last < 6) return;
  const ids = sh.getRange(6, COL_EXPENSE_ID, last - 5, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === expenseId) {
      sh.getRange(6 + i, 1, 1, COL_EXPENSE_ID).clearContent();
      sh.getRange(6 + i, COL_EXPENSE_TIME).clearContent();
      sh.getRange(6 + i, COL_EXPENSE_STAFF).clearContent();
    }
  }
}

/** A列を基準に次の空き行を返す */
function nextEmptyRow(sh, startRow) {
  const vals = sh.getRange(startRow, 1, sh.getMaxRows() - startRow + 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] !== null) return startRow + i + 1;
  }
  return startRow;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 【初回に1回だけ手動実行】セットアップ
 *  1) 集計用の名前付き範囲を5005行目まで拡張
 *     (xlsx由来の範囲は売上606行・経費407行で止まっており、
 *      その先に書かれたデータが月次集計から漏れるため)
 *  2) Daily Product Sales の R5 に「POS Order ID」ヘッダーを追加
 * エディタ上でこの関数を選んで ▶ 実行。
 */
function setupOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const NEW_LAST = 5005;
  ss.getNamedRanges().forEach(function (nr) {
    const r = nr.getRange();
    const sheetName = r.getSheet().getName();
    let hit = false;
    if (sheetName === SHEET_SALES && r.getLastRow() === 606) hit = true;
    if (sheetName === SHEET_EXPENSE && r.getLastRow() === 407) hit = true;
    if (hit) {
      const sh = r.getSheet();
      nr.setRange(sh.getRange(r.getRow(), r.getColumn(),
                              NEW_LAST - r.getRow() + 1, r.getNumColumns()));
      Logger.log('extended: ' + nr.getName());
    }
  });
  const sales = ss.getSheetByName(SHEET_SALES);
  sales.getRange(5, COL_ORDER_ID).setValue('POS Order ID\n(自動・編集不可)');
  const exp = ss.getSheetByName(SHEET_EXPENSE);
  exp.getRange(5, COL_EXPENSE_ID).setValue('POS Expense ID\n(自動・編集不可)');
  Logger.log('setup done');
}
