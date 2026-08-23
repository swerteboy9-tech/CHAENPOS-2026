/* CHAEN POS → Google Sheets 連携モジュール (v1.18)
   デプロイ後に URL / TOKEN を設定。STOREは店舗が増えたら端末ごとに変更。 */
const SheetSync = (() => {
  const CONFIG = {
    URL: 'https://script.google.com/macros/s/AKfycbzm3znIf4AtO3u3tG7EZQ634M_7tHNGk3O8rWXNSLc_bziA0VTLIJKmKjEZrLOIgRV80Q/exec', // ← Apps ScriptのウェブアプリURL
    TOKEN: 'CHAENPOS-2026',                            // ← Code.gsのTOKENと同じ値
    STORE: 'S001',                                            // ← Store MasterのStore ID
    ENABLED: true,
  };

  /* POSの 商品名|サイズ → 管理シートのRecipe ID 対応表
     商品を追加したら Recipe Master と両方に登録すること */
  const RECIPE_IDS = {
    'Signature Matcha Latte|12oz': 'SIG12',
    'Matcha Cloud Latte|16oz': 'CLO16',
    'Matcha Cloud Latte|22oz': 'CLO22',
    'Einspanner Matcha Latte|16oz': 'EIN16',
    'Einspanner Matcha Latte|22oz': 'EIN22',
    'Matcha Mango|16oz': 'MAN16',
    'Matcha Mango|22oz': 'MAN22',
    'Matcha Strawberry|16oz': 'STR16',
    'Matcha Strawberry|22oz': 'STR22',
    'Matcha Oreo|16oz': 'ORE16',
    'Matcha Oreo|22oz': 'ORE22',
    'Matcha Sea Salt|16oz': 'SEA16',
    'Matcha Sea Salt|22oz': 'SEA22',
    'Nutella Matcha Latte|16oz': 'NUT16',
    'Nutella Matcha Latte|22oz': 'NUT22',
    'Matcha Frappe|16oz': 'FRA16',
    'Matcha Frappe|22oz': 'FRA22',
    'Churros|5 sticks': 'CHU5',
  };

  /* POSの経費カテゴリー → Expense Logのカテゴリー(プルダウンと完全一致させること) */
  const EXPENSE_CATS = {
    'Ice': 'Ice / 氷',
    'Water': 'Ingredients / 材料',
    'Breakfast / Meal': 'Other / その他',
    'Tissue / Supplies': 'Packaging / 資材',
    'Staff Advance': 'Labor / 人件費',
    'Ingredient Purchase': 'Ingredients / 材料',
    'Delivery': 'Transport / 交通',
    'Other': 'Other / その他',
  };

  const QUEUE_KEY = 'chaen_sheet_queue';

  function itemsToRows(items) {
    return items.map(it => ({
      recipeId: RECIPE_IDS[`${it.name}|${it.size}`] || `??${it.name}`,
      qty: it.qty,
      promo: it.discount ? it.qty : 0, // -₱30割引を使った杯数
    }));
  }

  function post(payload) {
    if (!CONFIG.ENABLED) return;
    payload.token = CONFIG.TOKEN;
    // Content-Typeヘッダは付けない(text/plain扱い→CORSプリフライト回避)
    fetch(CONFIG.URL, { method: 'POST', body: JSON.stringify(payload) })
      .then(r => r.json())
      .then(j => { if (!j.ok) throw new Error(j.error || 'sheet error'); flushQueue(); })
      .catch(err => { enqueue(payload); console.warn('SheetSync queued:', err); });
  }

  function enqueue(payload) {
    const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
    q.push(payload);
    localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
  }

  async function flushQueue() {
    const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
    if (!q.length) return;
    localStorage.setItem(QUEUE_KEY, '[]');
    for (const p of q) { // 順序保持のため直列送信(追加→編集の順を守る)
      try {
        const r = await fetch(CONFIG.URL, { method: 'POST', body: JSON.stringify(p) });
        const j = await r.json();
        if (!j.ok) throw new Error(j.error);
      } catch (e) { enqueue(p); }
    }
  }

  window.addEventListener('online', flushQueue);
  window.addEventListener('load', flushQueue);

  return {
    orderAdd(rec) {
      post({ type: 'order_add', orderId: rec.id, store: CONFIG.STORE,
             date: rec.date, pay: rec.payment, items: itemsToRows(rec.items) });
    },
    orderEdit(rec) {
      post({ type: 'order_edit', orderId: rec.id, store: CONFIG.STORE,
             date: rec.date, pay: rec.payment, items: itemsToRows(rec.items) });
    },
    orderDelete(id) {
      post({ type: 'order_delete', orderId: id });
    },
    expenseAdd(rec) {
      post({ type: 'expense', expenseId: rec.id, store: CONFIG.STORE, date: rec.date,
             category: EXPENSE_CATS[rec.category] || 'Other / その他',
             item: rec.category, amount: rec.amount, pay: rec.payment,
             notes: rec.memo || '' });
    },
    expenseEdit(rec) {
      post({ type: 'expense_edit', expenseId: rec.id, store: CONFIG.STORE, date: rec.date,
             category: EXPENSE_CATS[rec.category] || 'Other / その他',
             item: rec.category, amount: rec.amount, pay: rec.payment,
             notes: rec.memo || '' });
    },
    expenseDelete(id) {
      post({ type: 'expense_delete', expenseId: id });
    },
  };
})();
