/* CHAEN POS → Google Sheets 連携モジュール (v1.22)
   デプロイ後に URL / TOKEN を設定。STOREは店舗が増えたら端末ごとに変更。
   v1.20: 同じレジをシフト交代で共有する運用に対応。注文・経費に時刻を送信し、
   fetchTodayData() で本日分の注文・経費をまとめて取得できるようにした。
   v1.21: 注文に対応スタッフ名(rec.staff)を送信するようにした。
   v1.22: 経費にも記録スタッフ名(rec.staff)を送信するようにした。 */
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

  // RECIPE_IDS の逆引き('SIG12' → {name:'Signature Matcha Latte', size:'12oz'})
  // 他端末の注文をPOS画面用に復元する際に使う
  const RECIPE_LOOKUP = Object.fromEntries(
    Object.entries(RECIPE_IDS).map(([key, id]) => {
      const [name, size] = key.split('|');
      return [id, { name, size }];
    })
  );

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

  // 同じSTOREの「本日分」の注文・経費をまとめて取得
  // (シフト交代で端末が変わっても、その日のデータをこの端末にも取り込むため)
  // 失敗時は空を返す(通信エラーでPOS本体の動作を止めないため)
  async function fetchTodayData() {
    if (!CONFIG.ENABLED) return { orders: [], expenses: [] };
    try {
      const url = `${CONFIG.URL}?action=today&token=${encodeURIComponent(CONFIG.TOKEN)}&store=${encodeURIComponent(CONFIG.STORE)}`;
      const r = await fetch(url);
      const j = await r.json();
      if (!j.ok) throw new Error(j.error || 'fetch error');
      return { orders: j.orders || [], expenses: j.expenses || [] };
    } catch (err) {
      console.warn('SheetSync fetchTodayData failed:', err);
      return { orders: [], expenses: [] };
    }
  }

  window.addEventListener('online', flushQueue);
  window.addEventListener('load', flushQueue);

  return {
    recipeLookup: RECIPE_LOOKUP,
    fetchTodayData,
    orderAdd(rec) {
      post({ type: 'order_add', orderId: rec.id, store: CONFIG.STORE,
             date: rec.date, time: rec.time || '', staff: rec.staff || '', pay: rec.payment, items: itemsToRows(rec.items) });
    },
    orderEdit(rec) {
      post({ type: 'order_edit', orderId: rec.id, store: CONFIG.STORE,
             date: rec.date, time: rec.time || '', staff: rec.staff || '', pay: rec.payment, items: itemsToRows(rec.items) });
    },
    orderDelete(id) {
      post({ type: 'order_delete', orderId: id });
    },
    expenseAdd(rec) {
      post({ type: 'expense', expenseId: rec.id, store: CONFIG.STORE, date: rec.date, time: rec.time || '', staff: rec.staff || '',
             category: EXPENSE_CATS[rec.category] || 'Other / その他',
             item: rec.category, amount: rec.amount, pay: rec.payment,
             notes: rec.memo || '' });
    },
    expenseEdit(rec) {
      post({ type: 'expense_edit', expenseId: rec.id, store: CONFIG.STORE, date: rec.date, time: rec.time || '', staff: rec.staff || '',
             category: EXPENSE_CATS[rec.category] || 'Other / その他',
             item: rec.category, amount: rec.amount, pay: rec.payment,
             notes: rec.memo || '' });
    },
    expenseDelete(id) {
      post({ type: 'expense_delete', expenseId: id });
    },
  };
})();
