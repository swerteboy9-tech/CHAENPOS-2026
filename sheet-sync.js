/* CHAEN POS → Google Sheets 連携モジュール (v1.24)
   デプロイ後に URL / TOKEN を設定。STOREは店舗が増えたら端末ごとに変更。
   v1.20: 同じレジをシフト交代で共有する運用に対応。注文・経費に時刻を送信し、
   fetchTodayData() で本日分の注文・経費をまとめて取得できるようにした。
   v1.21: 注文に対応スタッフ名(rec.staff)を送信するようにした。
   v1.22: 経費にも記録スタッフ名(rec.staff)を送信するようにした。
   v1.23: 商品マスターをシートに一本化。fetchProducts() で Recipe Master の
          商品リストを取得し、localStorageにキャッシュする。取得できた場合は
          Recipe ID の対応表もシート由来のものを優先して使う(下の RECIPE_IDS は
          初回起動時・通信不能時のフォールバックとして残す)。
   v1.24: 在庫管理に対応。fetchPresets() で経費プリセット、fetchInventory() で
          残在庫を取得する(どちらもlocalStorageにキャッシュ)。経費の送信に
          材料ID・入庫数量・個数・仕入単位・下書きフラグを追加した。 */
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
  const MENU_KEY  = 'chaen_menu_cache';
  const PRESET_KEY = 'chaen_preset_cache';
  const INV_KEY    = 'chaen_inventory_cache';

  /* シートから取得した 商品名|サイズ → Recipe ID。取得できるまでは null。 */
  let dynamicIds = null;
  /* シートから取得した Recipe ID → {name,size}。 */
  let dynamicLookup = null;

  function recipeIdFor(name, size) {
    const key = `${name}|${size}`;
    if (dynamicIds && dynamicIds[key]) return dynamicIds[key];
    return RECIPE_IDS[key] || `??${name}`;
  }

  function itemsToRows(items) {
    return items.map(it => ({
      recipeId: recipeIdFor(it.name, it.size),
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

  /* Recipe Master の商品リストを取得して localStorage にキャッシュする。
     戻り値 {updatedAt, fetchedAt, products:[{id,name,size,category,price}]}
     通信失敗時は例外を投げる(呼び出し側でキャッシュにフォールバックする)。 */
  async function fetchProducts() {
    if (!CONFIG.ENABLED) throw new Error('sheet sync disabled');
    const url = `${CONFIG.URL}?action=products&token=${encodeURIComponent(CONFIG.TOKEN)}`;
    const r = await fetch(url);
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'products fetch error');
    const products = j.products || [];
    if (!products.length) throw new Error('empty product list');
    const payload = { updatedAt: j.updatedAt || '', fetchedAt: new Date().toISOString(), products };
    try { localStorage.setItem(MENU_KEY, JSON.stringify(payload)); } catch (e) { /* 容量超過は無視 */ }
    applyProducts(products);
    return payload;
  }

  /* 経費プリセットを取得してキャッシュする */
  async function fetchPresets() {
    if (!CONFIG.ENABLED) throw new Error('sheet sync disabled');
    const url = `${CONFIG.URL}?action=expensepresets&token=${encodeURIComponent(CONFIG.TOKEN)}`;
    const r = await fetch(url);
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'presets fetch error');
    const presets = j.presets || [];
    if (!presets.length) throw new Error('empty preset list');
    const payload = { fetchedAt: new Date().toISOString(), presets };
    try { localStorage.setItem(PRESET_KEY, JSON.stringify(payload)); } catch (e) {}
    return payload;
  }
  function cachedPresets() {
    try {
      const v = JSON.parse(localStorage.getItem(PRESET_KEY) || 'null');
      if (v && Array.isArray(v.presets) && v.presets.length) return v;
    } catch (e) {}
    return null;
  }

  /* 在庫を取得してキャッシュする */
  async function fetchInventory() {
    if (!CONFIG.ENABLED) throw new Error('sheet sync disabled');
    const url = `${CONFIG.URL}?action=inventory&token=${encodeURIComponent(CONFIG.TOKEN)}`;
    const r = await fetch(url);
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'inventory fetch error');
    const items = j.items || [];
    const payload = { fetchedAt: new Date().toISOString(), items };
    try { localStorage.setItem(INV_KEY, JSON.stringify(payload)); } catch (e) {}
    return payload;
  }
  function cachedInventory() {
    try {
      const v = JSON.parse(localStorage.getItem(INV_KEY) || 'null');
      if (v && Array.isArray(v.items)) return v;
    } catch (e) {}
    return null;
  }

  /* キャッシュ済みの商品リストを返す(なければ null) */
  function cachedProducts() {
    try {
      const v = JSON.parse(localStorage.getItem(MENU_KEY) || 'null');
      if (v && Array.isArray(v.products) && v.products.length) { applyProducts(v.products); return v; }
    } catch (e) { /* 壊れたキャッシュは無視 */ }
    return null;
  }

  /* 商品リストから Recipe ID の対応表を作り直す */
  function applyProducts(products) {
    const ids = {}, lookup = {};
    products.forEach(p => {
      ids[`${p.name}|${p.size}`] = p.id;
      lookup[p.id] = { name: p.name, size: p.size };
    });
    dynamicIds = ids;
    dynamicLookup = lookup;
  }

  /* Recipe ID → {name,size}。シート由来を優先し、無ければ内蔵表を使う。 */
  function lookupRecipe(id) {
    return (dynamicLookup && dynamicLookup[id]) || RECIPE_LOOKUP[id] || null;
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

  /* 経費の送信内容を組み立てる。
     rec.sheetCategory があればプリセット経由(シートのカテゴリー名そのまま)、
     無ければ従来の EXPENSE_CATS 変換を使う。 */
  function expensePayload(type, rec) {
    const viaPreset = !!rec.sheetCategory;
    return {
      type: type, expenseId: rec.id, store: CONFIG.STORE,
      date: rec.date, time: rec.time || '', staff: rec.staff || '',
      category: viaPreset ? rec.sheetCategory : (EXPENSE_CATS[rec.category] || 'Other / その他'),
      item: viaPreset ? (rec.item || rec.category) : rec.category,
      amount: rec.amount, pay: rec.payment, notes: rec.memo || '',
      ingId: rec.ingId || '', qty: rec.qty || '',
      packs: (rec.packs === undefined ? '' : rec.packs),
      buyUnit: rec.buyUnit || '',
      unitPrice: rec.unitPrice || '',
      draft: !!rec.draft
    };
  }

  window.addEventListener('online', flushQueue);
  window.addEventListener('load', flushQueue);

  return {
    recipeLookup: RECIPE_LOOKUP,   // 内蔵フォールバック表(後方互換のため残す)
    lookupRecipe,                  // シート由来を優先した逆引き
    fetchProducts,
    cachedProducts,
    fetchPresets,
    cachedPresets,
    fetchInventory,
    cachedInventory,
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
      post(expensePayload('expense', rec));
    },
    expenseEdit(rec) {
      post(expensePayload('expense_edit', rec));
    },
    expenseDelete(id) {
      post({ type: 'expense_delete', expenseId: id });
    },
  };
})();
