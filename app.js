const PRODUCTS=[
{name:'Signature Matcha Latte',size:'12oz',price:99,icon:'🍵'},
{name:'Matcha Mango',size:'16oz',price:150,icon:'🥭'},
{name:'Matcha Cloud Latte',size:'16oz',price:160,icon:'☁️'},
{name:'Matcha Frappe',size:'16oz',price:160,icon:'🥤'},
{name:'Matcha Sea Salt',size:'16oz',price:160,icon:'🧂'},
{name:'Matcha Strawberry',size:'16oz',price:180,icon:'🍓'},
{name:'Matcha Oreo',size:'16oz',price:210,icon:'🍪'},
{name:'Einspanner Matcha Latte',size:'16oz',price:190,icon:'🥛'},
{name:'Nutella Matcha Latte',size:'16oz',price:270,icon:'🍫'},
{name:'Churros',size:'5 sticks',price:110,icon:'🥖'}
];
const EXPENSES=['Ice','Water','Breakfast / Meal','Tissue / Supplies','Staff Advance','Ingredient Purchase','Delivery','Other'];
const state={view:'home',cart:[],selectedProduct:null,selectedSize:null,selectedDiscount:0,selectedQty:0,payment:'CASH',historyTab:'sales',historyDetailDate:null,editingOrderId:null,editingExpenseId:null,reportMonth:null,recipeId:null,orderDateTime:null,quickOptions:{}};
const $=s=>document.querySelector(s);
const peso=n=>'₱'+Number(n||0).toLocaleString('en-PH');
const pad=n=>String(n).padStart(2,'0');
const dateKey=(d=new Date())=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const timeKey=(d=new Date())=>`${pad(d.getHours())}:${pad(d.getMinutes())}`;
const dateTimeLocalValue=(d=new Date())=>`${dateKey(d)}T${timeKey(d)}`;
const load=(k,def=[])=>JSON.parse(localStorage.getItem(k)||JSON.stringify(def));
const save=(k,v)=>localStorage.setItem(k,JSON.stringify(v));
const orders=()=>load('chaen_orders'); const expenses=()=>load('chaen_expenses'); const closings=()=>load('chaen_closings'); const auditLog=()=>load('chaen_audit_log');
const staffList=()=>load('chaen_staff'); const currentStaff=()=>localStorage.getItem('chaen_current_staff')||'';
function setCurrentStaff(name){localStorage.setItem('chaen_current_staff',name)}
function addStaffMember(name){const list=staffList();if(!list.includes(name)){list.push(name);save('chaen_staff',list)}}
function addAudit(entry){const arr=auditLog();arr.push({id:crypto.randomUUID(),at:new Date().toISOString(),...entry});save('chaen_audit_log',arr)}
function markClosingNeedsReview(date){const arr=closings();let changed=false;arr.forEach(c=>{if(c.date===date){c.needsReview=true;changed=true}});if(changed)save('chaen_closings',arr)}
function todayOrders(){return orders().filter(x=>x.date===dateKey())} function todayExpenses(){return expenses().filter(x=>x.date===dateKey())}
// totals() は orders()/expenses() をそのまま使う。シフト交代で同じレジを
// 共有する運用のため、他端末の注文・経費も syncFromSheet() でこの端末の
// ローカル保存に取り込んでから計算する(=自然に合算された金額になる)。
function totals(){const os=todayOrders(), es=todayExpenses(); const sales=os.reduce((a,b)=>a+b.total,0),cash=os.filter(x=>x.payment==='CASH').reduce((a,b)=>a+b.total,0),gcash=sales-cash,exp=es.reduce((a,b)=>a+b.amount,0),cashExp=es.filter(x=>x.payment==='CASH').reduce((a,b)=>a+b.amount,0),cups=os.reduce((a,b)=>a+b.items.reduce((s,i)=>s+i.qty,0),0);return{sales,cash,gcash,exp,cashExp,net:sales-exp,cups,expectedCash:cash-cashExp}}

// ── シフト交代同期:他端末が入力した本日分の注文・経費を、この端末のローカル保存に取り込む ──
function buildImportedOrder(raw){
  const items=raw.items.map(it=>{
    const meta=SheetSync.recipeLookup[it.recipeId]||{name:it.recipeId,size:''};
    const prod=PRODUCTS.find(p=>p.name===meta.name);
    const price=prod?prod.price+(meta.size==='22oz'?30:0):0;
    const discount=Number(it.promo)>0?30:0;
    return{name:meta.name,size:meta.size,price,qty:Number(it.qty)||0,discount};
  });
  const total=items.reduce((a,b)=>a+(b.price-b.discount)*b.qty,0);
  return{id:raw.orderId,date:raw.date,time:raw.time||'',items,total,payment:raw.pay,backdated:false,staff:raw.staff||'',importedFromOtherDevice:true};
}
function buildImportedExpense(raw){
  return{id:raw.expenseId,date:raw.date,time:raw.time||'',category:raw.category,amount:Number(raw.amount)||0,payment:raw.pay,memo:raw.notes||'',staff:raw.staff||'',importedFromOtherDevice:true};
}
async function syncFromSheet(){
  const {orders:rawOrders,expenses:rawExpenses}=await SheetSync.fetchTodayData();
  let changed=false;
  if(rawOrders.length){
    const arr=orders();
    const localIds=new Set(arr.map(o=>o.id));
    rawOrders.forEach(o=>{if(!localIds.has(o.orderId)){arr.push(buildImportedOrder(o));changed=true}});
    if(changed)save('chaen_orders',arr);
  }
  if(rawExpenses.length){
    const arr=expenses();
    const localIds=new Set(arr.map(e=>e.id));
    let expChanged=false;
    rawExpenses.forEach(e=>{if(!localIds.has(e.expenseId)){arr.push(buildImportedExpense(e));expChanged=true}});
    if(expChanged){save('chaen_expenses',arr);changed=true}
  }
  if(changed&&(state.view==='home'||state.view==='history'||state.view==='sales'||state.view==='close'))render();
}
function setView(v){state.view=v;render()}
function cartQty(){return state.cart.reduce((sum,item)=>sum+Number(item.qty||0),0)}
function productCartQty(name){return state.cart.filter(item=>item.name===name).reduce((sum,item)=>sum+Number(item.qty||0),0)}
function updateOrderBadge(){const badge=document.querySelector('[data-order-badge]');if(!badge)return;const qty=cartQty();badge.textContent=qty>99?'99+':String(qty);badge.hidden=qty===0;badge.setAttribute('aria-label',`Current order quantity ${qty}`)}
function navActive(){document.querySelectorAll('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.view===state.view));updateOrderBadge()}
function syncTopbarHeight(){const h=document.querySelector('.topbar')?.offsetHeight||68;document.documentElement.style.setProperty('--topbar-h',`${h}px`)}
function render(){const app=$('#app'); if(state.view==='home') app.innerHTML=homeView(); if(state.view==='order') app.innerHTML=orderView(); if(state.view==='expense') app.innerHTML=expenseView(); if(state.view==='sales') app.innerHTML=salesView(); if(state.view==='close') app.innerHTML=closeView(); if(state.view==='report') app.innerHTML=reportView(); if(state.view==='recipe') app.innerHTML=recipeView(); if(state.view==='history') app.innerHTML=historyView(); syncTopbarHeight(); bind(); navActive();}
function staffCardHTML(){const list=staffList(),cur=currentStaff();return `<div class="card"><div class="section-title">👤 Staff on Duty</div><select id="staffSelect" class="input"><option value="">-- Select --</option>${list.map(n=>`<option value="${n}" ${n===cur?'selected':''}>${n}</option>`).join('')}</select><button class="btn" data-add-staff style="width:100%;margin-top:8px">+ Add Staff</button></div>`}
function homeView(){const t=totals();return `${staffCardHTML()}<div class="card hero"><div class="label">TODAY SALES</div><div class="big">${peso(t.sales)}</div><div class="row"><span>${t.cups} items sold</span><span>${dateKey()}</span></div></div><div class="grid2"><button class="btn primary" data-go="order">🧋 ORDER<br><small>Take an order</small></button><button class="btn alt" data-go="expense">🧺 EXPENSE<br><small>Record a cost</small></button><button class="btn" data-go="report">📊 REPORT<br><small>This month</small></button><button class="btn action-red" data-go="close">🔒 CLOSE DAY<br><small>Count the cash</small></button></div><button class="btn btn-wide" data-go="recipe">📖 RECIPES<br><small>What goes in each cup</small></button><div class="card"><div class="section-title">Today’s Summary <small class="muted">(shared register — includes all shifts)</small></div><div class="row"><span>Cash</span><b>${peso(t.cash)}</b></div><div class="row"><span>GCash</span><b>${peso(t.gcash)}</b></div><div class="row"><span>Expenses</span><b>${peso(t.exp)}</b></div><div class="row"><span>Net</span><b>${peso(t.net)}</b></div></div>`}
function quickOption(i){
  const p=PRODUCTS[i];
  if(!state.quickOptions[i]) state.quickOptions[i]={size:p.name==='Signature Matcha Latte'?'12oz':p.name==='Churros'?'5 sticks':'16oz',qty:0,discountPct:0};
  return state.quickOptions[i];
}
function quickPrice(p,size){return p.price+(size==='22oz'?30:0)}
function quickCard(p,i){
  const q=quickOption(i),sizes=p.name==='Signature Matcha Latte'?['12oz']:p.name==='Churros'?['5 sticks']:['16oz','22oz'];
  const price=quickPrice(p,q.size),discount=(q.size==='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros')?q.discountPct:0,lineTotal=(price-discount)*q.qty;
  return `<div class="card quick-product-card">
    <div class="product quick-product-head"><div class="picon">${p.icon}</div><div><div class="pname">${p.name}</div><small>${p.size}</small></div><div class="price">${peso(price)}</div></div>
    <div class="quick-controls">
      <div class="quick-field"><span class="quick-label">Size</span><div class="quick-chips">${sizes.map(size=>`<button class="quick-chip ${q.size===size?'active':''}" data-quick-size="${i}" data-size-value="${size}">${size}</button>`).join('')}</div></div>
      <div class="quick-field"><span class="quick-label">Quantity</span><div class="quick-qty"><button data-quick-minus="${i}">−</button><b>${q.qty}</b><button data-quick-plus="${i}">＋</button></div></div>
      <div class="quick-field"><span class="quick-label">Discount</span><div class="quick-chips"><button class="quick-chip ${discount===0?'active':''}" data-quick-discount="${i}" data-discount-value="0">None</button>${(q.size==='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros')?`<button class="quick-chip ${discount===30?'active discount':''}" data-quick-discount="${i}" data-discount-value="30">₱30 OFF</button>`:''}</div>${q.size!=='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros'?'<small class="muted">₱30 OFF applies to 22oz only</small>':''}</div>
    </div>
    <div class="quick-product-footer"><div><small>Item Total</small><div class="quick-line-total">${peso(lineTotal)}</div></div><button class="btn action-red quick-add-btn" data-quick-add="${i}">＋ ADD TO ORDER</button></div>
  </div>`
}
function orderView(){
  if(state.selectedProduct)return productDetailView();
  const cartTotal=state.cart.reduce((a,b)=>a+(b.price-b.discount)*b.qty,0);
  const currentQty=cartQty();
  return `${state.cart.length?`<div class="sticky-order-summary"><div class="sticky-order-info"><b>Current Order (${currentQty})</b><span class="sticky-order-total">${peso(cartTotal)}</span></div><button class="btn action-red sticky-checkout-btn" data-checkout aria-label="Review order and proceed to payment">🛒 CHECKOUT</button></div>`:''}
  <div class="section-title order-section-title">Select Product</div>
  <div class="product-picker-help">Tap a product to add to order</div>
  <div class="product-picker-grid">${PRODUCTS.map((p,i)=>{
    const qty=productCartQty(p.name);
    return `<button class="product-picker-card" data-product="${i}" aria-label="Select ${p.name}">
      <span class="product-picker-icon-wrap">
        <span class="product-picker-icon">${p.icon}</span>
        ${qty>0?`<span class="product-qty-badge" aria-label="${qty} in current order">${qty>99?'99+':qty}</span>`:''}
      </span>
      <span class="product-picker-name">${p.name}</span>
    </button>`;
  }).join('')}</div>`
}
function productDetailView(){const p=state.selectedProduct; const sizes=p.name==='Signature Matcha Latte'?['12oz']:p.name==='Churros'?['5 sticks']:['16oz','22oz']; const size=state.selectedSize||sizes[0]; const base=p.price+(size==='22oz'?30:0);return `<button class="btn product-detail-back" data-back>← Back to Products</button><div class="section-title product-detail-title">Order Details</div><div class="card product-detail-card"><div class="product"><div class="picon">${p.icon}</div><div><div class="pname">${p.name}</div><div class="price">${peso(base)}</div></div></div><div class="field"><label>Size</label><div class="chips">${sizes.map(s=>`<button class="chip ${s===size?'active':''}" data-size="${s}">${s}</button>`).join('')}</div></div><div class="field"><label>Quantity</label><div class="qty-control"><button class="qty-btn" data-qty-minus aria-label="Decrease quantity">−</button><input class="qty-input" id="productQty" type="number" inputmode="numeric" min="0" max="99" value="${state.selectedQty}" /><button class="qty-btn" data-qty-plus aria-label="Increase quantity">＋</button></div><small class="muted">Change this when ordering multiple units of the same item.</small></div><div class="field"><label>Discount (per item)</label><div class="chips">${(size==='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros'?[0,30]:[0]).map(d=>`<button class="chip ${d===state.selectedDiscount?'active':''}" data-discount="${d}">${d?peso(d)+' OFF':'None'}</button>`).join('')}</div>${(size!=='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros')?'<small class="muted">₱30 OFF applies to 22oz only</small>':''}</div><div class="row" style="margin:12px 0"><span>Item Total</span><b class="total">${peso((base-state.selectedDiscount)*state.selectedQty)}</b></div><button class="btn primary" style="width:100%" data-add>ADD TO ORDER</button></div>`}
function checkoutView(){const subtotal=state.cart.reduce((a,b)=>a+b.price*b.qty,0),disc=state.cart.reduce((a,b)=>a+b.discount*b.qty,0),total=subtotal-disc;const dt=state.orderDateTime||dateTimeLocalValue();const isEdit=!!state.editingOrderId;return `<div class="section-title">${isEdit?'Edit Order':'Order Review'}</div>${isEdit?`<div class="notice">You are editing a past order. Saving will automatically recalculate sales for that date.</div>`:''}<div class="card list">${state.cart.map((i,idx)=>`<div class="item"><div class="row"><div><b>${i.name}</b><br><small>${i.size}${i.discount?` / ${i.discountLabel||peso(i.discount)+' OFF'} each`:''}</small><div class="cart-qty"><button data-cart-minus="${idx}">−</button><span>${i.qty}</span><button data-cart-plus="${idx}">＋</button></div></div><div class="right"><b>${peso((i.price-i.discount)*i.qty)}</b><br><button class="btn danger" style="padding:5px 8px" data-remove="${idx}">Remove</button></div></div></div>`).join('')}</div><button class="btn" style="width:100%;margin-bottom:12px" data-add-more>＋ Add / Change Items</button><div class="card"><div class="row"><span>Subtotal</span><b>${peso(subtotal)}</b></div><div class="row"><span>Discount</span><b>-${peso(disc)}</b></div><div class="row"><span>Total</span><span class="total">${peso(total)}</span></div><div class="field"><label>Order Date & Time</label><input class="input" id="orderDateTime" type="datetime-local" value="${dt}" /><small class="muted">For a missed entry, change this to the actual order date and time.</small></div><div class="field"><label>Payment Method</label><div class="grid2"><button class="btn ${state.payment==='CASH'?'primary':''}" data-pay="CASH">CASH</button><button class="btn ${state.payment==='GCASH'?'primary':''}" data-pay="GCASH">GCASH</button></div></div><button class="btn ${isEdit?'primary':'action-red'}" style="width:100%" data-complete>${isEdit?'SAVE CHANGES':'COMPLETE ORDER'}</button>${isEdit?`<button class="btn" style="width:100%;margin-top:8px" data-cancel-edit>CANCEL EDIT</button>`:''}</div>`}
function expenseView(){const ed=state.editingExpenseId?expenses().find(x=>x.id===state.editingExpenseId):null;return `<div class="section-title">${ed?'Edit Expense':'Record Expense'}</div>${ed?`<div class="card" style="padding:10px 14px;margin-bottom:10px"><small class="muted">Editing: ${ed.date} ${ed.time} / -${peso(ed.amount)}</small></div>`:''}<div class="card"><div class="field"><label>Category</label><select id="expCat">${EXPENSES.map(x=>`<option ${ed&&ed.category===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Amount</label><input class="input" id="expAmount" inputmode="decimal" placeholder="e.g. 120" value="${ed?ed.amount:''}" /></div><div class="field"><label>Payment Method</label><select id="expPay"><option ${ed&&ed.payment==='CASH'?'selected':''}>CASH</option><option ${ed&&ed.payment==='GCASH'?'selected':''}>GCASH</option></select></div><div class="field"><label>Memo</label><input class="input" id="expMemo" placeholder="e.g. Ice 20kg" value="${ed?String(ed.memo||'').replaceAll('"','&quot;'):''}" /></div><button class="btn primary" style="width:100%" data-save-exp>${ed?'UPDATE EXPENSE':'SAVE EXPENSE'}</button>${ed?'<button class="btn" style="width:100%;margin-top:8px" data-cancel-exp-edit>Cancel Edit</button>':''}</div>`}
function salesView(){const t=totals();return `<div class="section-title">Today’s Sales</div><div class="grid2"><div class="card"><div class="label">TOTAL SALES</div><div class="total">${peso(t.sales)}</div></div><div class="card"><div class="label">ITEMS SOLD</div><div class="total">${t.cups}</div></div><div class="card"><div class="label">CASH</div><div class="total">${peso(t.cash)}</div></div><div class="card"><div class="label">GCASH</div><div class="total">${peso(t.gcash)}</div></div><div class="card"><div class="label">EXPENSE</div><div class="total">${peso(t.exp)}</div></div><div class="card"><div class="label">NET</div><div class="total">${peso(t.net)}</div></div></div>`}
function closeView(){const t=totals();const last=closings().filter(x=>x.date===dateKey()).at(-1);return `<div class="section-title">Close Day & Reconcile</div><div class="notice" style="margin-bottom:12px">CLOSE DAY does not delete sales history. It only saves the end-of-day reconciliation.</div><div class="card"><div class="row"><span>Total Sales</span><b>${peso(t.sales)}</b></div><div class="row"><span>Total Expense</span><b>${peso(t.exp)}</b></div><div class="row"><span>Net Sales</span><b>${peso(t.net)}</b></div><hr><div class="row"><span>Cash Sales</span><b>${peso(t.cash)}</b></div><div class="row"><span>Cash Expense</span><b>${peso(t.cashExp)}</b></div><div class="row"><span>Expected Cash</span><b>${peso(t.expectedCash)}</b></div><div class="field"><label>Actual Cash</label><input class="input" id="actualCash" inputmode="decimal" value="${last?.actualCash??t.expectedCash}" /></div><button class="btn action-red" style="width:100%" data-close-day>CLOSE DAY</button>${last?`<div class="notice" style="margin-top:12px">Previous reconciliation difference: <b>${peso(last.difference)}</b></div>`:''}</div>`}
function historyProductIcon(name){return PRODUCTS.find(p=>p.name===name)?.icon||'•'}
function dailySalesGroups(){const map=new Map();orders().forEach(o=>{if(!map.has(o.date))map.set(o.date,{date:o.date,orders:[],total:0,items:0});const g=map.get(o.date);g.orders.push(o);g.total+=Number(o.total||0);g.items+=o.items.reduce((sum,i)=>sum+Number(i.qty||0),0)});return [...map.values()].sort((a,b)=>b.date.localeCompare(a.date))}
function historySalesHTML(){const groups=dailySalesGroups();return groups.length?groups.map(g=>`<div class="daily-sales-card"><div><div class="daily-sales-date">${g.date}</div><div class="daily-sales-meta">${g.orders.length} orders / ${g.items} items</div></div><div class="daily-sales-right"><div class="daily-sales-total">${peso(g.total)}</div><button class="btn detail-btn" data-history-detail="${g.date}">DETAILS</button></div></div>`).join(''):'<div class="muted">No sales history yet.</div>'}
function historySalesDetailHTML(date){const os=orders().filter(o=>o.date===date).sort((a,b)=>a.time.localeCompare(b.time));const total=os.reduce((sum,o)=>sum+Number(o.total||0),0);const itemCount=os.reduce((sum,o)=>sum+o.items.reduce((s,i)=>s+Number(i.qty||0),0),0);return `<button class="btn history-back" data-history-back>← Back to Sales History</button><div class="daily-detail-head card"><div><div class="label">${date}</div><div class="daily-detail-total">${peso(total)}</div></div><div class="right"><b>${os.length} orders</b><br><small>${itemCount} items</small></div></div><div class="list history-list">${os.map((o,idx)=>`<div class="history-order"><div class="history-order-head"><div><b>ORDER ${idx+1} / ${o.time}</b><br><small>${o.payment}${o.staff?' · 👤 '+o.staff:''}</small>${o.backdated?` <span class="history-badge">BACKDATED</span>`:''}${o.editedAt?` <span class="history-badge edited">EDITED</span>`:''}</div><div class="history-order-actions"><div class="history-order-total">${peso(o.total)}</div><button class="mini-edit" data-edit-order="${o.id}">Edit</button><button class="mini-delete" data-delete-order="${o.id}">Delete</button></div></div><div class="history-lines">${o.items.map(i=>`<div class="history-line"><div class="history-icon">${historyProductIcon(i.name)}</div><div class="history-line-name"><b>${i.name}</b><small>${i.size}</small></div><div class="history-qty">× ${i.qty}</div><div class="history-line-total">${peso((i.price-i.discount)*i.qty)}</div></div>`).join('')}</div></div>`).join('')}</div>`}
function historyExpensesHTML(){const es=expenses().slice().reverse().slice(0,30);return es.length?es.map(e=>`<div class="item"><div class="row"><div><b>${e.category}</b><br><small>${e.date} ${e.time} / ${e.payment}${e.staff?' · 👤 '+e.staff:''} ${e.memo?'/ '+e.memo:''}</small></div><div style="text-align:right"><b>-${peso(e.amount)}</b><div class="history-order-actions" style="margin-top:4px"><button class="mini-edit" data-edit-expense="${e.id}">Edit</button><button class="mini-delete" data-delete-expense="${e.id}">Delete</button></div></div></div></div>`).join(''):'<div class="muted">No expense history yet.</div>'}
function historyView(){const salesActive=state.historyTab==='sales';const detail=salesActive&&state.historyDetailDate;return `<div class="section-title">History</div><div class="card"><div class="chips history-tabs"><button class="chip ${salesActive?'active':''}" data-history-tab="sales">SALES</button><button class="chip ${!salesActive?'active':''}" data-history-tab="expense">EXPENSE</button></div><div id="historyList">${salesActive?(detail?historySalesDetailHTML(state.historyDetailDate):`<div class="list history-list">${historySalesHTML()}</div>`):`<div class="list history-list">${historyExpensesHTML()}</div>`}</div></div><button class="btn" style="width:100%" data-export>Export CSV</button>`}
function beginEditOrder(id){const o=orders().find(x=>x.id===id);if(!o)return; if(state.cart.length&&!confirm('There is an unfinished current order. Editing a past order will replace the current cart. Continue?'))return;state.editingOrderId=o.id;state.cart=JSON.parse(JSON.stringify(o.items));state.payment=o.payment;state.orderDateTime=`${o.date}T${o.time}`;state.selectedProduct=null;state.view='order';render();setTimeout(()=>{document.querySelector('[data-checkout]')?.click()},0)}
function cancelOrderEdit(){state.editingOrderId=null;state.orderDateTime=null;state.cart=[];state.payment='CASH';setView('history')}
function bind(){updateOrderBadge();const ss=$('#staffSelect');if(ss)ss.onchange=()=>{setCurrentStaff(ss.value);render()};const as=$('[data-add-staff]');if(as)as.onclick=()=>{const name=(prompt('Staff name')||'').trim();if(!name)return;addStaffMember(name);setCurrentStaff(name);render()};document.querySelectorAll('[data-go]').forEach(b=>b.onclick=()=>{state.editingExpenseId=null;state.recipeId=null;setView(b.dataset.go)});document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{state.editingExpenseId=null;state.recipeId=null;setView(b.dataset.view)});document.querySelectorAll('[data-quick-pay]').forEach(b=>b.onclick=()=>{state.payment=b.dataset.quickPay;render()});document.querySelectorAll('[data-quick-size]').forEach(b=>b.onclick=()=>{const i=+b.dataset.quickSize,q=quickOption(i);q.size=b.dataset.sizeValue;if(q.size!=='22oz')q.discountPct=0;render()});document.querySelectorAll('[data-quick-minus]').forEach(b=>b.onclick=()=>{const i=+b.dataset.quickMinus;const q=quickOption(i);q.qty=Math.max(0,q.qty-1);render()});document.querySelectorAll('[data-quick-plus]').forEach(b=>b.onclick=()=>{const i=+b.dataset.quickPlus;const q=quickOption(i);q.qty=Math.min(99,q.qty+1);render()});document.querySelectorAll('[data-quick-discount]').forEach(b=>b.onclick=()=>{const i=+b.dataset.quickDiscount;quickOption(i).discountPct=+b.dataset.discountValue;render()});document.querySelectorAll('[data-quick-add]').forEach(b=>b.onclick=()=>{const i=+b.dataset.quickAdd,p=PRODUCTS[i],q=quickOption(i);if(q.qty<=0)return alert('Please set quantity to 1 or more.');const price=quickPrice(p,q.size),discount=(q.size==='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros')?q.discountPct:0,item={name:p.name,size:q.size,price,qty:q.qty,discount,discountLabel:discount?`₱30 OFF`:''};const same=state.cart.find(x=>x.name===item.name&&x.size===item.size&&x.discount===item.discount);if(same)same.qty=Math.min(99,same.qty+item.qty);else state.cart.push(item);state.quickOptions[i]={size:p.name==='Signature Matcha Latte'?'12oz':p.name==='Churros'?'5 sticks':'16oz',qty:0,discountPct:0};render()});document.querySelectorAll('[data-product]').forEach(el=>el.onclick=()=>{state.selectedProduct=PRODUCTS[+el.dataset.product];state.selectedSize=null;state.selectedDiscount=0;state.selectedQty=0;render()});const back=$('[data-back]');if(back)back.onclick=()=>{state.selectedProduct=null;state.selectedQty=0;render()};document.querySelectorAll('[data-size]').forEach(b=>b.onclick=()=>{state.selectedSize=b.dataset.size;if(state.selectedSize!=='22oz')state.selectedDiscount=0;render()});document.querySelectorAll('[data-discount]').forEach(b=>b.onclick=()=>{state.selectedDiscount=+b.dataset.discount;render()});const qm=$('[data-qty-minus]'),qp=$('[data-qty-plus]'),qi=$('#productQty');if(qm)qm.onclick=()=>{state.selectedQty=Math.max(0,Number(state.selectedQty||0)-1);render()};if(qp)qp.onclick=()=>{state.selectedQty=Math.min(99,Number(state.selectedQty||0)+1);render()};if(qi)qi.onchange=()=>{state.selectedQty=Math.max(0,Math.min(99,Number(qi.value)||0));render()};const add=$('[data-add]');if(add)add.onclick=()=>{const p=state.selectedProduct;const size=state.selectedSize||(p.name==='Signature Matcha Latte'?'12oz':p.name==='Churros'?'5 sticks':'16oz');const price=p.price+(size==='22oz'?30:0);const qty=Math.max(0,Math.min(99,Number(state.selectedQty)||0));if(qty<=0)return alert('Please set quantity to 1 or more.');const discount=(size==='22oz'&&p.name!=='Signature Matcha Latte'&&p.name!=='Churros')?state.selectedDiscount:0;state.cart.push({name:p.name,size,price,qty,discount,discountLabel:discount?'₱30 OFF':''});state.selectedProduct=null;state.selectedSize=null;state.selectedDiscount=0;state.selectedQty=0;render()};const co=$('[data-checkout]');if(co)co.onclick=()=>{$('#app').innerHTML=checkoutView();bind()};document.querySelectorAll('[data-pay]').forEach(b=>b.onclick=()=>{state.orderDateTime=$('#orderDateTime')?.value||state.orderDateTime;state.payment=b.dataset.pay;$('#app').innerHTML=checkoutView();bind()});document.querySelectorAll('[data-remove]').forEach(b=>b.onclick=()=>{state.orderDateTime=$('#orderDateTime')?.value||state.orderDateTime;state.cart.splice(+b.dataset.remove,1);$('#app').innerHTML=checkoutView();bind()});document.querySelectorAll('[data-cart-minus]').forEach(b=>b.onclick=()=>{state.orderDateTime=$('#orderDateTime')?.value||state.orderDateTime;const i=+b.dataset.cartMinus;state.cart[i].qty=Math.max(1,state.cart[i].qty-1);$('#app').innerHTML=checkoutView();bind()});document.querySelectorAll('[data-cart-plus]').forEach(b=>b.onclick=()=>{state.orderDateTime=$('#orderDateTime')?.value||state.orderDateTime;const i=+b.dataset.cartPlus;state.cart[i].qty=Math.min(99,state.cart[i].qty+1);$('#app').innerHTML=checkoutView();bind()});const addMore=$('[data-add-more]');if(addMore)addMore.onclick=()=>{state.orderDateTime=$('#orderDateTime')?.value||state.orderDateTime;state.selectedProduct=null;render()};const comp=$('[data-complete]');if(comp)comp.onclick=()=>{if(!state.cart.length)return alert('Please add an item.');const subtotal=state.cart.reduce((a,b)=>a+b.price*b.qty,0),disc=state.cart.reduce((a,b)=>a+b.discount*b.qty,0);const input=$('#orderDateTime');const raw=input?.value||state.orderDateTime||dateTimeLocalValue();const [d,tRaw]=raw.split('T');const t=(tRaw||'00:00').slice(0,5);const now=new Date();const chosen=new Date(`${d}T${t}:00`);const backdated=(now-chosen)>10*60*1000;const arr=orders();if(state.editingOrderId){const idx=arr.findIndex(x=>x.id===state.editingOrderId);if(idx<0)return alert('The order to edit could not be found.');const before=JSON.parse(JSON.stringify(arr[idx]));arr[idx]={...arr[idx],date:d,time:t,items:JSON.parse(JSON.stringify(state.cart)),total:subtotal-disc,payment:state.payment,backdated:arr[idx].backdated||backdated,editedAt:now.toISOString()};save('chaen_orders',arr);SheetSync.orderEdit(arr[idx]);markClosingNeedsReview(before.date);markClosingNeedsReview(d);addAudit({action:'order_edit',orderId:state.editingOrderId,before,after:arr[idx]});state.editingOrderId=null;state.orderDateTime=null;state.cart=[];state.payment='CASH';alert('Order updated. Sales for that date have been recalculated automatically.');setView('history')}else{if(!currentStaff())return alert('Please select the staff on duty before completing the order.');const rec={id:crypto.randomUUID(),date:d,time:t,items:JSON.parse(JSON.stringify(state.cart)),total:subtotal-disc,payment:state.payment,backdated,staff:currentStaff()};arr.push(rec);save('chaen_orders',arr);SheetSync.orderAdd(rec);if(backdated)markClosingNeedsReview(d);addAudit({action:backdated?'order_backdated_add':'order_add',orderId:rec.id,after:rec});state.cart=[];state.orderDateTime=null;alert(backdated?'Recorded as a backdated order.':'Order recorded.');setView('home')}};const se=$('[data-save-exp]');if(se)se.onclick=()=>{const amount=Number($('#expAmount').value);if(!amount)return alert('Please enter an amount.');const now=new Date();const arr=expenses();if(state.editingExpenseId){const idx=arr.findIndex(x=>x.id===state.editingExpenseId);if(idx<0)return alert('Expense not found.');const before=JSON.parse(JSON.stringify(arr[idx]));arr[idx]={...arr[idx],category:$('#expCat').value,amount,payment:$('#expPay').value,memo:$('#expMemo').value,editedAt:now.toISOString()};save('chaen_expenses',arr);SheetSync.expenseEdit(arr[idx]);addAudit({action:'expense_edit',expenseId:state.editingExpenseId,before,after:arr[idx]});state.editingExpenseId=null;state.historyTab='expense';alert('Expense updated.');setView('history')}else{if(!currentStaff())return alert('Please select the staff on duty before recording an expense.');const expRec={id:crypto.randomUUID(),date:dateKey(now),time:now.toLocaleTimeString('ja-JP',{hour:'2-digit',minute:'2-digit'}),category:$('#expCat').value,amount,payment:$('#expPay').value,memo:$('#expMemo').value,staff:currentStaff()};arr.push(expRec);save('chaen_expenses',arr);SheetSync.expenseAdd(expRec);alert('Expense recorded.');setView('home')}};const cxe=$('[data-cancel-exp-edit]');if(cxe)cxe.onclick=()=>{state.editingExpenseId=null;state.historyTab='expense';setView('history')};const cd=$('[data-close-day]');if(cd)cd.onclick=()=>{const t=totals(),actual=Number($('#actualCash').value||0),difference=actual-t.expectedCash;const arr=closings();const now=new Date();const rec={id:crypto.randomUUID(),date:dateKey(now),time:timeKey(now),actualCash:actual,expectedCash:t.expectedCash,difference,totalSales:t.sales,totalExpense:t.exp,netSales:t.net};arr.push(rec);save('chaen_closings',arr);addAudit({action:'close_day',closingId:rec.id,after:rec});alert(`Day closed\nDifference: ${peso(difference)}\nSales history has been retained.`);render()};document.querySelectorAll('[data-history-tab]').forEach(b=>b.onclick=()=>{state.historyTab=b.dataset.historyTab;state.historyDetailDate=null;render()});document.querySelectorAll('[data-history-detail]').forEach(b=>b.onclick=()=>{state.historyDetailDate=b.dataset.historyDetail;render()});const historyBack=$('[data-history-back]');if(historyBack)historyBack.onclick=()=>{state.historyDetailDate=null;render()};document.querySelectorAll('[data-edit-order]').forEach(b=>b.onclick=()=>beginEditOrder(b.dataset.editOrder));document.querySelectorAll('[data-delete-order]').forEach(b=>b.onclick=()=>{const id=b.dataset.deleteOrder;const arr=orders();const idx=arr.findIndex(x=>x.id===id);if(idx<0)return alert('Order not found.');const o=arr[idx];if(!confirm(`Delete this order?\n${o.date} ${o.time} / ${peso(o.total)}\n\nThis will recalculate the sales total for that date.`))return;arr.splice(idx,1);save('chaen_orders',arr);SheetSync.orderDelete(id);markClosingNeedsReview(o.date);addAudit({action:'order_delete',orderId:id,before:o});alert('Order deleted. Sales for that date have been recalculated.');render()});document.querySelectorAll('[data-recipe-open]').forEach(b=>b.onclick=()=>{state.recipeId=b.dataset.recipeOpen;render()});document.querySelectorAll('[data-recipe-size]').forEach(b=>b.onclick=()=>{state.recipeId=b.dataset.recipeSize;render()});const rb=$('[data-recipe-back]');if(rb)rb.onclick=()=>{state.recipeId=null;render()};document.querySelectorAll('[data-month-shift]').forEach(b=>b.onclick=()=>{const cur=state.reportMonth||monthKey();const next=shiftMonth(cur,+b.dataset.monthShift);if(next>monthKey())return;state.reportMonth=next;render()});const et=$('[data-edit-target]');if(et)et.onclick=()=>{const v=prompt('Monthly sales goal (PHP)',String(monthTarget()));if(v===null)return;const n=Number(v);if(!n||n<0)return alert('Enter a number greater than 0.');localStorage.setItem(REPORT_TARGET_KEY,String(n));render()};const oc=$('[data-open-cash]');if(oc)oc.onchange=()=>{setOpeningBalance(state.reportMonth||monthKey(),{cash:Number(oc.value)||0});render()};const og=$('[data-open-gcash]');if(og)og.onchange=()=>{setOpeningBalance(state.reportMonth||monthKey(),{gcash:Number(og.value)||0});render()};document.querySelectorAll('[data-edit-expense]').forEach(b=>b.onclick=()=>{state.editingExpenseId=b.dataset.editExpense;setView('expense')});document.querySelectorAll('[data-delete-expense]').forEach(b=>b.onclick=()=>{const id=b.dataset.deleteExpense;const arr=expenses();const idx=arr.findIndex(x=>x.id===id);if(idx<0)return alert('Expense not found.');const e=arr[idx];if(!confirm(`Delete this expense?\n${e.date} ${e.time} / ${e.category} / ${peso(e.amount)}`))return;arr.splice(idx,1);save('chaen_expenses',arr);SheetSync.expenseDelete(id);addAudit({action:'expense_delete',expenseId:id,before:e});alert('Expense deleted.');render()});const cancelEdit=$('[data-cancel-edit]');if(cancelEdit)cancelEdit.onclick=cancelOrderEdit;const dtInput=$('#orderDateTime');if(dtInput)dtInput.onchange=()=>{state.orderDateTime=dtInput.value};const ex=$('[data-export]');if(ex)ex.onclick=exportCSV;}
function exportCSV(){const rows=[['type','date','time','description','payment','amount']];orders().forEach(o=>rows.push(['sale',o.date,o.time,o.items.map(i=>`${i.name} ${i.size}`).join(' + '),o.payment,o.total]));expenses().forEach(e=>rows.push(['expense',e.date,e.time,`${e.category} ${e.memo||''}`.trim(),e.payment,-e.amount]));const csv='\ufeff'+rows.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=`chaen-pos-${dateKey()}.csv`;a.click();URL.revokeObjectURL(a.href)}
setInterval(()=>{$('#clock').textContent=new Date().toLocaleString('ja-JP',{month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'})},1000);render();if('serviceWorker'in navigator)navigator.serviceWorker.register('sw.js');syncFromSheet();setInterval(syncFromSheet,30000);

/* ===== v1.19: REPORT (monthly dashboard) ===== */
const REPORT_TARGET_KEY='chaen_month_target';
const OPENING_KEY='chaen_opening_balance';
function monthKey(d=new Date()){return `${d.getFullYear()}-${pad(d.getMonth()+1)}`}
function monthLabel(mk){const [y,m]=mk.split('-');return `${['January','February','March','April','May','June','July','August','September','October','November','December'][+m-1]} ${y}`}
function shiftMonth(mk,delta){const [y,m]=mk.split('-').map(Number);const d=new Date(y,m-1+delta,1);return monthKey(d)}
function monthTarget(){return Number(localStorage.getItem(REPORT_TARGET_KEY)||45000)}
function openingBalance(mk){const all=load(OPENING_KEY,{});return all[mk]||{cash:0,gcash:0}}
function setOpeningBalance(mk,patch){const all=load(OPENING_KEY,{});all[mk]={...openingBalance(mk),...patch};save(OPENING_KEY,all)}
function monthStats(mk){
  const os=orders().filter(o=>o.date.startsWith(mk));
  const es=expenses().filter(e=>e.date.startsWith(mk));
  const sales=os.reduce((a,b)=>a+b.total,0);
  const cashSales=os.filter(o=>o.payment==='CASH').reduce((a,b)=>a+b.total,0);
  const gcashSales=sales-cashSales;
  const exp=es.reduce((a,b)=>a+b.amount,0);
  const cashExp=es.filter(e=>e.payment==='CASH').reduce((a,b)=>a+b.amount,0);
  const gcashExp=exp-cashExp;
  const items=os.reduce((a,o)=>a+o.items.reduce((s,i)=>s+i.qty,0),0);
  const days=new Set(os.map(o=>o.date)).size;
  const open=openingBalance(mk);
  return {os,es,sales,cashSales,gcashSales,exp,cashExp,gcashExp,items,days,
    net:sales-exp,orderCount:os.length,
    avgTicket:os.length?Math.round(sales/os.length):0,
    cashBal:open.cash+cashSales-cashExp,
    gcashBal:open.gcash+gcashSales-gcashExp,open};
}
function productRanking(mk){
  const map={};
  orders().filter(o=>o.date.startsWith(mk)).forEach(o=>o.items.forEach(i=>{
    const k=i.name;
    if(!map[k])map[k]={name:k,qty:0,sales:0,icon:(PRODUCTS.find(p=>p.name===k)||{}).icon||'🍵'};
    map[k].qty+=i.qty;
    map[k].sales+=(i.price-i.discount)*i.qty;
  }));
  return Object.values(map).sort((a,b)=>b.qty-a.qty);
}
function expenseRanking(mk){
  const map={};
  expenses().filter(e=>e.date.startsWith(mk)).forEach(e=>{map[e.category]=(map[e.category]||0)+e.amount});
  return Object.entries(map).map(([name,amount])=>({name,amount})).sort((a,b)=>b.amount-a.amount);
}
function reportView(){
  const mk=state.reportMonth||monthKey();
  const s=monthStats(mk), target=monthTarget();
  const pct=target>0?Math.min(100,Math.round(s.sales/target*100)):0;
  const ranks=productRanking(mk), maxQty=ranks.length?ranks[0].qty:1;
  const eranks=expenseRanking(mk), maxExp=eranks.length?eranks[0].amount:1;
  const isThisMonth=mk===monthKey();
  return `
  <div class="month-switch">
    <button class="month-arrow" data-month-shift="-1" aria-label="Previous month">‹</button>
    <div class="month-now"><b>${monthLabel(mk)}</b><small>${s.days} days with sales</small></div>
    <button class="month-arrow" data-month-shift="1" ${isThisMonth?'disabled':''} aria-label="Next month">›</button>
  </div>

  <div class="card cup-card">
    <div class="cup" style="--fill:${pct}%"><div class="cup-liquid"></div><div class="cup-foam"></div></div>
    <div class="cup-info">
      <div class="label">SALES THIS MONTH</div>
      <div class="cup-total">${peso(s.sales)}</div>
      <div class="cup-goal">${pct}% of ${peso(target)} goal</div>
      <button class="mini-edit" data-edit-target>Change goal</button>
    </div>
  </div>

  <div class="stat-grid">
    <div class="stat stat-sales"><span class="stat-emoji">🧾</span><div class="label">ORDERS</div><b>${s.orderCount}</b><small>${s.items} items</small></div>
    <div class="stat stat-ticket"><span class="stat-emoji">🎯</span><div class="label">AVG TICKET</div><b>${peso(s.avgTicket)}</b><small>per order</small></div>
    <div class="stat stat-exp"><span class="stat-emoji">🧺</span><div class="label">EXPENSES</div><b>${peso(s.exp)}</b><small>${s.es.length} records</small></div>
    <div class="stat stat-net"><span class="stat-emoji">${s.net>=0?'🌱':'🍂'}</span><div class="label">NET</div><b>${peso(s.net)}</b><small>sales − expenses</small></div>
  </div>

  <div class="section-title">Balance</div>
  <div class="wallet-grid">
    <div class="wallet wallet-cash">
      <div class="wallet-head"><span>💵</span><b>CASH</b></div>
      <div class="wallet-total">${peso(s.cashBal)}</div>
      <div class="wallet-line"><span>Opening</span><input class="wallet-input" data-open-cash inputmode="decimal" value="${s.open.cash}" /></div>
      <div class="wallet-line"><span>In</span><b class="pos">+${peso(s.cashSales)}</b></div>
      <div class="wallet-line"><span>Out</span><b class="neg">−${peso(s.cashExp)}</b></div>
    </div>
    <div class="wallet wallet-gcash">
      <div class="wallet-head"><span>📱</span><b>GCASH</b></div>
      <div class="wallet-total">${peso(s.gcashBal)}</div>
      <div class="wallet-line"><span>Opening</span><input class="wallet-input" data-open-gcash inputmode="decimal" value="${s.open.gcash}" /></div>
      <div class="wallet-line"><span>In</span><b class="pos">+${peso(s.gcashSales)}</b></div>
      <div class="wallet-line"><span>Out</span><b class="neg">−${peso(s.gcashExp)}</b></div>
    </div>
  </div>

  <div class="section-title">Best sellers</div>
  <div class="card rank-card">${ranks.length?ranks.map((r,i)=>`
    <div class="rank">
      <span class="rank-no rank-no-${i<3?i+1:'x'}">${i+1}</span>
      <span class="rank-icon">${r.icon}</span>
      <div class="rank-body">
        <div class="rank-top"><b>${r.name}</b><span class="rank-qty">${r.qty}</span></div>
        <div class="rank-bar"><i style="width:${Math.max(6,Math.round(r.qty/maxQty*100))}%"></i></div>
        <small>${peso(r.sales)}</small>
      </div>
    </div>`).join(''):'<div class="empty">No sales yet this month. Take the first order and it shows up here.</div>'}</div>

  <div class="section-title">Where the money went</div>
  <div class="card rank-card">${eranks.length?eranks.map(r=>`
    <div class="rank rank-exp">
      <span class="rank-icon">🧺</span>
      <div class="rank-body">
        <div class="rank-top"><b>${r.name}</b><span class="rank-qty">${peso(r.amount)}</span></div>
        <div class="rank-bar rank-bar-exp"><i style="width:${Math.max(6,Math.round(r.amount/maxExp*100))}%"></i></div>
      </div>
    </div>`).join(''):'<div class="empty">No expenses recorded this month.</div>'}</div>
  `;
}

/* ===== v1.20: RECIPE (build & cost breakdown) ===== */
function recipeGroups(){
  const g={};
  RECIPES.forEach(r=>{(g[r.name]=g[r.name]||[]).push(r)});
  return g;
}
function recipeById(id){return RECIPES.find(r=>r.id===id)}
function recipeView(){
  const sel=state.recipeId?recipeById(state.recipeId):null;
  if(!sel) return recipeListView();
  const groups=recipeGroups()[sel.name]||[sel];
  const pourable=sel.items.filter(i=>i.q>0);
  const volume=pourable.reduce((a,b)=>a+b.q,0);
  const totalCost=+(sel.ingTotal+sel.packCost).toFixed(2);
  const layers=pourable.slice().sort((a,b)=>b.q-a.q);
  return `
  <button class="btn history-back" data-recipe-back>‹ All recipes</button>
  <div class="section-title">${sel.name}</div>
  ${groups.length>1?`<div class="chips" style="margin:-4px 0 14px">${groups.map(g=>`<button class="chip ${g.id===sel.id?'active':''}" data-recipe-size="${g.id}">${g.size}</button>`).join('')}</div>`:''}

  <div class="card build-card">
    <div class="build-cup">
      <div class="build-cup-body">
        ${layers.map(i=>{const s=ingStyle(i.n);const pct=i.q/volume*100;return `<div class="build-layer" style="height:${pct}%;background:${s.c}" title="${i.n}"></div>`}).join('')}
      </div>
      <small>${volume} g/ml in the cup</small>
    </div>
    <div class="build-legend">
      ${layers.map(i=>{const s=ingStyle(i.n);const pct=Math.round(i.q/volume*100);return `
        <div class="legend-row">
          <span class="legend-dot" style="background:${s.c}"></span>
          <span class="legend-name">${s.e} ${i.n}</span>
          <b class="legend-qty">${i.q}${i.u}</b>
          <span class="legend-pct">${pct}%</span>
        </div>`}).join('')}
      ${sel.items.filter(i=>!i.q).map(i=>`<div class="legend-row legend-row-muted"><span class="legend-dot" style="background:${ingStyle(i.n).c}"></span><span class="legend-name">${ingStyle(i.n).e} ${i.n}</span><b class="legend-qty">to taste</b><span class="legend-pct">—</span></div>`).join('')}
      ${!pourable.length?'<div class="empty">No ingredients entered yet. Add them in Recipe Master and this fills in.</div>':''}
    </div>
  </div>

  <div class="section-title">Cost per cup</div>
  <div class="card">
    <div class="cost-bar">
      ${sel.items.filter(i=>i.c>0).sort((a,b)=>b.c-a.c).map(i=>`<i style="width:${i.c/totalCost*100}%;background:${ingStyle(i.n).c}" title="${i.n}"></i>`).join('')}
      <i style="width:${sel.packCost/totalCost*100}%;background:#C6BCA6" title="Packaging"></i>
    </div>
    ${sel.items.filter(i=>i.c>0).sort((a,b)=>b.c-a.c).map(i=>`
      <div class="cost-row"><span class="legend-dot" style="background:${ingStyle(i.n).c}"></span><span>${i.n}</span><b>${peso(i.c)}</b></div>`).join('')}
    <div class="cost-row"><span class="legend-dot" style="background:#C6BCA6"></span><span>Packaging${sel.pack.length?` · ${sel.pack.length} items`:''}</span><b>${peso(sel.packCost)}</b></div>
    <hr>
    <div class="row"><span>Selling price</span><b class="cost-price">${peso(sel.price)}</b></div>
    <div class="row"><span>Total cost</span><b>${peso(totalCost)}</b></div>
    <div class="row"><span>Gross profit</span><b class="cost-profit">${peso(sel.gross)}</b></div>
    <div class="row"><span>Food cost</span><b class="fc-pill ${sel.fcPct>35?'fc-high':sel.fcPct<15?'fc-low':''}">${sel.fcPct}%</b></div>
    ${sel.fcPct<15?'<div class="notice" style="margin-top:12px">Food cost looks unusually low — check that every ingredient is entered in Recipe Master.</div>':''}
    ${sel.fcPct>35?'<div class="notice" style="margin-top:12px">Food cost is above 35%. Worth reviewing the price or the portion sizes.</div>':''}
  </div>
  ${sel.pack.length?`<div class="section-title">Packaging</div><div class="card"><div class="pack-list">${sel.pack.map(p=>`<span class="pack-chip">${p}</span>`).join('')}</div></div>`:''}
  `;
}
function recipeListView(){
  const groups=recipeGroups();
  return `<div class="section-title">Recipes</div>
  <p class="product-picker-help">Tap a drink to see what goes in the cup and where the cost sits.</p>
  <div class="recipe-grid">
    ${Object.entries(groups).map(([name,list])=>{
      const first=list[0];
      const icon=(PRODUCTS.find(p=>p.name===name)||{}).icon||'🍵';
      const fc=Math.round(list.reduce((a,b)=>a+b.fcPct,0)/list.length);
      return `<button class="recipe-tile" data-recipe-open="${first.id}">
        <span class="recipe-tile-icon">${icon}</span>
        <span class="recipe-tile-body">
          <b>${name}</b>
          <small>${list.map(l=>l.size).join(' · ')}</small>
          <span class="fc-pill ${fc>35?'fc-high':fc<15?'fc-low':''}">${fc}% food cost</span>
        </span>
      </button>`}).join('')}
  </div>`;
}
