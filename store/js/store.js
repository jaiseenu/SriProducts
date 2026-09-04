/* ==========================================================================
   Sri Products — Storefront app
   Vanilla JS, hash router, no build step — same pattern as the internal
   Business Manager app, kept separate on purpose (public traffic vs.
   staff-only traffic shouldn't share a codebase or a set of keys).
   Phase 1 scope: browse catalog, cart, guest checkout, pickup only,
   pay-on-pickup only. No login, no delivery, no online payment yet.
   ========================================================================== */

const Store = (() => {
  const root = document.getElementById('app');
  const CART_KEY = 'sri_store_cart';

  // ---------- utilities ----------

  function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function money(n) {
    const v = Number(n || 0);
    return '₹' + v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function navigate(hash) { location.hash = hash; }
  function toast(msg) {
    const el = document.createElement('div');
    el.className = 'toast';
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 3000);
  }
  function bindGoAttrs() { document.querySelectorAll('[data-go]').forEach(el => el.onclick = () => navigate(el.dataset.go)); }
  function errorState(msg) { return `<div class="empty-state"><div class="empty-title">Something went wrong</div><p>${esc(msg)}</p></div>`; }
  function statusBadge(status) {
    const cls = { pending: 'badge-pending', confirmed: 'badge-confirmed', ready: 'badge-ready', completed: 'badge-completed', cancelled: 'badge-cancelled' }[status] || 'badge-pending';
    const label = { pending: 'Pending', confirmed: 'Confirmed', ready: 'Ready for pickup', completed: 'Completed', cancelled: 'Cancelled' }[status] || status;
    return `<span class="badge ${cls}">${esc(label)}</span>`;
  }

  function shell(title, bodyHtml, opts) {
    opts = opts || {};
    const back = opts.back !== false;
    const cartCount = getCart().reduce((n, i) => n + i.quantity, 0);
    root.innerHTML = `
      <div class="top-bar">
        ${back ? '<button class="back-btn" id="backBtn">&#8592;</button>' : ''}
        <h1>${esc(title)}</h1>
        <button class="cart-btn" id="cartBtn">Cart${cartCount ? `<span class="cart-badge">${cartCount}</span>` : ''}</button>
      </div>
      <div class="screen">${bodyHtml}</div>
      <div class="footer-nav"><a data-go="#/shop">Shop</a><a data-go="#/track">Track an order</a></div>
    `;
    const backBtn = document.getElementById('backBtn');
    if (backBtn) backBtn.onclick = () => history.back();
    document.getElementById('cartBtn').onclick = () => navigate('#/cart');
    bindGoAttrs();
  }

  // ---------- cart (localStorage; guest checkout, no account needed) ----------

  function getCart() {
    try { return JSON.parse(localStorage.getItem(CART_KEY) || '[]'); } catch (e) { return []; }
  }
  function saveCart(cart) { localStorage.setItem(CART_KEY, JSON.stringify(cart)); }
  function addToCart(item, quantity) {
    const cart = getCart();
    const existing = cart.find(c => c.itemId === item.id);
    if (existing) existing.quantity += quantity;
    else cart.push({ itemId: item.id, name: item.name, unit: item.unit, price: Number(item.price), quantity });
    saveCart(cart);
  }
  function updateCartQty(itemId, quantity) {
    let cart = getCart();
    if (quantity <= 0) cart = cart.filter(c => c.itemId !== itemId);
    else { const row = cart.find(c => c.itemId === itemId); if (row) row.quantity = quantity; }
    saveCart(cart);
  }
  function clearCart() { saveCart([]); }
  function cartSubtotal(cart) { return cart.reduce((sum, c) => sum + c.quantity * c.price, 0); }

  // ---------- CATALOG ----------

  async function screenShop() {
    shell('Sri Products', `<div class="empty-state">Loading products…</div>`, { back: false });
    if (!Sb.ready) { document.querySelector('.screen').innerHTML = errorState('Store is not configured yet. Edit js/config.js with your Supabase project URL and anon key.'); return; }
    const { data, error } = await Sb.client.from('v_catalog').select('*').order('name');
    if (error) { document.querySelector('.screen').innerHTML = errorState(error.message); return; }
    const cart = getCart();
    document.querySelector('.screen').innerHTML = `
      <p class="muted" style="margin-bottom:14px;">Quality never compromised — order online, pick up in store.</p>
      <div class="product-grid">
        ${data.map(p => {
          const inCart = cart.find(c => c.itemId === p.id);
          const qty = inCart ? inCart.quantity : 0;
          const outOfStock = Number(p.available_qty) <= 0;
          return `
          <div class="product-card" data-item='${esc(JSON.stringify(p))}'>
            <div class="p-name">${esc(p.name)}</div>
            <div class="p-unit">per ${esc(p.unit)}</div>
            <div class="p-price">${p.price != null ? money(p.price) : 'Price unavailable'}</div>
            <div class="p-stock ${outOfStock ? 'out' : ''}">${outOfStock ? 'Out of stock' : Number(p.available_qty) + ' ' + esc(p.unit) + ' available'}</div>
            ${outOfStock || p.price == null ? '' : `
              <div class="qty-stepper" data-qty-for="${esc(p.id)}">
                <button data-step="-1">−</button>
                <div class="qty-val">${qty}</div>
                <button data-step="1">+</button>
              </div>
              <button class="add-btn" data-add="${esc(p.id)}" ${qty === 0 ? '' : 'style="display:none"'}>Add to cart</button>
            `}
          </div>`;
        }).join('') || '<div class="empty-state">No products available right now.</div>'}
      </div>
    `;
    bindGoAttrs();

    document.querySelectorAll('.product-card').forEach(card => {
      const product = JSON.parse(card.dataset.item);
      const stepper = card.querySelector('.qty-stepper');
      if (!stepper) return;
      const valEl = stepper.querySelector('.qty-val');
      const addBtn = card.querySelector('.add-btn');
      let pending = 0;
      stepper.querySelectorAll('button').forEach(btn => {
        btn.onclick = () => {
          const delta = Number(btn.dataset.step);
          pending = Math.max(0, pending + delta);
          if (pending > Number(product.available_qty)) pending = Number(product.available_qty);
          valEl.textContent = pending;
          addBtn.style.display = pending > 0 ? '' : 'none';
        };
      });
      addBtn.onclick = () => {
        addToCart(product, pending);
        toast(`Added ${pending} ${product.unit} of ${product.name} to cart.`);
        screenShop(); // re-render: resets the stepper and refreshes the cart badge count
      };
    });
  }

  // ---------- CART ----------

  function screenCart() {
    const cart = getCart();
    shell('Your cart', `
      <div class="panel" id="cartList">${cart.length ? cart.map(c => `
        <div class="list-row">
          <div>
            <div class="row-title">${esc(c.name)}</div>
            <div class="row-sub">${money(c.price)} / ${esc(c.unit)}</div>
          </div>
          <div style="text-align:right">
            <div class="qty-stepper" style="width:110px" data-item="${esc(c.itemId)}">
              <button data-step="-1">−</button>
              <div class="qty-val">${c.quantity}</div>
              <button data-step="1">+</button>
            </div>
            <button class="remove-link" data-remove="${esc(c.itemId)}" style="margin-top:6px">Remove</button>
          </div>
        </div>`).join('') : '<div class="empty-state">Your cart is empty. <div style="margin-top:10px"><a data-go="#/shop" style="color:var(--gold-600);font-weight:700;cursor:pointer;">Browse products &#8250;</a></div></div>'}</div>
      ${cart.length ? `
        <div class="totals-panel">
          <div class="totals-row grand"><span>Subtotal</span><span>${money(cartSubtotal(cart))}</span></div>
        </div>
        <div class="btn-row"><button class="btn btn-primary" id="checkoutBtn">Checkout</button></div>
      ` : ''}
    `, { back: false });
    bindGoAttrs();
    if (!cart.length) return;

    document.querySelectorAll('#cartList .qty-stepper').forEach(stepper => {
      const itemId = stepper.dataset.item;
      stepper.querySelectorAll('button').forEach(btn => {
        btn.onclick = () => {
          const row = getCart().find(c => c.itemId === itemId);
          if (!row) return;
          updateCartQty(itemId, row.quantity + Number(btn.dataset.step));
          screenCart();
        };
      });
    });
    document.querySelectorAll('[data-remove]').forEach(btn => {
      btn.onclick = () => { updateCartQty(btn.dataset.remove, 0); screenCart(); };
    });
    document.getElementById('checkoutBtn').onclick = () => navigate('#/checkout');
  }

  // ---------- CHECKOUT ----------

  async function screenCheckout() {
    const cart = getCart();
    if (!cart.length) { navigate('#/shop'); return; }
    shell('Checkout', `<div class="empty-state">Loading…</div>`);
    let locations = [];
    if (Sb.ready) {
      const { data, error } = await Sb.client.from('v_pickup_locations').select('*');
      if (!error) locations = data;
    }
    document.querySelector('.screen').innerHTML = `
      <div class="section-label">Your order</div>
      <div class="panel">${cart.map(c => `
        <div class="list-row"><div class="row-title">${esc(c.name)}</div><div class="amount">${c.quantity} ${esc(c.unit)} × ${money(c.price)}</div></div>
      `).join('')}</div>
      <div class="totals-panel"><div class="totals-row grand"><span>Subtotal</span><span>${money(cartSubtotal(cart))}</span></div></div>

      <div class="section-label">Pickup details</div>
      <div class="field"><label>Your name</label><input id="custName" placeholder="Full name"></div>
      <div class="field"><label>Phone number</label><input id="custPhone" type="tel" placeholder="10-digit mobile number"></div>
      <div class="field"><label>Pickup location</label>
        <select id="pickupLoc">${locations.map(l => `<option value="${esc(l.id)}">${esc(l.name)} — ${esc(l.address)}</option>`).join('') || '<option value="">No pickup locations configured</option>'}</select>
      </div>
      <div class="field"><label>Notes (optional)</label><textarea id="notes" rows="2" placeholder="Anything we should know?"></textarea></div>
      <div class="field hint">Pay in person when you collect your order — no online payment yet.</div>
      <button class="btn btn-primary" id="placeOrderBtn">Place order (pay at pickup)</button>
    `;
    document.getElementById('placeOrderBtn').onclick = async (e) => {
      const name = document.getElementById('custName').value.trim();
      const phone = document.getElementById('custPhone').value.trim();
      const pickupLoc = document.getElementById('pickupLoc').value;
      const notes = document.getElementById('notes').value.trim();
      if (!name) { toast('Enter your name.'); return; }
      if (!/^\d{10}$/.test(phone.replace(/\D/g, ''))) { toast('Enter a valid 10-digit phone number.'); return; }
      if (!pickupLoc) { toast('Select a pickup location.'); return; }
      e.target.disabled = true; e.target.textContent = 'Placing order…';
      try {
        const { data, error } = await Sb.client.rpc('create_order', {
          p_customer_name: name,
          p_customer_phone: phone,
          p_pickup_location_id: pickupLoc,
          p_notes: notes,
          p_items: cart.map(c => ({ item_id: c.itemId, quantity: c.quantity }))
        });
        if (error) throw error;
        const order = Array.isArray(data) ? data[0] : data;
        clearCart();
        navigate('#/order/' + order.order_number + '/' + encodeURIComponent(phone));
      } catch (err) {
        toast(err.message || 'Could not place order. Please try again.');
        e.target.disabled = false; e.target.textContent = 'Place order (pay at pickup)';
      }
    };
  }

  // ---------- ORDER CONFIRMATION / STATUS ----------

  async function screenOrderStatus(orderNumber, phone) {
    shell('Order status', `<div class="empty-state">Loading…</div>`, { back: false });
    try {
      const { data, error } = await Sb.client.rpc('get_order_status', { p_order_number: orderNumber, p_phone: phone });
      if (error) throw error;
      const order = Array.isArray(data) ? data[0] : data;
      if (!order) { document.querySelector('.screen').innerHTML = errorState('Order not found. Check your order number and phone number.'); return; }
      const items = order.items || [];
      document.querySelector('.screen').innerHTML = `
        <div class="empty-state" style="padding:8px 0 20px;">
          <div class="empty-title">Order ${esc(order.order_number)}</div>
          ${statusBadge(order.status)}
        </div>
        <div class="section-label">Items</div>
        <div class="panel">${items.map(i => `<div class="list-row"><div class="row-title">${esc(i.name)}</div><div class="amount">${i.quantity} ${esc(i.unit)} — ${money(i.lineTotal)}</div></div>`).join('')}</div>
        <div class="totals-panel"><div class="totals-row grand"><span>Total</span><span>${money(order.subtotal)}</span></div></div>
        <div class="field hint" style="margin-top:14px;">Pay at pickup. Bring this order number.</div>
        ${order.status === 'pending' ? '<button class="btn btn-danger" id="cancelBtn" style="margin-top:10px;">Cancel this order</button>' : ''}
      `;
      const cancelBtn = document.getElementById('cancelBtn');
      if (cancelBtn) cancelBtn.onclick = async () => {
        if (!window.confirm('Cancel this order?')) return;
        try { await Sb.client.rpc('cancel_order', { p_order_number: orderNumber, p_phone: phone }); toast('Order cancelled.'); screenOrderStatus(orderNumber, phone); }
        catch (err) { toast(err.message); }
      };
    } catch (err) {
      document.querySelector('.screen').innerHTML = errorState(err.message);
    }
  }

  function screenTrack() {
    shell('Track an order', `
      <div class="field"><label>Order number</label><input id="trackOrder" placeholder="ORD-000001"></div>
      <div class="field"><label>Phone number used at checkout</label><input id="trackPhone" type="tel" placeholder="10-digit mobile number"></div>
      <button class="btn btn-primary" id="trackBtn">Check status</button>
    `, { back: false });
    document.getElementById('trackBtn').onclick = () => {
      const orderNumber = document.getElementById('trackOrder').value.trim();
      const phone = document.getElementById('trackPhone').value.trim();
      if (!orderNumber || !phone) { toast('Enter both fields.'); return; }
      navigate('#/order/' + encodeURIComponent(orderNumber) + '/' + encodeURIComponent(phone));
    };
  }

  // ---------- ROUTER ----------

  function route() {
    const hash = location.hash || '#/shop';
    const parts = hash.replace(/^#\//, '').split('/');
    if (parts[0] === 'shop' || parts[0] === '') return screenShop();
    if (parts[0] === 'cart') return screenCart();
    if (parts[0] === 'checkout') return screenCheckout();
    if (parts[0] === 'track') return screenTrack();
    if (parts[0] === 'order' && parts[1]) return screenOrderStatus(decodeURIComponent(parts[1]), decodeURIComponent(parts[2] || ''));
    return screenShop();
  }

  function init() {
    window.addEventListener('hashchange', route);
    route();
  }

  return { init };
})();

Store.init();
