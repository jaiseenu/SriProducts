/* ==========================================================================
   Sri Products — Storefront app
   Vanilla JS, hash router, no build step — same pattern as the internal
   Business Manager app, kept separate on purpose (public traffic vs.
   staff-only traffic shouldn't share a codebase or a set of keys).
   Phase 1 + 2 scope: browse catalog, cart, guest checkout, pickup or
   delivery with flat/free-threshold shipping, pay in person either
   way. No login, no online payment gateway yet.
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
    let settings = { delivery_fee: 49, free_delivery_threshold: 999 };
    if (Sb.ready) {
      const [locRes, settingsRes] = await Promise.all([
        Sb.client.from('v_pickup_locations').select('*'),
        Sb.client.from('v_store_settings').select('*')
      ]);
      if (!locRes.error) locations = locRes.data;
      if (!settingsRes.error) settingsRes.data.forEach(s => { settings[s.key] = Number(s.value); });
    }
    const subtotal = cartSubtotal(cart);

    function shippingFeeFor(fulfillment) {
      if (fulfillment !== 'delivery') return 0;
      return subtotal >= settings.free_delivery_threshold ? 0 : settings.delivery_fee;
    }
    function renderTotals(fulfillment) {
      const fee = shippingFeeFor(fulfillment);
      return `
        <div class="totals-row"><span>Subtotal</span><span>${money(subtotal)}</span></div>
        ${fulfillment === 'delivery' ? `<div class="totals-row"><span>Delivery</span><span>${fee === 0 ? 'Free' : money(fee)}</span></div>` : ''}
        <div class="totals-row grand"><span>Total</span><span>${money(subtotal + fee)}</span></div>
      `;
    }
    function renderFulfillmentFields(fulfillment) {
      if (fulfillment === 'pickup') {
        return `<div class="field"><label>Pickup location</label>
          <select id="pickupLoc">${locations.map(l => `<option value="${esc(l.id)}">${esc(l.name)} — ${esc(l.address)}</option>`).join('') || '<option value="">No pickup locations configured</option>'}</select>
        </div>`;
      }
      return `
        <div class="field"><label>Address line 1</label><input id="addr1" placeholder="House/flat no., street"></div>
        <div class="field"><label>Address line 2 (optional)</label><input id="addr2" placeholder="Area, apartment name"></div>
        <div class="field"><label>City</label><input id="addrCity" placeholder="City"></div>
        <div class="field"><label>Pincode</label><input id="addrPincode" inputmode="numeric" placeholder="6-digit pincode"></div>
        <div class="field"><label>Landmark (optional)</label><input id="addrLandmark" placeholder="Nearby landmark"></div>
        <div class="field hint">${subtotal >= settings.free_delivery_threshold ? 'Free delivery on this order.' : `Free delivery on orders over ${money(settings.free_delivery_threshold)}.`}</div>
      `;
    }

    let fulfillment = 'pickup';
    let payNow = false; // false = pay in person (pickup/delivery), true = pay online now

    function placeOrderLabel() {
      if (payNow) return 'Pay online now';
      return 'Place order (pay ' + (fulfillment === 'pickup' ? 'at pickup' : 'on delivery') + ')';
    }

    document.querySelector('.screen').innerHTML = `
      <div class="section-label">Your order</div>
      <div class="panel">${cart.map(c => `
        <div class="list-row"><div class="row-title">${esc(c.name)}</div><div class="amount">${c.quantity} ${esc(c.unit)} × ${money(c.price)}</div></div>
      `).join('')}</div>
      <div class="totals-panel" id="totalsPanel">${renderTotals(fulfillment)}</div>

      <div class="section-label">How would you like to get it?</div>
      <div class="btn-row" style="margin-top:0;">
        <button class="btn btn-primary" id="fulfillPickup">Pickup</button>
        <button class="btn btn-secondary" id="fulfillDelivery">Delivery</button>
      </div>

      <div class="section-label">How would you like to pay?</div>
      <div class="btn-row" style="margin-top:0;">
        <button class="btn btn-primary" id="payInPerson">Pay in person</button>
        <button class="btn btn-secondary" id="payOnline">Pay online now</button>
      </div>

      <div class="section-label">Your details</div>
      <div class="field"><label>Your name</label><input id="custName" placeholder="Full name"></div>
      <div class="field"><label>Phone number</label><input id="custPhone" type="tel" placeholder="10-digit mobile number"></div>
      <div id="fulfillmentFields">${renderFulfillmentFields(fulfillment)}</div>
      <div class="field"><label>Notes (optional)</label><textarea id="notes" rows="2" placeholder="Anything we should know?"></textarea></div>
      <div class="field hint" id="payHint">Pay in person when you collect your order.</div>
      <button class="btn btn-primary" id="placeOrderBtn">${placeOrderLabel()}</button>
    `;

    function setFulfillment(next) {
      fulfillment = next;
      document.getElementById('fulfillPickup').className = 'btn ' + (next === 'pickup' ? 'btn-primary' : 'btn-secondary');
      document.getElementById('fulfillDelivery').className = 'btn ' + (next === 'delivery' ? 'btn-primary' : 'btn-secondary');
      document.getElementById('fulfillmentFields').innerHTML = renderFulfillmentFields(next);
      document.getElementById('totalsPanel').innerHTML = renderTotals(next);
      updatePayHint();
      document.getElementById('placeOrderBtn').textContent = placeOrderLabel();
    }
    function setPayNow(next) {
      payNow = next;
      document.getElementById('payInPerson').className = 'btn ' + (!next ? 'btn-primary' : 'btn-secondary');
      document.getElementById('payOnline').className = 'btn ' + (next ? 'btn-primary' : 'btn-secondary');
      updatePayHint();
      document.getElementById('placeOrderBtn').textContent = placeOrderLabel();
    }
    function updatePayHint() {
      document.getElementById('payHint').textContent = payNow
        ? 'You will be taken to a secure Razorpay payment window after placing this order.'
        : (fulfillment === 'pickup' ? 'Pay in person when you collect your order.' : 'Pay in cash/UPI to the delivery person.');
    }
    document.getElementById('fulfillPickup').onclick = () => setFulfillment('pickup');
    document.getElementById('fulfillDelivery').onclick = () => setFulfillment('delivery');
    document.getElementById('payInPerson').onclick = () => setPayNow(false);
    document.getElementById('payOnline').onclick = () => setPayNow(true);

    document.getElementById('placeOrderBtn').onclick = async (e) => {
      const name = document.getElementById('custName').value.trim();
      const phone = document.getElementById('custPhone').value.trim();
      const notes = document.getElementById('notes').value.trim();
      if (!name) { toast('Enter your name.'); return; }
      if (!/^\d{10}$/.test(phone.replace(/\D/g, ''))) { toast('Enter a valid 10-digit phone number.'); return; }

      const payload = {
        p_fulfillment_type: fulfillment,
        p_customer_name: name,
        p_customer_phone: phone,
        p_pickup_location_id: null,
        p_delivery_address_line1: null,
        p_delivery_address_line2: null,
        p_delivery_city: null,
        p_delivery_pincode: null,
        p_delivery_landmark: null,
        p_notes: notes,
        p_items: cart.map(c => ({ item_id: c.itemId, quantity: c.quantity })),
        p_pay_online: payNow
      };
      if (fulfillment === 'pickup') {
        const pickupLoc = document.getElementById('pickupLoc').value;
        if (!pickupLoc) { toast('Select a pickup location.'); return; }
        payload.p_pickup_location_id = pickupLoc;
      } else {
        const addr1 = document.getElementById('addr1').value.trim();
        const city = document.getElementById('addrCity').value.trim();
        const pincode = document.getElementById('addrPincode').value.trim();
        if (!addr1) { toast('Enter your address.'); return; }
        if (!city) { toast('Enter your city.'); return; }
        if (!/^\d{6}$/.test(pincode)) { toast('Enter a valid 6-digit pincode.'); return; }
        payload.p_delivery_address_line1 = addr1;
        payload.p_delivery_address_line2 = document.getElementById('addr2').value.trim();
        payload.p_delivery_city = city;
        payload.p_delivery_pincode = pincode;
        payload.p_delivery_landmark = document.getElementById('addrLandmark').value.trim();
      }

      e.target.disabled = true; e.target.textContent = 'Placing order…';
      try {
        const { data, error } = await Sb.client.rpc('create_order', payload);
        if (error) throw error;
        const order = Array.isArray(data) ? data[0] : data;
        clearCart();
        if (payNow) {
          e.target.textContent = 'Opening payment…';
          await startRazorpayPayment(order.order_number, phone, name);
        } else {
          navigate('#/order/' + order.order_number + '/' + encodeURIComponent(phone));
        }
      } catch (err) {
        toast(err.message || 'Could not place order. Please try again.');
        e.target.disabled = false; e.target.textContent = placeOrderLabel();
      }
    };
  }

  // ---------- RAZORPAY ----------
  // Payment confirmation is never trusted from this callback alone —
  // verify-payment re-checks the signature server-side before marking
  // anything paid, and the razorpay-webhook Edge Function is the
  // durable fallback if the browser closes before this callback runs.

  async function startRazorpayPayment(orderNumber, phone, name) {
    try {
      const { data, error } = await Sb.client.functions.invoke('create-razorpay-order', { body: { order_number: orderNumber } });
      if (error) throw error;
      if (data.error) throw new Error(data.error);

      const rzp = new Razorpay({
        key: data.key_id,
        amount: data.amount,
        currency: data.currency,
        order_id: data.razorpay_order_id,
        name: 'Sri Products',
        description: 'Order ' + orderNumber,
        prefill: { name: name, contact: phone },
        theme: { color: '#101d33' },
        handler: async (response) => {
          try {
            const { data: verifyData, error: verifyError } = await Sb.client.functions.invoke('verify-payment', {
              body: {
                order_number: orderNumber,
                razorpay_order_id: response.razorpay_order_id,
                razorpay_payment_id: response.razorpay_payment_id,
                razorpay_signature: response.razorpay_signature
              }
            });
            if (verifyError || verifyData.error) toast('Payment received — confirming with the shop, check "Track an order" shortly.');
          } catch (e) { /* webhook will still confirm this independently */ }
          navigate('#/order/' + orderNumber + '/' + encodeURIComponent(phone));
        },
        modal: {
          ondismiss: () => {
            toast('Payment not completed. Your order is saved — you can pay from "Track an order".');
            navigate('#/order/' + orderNumber + '/' + encodeURIComponent(phone));
          }
        }
      });
      rzp.open();
    } catch (err) {
      toast(err.message || 'Could not start online payment. Your order is saved as pending.');
      navigate('#/order/' + orderNumber + '/' + encodeURIComponent(phone));
    }
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
      const isDelivery = order.fulfillment_type === 'delivery';
      const addressBlock = isDelivery ? `
        <div class="section-label">Delivery address</div>
        <div class="panel"><div class="list-row"><div>
          <div class="row-title">${esc(order.delivery_address_line1)}${order.delivery_address_line2 ? ', ' + esc(order.delivery_address_line2) : ''}</div>
          <div class="row-sub">${esc(order.delivery_city)} — ${esc(order.delivery_pincode)}${order.delivery_landmark ? ' · Near ' + esc(order.delivery_landmark) : ''}</div>
        </div></div></div>
      ` : '';
      document.querySelector('.screen').innerHTML = `
        <div class="empty-state" style="padding:8px 0 20px;">
          <div class="empty-title">Order ${esc(order.order_number)}</div>
          ${statusBadge(order.status)}
          <div class="muted" style="margin-top:6px;">${isDelivery ? 'Delivery' : 'Pickup'} · Pay ${isDelivery ? 'on delivery' : 'at pickup'}</div>
        </div>
        <div class="section-label">Items</div>
        <div class="panel">${items.map(i => `<div class="list-row"><div class="row-title">${esc(i.name)}</div><div class="amount">${i.quantity} ${esc(i.unit)} — ${money(i.lineTotal)}</div></div>`).join('')}</div>
        <div class="totals-panel">
          <div class="totals-row"><span>Subtotal</span><span>${money(order.subtotal)}</span></div>
          ${isDelivery ? `<div class="totals-row"><span>Delivery</span><span>${Number(order.shipping_fee) === 0 ? 'Free' : money(order.shipping_fee)}</span></div>` : ''}
          <div class="totals-row grand"><span>Total</span><span>${money(order.grand_total)}</span></div>
        </div>
        ${addressBlock}
        <div class="field hint" style="margin-top:14px;">${
          order.payment_method === 'online' && order.payment_status !== 'paid' ? 'Online payment not completed yet.'
          : isDelivery ? 'Have cash/UPI ready for the delivery person.' : 'Pay at pickup. Bring this order number.'
        }</div>
        ${order.payment_method === 'online' && order.payment_status === 'unpaid' && order.status === 'pending' ? '<button class="btn btn-primary" id="payNowBtn" style="margin-top:10px;">Pay now</button>' : ''}
        ${order.status === 'pending' ? '<button class="btn btn-danger" id="cancelBtn" style="margin-top:10px;">Cancel this order</button>' : ''}
      `;
      const payNowBtn = document.getElementById('payNowBtn');
      if (payNowBtn) payNowBtn.onclick = () => startRazorpayPayment(orderNumber, phone, '');
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
