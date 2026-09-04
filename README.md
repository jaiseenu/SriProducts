# Sri Products — Storefront (Phase 1)

Public storefront: browse products, add to cart, guest checkout,
**pickup only, pay-at-pickup only**. No login, no delivery, no online
payment yet — those are later phases. Runs on Supabase (Postgres),
separate from the internal Business Manager app for now.

## 1. Set up Supabase

1. Create a project at [supabase.com](https://supabase.com) (free tier).
2. Open the SQL Editor and run `../supabase/schema.sql` (in the repo
   root's `supabase/` folder, one level up from this `store/` folder).
   It creates the tables, the public read-only views, the
   `create_order` / `get_order_status` / `cancel_order` functions, the
   RLS policies, and inserts 3 sample products so you have something
   to test with immediately.
3. Project Settings → API → copy the **Project URL** and **anon public**
   key. Do **not** use the `service_role` key anywhere in this folder —
   it bypasses Row Level Security and must never ship in frontend code.

## 2. Configure the frontend

Edit `js/config.js`:
```js
window.STORE_CONFIG = {
  SUPABASE_URL: 'https://xxxx.supabase.co',
  SUPABASE_ANON_KEY: 'eyJ...'
};
```

## 3. Run it locally / deploy it

No build step. Either:
- Open `index.html` directly, or serve the folder with any static
  server (`npx serve .`, VS Code "Live Server", etc.) — a real HTTP
  server is safer than `file://` for CORS behavior with Supabase.
- Deploy the `store/` folder to GitHub Pages, Netlify, Vercel, or
  Cloudflare Pages — it's fully static.

Keep this on a different path/subdomain than the internal Business
Manager (e.g. `shop.sriproducts.com` vs. the staff app) since they're
different audiences and, later, different auth models.

## 4. Test the end-to-end flow

1. Load the site — you should see the 3 sample products from the seed
   data with prices and "100 [unit] available".
2. Add a couple to your cart, go to Checkout, fill in a name + 10-digit
   phone + pickup location, place the order.
3. You'll land on the order confirmation screen with an order number
   (e.g. `ORD-000001`) and a "Cancel this order" option (only while
   status is `pending`).
4. In the Supabase dashboard → Table Editor → `orders`, confirm the row
   exists with the right subtotal, and that `stock_balances.reserved`
   went up by the quantity ordered for each item.
5. Use "Track an order" from the footer nav with the same order number
   + phone to confirm guest lookup works.
6. Try a mismatched phone number on that same order number — it should
   say "Order not found" rather than revealing anything about the real
   order (this is deliberate, see `get_order_status` in schema.sql).

## 5. Managing orders (Phase 1 has no admin screen yet)

Until the Business Manager is migrated onto this same Supabase project
and gets a real Orders screen, **use the Supabase dashboard's Table
Editor directly** as your interim admin tool:
- `orders` — update `status` (`pending` → `confirmed` → `ready` →
  `completed`) and `payment_status` as you take payment and hand over
  the order.
- `items` / `price_history` / `stock_balances` — add real products,
  set real prices, set real stock levels. Delete the sample rows from
  `schema.sql`'s seed data first.
- `pickup_locations` — replace the placeholder address with your real
  one(s).

This is manual and not meant to be permanent — it's the fastest way to
get Phase 1 live without also building an admin UI before you've
validated anyone wants to order online at all.

## What's deliberately not in Phase 1

- No delivery, no shipping cost calculation.
- No online payment (Razorpay) — pay-at-pickup only.
- No customer accounts / login — guest checkout, order lookup by
  order number + phone.
- No B2B pricing tiers (`customer_price_overrides` table exists in the
  schema but nothing writes to it yet).
- No connection to the existing Google Sheets Business Manager — this
  is a standalone Supabase project. Migrating the Business Manager
  onto the same database (so online orders and phone/counter sales
  share one system) is a separate, later step.
- Stock reservations (`stock_balances.reserved`) are held forever once
  an order is placed and never auto-released except by explicit
  cancellation — there's no "pending order timeout" job yet. Since
  there's no online payment to abandon mid-checkout in Phase 1, this is
  a smaller risk than it would be once online prepay is added, but a
  customer who places an order and never shows up will hold that stock
  reserved until someone manually cancels or completes the order.
