# PomProc — Production Architecture Blueprint & Technical Specification

**From:** Principal Full-Stack Architect
**Subject:** Migrating PomProc v3.x (single-file React prototype) → multi-tenant, cloud-backed PWA
**Scope:** Facility-management site ordering, warehouse fulfillment, equipment tracking, multi-entity billing, inventory analytics
**Covers v3.x additions:** notifications (WhatsApp/email), pre-defined **Kits**, order **payment status** & extended billing, **order amendments** (dated add-to-demande), and **inter-entity (FFA/DC/IIS) flow analytics**. Companion doc: `PomProc_Data_Engine_Spec.md` (relational schema, image pipeline, Excel import).

---

## 0. What the prototype actually is (grounding the design)

Before proposing anything, here is what v3.1 *really* does, read from the source. The blueprint below preserves this behaviour exactly — the goal is to lift it onto a real backend without changing the domain logic your users already rely on.

| Concern | Prototype reality (from `pomproc__3_.jsx`) | Implication for production |
|---|---|---|
| Persistence | Two keys in `window.storage` (shared artifact KV): `pomproc:data3` (catalog + users + templates + counters) and `pomproc:orders3` (all orders). Everything is one big JSON blob per key. | Replace with Postgres tables; the two blobs become ~10 normalized tables. |
| Auth | Plaintext passwords compared client-side (`x.password === p`). Self-registration creates `role: "user"`. | **Security-critical rewrite.** Hash with Argon2id, move comparison server-side. |
| Roles | `user` (demandeur), `magasinier`, `admin`. Warehouse views gated by `role === "magasinier" || "admin"`. | Becomes RBAC middleware. |
| Reference codes | `YYMM-XXX`, generated from `data.counters[ym]` incremented client-side. **Race-prone.** | Atomic Postgres sequence per month (see §2 `order_counters`). |
| Entities | `FFA`, `DC` are billable (`bill: true`); `IIS` is internal (`bill: false`). Order tracks `billFFA/billDC` (live) plus `billFFA0/billDC0` (initial snapshot at submit). | Deltas = live − snapshot. Preserve both. |
| Line identity | `lineKey = productId` or `productId::variant`. The `prep` object is keyed by this. | Composite natural key on `order_items`; `prep` fields fold into the same row. |
| Associations | Products carry `assoc: [ids]`; `assocClosure()` does a BFS to chain Generator→Gas Can→Hose. | Self-referencing `product_associations` + recursive CTE. |
| Variants | Category-level or product-level `variants: [{name, options[]}]` (e.g. Taille × Couleur). Matrix expanded client-side. | `JSONB variants_schema`; selected variant stored as `variant_label` on the item. |
| Photos | `prepPhotos: [dataURL]` — base64 thumbnails compressed to ~1000px/0.72 inline in the JSON. | Move to S3; store URLs in `order_photos`. This alone will shrink the orders payload by 10–100×. |
| Prep state | `prep[lineKey] = { state: ok\|missing\|replaced, qtyPrep, repl, replRef, replPrice, replChosen }`; `assetNo` lives on the line. | Columns on `order_items`. |
| Analytics | Computed live in the browser over all orders: frequency, effective qty, pairwise co-occurrence matrix, replacement/missing tallies, zero-demand ("never"), category donut, **inter-entity flow** (per-entity consumption, owner→billed matrix, top cross-entity items). | Push down into SQL aggregations / materialized views (see §4, incl. §4.4 flow). |
| Kits | `data.kits = [{id, name, items:[{productId, variant, qty, mode, entity}]}]`. Built in `AdminKits`, unpacked additively into a draft by `applyKit()`; kit items carry their **own** `entity` (unlike template lines). | `kits` + `kit_items` tables (see §2). |
| Payment | Orders carry `paymentStatus` (`non_paye`\|`paye`), toggled **admin-only** in `Billing.togglePay`. Billing aggregates by month / demandeur / **chef d'équipe** (`responsable`) / magasinier, with grand total + paid/unpaid split; Excel export adds a récap sheet + per-chef sheet + payment column. | `payment_status` enum+column; RBAC-gated PATCH; billing aggregations in §4. |
| Amendments | Both demandeur and magasinier/admin can **add items to an existing/finished demande**. Appended lines carry `addedAt`/`addedBy`/`addedRole`; originals are frozen (append-only). Adding to a `terminee` order **reopens** it (`en_cours`) for prep. Sheet groups lines by date; a blank ruled **notes page** is appended to the printout. | `added*` columns on `order_items`, `amendedAt` on `orders`; append endpoint in §3.3. |

The single most important structural change is **photos and the orders blob**. Today every order embeds its proof photos as base64 inside one JSON array that is re-read and re-written wholesale on every edit. That does not survive contact with real warehouse volume. Everything else is a clean normalization exercise.

---

## 1. System Architecture & Component Topology

### 1.1 Topology

```
                          ┌─────────────────────────────────────────────┐
                          │                  CLIENTS                      │
                          │  Demandeur (mobile PWA)  ·  Magasinier (tablet)│
                          │  Admin (desktop)  —  installable, offline-cap │
                          └───────────────┬─────────────────────────────┘
                                          │ HTTPS / WSS (Wi-Fi & 5G)
                                          ▼
                          ┌─────────────────────────────────────────────┐
                          │            EDGE / CDN (Vercel/CloudFront)     │
                          │   Static PWA shell, service worker, images    │
                          └───────────────┬─────────────────────────────┘
                                          │
             ┌────────────────────────────┼────────────────────────────┐
             ▼                            ▼                            ▼
  ┌────────────────────┐      ┌────────────────────────┐   ┌────────────────────┐
  │   Next.js App       │      │   API layer (tRPC/REST) │   │  Realtime service   │
  │   (App Router, RSC) │◄────►│   Node.js + TypeScript   │   │  (WS / Supabase RT) │
  │   PWA + Workbox      │      │   Zod validation, RBAC   │   │  order.* channels    │
  └────────────────────┘      └───────────┬────────────┘   └─────────┬──────────┘
                                          │                          │
              ┌───────────────────────────┼──────────────┬──────────┘
              ▼                           ▼              ▼
   ┌──────────────────┐      ┌────────────────────┐  ┌──────────────────────┐
   │  PostgreSQL       │      │  Object Storage     │  │  Async workers        │
   │  (Prisma/Drizzle) │      │  S3 / Cloudinary    │  │  (BullMQ / Inngest)   │
   │  RLS multi-tenant │      │  proof photos       │  │  PDF · XLSX · WhatsApp │
   │  matview analytics│      │  labels / exports   │  │  · email · webhooks   │
   └──────────────────┘      └────────────────────┘  └──────────────────────┘
```

### 1.2 Stack decisions (with the "why")

- **Frontend — Next.js (App Router) + TypeScript, PWA via `next-pwa`/Workbox.** Chosen over plain Vite because you get server components (fast first paint on 5G), route handlers (colocate the API), and streaming. Keep the *entire current UI* — it's already mobile-first and componentized (`OrderWizard`, `Magasin`, `Prepare`, `Sheet`, `Billing`, `Analytics`, `Admin`). Migrate components largely as-is; swap `window.storage` calls for data-layer hooks.
- **Styling.** The prototype ships one big CSS string. Port it to CSS Modules (least churn) or Tailwind (if you want a design-token system). Not on the critical path.
- **API — start with tRPC, expose REST/GraphQL only if third parties need it.** Because client and server are both TypeScript, tRPC gives end-to-end type safety with zero schema duplication and the fastest path from the prototype's function calls to typed procedures. The tables in §3 are written framework-neutrally (as REST) so they also serve a GraphQL/OpenAPI implementation.
- **Realtime — Supabase Realtime *or* a small Socket.IO service.** If you adopt Supabase for Postgres, its Realtime (Postgres logical replication → WS) gives you `order.created` / queue-take events for free. If you self-host Postgres, run a thin Socket.IO gateway that publishes domain events emitted by the API.
- **DB — PostgreSQL + Prisma.** Prisma for the schema/migrations/DX; drop to raw SQL (via `prisma.$queryRaw`) for the analytics aggregations in §4, which are not things an ORM should express. Drizzle is a fine alternative if you prefer SQL-first.
- **Storage — S3 (or Cloudinary if you want automatic transforms).** Photos uploaded direct-to-S3 via presigned URLs; only the resulting key is sent to the API.
- **Documents — Puppeteer worker for PDF; SheetJS (already a dependency: `import * as XLSX`) for XLSX.** The prototype already prints via CSS `@media print` and exports XLSX in-browser. Server-side rendering of the dual A4 dispatch sheet and the rotated box label (both defined in `Sheet`) makes them reproducible and emailable.

### 1.3 Multi-tenancy

The prototype is single-org. To make it multi-tenant, add a `tenants` table and a `tenant_id` FK on **every** table below. Enforce isolation with **PostgreSQL Row-Level Security** (a policy of `tenant_id = current_setting('app.tenant_id')::uuid`), and set `app.tenant_id` from the authenticated session at the start of every request/transaction. RLS is the safety net that a forgotten `WHERE tenant_id = …` can't defeat.

---

## 2. Database Schema (Prisma)

This is a faithful normalization of the two JSON blobs. Notes call out where a prototype field maps in. All tables carry `tenant_id` (omitted from prose for brevity but present in the schema).

```prisma
// schema.prisma
generator client { provider = "prisma-client-js" }
datasource db { provider = "postgresql"; url = env("DATABASE_URL") }

enum Role          { user magasinier admin }
enum Entity        { FFA IIS DC }
enum OrderStatus   { envoyee en_cours terminee }        // À préparer / En préparation / Terminée
enum PaymentStatus { non_paye paye }                     // orders.paymentStatus (admin-toggled)
enum UnitMode      { unite paquet }
enum ItemState     { ok missing replaced }               // prep.state
enum PrepItemFlag  { pending ok missing replaced }

model Tenant {
  id        String   @id @default(uuid())
  name      String
  slug      String   @unique
  createdAt DateTime @default(now())
  users     User[]
  // …all other relations
}

model User {
  id           String   @id @default(uuid())
  tenantId     String
  username     String
  passwordHash String                                   // was plaintext `password` — MUST hash (Argon2id)
  role         Role     @default(user)
  workplace    String
  email        String?
  phone        String?                                  // WhatsApp intl format, digits only (see waLink)
  createdAt    DateTime @default(now())

  tenant             Tenant     @relation(fields: [tenantId], references: [id])
  ordersRequested    Order[]    @relation("requester")
  ordersAssigned     Order[]    @relation("assignee")
  templates          Template[]

  @@unique([tenantId, username])
  @@index([tenantId, role])
}

model Category {
  id             String   @id @default(uuid())
  tenantId       String
  name           String
  displayOrder   Int                                    // was `order`
  requiresAssetNum Boolean @default(false)              // was `num` — drives the "N° matériel" field
  color          String
  icon           String
  variantsSchema Json?                                  // [{name, options[]}]  (was `variants`)

  products Product[]
  @@index([tenantId, displayOrder])
}

model Product {
  id             String   @id @default(uuid())
  tenantId       String
  categoryId     String
  ref            String                                 // supplier ref — NOT unique (prototype has dup refs)
  name           String                                 // was `nom`
  unit           String   @default("Pièce")             // was `unite`
  packageQty     Int      @default(1)                   // was `qteEmballage`
  price          Decimal  @default(0) @db.Decimal(12,4) // was `prix` (note 0.328 etc — keep 4dp)
  ownerEntity    Entity   @default(IIS)                 // was `entity`
  imageUrl       String?                                // was `image` (base64/URL) → S3 URL
  variantsSchema Json?                                  // product-level override of category variants
  active         Boolean  @default(true)

  category      Category  @relation(fields: [categoryId], references: [id])
  parentsOf     ProductAssociation[] @relation("parent")
  childrenOf    ProductAssociation[] @relation("child")
  orderItems    OrderItem[]
  templateItems TemplateItem[]

  @@index([tenantId, categoryId])
  @@index([tenantId, ref])
  @@index([tenantId, name])
}

// Self-referencing closure edge. assocClosure() BFS becomes a recursive CTE (see §3).
model ProductAssociation {
  id                  String  @id @default(uuid())
  tenantId            String
  parentProductId     String
  associatedProductId String
  parent     Product @relation("parent", fields: [parentProductId],     references: [id], onDelete: Cascade)
  associated Product @relation("child",  fields: [associatedProductId], references: [id], onDelete: Cascade)

  @@unique([parentProductId, associatedProductId])
  @@index([tenantId, parentProductId])
}

model Template {
  id            String  @id @default(uuid())
  tenantId      String
  name          String
  ownerUserId   String?                                 // was `ownerUsername`
  shared        Boolean @default(false)                 // globally visible when true
  createdAt     DateTime @default(now())

  owner User?           @relation(fields: [ownerUserId], references: [id])
  items TemplateItem[]
  @@index([tenantId, shared])
}

model TemplateItem {
  id         String   @id @default(uuid())
  templateId String
  productId  String
  qty        Int      @default(1)
  mode       UnitMode @default(unite)
  template   Template @relation(fields: [templateId], references: [id], onDelete: Cascade)
  product    Product  @relation(fields: [productId],  references: [id])
  @@index([templateId])
}

// Pre-defined bundles. Unpacked additively into a draft by applyKit(); unlike
// template items, kit items carry their own billedEntity. (See Data Engine Spec §1.)
model Kit {
  id        String    @id @default(uuid())
  tenantId  String
  name      String
  active    Boolean   @default(true)
  createdAt DateTime  @default(now())
  items     KitItem[]
  @@unique([tenantId, name])
  @@index([tenantId, active])
}

model KitItem {
  id           String   @id @default(uuid())
  kitId        String
  productId    String
  variantLabel String?                                  // it.variant ("" → null)
  qty          Int      @default(1)
  mode         UnitMode @default(unite)
  billedEntity Entity   @default(IIS)                   // kit items DO carry entity
  position     Int      @default(0)
  kit     Kit     @relation(fields: [kitId],     references: [id], onDelete: Cascade)
  product Product @relation(fields: [productId], references: [id])
  @@unique([kitId, productId, variantLabel])
  @@index([kitId])
}

model Order {
  id               String      @id @default(uuid())
  tenantId         String
  referenceCode    String                               // YYMM-XXX  (see OrderCounter)
  requesterId      String
  siteName         String                               // was `nom`
  siteNumber       String                               // was `numero`
  siteSupervisor   String?                              // was `responsable`
  departDate       DateTime?                            // was `departDate`
  departTime       String?                              // free-text "HH:mm" in prototype
  templateName     String?
  status           OrderStatus  @default(envoyee)
  paymentStatus    PaymentStatus @default(non_paye)     // admin-toggled in Billing
  amendedAt        DateTime?                             // last time items were appended (amend flow)
  assignedUserId   String?                              // was `assignedTo` (magasinier)
  completedBy      String?                              // free-text name at sign-off
  completedAt      DateTime?
  notes            String?                              // was `magasinierNote`

  totalAmount      Decimal @default(0) @db.Decimal(12,2)
  billFfa          Decimal @default(0) @db.Decimal(12,2) // live
  billDc           Decimal @default(0) @db.Decimal(12,2)
  initialBillFfa   Decimal @default(0) @db.Decimal(12,2) // snapshot at submit (was billFFA0)
  initialBillDc    Decimal @default(0) @db.Decimal(12,2) // was billDC0
  createdAt        DateTime @default(now())

  requester User        @relation("requester", fields: [requesterId],   references: [id])
  assignee  User?       @relation("assignee",  fields: [assignedUserId], references: [id])
  items     OrderItem[]
  photos    OrderPhoto[]

  @@unique([tenantId, referenceCode])
  @@index([tenantId, status, createdAt])                // powers Magasin queue filters
  @@index([tenantId, requesterId])
  @@index([tenantId, paymentStatus])                    // powers Billing paid/unpaid filters
}

model OrderItem {
  id             String     @id @default(uuid())
  orderId        String
  productId      String                                 // for non-catalog extras, see note
  variantLabel   String?                                // the ::variant part of lineKey
  requestedQty   Int
  deliveredQty   Int?                                   // prep.qtyPrep (null = untreated → falls back to requested)
  unitMode       UnitMode   @default(unite)
  billedEntity   Entity                                 // per-line (was `entity`)
  assetNumber    String?                                // was line.assetNo (only when category.requiresAssetNum)
  state          PrepItemFlag @default(pending)         // prep.state
  replacementName  String?                              // prep.repl
  replacementRef   String?                              // prep.replRef
  replacementPrice Decimal? @db.Decimal(12,4)           // prep.replPrice
  // amendment provenance (append-only add-to-demande) — null on original lines
  addedAt          DateTime?                             // line.addedAt (batch date on the sheet)
  addedByUserId    String?                              // line.addedBy → user
  addedByRole      Role?                                // line.addedRole
  // Non-catalog "extras": snapshot name/ref/price/unit so the line survives catalog edits.
  extraName      String?
  extraRef       String?
  extraPrice     Decimal? @db.Decimal(12,4)
  extraUnit      String?

  order   Order   @relation(fields: [orderId],   references: [id], onDelete: Cascade)
  product Product @relation(fields: [productId], references: [id])

  @@unique([orderId, productId, variantLabel])          // == lineKey uniqueness
  @@index([orderId])
  @@index([productId, state])                           // powers shortage/substitution analytics
}

model OrderPhoto {
  id        String   @id @default(uuid())
  orderId   String
  url       String                                      // S3 key/URL (was inline base64 in prepPhotos)
  width     Int?
  height    Int?
  uploadedAt DateTime @default(now())
  order     Order    @relation(fields: [orderId], references: [id], onDelete: Cascade)
  @@index([orderId])
}

// Atomic monthly sequence. Replaces client-side data.counters[ym]++.
model OrderCounter {
  tenantId  String
  monthKey  String                                      // "YYMM"
  lastValue Int    @default(0)
  @@id([tenantId, monthKey])
}
```

**Reference-code generation (atomic).** The prototype's `counters[ym] + 1` in the browser will collide the moment two people submit in the same month. Do it in one round-trip inside the submit transaction:

```sql
INSERT INTO "OrderCounter" ("tenantId","monthKey","lastValue")
VALUES ($1, $2, 1)
ON CONFLICT ("tenantId","monthKey")
DO UPDATE SET "lastValue" = "OrderCounter"."lastValue" + 1
RETURNING "lastValue";
-- reference_code := monthKey || '-' || lpad(lastValue::text, 3, '0')
```

**Non-catalog extras.** The prototype merges `draft.extras` into the product map at runtime. In SQL, keep them as `OrderItem` rows with a null-ish `productId` sentinel (or a nullable FK + the `extra*` snapshot columns above). Snapshotting name/price protects historical orders from later catalog changes — a correctness win the prototype doesn't have.

---

## 3. API Routing Table & Payloads

REST signatures (map 1:1 to tRPC procedures). All responses are tenant-scoped by RLS; all mutating routes pass through `authGuard` then `roleGuard(...)`.

### 3.1 Auth & RBAC

| Method | Path | Role | Body → Response |
|---|---|---|---|
| POST | `/api/auth/register` | public | `{username,password,workplace,email,phone}` → `201 {user}` (role forced to `user`) |
| POST | `/api/auth/login` | public | `{username,password}` → `{accessToken, refreshToken, user}` |
| POST | `/api/auth/refresh` | public | `{refreshToken}` → `{accessToken}` |
| POST | `/api/auth/logout` | auth | `—` → `204` |
| GET | `/api/me` | auth | → `{id,username,role,workplace,…}` |

Middleware:

```ts
// authGuard: verify short-lived JWT (or session cookie), load user, SET app.tenant_id
// roleGuard(...allowed): 403 unless req.user.role ∈ allowed
router.use('/api/warehouse', authGuard, roleGuard('magasinier','admin'));
router.use('/api/admin',     authGuard, roleGuard('admin'));
```

Passwords: **Argon2id**, verify server-side. Tokens: 15-min access JWT + rotating refresh token (httpOnly cookie). This replaces the prototype's plaintext client-side check entirely.

### 3.2 Catalog & associations (read-heavy, cache aggressively)

| Method | Path | Role | Purpose |
|---|---|---|---|
| GET | `/api/categories` | auth | ordered by `displayOrder`; includes `variantsSchema` |
| GET | `/api/products?catId=&q=&freq=1` | auth | search/filter; `freq=1` returns top-24 by historical volume (the wizard's "Fréquents" chip) |
| GET | `/api/products/:id/associations` | auth | **closure** of associated products |
| POST/PUT/DELETE | `/api/admin/products…` | admin | CRUD + Excel import/export |
| POST/PUT/DELETE | `/api/admin/categories…` | admin | CRUD |
| GET | `/api/kits` | auth | active kits + items (wizard "🎁 Kits" chip) |
| POST/PUT/DELETE | `/api/admin/kits…` | admin | Kit builder CRUD (name + items with qty/mode/variant/entity) |

The association closure (prototype's `assocClosure` BFS) as a recursive CTE:

```sql
WITH RECURSIVE chain AS (
  SELECT associated_product_id AS pid
  FROM product_associations WHERE parent_product_id = $1
  UNION
  SELECT pa.associated_product_id
  FROM product_associations pa
  JOIN chain c ON pa.parent_product_id = c.pid
)
SELECT * FROM products WHERE id IN (SELECT pid FROM chain);
```

### 3.3 Wizard & order submission (atomic)

```
POST /api/orders
Body: {
  siteName, siteNumber, siteSupervisor, departDate, departTime, templateName,
  lines: [{ productId, variantLabel?, qty, mode: "unite"|"paquet", billedEntity }],
  extras: [{ name, ref, price, unit, qty, mode, billedEntity }]
}
→ 201 { order }   // with generated referenceCode + snapshotted bills
```

Server does, in **one transaction**:
1. Bump `OrderCounter` → `referenceCode`.
2. (Optional) resolve association closure for each line to *suggest* additions — but only persist what the client sent; the prompt happens in the UI.
3. Compute per-line cost: `qty × (mode==='paquet' ? packageQty : 1) × price` (mirrors `lineTotal`).
4. Sum into `billFfa`/`billDc` by `billedEntity` (IIS is non-billing), snapshot to `initialBill*`.
5. Insert order + items.
6. Emit `order.created` on the realtime channel.

Templates, past orders, and **kits** are all client-side *starting points*: the wizard expands a kit's items (with their per-item entity) into the draft via `applyKit()` and submits the resulting `lines[]` — the server persists lines, not a kit reference.

**Amending an order (append-only add-to-demande).** The prototype's `addToOrder()` appends items — including to a *finished* demande — stamping each with date, author, and role, and reopening a completed order for prep.

```
POST /api/orders/:id/items          role: requester (own order) | magasinier | admin
Body: { lines: [...], extras: [...] }          // same shape as submission
→ 200 { order }
```
Server, in one transaction: stamp each new item `addedAt=now`, `addedByUserId`, `addedByRole`; **never touch existing items** (audit-trail integrity); add their cost into `totalAmount` and both `bill*`/`initialBill*`; set `amendedAt=now`; if `status='terminee'` flip to `en_cours` (there is new work to prepare); emit `order.updated`. The dispatch sheet renders originals under the creation date and each later batch under its `addedAt` date; the demandeur or a magasinier/admin can trigger this from the order sheet / prep screen.

### 3.4 Warehouse queue & preparation (realtime + locking)

| Method | Path | Role | Purpose |
|---|---|---|---|
| GET | `/api/warehouse/queue?range=today\|week\|month\|all` | mag/admin | board rows; uses `@@index(tenantId,status,createdAt)` |
| POST | `/api/warehouse/orders/:id/take` | mag/admin | claim: sets `status=en_cours`, `assignedUserId` **iff currently `envoyee` and unassigned** (optimistic lock) |
| POST | `/api/warehouse/orders/:id/release` | mag/admin | give back control |
| PATCH | `/api/warehouse/items/:itemId` | mag/admin | set `{state, deliveredQty, replacementName, replacementRef, replacementPrice, assetNumber, billedEntity}`; server recomputes bills (mirrors `recalc`/`eff`) |
| POST | `/api/warehouse/orders/:id/photos` | mag/admin | returns presigned S3 URL; client uploads direct, then confirms → `OrderPhoto` row |
| POST | `/api/warehouse/orders/:id/complete` | mag/admin | `{completedBy, notes}` → `status=terminee`, `completedAt=now`; locks order for non-admins |

**Take/lock** (atomic claim — the multi-magasinier safety the prototype lacks):

```sql
UPDATE orders SET status='en_cours', assigned_user_id=$user
WHERE id=$id AND status='envoyee' AND assigned_user_id IS NULL
RETURNING *;                       -- 0 rows ⇒ 409 Conflict "already taken"
```

**Effective cost** on each item PATCH (server-side port of `eff()`):
```
factor   = unitMode === 'paquet' ? packageQty : 1
qty      = state === 'missing' ? 0 : (deliveredQty ?? requestedQty)
price    = (state === 'replaced' && replacementPrice != null) ? replacementPrice : product.price
lineCost = qty × factor × price
```
Recompute `billFfa/billDc/totalAmount`, persist, emit `order.updated`.

Realtime channels: `tenant:{id}:orders` with events `order.created`, `order.taken`, `order.updated`, `order.completed`. The Magasin board subscribes and live-reorders; a demandeur subscribes to their own order for status.

### 3.5 Documents & export

| Method | Path | Role | Purpose |
|---|---|---|---|
| GET | `/api/orders/:id/sheet.pdf` | mag/admin | dual A4 dispatch sheet (Admin copy + return copy, signature blocks) via Puppeteer; items grouped by date with an **"Ajouts après création"** block, plus a trailing blank ruled **notes page** for handwriting |
| GET | `/api/orders/:id/label.pdf` | mag/admin | transport box tag, rotated 90° (CSS `transform: rotate(90deg)` in print layout) |
| GET | `/api/exports/orders.xlsx?range=` | mag/admin | SheetJS workbook (same columns as prototype's `rows.push([...])`) |
| GET | `/api/exports/billing.xlsx?month=` | admin | récap sheet (grand total, FFA/DC, paid/unpaid) + one sheet each **par mois / par demandeur / par chef d'équipe / par magasinier** + détail (incl. `paymentStatus`) + modifications |
| GET | `/api/exports/analytics.xlsx?range=` | mag/admin | top products, category share, co-occurrence, replacements/missing, never-ordered, **+ inter-entity flow sheets** (per-entity, owner→billed matrix, top cross-entity) |

PDF/XLSX run as **async jobs** for large ranges (enqueue → poll/`204 + Location` → download from S3). Small single-order PDFs render inline.

### 3.6 Notifications

| Method | Path | Role | Purpose |
|---|---|---|---|
| GET | `/api/orders/:id/notify/whatsapp` | auth | returns a `wa.me/<digits>?text=<encoded>` deep link (mirrors `waLink`+`msgDone`/`msgNewDemande`) |
| GET | `/api/orders/:id/notify/mailto` | auth | returns `mailto:` link, or |
| POST | `/api/orders/:id/notify/email` | mag/admin | server-sent email (SendGrid/SES) to supervisor on completion, with change summary |

Message bodies reuse the prototype's exact French templates (`msgNewDemande`, `msgDone`, `changeSummary`) so field staff see no difference. WhatsApp stays a deep link (no Business API cost) unless you later want delivery receipts.

### 3.7 Billing & payment

| Method | Path | Role | Purpose |
|---|---|---|---|
| GET | `/api/billing?month=&groupBy=month\|demandeur\|chef\|magasinier` | admin | aggregated FFA/DC totals + counts + Δ (live−snapshot) + paid/unpaid per group; grand total in the envelope |
| PATCH | `/api/orders/:id/payment` | **admin** | `{paymentStatus: "paye"\|"non_paye"}` → toggles `orders.payment_status` (mirrors `Billing.togglePay`, admin-only) |

Non-admins see payment as read-only. `chef` groups by `siteSupervisor` (the demande's `responsable`); `magasinier` restricts to `terminee` orders grouped by `completedBy`. IIS lines are never billed.

---

## 4. Analytics Engine

Port the browser aggregations to SQL. For dashboards that tolerate slight staleness, back each with a **materialized view** refreshed on a schedule (or `REFRESH … CONCURRENTLY` after order completion).

### 4.1 Pareto / category usage (from the `freq`,`qty`,`catAgg` tallies)

```sql
SELECT p.id, p.name, p.ref, c.name AS category,
       COUNT(oi.id)                                   AS order_line_count,   -- freq
       SUM(oi.requested_qty *
           CASE WHEN oi.unit_mode='paquet' THEN p.package_qty ELSE 1 END) AS total_units, -- qty
       SUM(SUM(oi.requested_qty *                                              -- cumulative %
           CASE WHEN oi.unit_mode='paquet' THEN p.package_qty ELSE 1 END))
         OVER (ORDER BY SUM(oi.requested_qty) DESC)
       / NULLIF(SUM(SUM(oi.requested_qty)) OVER (),0) AS cumulative_share
FROM order_items oi
JOIN products p ON p.id = oi.product_id
JOIN categories c ON c.id = p.category_id
JOIN orders o ON o.id = oi.order_id
WHERE o.tenant_id = $1 AND o.created_at >= $2
GROUP BY p.id, c.name
ORDER BY total_units DESC;
```
The window function gives the cumulative-share column the prototype draws as the right-hand "effet Pareto" bar. Category donut = the same grouped by `c.id`.

### 4.2 Co-occurrence / basket analysis (the pairwise matrix)

The prototype counts, per order, every unordered pair of distinct products. In SQL, self-join the distinct items of each order:

```sql
SELECT a.product_id AS a, b.product_id AS b, COUNT(*) AS pair_count
FROM (SELECT DISTINCT order_id, product_id FROM order_items) a
JOIN (SELECT DISTINCT order_id, product_id FROM order_items) b
  ON a.order_id = b.order_id AND a.product_id < b.product_id   -- unordered, no self-pairs
JOIN orders o ON o.id = a.order_id
WHERE o.tenant_id = $1
GROUP BY a.product_id, b.product_id
ORDER BY pair_count DESC
LIMIT 50;
```
Top pairs feed the heatmap (restrict `a,b` to the top-N most frequent products for a readable grid) and drive **auto-template suggestions**: a pair with high `pair_count` and high *confidence* (`pair_count / support(a)`) is a candidate bundle. Add lift if you want proper market-basket ranking:
`lift = P(a,b) / (P(a)·P(b))`.

### 4.3 Substitution & shortage tracking (from `prep` states)

```sql
SELECT p.id, p.name, p.ref,
       COUNT(*) FILTER (WHERE oi.state='missing')  AS times_missing,
       COUNT(*) FILTER (WHERE oi.state='replaced') AS times_replaced,
       MODE() WITHIN GROUP (ORDER BY oi.replacement_name)
         FILTER (WHERE oi.state='replaced')        AS most_common_substitute
FROM order_items oi
JOIN products p ON p.id = oi.product_id
JOIN orders o ON o.id = oi.order_id
WHERE o.tenant_id = $1
GROUP BY p.id
HAVING COUNT(*) FILTER (WHERE oi.state IN ('missing','replaced')) > 0
ORDER BY times_missing DESC, times_replaced DESC;
```

**Zero-demand ("never" / catalog pruning):**
```sql
SELECT p.* FROM products p
WHERE p.tenant_id = $1 AND p.active
  AND NOT EXISTS (SELECT 1 FROM order_items oi WHERE oi.product_id = p.id);
```

**Billing deltas** (the FFA/DC reconciliation): `delta_ffa = bill_ffa − initial_bill_ffa` per order; the substitutions and out-of-stock captured above are the *reasons* for each delta. Roll up per month/entity for the Billing view.

### 4.4 Inter-entity flow (FFA / DC / IIS movement)

Each line has two entities: the product's `ownerEntity` (who owns the stock) and the line's `billedEntity` (who it's charged to / bought from — mutable during prep). Analysing one against the other surfaces cross-entity logistics: IIS sites consuming items sourced from FFA or DC. Cost per line mirrors `lineTotal`: `requested_qty × (unit_mode='paquet' ? package_qty : 1) × price`.

**Consumption per (billed) entity** — the per-entity bars:
```sql
SELECT oi.billed_entity AS entity,
       COUNT(*)                                       AS line_count,
       SUM(oi.requested_qty *
           CASE WHEN oi.unit_mode='paquet' THEN COALESCE(p.package_qty,1) ELSE 1 END) AS units,
       SUM(oi.requested_qty *
           CASE WHEN oi.unit_mode='paquet' THEN COALESCE(p.package_qty,1) ELSE 1 END * p.price) AS cost
FROM order_items oi
JOIN orders o   ON o.id = oi.order_id
LEFT JOIN products p ON p.id = oi.product_id
WHERE o.tenant_id = $1 AND o.created_at >= $2
GROUP BY oi.billed_entity;
```

**Owner → billed flow matrix** ("who takes the most from whom" — off-diagonal cells are transfers):
```sql
SELECT COALESCE(p.owner_entity,'IIS') AS owner, oi.billed_entity AS billed,
       SUM(oi.requested_qty *
           CASE WHEN oi.unit_mode='paquet' THEN COALESCE(p.package_qty,1) ELSE 1 END * p.price) AS cost
FROM order_items oi
JOIN orders o ON o.id = oi.order_id
LEFT JOIN products p ON p.id = oi.product_id
WHERE o.tenant_id = $1
GROUP BY owner, oi.billed_entity;
```

**Top cross-entity items** (sourced outside IIS): the same aggregation filtered `WHERE oi.billed_entity <> 'IIS'`, grouped by product, ordered by cost — feeds the "Top articles transférés" chart and the *Top inter-entités* Excel sheet. All three power the analytics dashboard's flux section and its export.

---

## 5. PWA Caching & Offline Strategy

Field supervisors on flaky 5G must be able to *browse the catalog and draft an order* offline, then sync. Warehouse actions require connectivity (they mutate shared state) but should degrade gracefully.

**Service worker (Workbox) routing:**
- **App shell / JS / CSS** — precache, cache-first, versioned by build hash.
- **Catalog (`/api/categories`, `/api/products`)** — stale-while-revalidate. This is the big one: the whole catalog is a few hundred rows; cache it so the wizard opens instantly and works offline.
- **Product images** — cache-first with a capped `ExpirationPlugin` (e.g. 200 entries / 30 days).
- **Order reads** — network-first, fall back to cache (supervisor can review their recent orders offline).
- **Order writes** — network-only; if offline, enqueue (see below).

**Offline order drafting & submission:**
- Draft state lives in **IndexedDB** (via `idb`), not the ephemeral React state the prototype uses — so a dropped connection or app close never loses a half-built order.
- On submit while offline, use **Workbox Background Sync** (a `Queue`) to hold the `POST /api/orders`. When connectivity returns, the SW replays it. The **reference code is assigned by the server on replay** (never offline — that's what guarantees uniqueness), and the client shows a "pending sync" state until it gets the real `referenceCode` back.
- Idempotency: each queued submit carries a client-generated `Idempotency-Key` (UUID) so a double-replay can't create two orders.

**Conflict handling:** orders are append-mostly; the only real contention is *take/complete*, which is server-authoritative (the atomic `UPDATE … WHERE status='envoyee'`). A magasinier who acted on a stale board simply gets a `409` and the board refreshes over realtime.

**Installability:** Web App Manifest (name, icons, `display: standalone`, theme color matching the brand). Prompt install after first successful order. `capture="environment"` on the photo input (already in the prototype) keeps the direct-camera flow.

---

## 6. Data Migration Plan (localStorage/JSON → PostgreSQL)

The current data lives in two `window.storage` keys. Migration is a one-shot ETL script; run it against a copy first.

**Step 0 — Extract.** Export both blobs to files: `data3.json` (`{users, productCats, products, templates, counters}`) and `orders3.json` (`[order, …]`). In the browser console: `await window.storage.get('pomproc:data3', true)` etc., or add a temporary "Export all" admin button.

**Step 1 — Provision.** Create the tenant row, run `prisma migrate deploy` to build the schema empty.

**Step 2 — Reference data (order matters — respect FKs):**
1. `categories` ← `productCats` (`order→displayOrder`, `num→requiresAssetNum`, `variants→variantsSchema`). Keep old ids in a `legacy_id → uuid` map.
2. `products` ← `products` (`nom→name`, `prix→price`, `qteEmballage→packageQty`, `entity→ownerEntity`, `image→imageUrl`). Map `catId` via the category id-map. **Do not** trust `ref` as unique — the seed has duplicates; keep it as a plain indexed column.
3. `product_associations` ← for each product with `assoc: [ids]`, insert one edge per associated id (parent = product, associated = id), remapping both through the product id-map.
4. `users` ← `users`. **Passwords cannot be carried over as-is** (they were plaintext). Two options: (a) set a random hash and force a reset email on first login, or (b) hash the existing plaintext once during migration and require a reset at leisure. Prefer (a).
5. `templates` + `template_items` ← `templates` (each `lines[]` entry → a `TemplateItem`; `ownerUsername` → `ownerUserId` via user map).
6. `kits` + `kit_items` ← `data.kits` (each `items[]` entry → a `KitItem`, **preserving its own `entity`** → `billedEntity`; remap `productId` via the product id-map; keep author order in `position`).

**Step 3 — Counters.** For each `counters[ym]`, insert an `OrderCounter{monthKey: ym, lastValue}`. This preserves the sequence so new orders continue from the right number instead of colliding with historical refs.

**Step 4 — Orders + items + photos** (per order):
1. Insert `Order` (`nom→siteName`, `numero→siteNumber`, `responsable→siteSupervisor`, `billFFA0→initialBillFfa`, etc.; `assignedTo`/`completedBy` are free-text names — match to `userId` where possible, else keep the string in `completedBy`). Carry `paymentStatus` (default `non_paye` if absent) and `amendedAt` if present.
   - For each `line`, also carry the amendment provenance if present: `addedAt`/`addedBy`(→`addedByUserId`)/`addedRole`(→`addedByRole`). Absent ⇒ original line (null).
2. For each `line`, insert an `OrderItem`. **Fold in the matching `prep[lineKey]`**: `state`, `qtyPrep→deliveredQty`, `repl→replacementName`, `replRef→replacementRef`, `replPrice→replacementPrice`; `assetNo→assetNumber`. `lineKey = productId` or `productId::variant` → split into `productId` + `variantLabel`.
3. For each `extras` entry, insert an `OrderItem` using the `extra*` snapshot columns (no product FK).
4. For each `prepPhotos` dataURL: decode base64 → upload to S3 → insert `OrderPhoto{url}`. This is where the payload shrinks dramatically.

**Step 5 — Verify.** Reconcile counts (orders in = orders out, sum of `line+extras` = `order_items`), and re-run the §4 analytics queries; the top-products / co-occurrence / missing-replaced numbers must match the prototype's live dashboard for the same date range. That equality is your acceptance test.

**Step 6 — Cutover.** Freeze writes on the prototype, run the ETL once more for the delta, flip DNS to the Next.js app. Keep the exported JSON as an immutable backup.

---

## 7. Suggested build order (de-risked)

1. Schema + migrations + seed the existing catalog; RLS policies.
2. Auth (hashing, tokens, guards) — nothing else is safe until this exists.
3. Read paths: catalog, queue, order detail — port `OrderWizard` and `Magasin` to the API.
4. Write paths: atomic submit (counters), take/lock, item PATCH with server-side `eff`/`recalc`.
5. Photos → S3; realtime channels.
6. Documents (PDF/XLSX) + notifications.
7. Analytics matviews + Billing.
8. PWA/offline + install.
9. Data migration + cutover.

Each step is shippable and testable on its own; the app stays usable throughout because the UI already exists.
