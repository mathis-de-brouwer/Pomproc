# PomProc — Database Engine, Image Pipeline & Excel Data Engine

**Companion to:** `PomProc_Architecture_Blueprint.md`
**Scope of this document:** (1) fully normalized relational schema, (2) end-to-end image asset architecture, (3) Excel batch import/sync engine, (4) ready-to-build sample import templates.
**Grounding:** every field below is traced to `pomproc_2_.jsx` (the v3.1 prototype). Where this spec departs from the earlier blueprint, it is called out explicitly.

---

## 0. What changed after re-reading the source

Several things in the running prototype are **not** in the earlier blueprint and materially affect the schema. They are corrected here. *(The last two rows — order amendments and inter-entity flow — were added in the v3.x feature round after this document's first draft.)*

| Finding in `pomproc_2_.jsx` | Blueprint status | Decision in this spec |
|---|---|---|
| **Kits are a first-class feature.** `data.kits = [{id, name, items:[{productId, variant, qty, mode, entity}]}]`, managed in `AdminKits`, applied additively via `applyKit()`. Kit items carry their own `entity` (unlike template lines). | Omitted entirely. | Add `kits` + `kit_items` as normalized tables. |
| **Orders carry `paymentStatus`** (`"non_paye"`/`"paye"`, toggled by admin in `Billing.togglePay`). | Missing. | Add `payment_status` enum + column on `orders`. |
| **`ref` is NOT unique.** The seed has genuine duplicates — e.g. `710522` (p1 & p18), `270422` (p29 & p36), `140800` (p30 & p38). Matching an import purely on `ref` is ambiguous. | Noted "keep ref as plain indexed column." | Introduce a **`sku`** column: a stable, tenant-unique business key that becomes the primary Excel match target. `ref` stays as the (non-unique) supplier reference. |
| **Orders can be amended after creation (append-only).** Both the demandeur and a magasinier/admin can add items to an existing — even *finished* — demande. Each added line carries `addedAt`/`addedBy`/`addedRole`; original lines are never mutated. Adding to a `terminee` order **reopens** it to `en_cours` so the additions get prepared. The dispatch sheet groups items by date ("5 le 14/09, +2 le 16/09"). | Absent. | Add `addedAt`/`addedByUserId`/`addedByRole` to `order_items` (null ⇒ original line) + `amendedAt` on `orders`. The amend endpoint is server-authoritative for the reopen + bill recompute. |
| **Inter-entity logistics analytics (FFA/DC/IIS flow).** Each line's `billedEntity` (who it's charged to / bought from) is analysed against the product's `ownerEntity` to show cross-entity movement — top transferred items, consumption per entity, an owner→billed flow matrix ("who takes the most from whom"). | Absent. | No new columns (both entities already stored); add the §4 flow aggregations + Excel sheets. |

Two further design choices this document makes:

- **`order_prep_logs` is an append-only event log**, not a rename of `order_items`. The *current* prep state still lives as columns on `order_items` (fast reads for the dispatch sheet and analytics). The log records *every* prep action (who set what, when, old→new) so the warehouse has an auditable trail — this is what turns a `prep` object mutation into accountability.
- **`templates` gets a `template_items` child table.** The requested table list included `kit_items` but not `template_items`; storing kit lines relationally while leaving template lines as a blob would be inconsistent, so both are normalized the same way. `template_items` is flagged below as an addition to the requested ten.

---

# 1. Relational Database Schema

## 1.1 Enumerations

```prisma
enum Role          { user magasinier admin }
enum Entity        { FFA IIS DC }                 // FFA/DC bill=true, IIS bill=false
enum OrderStatus   { envoyee en_cours terminee }  // À préparer / En préparation / Terminée
enum PaymentStatus { non_paye paye }              // orders.paymentStatus
enum UnitMode      { unite paquet }               // line.mode
enum PrepState     { pending ok missing replaced }// prep.state (+ pending = untreated)
enum PrepAction    { state_change qty_change replace asset_set photo_add note }
enum PhotoKind     { proof damage signature other }
```

## 1.2 Prisma models (the requested ten + required companions)

All tables carry `tenantId` for multi-tenant RLS (see blueprint §1.3). Money is `Decimal(12,4)` at the line/price level (the seed has 4-dp prices like `0.328`) and `Decimal(12,2)` at order roll-ups.

```prisma
generator client { provider = "prisma-client-js" }
datasource db    { provider = "postgresql"; url = env("DATABASE_URL") }

// ─────────────────────────────── users ───────────────────────────────
model User {
  id           String   @id @default(uuid())
  tenantId     String
  username     String
  passwordHash String                          // was plaintext `password` → Argon2id
  role         Role     @default(user)
  workplace    String
  email        String?
  phone        String?                          // WhatsApp intl, digits only
  active       Boolean  @default(true)
  createdAt    DateTime @default(now())

  tenant          Tenant       @relation(fields: [tenantId], references: [id], onDelete: Cascade)
  ordersRequested Order[]      @relation("requester")
  ordersAssigned  Order[]      @relation("assignee")
  templates       Template[]
  prepLogs        OrderPrepLog[]

  @@unique([tenantId, username])
  @@index([tenantId, role])
}

// ────────────────────────────── categories ───────────────────────────
model Category {
  id               String   @id @default(uuid())
  tenantId         String
  name             String
  displayOrder     Int                          // was `order`
  requiresAssetNum Boolean  @default(false)      // was `num` → drives "N° matériel"
  color            String
  icon             String
  variantsSchema   Json?                         // [{name, options[]}]  (was `variants`)
  active           Boolean  @default(true)

  products Product[]

  @@unique([tenantId, name])                     // category names are the Excel match key
  @@index([tenantId, displayOrder])
}

// ─────────────────────────────── products ────────────────────────────
model Product {
  id             String   @id @default(uuid())
  tenantId       String
  categoryId     String
  sku            String                          // NEW stable unique business key (import target)
  ref            String                          // supplier ref — NOT unique (dup refs in seed)
  name           String                          // was `nom`
  unit           String   @default("Pièce")      // was `unite`
  packageQty     Int      @default(1)            // was `qteEmballage`
  price          Decimal  @default(0) @db.Decimal(12,4) // was `prix`
  ownerEntity    Entity   @default(IIS)          // was `entity`
  imageUrl       String?                          // was `image` (base64/URL) → CDN URL
  imageKey       String?                          // object-store key (source of truth)
  variantsSchema Json?                            // product-level override of category variants
  isCustom       Boolean  @default(false)         // was extras `custom:true`
  active         Boolean  @default(true)
  createdAt      DateTime @default(now())
  updatedAt      DateTime @updatedAt

  category      Category             @relation(fields: [categoryId], references: [id], onDelete: Restrict)
  parentsOf     ProductAssociation[] @relation("parent")
  childrenOf    ProductAssociation[] @relation("child")
  orderItems    OrderItem[]
  templateItems TemplateItem[]
  kitItems      KitItem[]

  @@unique([tenantId, sku])                       // deterministic import matching
  @@index([tenantId, categoryId])
  @@index([tenantId, ref])                        // non-unique, still searched
  @@index([tenantId, name])
}

// self-referencing closure edge — assocClosure() BFS → recursive CTE
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

// ──────────────────────── kits + kit_items (NEW) ──────────────────────
model Kit {
  id        String    @id @default(uuid())
  tenantId  String
  name      String
  active    Boolean   @default(true)
  createdAt DateTime  @default(now())
  items     KitItem[]

  @@unique([tenantId, name])                      // kit name is the Excel group key
  @@index([tenantId, active])
}

model KitItem {
  id           String   @id @default(uuid())
  kitId        String
  productId    String
  variantLabel String?                            // it.variant ("" → null)
  qty          Int      @default(1)
  mode         UnitMode @default(unite)
  billedEntity Entity   @default(IIS)             // kit items DO carry entity (unlike templates)
  position     Int      @default(0)               // preserves author ordering

  kit     Kit     @relation(fields: [kitId],     references: [id], onDelete: Cascade)
  product Product @relation(fields: [productId], references: [id], onDelete: Restrict)

  @@unique([kitId, productId, variantLabel])
  @@index([kitId])
}

// ──────────────────── templates + template_items ──────────────────────
// (template_items is an addition to the requested ten — see §0)
model Template {
  id          String         @id @default(uuid())
  tenantId    String
  name        String
  ownerUserId String?                             // was `ownerUsername`
  shared      Boolean        @default(false)
  createdAt   DateTime       @default(now())

  owner User?           @relation(fields: [ownerUserId], references: [id], onDelete: SetNull)
  items TemplateItem[]

  @@index([tenantId, shared])
}

model TemplateItem {
  id           String   @id @default(uuid())
  templateId   String
  productId    String
  variantLabel String?
  qty          Int      @default(1)
  mode         UnitMode @default(unite)           // template lines have NO entity (resolved on apply)
  position     Int      @default(0)

  template Template @relation(fields: [templateId], references: [id], onDelete: Cascade)
  product  Product  @relation(fields: [productId],  references: [id], onDelete: Restrict)

  @@unique([templateId, productId, variantLabel])
  @@index([templateId])
}

// ─────────────────────────────── orders ──────────────────────────────
model Order {
  id             String        @id @default(uuid())
  tenantId       String
  referenceCode  String                            // YYMM-XXX (see OrderCounter)
  requesterId    String
  siteName       String                            // was `nom`
  siteNumber     String                            // was `numero`
  siteSupervisor String?                           // was `responsable`
  departDate     DateTime?
  departTime     String?                           // free-text "HH:mm"
  templateName   String?
  status         OrderStatus   @default(envoyee)
  paymentStatus  PaymentStatus @default(non_paye)  // NEW (was paymentStatus)
  assignedUserId String?                           // was `assignedTo`
  completedBy    String?                           // free-text sign-off name
  completedAt    DateTime?
  note           String?                           // was `magasinierNote`

  totalAmount    Decimal @default(0) @db.Decimal(12,2)
  billFfa        Decimal @default(0) @db.Decimal(12,2)  // live
  billDc         Decimal @default(0) @db.Decimal(12,2)
  initialBillFfa Decimal @default(0) @db.Decimal(12,2)  // snapshot at submit (billFFA0)
  initialBillDc  Decimal @default(0) @db.Decimal(12,2)  // billDC0
  createdAt      DateTime @default(now())
  amendedAt      DateTime?                              // last time items were appended (see amend flow)

  requester User          @relation("requester", fields: [requesterId],   references: [id], onDelete: Restrict)
  assignee  User?         @relation("assignee",  fields: [assignedUserId], references: [id], onDelete: SetNull)
  items     OrderItem[]
  photos    OrderPhoto[]
  prepLogs  OrderPrepLog[]

  @@unique([tenantId, referenceCode])
  @@index([tenantId, status, createdAt])           // powers Magasin queue
  @@index([tenantId, requesterId])
  @@index([tenantId, paymentStatus])               // powers Billing filters
}

// ───────────────────────────── order_items ───────────────────────────
model OrderItem {
  id             String    @id @default(uuid())
  orderId        String
  productId      String?                           // null for non-catalog "extras"
  variantLabel   String?                           // the ::variant part of lineKey
  requestedQty   Int
  deliveredQty   Int?                              // prep.qtyPrep (null → falls back to requested)
  unitMode       UnitMode  @default(unite)
  billedEntity   Entity                            // per-line, mutable during prep
  assetNumber    String?                           // line.assetNo (when category.requiresAssetNum)

  // amendment provenance (append-only additions) — null on original lines
  addedAt        DateTime?                          // line.addedAt (batch date shown on the sheet)
  addedByUserId  String?                            // line.addedBy resolved to a user
  addedByRole    Role?                              // line.addedRole (demandeur or magasinier/admin)

  // current prep state (mirrors prep[lineKey]) — history lives in OrderPrepLog
  prepState        PrepState @default(pending)
  replacementName  String?                         // prep.repl
  replacementRef   String?                         // prep.replRef
  replacementPrice Decimal?  @db.Decimal(12,4)     // prep.replPrice

  // snapshot columns for extras (survive later catalog edits)
  extraName  String?
  extraRef   String?
  extraPrice Decimal? @db.Decimal(12,4)
  extraUnit  String?

  order   Order    @relation(fields: [orderId],   references: [id], onDelete: Cascade)
  product Product? @relation(fields: [productId], references: [id], onDelete: Restrict)

  @@unique([orderId, productId, variantLabel])     // == lineKey uniqueness
  @@index([orderId])
  @@index([productId, prepState])                  // powers shortage/substitution analytics
}

// ──────────────────── order_prep_logs (append-only) ───────────────────
model OrderPrepLog {
  id          String     @id @default(uuid())
  tenantId    String
  orderId     String
  orderItemId String?                              // null for order-level events (photo/note)
  actorId     String?                              // magasinier who acted
  action      PrepAction
  fromState   PrepState?
  toState     PrepState?
  fromQty     Int?
  toQty       Int?
  detail      Json?                                // {repl, replRef, replPrice} | {assetNumber} | {note} | {photoId}
  createdAt   DateTime   @default(now())

  order Order      @relation(fields: [orderId], references: [id], onDelete: Cascade)
  actor User?      @relation(fields: [actorId], references: [id], onDelete: SetNull)

  @@index([orderId, createdAt])
  @@index([tenantId, action, createdAt])
}

// ───────────────────────────── order_photos ──────────────────────────
model OrderPhoto {
  id         String    @id @default(uuid())
  tenantId   String
  orderId    String
  kind       PhotoKind @default(proof)             // was inline prepPhotos[]
  objectKey  String                                // private bucket key (source of truth)
  url        String?                               // last-minted signed URL cache (optional)
  width      Int?
  height     Int?
  byteSize   Int?
  contentType String?  @default("image/jpeg")
  uploadedById String?
  uploadedAt DateTime  @default(now())

  order Order @relation(fields: [orderId], references: [id], onDelete: Cascade)

  @@index([orderId])
}

// ─────────────── order_counters (atomic YYMM sequence) ────────────────
model OrderCounter {
  tenantId  String
  monthKey  String                                 // "YYMM"
  lastValue Int    @default(0)
  @@id([tenantId, monthKey])
}
```

## 1.3 Constraints Prisma can't express (raw SQL migration)

Add these in a follow-up SQL migration; they encode invariants the ORM leaves open.

```sql
-- Positive quantities and prices everywhere
ALTER TABLE "OrderItem"    ADD CONSTRAINT ck_oi_reqqty  CHECK ("requestedQty" > 0);
ALTER TABLE "OrderItem"    ADD CONSTRAINT ck_oi_delqty  CHECK ("deliveredQty" IS NULL OR "deliveredQty" >= 0);
ALTER TABLE "KitItem"      ADD CONSTRAINT ck_ki_qty     CHECK (qty > 0);
ALTER TABLE "TemplateItem" ADD CONSTRAINT ck_ti_qty     CHECK (qty > 0);
ALTER TABLE "Product"      ADD CONSTRAINT ck_p_price    CHECK (price >= 0);
ALTER TABLE "Product"      ADD CONSTRAINT ck_p_pkg      CHECK ("packageQty" >= 1);

-- An OrderItem is EITHER a catalog line (productId set) OR an extra (extraName set) — never neither
ALTER TABLE "OrderItem" ADD CONSTRAINT ck_oi_catalog_or_extra
  CHECK ( ("productId" IS NOT NULL) OR ("extraName" IS NOT NULL) );

-- A catalog line's uniqueness must treat NULL variant as "" (Postgres NULLs are distinct otherwise)
-- Use a generated column so the unique index is well-defined:
ALTER TABLE "OrderItem"    ADD COLUMN variant_key text GENERATED ALWAYS AS (COALESCE("variantLabel", '')) STORED;
CREATE UNIQUE INDEX uq_oi_linekey ON "OrderItem" ("orderId", "productId", variant_key);

-- reference codes are immutable once set (prevents accidental re-sequencing)
CREATE OR REPLACE FUNCTION lock_reference_code() RETURNS trigger AS $$
BEGIN IF NEW."referenceCode" <> OLD."referenceCode" THEN
  RAISE EXCEPTION 'referenceCode is immutable'; END IF; RETURN NEW; END; $$ LANGUAGE plpgsql;
CREATE TRIGGER trg_lock_ref BEFORE UPDATE ON "Order"
  FOR EACH ROW EXECUTE FUNCTION lock_reference_code();

-- JSONB shape guard for variantsSchema: must be an array (deep validation stays in Zod, §1.4)
ALTER TABLE "Product"  ADD CONSTRAINT ck_p_variants  CHECK (variants_schema IS NULL OR jsonb_typeof(variants_schema) = 'array');
ALTER TABLE "Category" ADD CONSTRAINT ck_c_variants  CHECK (variants_schema IS NULL OR jsonb_typeof(variants_schema) = 'array');

-- GIN index so you can query "which products expose a 'Taille' variant"
CREATE INDEX gin_p_variants ON "Product" USING gin (variants_schema jsonb_path_ops);
```

**Reference-code generation** (atomic, inside the submit transaction — replaces client-side `counters[ym]+1`):

```sql
INSERT INTO "OrderCounter" ("tenantId","monthKey","lastValue")
VALUES ($1, $2, 1)
ON CONFLICT ("tenantId","monthKey")
DO UPDATE SET "lastValue" = "OrderCounter"."lastValue" + 1
RETURNING "lastValue";
-- reference_code := monthKey || '-' || lpad(lastValue::text, 3, '0')   →  e.g. 2601-004
```

## 1.4 JSONB `variantsSchema` contract

Both `Category.variantsSchema` and `Product.variantsSchema` share one shape. `normGroups()` in the prototype already tolerates a legacy string-array form; **normalize to the object form on write** so the DB only ever holds one shape.

```jsonc
// Canonical form (matches c11 "Textile" in the seed)
[
  { "name": "Taille",  "options": ["S","M","L","XL","XXL","XXXL"] },
  { "name": "Couleur", "options": ["Blanc","Noir","Bleu","Rouge"] }
]
```

Validate at the app boundary (Zod), never trust the client:

```ts
const VariantGroup = z.object({
  name: z.string().min(1).max(40),
  options: z.array(z.string().min(1).max(60)).min(1).max(50),
});
export const VariantsSchema = z.array(VariantGroup).max(6);   // matrix guard
```

The selected variant on a line is stored flat as `OrderItem.variantLabel` (the `::variant` half of `lineKey`), e.g. `"M / Bleu"`. Resolution precedence when rendering the picker: **product-level `variantsSchema` overrides category-level** — exactly `variantGroupsOf(p, catById)`.

## 1.5 Cascade summary (what dies with what)

| Parent deleted | Cascades / behaviour |
|---|---|
| `Tenant` | Everything (`onDelete: Cascade` on every `tenantId` FK). |
| `Order` | `order_items`, `order_photos`, `order_prep_logs` cascade-delete. |
| `Kit` | `kit_items` cascade. `Template` → `template_items` cascade. |
| `Product` | **Restricted** while referenced by any order/template/kit item. Use `active=false` (soft delete) instead — historical orders must keep their lines. `product_associations` referencing it cascade. |
| `User` | `Restrict` if they requested an order (history integrity); `assignee`, template owner, prep-log actor all `SetNull`. Prefer `active=false`. |
| `Category` | `Restrict` while it has products (reassign or deactivate first). |

---

# 2. Image Asset Storage Architecture

There are **two image lifecycles** with opposite requirements. Conflating them is the main mistake to avoid.

| | **Product/catalog images** | **Order proof photos** |
|---|---|---|
| Source | Admin uploads, or Excel `ImageURL` batch | Magasinier camera on the tablet (`capture="environment"`) |
| Volume | ~hundreds, slow-changing | thousands/month, append-only |
| Sensitivity | Public-ish (product photos) | Internal evidence — must be access-controlled |
| Caching | Aggressive, immutable, CDN public | Short-lived signed URLs, no shared CDN cache |
| Table link | `products.imageKey` / `imageUrl` (1:1) | `order_photos` rows (1:N) |

## 2.1 Store selection

**Reference choice: S3-compatible object storage + CloudFront (or Cloudflare R2 + Cloudflare CDN).** Rationale: you control bucket policies and signing, R2 has no egress fees, and it decouples from the DB host.

- **If you adopt Supabase for Postgres** → use **Supabase Storage**. It gives you buckets with RLS-style policies tied to the same JWT you already use for the DB, so proof-photo authorization reuses your auth with zero extra plumbing. This is the lowest-infra path.
- **If you want zero image-processing code** → **Cloudinary**. Upload once; request `/w_320,q_auto,f_auto/` derivatives on the fly. You trade money and a third party for deleting the whole thumbnail worker. Good fit for catalog images, overkill (and a data-residency question) for proof photos.

The rest of this section describes the S3/R2 reference design; Supabase/Cloudinary are drop-in substitutions at the marked seams.

## 2.2 Buckets & key layout

Two buckets with different access models:

```
pomproc-public   (CDN-fronted, public read, immutable objects)
  t/{tenantId}/products/{productId}/{sha256-12}.webp          # display (max 1200px)
  t/{tenantId}/products/{productId}/{sha256-12}_thumb.webp    # grid thumb (320px)

pomproc-private  (NO public access; objects reached only via signed URLs)
  t/{tenantId}/orders/{orderId}/{photoId}.jpg                 # full proof (max 1600px)
  t/{tenantId}/orders/{orderId}/{photoId}_thumb.jpg           # 400px preview
```

- **Content-hashed keys** for product images make them immutable → `Cache-Control: public, max-age=31536000, immutable`. A product re-upload writes a new key and flips `products.imageKey`; the old CDN object can expire naturally. This is what lets the PWA cache product images cache-first (blueprint §5) without ever serving stale art.
- **Proof keys** use the `photoId` (the `OrderPhoto.id`) so the DB row and the object are 1:1 and either can be reconciled from the other.

## 2.3 Access rules

- **`pomproc-public`**: bucket policy allows `s3:GetObject` to the CDN origin only; the bucket itself is not directly listable. CDN serves it. No auth needed to *view* a product image (they're not secret), but *writes* go only through the API worker's IAM role.
- **`pomproc-private`**: **block all public access.** Reads happen exclusively through the API:
  1. Client requests `GET /api/orders/:id/photos`.
  2. API runs RBAC: caller must be `admin`/`magasinier` of the tenant, **or** the `requester` of that order.
  3. For each photo, API mints a **short-lived (e.g. 5-min) presigned GET URL** and returns it.
  4. Client `<img src>` loads it directly from storage. URLs are never persisted long-term (the `url` column is only a best-effort cache; `objectKey` is the source of truth).
- **Uploads are direct-to-storage via presigned PUT/POST**, so image bytes never transit the API server. The API only ever sees keys.

## 2.4 Compression & thumbnailing pipeline

The prototype compresses proof photos client-side to **1000 px / JPEG 0.72** (`fileToThumb(file, 1000, 0.72)` via canvas). Keep that as a **pre-upload step** — on flaky 5G it cuts upload payload 10–50× before a byte leaves the tablet — then do a **server-side derivative + hardening pass** for consistency.

**Proof photo flow (magasinier):**
```
[tablet] pick/capture image
   └─ client canvas resize → JPEG ~1000px q0.72         (fast, saves 5G bandwidth)
   └─ POST /api/orders/:id/photos:init  → API returns { photoId, presignedPut, key }
   └─ PUT bytes directly to pomproc-private/{key}
   └─ POST /api/orders/:id/photos:commit { photoId, w, h, size }
        └─ API inserts OrderPhoto row
        └─ API enqueues worker job (BullMQ/Inngest)
             worker (sharp):  strip EXIF/GPS · auto-orient · make _thumb 400px ·
                              re-encode to a sane ceiling (≤1600px, q0.8) · write derivatives
        └─ API writes OrderPrepLog{ action: photo_add, detail:{photoId} }
```

**Product image flow (admin / Excel):**
```
upload or fetch ImageURL
   └─ worker (sharp): validate it's a real image · strip metadata ·
                      produce .webp display(≤1200px q0.82) + _thumb(320px) ·
                      compute sha256 for the key
   └─ write both to pomproc-public
   └─ set products.imageKey + imageUrl (CDN URL of display)
```

**Hardening applied by the worker (both flows):** verify magic bytes (don't trust `Content-Type`), cap decoded dimensions to defuse decompression bombs, strip all EXIF (removes GPS from tablet photos — a privacy must), re-encode (neutralizes polyglot files). Reject anything over a size ceiling (e.g. 15 MB pre-compression).

## 2.5 Migration of existing inline photos

Historical orders embed `prepPhotos: [dataURL]` (base64 JPEG in the JSON). Migration (blueprint §6, step 4.4): for each dataURL → decode → PUT to `pomproc-private` under a fresh `photoId` → insert `OrderPhoto`. This is where the orders payload shrinks 10–100×. Product `image` fields (mostly empty strings in the seed, occasionally base64) migrate the same way into `pomproc-public`.

---

# 3. Excel Batch Import & Sync Engine

The prototype already imports `xlsx` (`import * as XLSX`) and exports sheets. This section defines the **inbound** engine: turning an uploaded workbook into safe, idempotent upserts. It runs in an async worker (imports can be large) and always supports a **dry-run**.

## 3.1 Pipeline overview

```
upload .xlsx ─▶ parse ─▶ normalize headers ─▶ per-row validate ─▶
  match existing ─▶ plan (create/update/skip/error) ─▶
    [dry-run: return plan]  |  [commit: one transaction] ─▶
      link images ─▶ write ImportBatch + per-row results ─▶ downloadable report
```

Track every run as an `ImportBatch` (id, tenantId, kind, filename, actorId, mode, counts, createdAt) with child `ImportRowResult` rows, so an import is auditable and re-runnable.

## 3.2 Step by step

**1 — Parse.** `XLSX.read(buf, {type:'array', cellDates:true})`. Resolve the target sheet by name (`Products`, `Categories`, `Kits`, `Templates`) or first sheet. `XLSX.utils.sheet_to_json(ws, {defval:null, raw:false})` → array of row objects (`raw:false` gives you trimmed strings, so `"7,71"` and locale numbers survive).

**2 — Normalize headers.** Map incoming headers case-insensitively and trim, against a synonym table so `Prix`/`Price`/`Prix (€)` all land on `price`. Unknown columns are ignored but listed in the report (catches typos). Numbers: accept both `.` and `,` decimals (`Number(String(v).replace(',','.'))`) — the seed itself has comma refs like `"1326,1"`, so **treat SKU/ref as text**, never parse them as numbers.

**3 — Validate each row (Zod).** Required fields present, price ≥ 0, packageQty ≥ 1, entity ∈ {FFA,IIS,DC}, unitMode ∈ {unite,paquet}. A row that fails validation becomes an `error` in the plan and never blocks the others.

**4 — Match existing records.** Precedence, first hit wins:
   1. **`SKU`** (exact, tenant-scoped) → the intended path. `@@unique([tenantId, sku])` makes this a single deterministic lookup.
   2. **`Reference` + `Name`** (both exact) → fallback when the sheet predates SKUs. Because `ref` is non-unique, ref *alone* is never a match key.
   3. If a bare `Reference` matches **more than one** product and no SKU is given → **ambiguous**: mark the row `error("ref matches N products; add a SKU column")`. This is the concrete payoff of §0's SKU decision.
   4. No match → **create** (subject to §3.3 category handling).

**5 — Handle missing categories.** Products reference a category by **name** (`CategoryName`). If it doesn't exist:
   - **`strict` mode (default):** row → `error("unknown category 'X'")`. Nothing is created implicitly.
   - **`autoCreate` mode (opt-in checkbox):** create the `Category` with sensible defaults (`displayOrder = max+1`, `requiresAssetNum=false`, a rotating color, `icon="mat"`), record it in the report as `category_created`, then attach the product. Never silently guess `requiresAssetNum`.

**6 — Apply pricing / inventory updates.** For matched products, diff each incoming field against the row and update only what changed (`price`, `name`, `unit`, `packageQty`, `ownerEntity`, `active`, `categoryId`). Every change is recorded field-level in `ImportRowResult.changes` (`{price: {from: 7.71, to: 7.96}}`) so Billing can see when a price moved and orders already submitted keep their **snapshotted** line prices (extras) / recompute deliberately (catalog). Blank cells mean "leave unchanged," not "set to null" — distinguish empty string from absent column.

**7 — Link batch image URLs.** Two supported modes:
   - **`ImageURL` column** (public http/https): the worker fetches, validates, runs the §2.4 product pipeline, sets `imageKey`/`imageUrl`. Fetch failures → row warning, product still upserts without an image.
   - **`ImageFile` column + companion upload**: the sheet names a filename (`gants-nitrile-m.jpg`) and the admin uploads a ZIP/multi-file alongside the workbook; the worker matches by filename, then runs the same pipeline. Unmatched filenames and unreferenced files both go in the report.

**8 — Commit.** In **one transaction** per sheet: upsert categories (if autoCreate), then products, then association edges (resolve `Associations` SKUs → ids in a second pass, since a referenced product may be created in the same batch), then kits/templates and their items. On any unexpected DB error the whole batch rolls back — a half-applied catalog is worse than a rejected one.

**9 — Report.** Write `ImportRowResult` per row and return a downloadable results workbook: original rows + `Status` (created/updated/skipped/error/warning) + `Detail` columns. This mirrors how `Magasin.exportXlsx` already builds sheets with `aoa_to_sheet`, so the export half is code you already have.

## 3.3 Idempotency & safety

- **Re-running the same file is a no-op**: matching is by stable SKU, and unchanged fields produce no writes. Running an import twice must not duplicate products, kit items, or associations (the `@@unique` keys guarantee it).
- **Kits/templates replace-vs-merge:** default to **replace** a kit's item set when its sheet is re-imported (a kit is a curated bundle; partial merges surprise people). Make it a per-import toggle.
- **Dry-run first** is the recommended UX: show "will create 4 products, update 12 prices, create 1 category, 2 errors" before anyone commits.
- **Associations & kit/template items** are resolved to `productId` only after all products in the batch exist, so a kit can reference a product introduced in the same file.

---

# 4. Sample Import Spreadsheets

Column headers below are the canonical names; the importer also accepts the listed synonyms. **Text columns** (`SKU`, `Reference`) are never numeric-parsed. Decimal cells accept `.` or `,`.

## 4.1 Products (`Products` sheet)

One row per product (or per variant SKU, if you choose to explode variants into their own SKUs).

| SKU | Reference | Name | CategoryName | Unit | PackageQty | Price | Entity | Active | ImageURL | Associations |
|---|---|---|---|---|---|---|---|---|---|---|
| MOP-UNIKO-40 | 710522 | Mop Uniko Microtex - 40 cm | Matériel | Pièce | 25 | 7.71 | IIS | yes | https://cdn.example.com/mop-uniko.jpg | |
| GANT-NIT-M | 119316 | Gants nitrile non poudrés BLEU 100p - MEDIUM | EPI | Boite | 10 | 4.98 | IIS | yes | | |
| MAC-KARCHER-500 | MAC-002 | Karcher 500 bar | Machines | Pièce | 1 | 0 | IIS | yes | | TUY-001 |
| TUY-001 | TUY-001 | Tuyau d'eau 20 m | Matériel | Pièce | 1 | 0 | IIS | yes | | RAC-001 |
| RAC-001 | RAC-001 | Raccord Gardena | Matériel | Pièce | 1 | 0 | IIS | yes | | |

- **SKU** — required, unique per tenant, stable. This is the match key. If you're seeding from the current data and have no SKUs yet, generate them once (e.g. `slug(name)+shortid`) and keep them forever.
- **Reference** — supplier ref; may repeat across rows (that's expected and allowed).
- **CategoryName** — must match an existing category unless `autoCreate` is on.
- **Unit / PackageQty / Price / Entity** — map to `unit`, `packageQty`, `price`, `ownerEntity`. `PackageQty > 1` enables the "×N" package toggle.
- **Active** — `yes/no` (also accepts `true/false`, `1/0`). Absent → treated as `yes` on create, unchanged on update.
- **Associations** — comma-separated **SKUs** of related products (drives the "À prévoir aussi" chain). Resolved after all rows exist, so forward references are fine.
- *(Optional)* **Variants** — for variant-bearing rows, a cell like `Taille: S, M, L, XL | Couleur: Blanc, Noir` parses to the §1.4 JSON (`|` separates groups, `:` names the group, `,` separates options).

## 4.2 Categories (`Categories` sheet)

| Name | DisplayOrder | RequiresAssetNum | Color | Icon | Variants |
|---|---|---|---|---|---|
| EPI | 1 | no | #E8792B | epi | |
| Consommables | 2 | no | #128A52 | conso | |
| Matériel | 3 | no | #2D6CDF | mat | |
| Outillage | 4 | yes | #0E9AA0 | out | |
| Machines | 5 | yes | #6C4BD6 | mac | |
| Carburant | 6 | yes | #D9453F | carb | |
| Textile | 11 | no | #9B4DCA | tex | Taille: S, M, L, XL, XXL, XXXL \| Couleur: Blanc, Noir, Bleu, Rouge |

- **Name** — unique per tenant; the key products join to.
- **DisplayOrder** — integer sort key (was `order`).
- **RequiresAssetNum** — `yes/no`; `yes` surfaces the "N° matériel" field on lines (was `num`). Set it deliberately (Outillage/Machines/Carburant/Échelles are `yes` in the seed).
- **Color** — hex; **Icon** — one of the app's icon keys (`epi, conso, mat, out, mac, carb, pou, vit, ech, epc, tex`).
- **Variants** — same `group: opts | group: opts` grammar as products. Category variants apply to every product in the category unless a product overrides them.

## 4.3 Kits (`Kits` sheet)

One row per **kit item**; rows sharing a `KitName` form one kit (grouped on import). Kit items carry their own `Entity`.

| KitName | ItemSKU | Variant | Qty | Mode | Entity |
|---|---|---|---|---|---|
| Kit démarrage chantier | GANT-NIT-M | | 2 | unite | FFA |
| Kit démarrage chantier | MAC-KARCHER-500 | | 1 | unite | FFA |
| Kit démarrage chantier | TUY-001 | | 1 | unite | FFA |
| Kit bureaux | MOP-UNIKO-40 | | 4 | unite | IIS |
| Kit bureaux | GANT-NIT-M | | 1 | paquet | IIS |
| Kit textile équipe | TEX-TSHIRT | M / Bleu | 6 | unite | DC |

- **KitName** — the group key; also `kits.name` (unique per tenant). Re-importing a KitName **replaces** that kit's items by default.
- **ItemSKU** — must resolve to a product (SKU match). Can reference a product created in the same Products import if both sheets are in one workbook.
- **Variant** — must be one of the product/category variant options (e.g. `M / Bleu`); blank for non-variant products.
- **Qty** — integer ≥ 1. **Mode** — `unite` or `paquet` (`paquet` multiplies by the product's PackageQty). **Entity** — `FFA/IIS/DC`.

## 4.4 Templates (`Templates` sheet) — *companion to §1's `template_items`*

Same shape as Kits, minus `Entity` (template lines resolve entity from the product on apply), plus ownership.

| TemplateName | Owner | Shared | ItemSKU | Variant | Qty | Mode |
|---|---|---|---|---|---|---|
| Bureaux | admin | yes | MOP-UNIKO-40 | | 4 | unite |
| Bureaux | admin | yes | GANT-NIT-M | | 1 | unite |
| Vitres | admin | yes | RAC-001 | | 2 | unite |
| Événementiel | admin | yes | GANT-NIT-M | | 2 | unite |

- **Owner** — username; resolved to `ownerUserId` (unknown owner → row warning, template still created with null owner if `Shared=yes`).
- **Shared** — `yes/no`; shared templates are visible to all users, private ones only to the owner.

---

## Appendix — field name crosswalk (prototype → schema → Excel header)

| Prototype (JSX) | Schema column | Excel header |
|---|---|---|
| `p.nom` | `Product.name` | Name |
| `p.prix` | `Product.price` | Price |
| `p.unite` | `Product.unit` | Unit |
| `p.qteEmballage` | `Product.packageQty` | PackageQty |
| `p.entity` | `Product.ownerEntity` | Entity |
| `p.ref` | `Product.ref` (+ new `sku`) | Reference (+ SKU) |
| `p.image` | `Product.imageKey`/`imageUrl` | ImageURL |
| `p.assoc[]` | `ProductAssociation` | Associations |
| `cat.order` | `Category.displayOrder` | DisplayOrder |
| `cat.num` | `Category.requiresAssetNum` | RequiresAssetNum |
| `cat.variants` | `Category.variantsSchema` | Variants |
| `kit.items[].entity` | `KitItem.billedEntity` | Entity |
| `line.assetNo` | `OrderItem.assetNumber` | — (order-time only) |
| `prep[k].state` | `OrderItem.prepState` + `OrderPrepLog` | — |
| `prep[k].qtyPrep` | `OrderItem.deliveredQty` | — |
| `prep[k].repl/replRef/replPrice` | `OrderItem.replacement*` | — |
| `o.billFFA0/billDC0` | `Order.initialBill*` | — |
| `o.paymentStatus` | `Order.paymentStatus` | — |
| `line.addedAt` | `OrderItem.addedAt` | — (order-time only) |
| `line.addedBy` / `line.addedRole` | `OrderItem.addedByUserId` / `addedByRole` | — |
| `o.amendedAt` | `Order.amendedAt` | — |
| `o.prepPhotos[]` | `OrderPhoto` rows | — |
