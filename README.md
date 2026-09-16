# PomProc — Facility Management Logistics & Material Order System

![Version](https://img.shields.io/badge/version-1.0.0--v1-blue.svg)
![Stack](https://img.shields.io/badge/stack-Next.js%2014%20|%20TypeScript%20|%20Prisma%20|%20PostgreSQL-000000.svg)
![Deployment](https://img.shields.io/badge/deploy-Railway-purple.svg)
![License](https://img.shields.io/badge/license-Proprietary-red.svg)

**PomProc** (PWA Operations & Material Procurement Engine) is an enterprise web application designed for facility management teams, site supervisors, and warehouse managers (magasiniers). It streamlines the end-to-end lifecycle of site equipment and material procurement: from mobile requisition in the field to real-time warehouse preparation queue management, dual-entity dispatching, multi-tenant billing splits (IIS / FFA / DC), consumption analytics, and catalog management.

Originally prototyped as a standalone single-file React application (`pomproc.jsx`), this repository contains the production-ready full-stack application built on **Next.js (App Router)**, **TypeScript**, **Prisma ORM**, and **PostgreSQL**, ready for seamless deployment on **Railway**.

---

## Table of Contents

- [Key Features](#-key-features)
- [System Architecture & Stack](#-system-architecture--stack)
- [Database Schema & Domain Model](#-database-schema--domain-model)
- [Repository Structure](#-repository-structure)
- [Environment Variables](#-environment-variables)
- [Getting Started & Local Setup](#-getting-started--local-setup)
- [Seeding the Database](#-seeding-the-database)
- [Deployment on Railway](#-deployment-on-railway)
- [Notification Engine Architecture](#-notification-engine-architecture)
- [Excel Import Engine](#-excel-import-engine)
- [Acceptance Testing & Verification Matrix](#-acceptance-testing--verification-matrix)
- [Phase 2 Roadmap](#-phase-2-roadmap)
- [License & Support](#-license--support)

---

## Key Features

### 1. Site Supervisor Requisition Wizard (Demandeur)
* **3-Step Order Flow**: Intuitive mobile-first interface for selecting materials, configuring quantities/variants, and reviewing items.
* **Pre-loaded Templates & Kits**: Quick-start orders using shared operational templates (*Bureaux*, *Événementiel*, *Vitres*) or saved custom kits.
* **Smart Associations ("À prévoir aussi")**: Automatic suggestions for complementary items (e.g., ordering a floor scrubber automatically prompts for matching detergent and pad sizes).
* **Live Line & Billing Calculations**: Real-time calculation of line totals and tax allocations per entity (FFA / DC / IIS).

### 2. Warehouse Preparation Queue (Magasinier)
* **Live Queue Polling (~10s)**: Automatic polling updates the warehouse board instantly as site teams submit new requests.
* **Atomic Contention Lock**: Server-enforced optimistic locking (`status = 'envoyee'`) prevents two warehouse workers from taking the same order simultaneously (returns HTTP 409 Conflict on stale actions).
* **Line-Item Fulfillment & Replacement**: Mark items as `OK`, `Missing`, or `Replaced` with substitute products, delivered quantities, and asset serial number tracking (`requiresAssetNum`).
* **Photo Proof of Preparation**: Capture and upload dispatch photographs directly from tablet/mobile camera (`bytea` binary DB storage in v1).
* **Order Amendments**: Site supervisors can append missing items to finalized orders (`terminee`), automatically reopening the order state (`en_cours`) and tracking amendment provenance per date.

### 3. Dual Dispatch Sheets & Label Printing
* **Print-Ready Delivery Receipts**: Browser-native styled dispatch sheets grouping items by baseline order and subsequent amendment dates.
* **Box Label Generator**: Print-ready barcode/QR labels formatted for warehouse shipping boxes.

### 4. Multi-Entity Billing Engine
* **Entity Split Engine**: Automatic routing and splitting of costs across **IIS**, **FFA**, and **DC** entity accounts.
* **Initial Snapshot vs. Live Settlement**: Retains `initialBillFfa` / `initialBillDc` snapshot values at order submission to compute billing deltas and variance against final fulfilled totals.
* **Flexible Grouping & Filters**: Breakdown by month, demandeur, chef d'équipe, or magasinier.
* **Payment State Management**: Admin-only toggling of `Paid` / `Unpaid` invoice statuses.
* **Native Excel Export**: One-click generation of styled `.xlsx` reports matching screen aggregations.

### 5. Analytics & Logistics Intelligence
* **Product Consumption Frequencies**: Track top-requested items across sites and timeframes.
* **Product Co-Occurrence Matrix**: Identify frequently paired materials to optimize warehouse shelf layouts.
* **Discrepancy & Replacement Tracking**: Monitor missing items and substitution rates to inform stock reorder points.
* **Entity Flow Visualization**: High-level distribution breakdown between FFA, DC, and IIS cost centers.

### 6. Catalog & User Administration
* **SKU-Based Excel Import**: Upsert products and categories via `.xlsx` templates with a dry-run preview prior to committing database transactions.
* **Role-Based Access Control (RBAC)**: Enforces `admin`, `magasinier`, and `user` (demandeur) permission levels across API routes and UI components.

---

## System Architecture & Stack

PomProc is designed as a single deployable web application backed by a managed PostgreSQL database, eliminating microservice complexity for v1 while preserving clean abstractions for future scale.

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           Client Devices                                │
│   ┌────────────────────────┐              ┌─────────────────────────┐   │
│   │ Site Phone (Demandeur) │              │ Warehouse Tablet (Mag)  │   │
│   └───────────┬────────────┘              └────────────┬────────────┘   │
└───────────────┼────────────────────────────────────────┼────────────────┘
                │ HTTP / REST / Web Push                 │ 10s Polling / REST
                ▼                                        ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                      Railway Web Service (Next.js)                      │
│                                                                         │
│   ┌────────────────────────┐              ┌─────────────────────────┐   │
│   │  React UI Components   │              │ App Router API Handlers │   │
│   └────────────────────────┘              └────────────┬────────────┘   │
│                                                        │                │
│   ┌────────────────────────────────────────────────────┴────────────┐   │
│   │                     Domain Logic & Reducers                     │   │
│   │   (Calculations, Billing Splits, Excel Import, Auth Guards)    │   │
│   └────────────────────────┬────────────────────────────────────────┘   │
│                            │                                            │
│   ┌────────────────────────┴────────────────────────────────────────┐   │
│   │                   Notification Dispatcher                       │   │
│   │         ┌──────────────┼───────────────┬────────────────┐       │   │
│   │         ▼              ▼               ▼                ▼       │   │
│   │     [In-App]       [Email]        [Web Push]      [WhatsApp]    │   │
│   │                    (Resend)        (VAPID)          (Stub)      │   │
│   └────────────────────────┬────────────────────────────────────────┘   │
└────────────────────────────┼────────────────────────────────────────────┘
                             │ Prisma ORM / Raw SQL ($queryRaw)
                             ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                     Railway PostgreSQL Database                         │
│                                                                         │
│   • Tenants & Users           • Categories & Products (SKU keyed)       │
│   • Orders & Items            • Order Prep Audit Logs & Photos (Bytea)  │
│   • Shared Templates & Kits   • Atomic Monthly Reference Counters       │
└─────────────────────────────────────────────────────────────────────────┘
```

### Core Technology Stack

| Component | Technology | Description |
| :--- | :--- | :--- |
| **Framework** | Next.js 14 (App Router) | Unified SSR/CSR framework and API route handlers |
| **Language** | TypeScript | Strict type checking across domain models and API endpoints |
| **Database** | PostgreSQL | Managed relational database hosted on Railway |
| **ORM** | Prisma ORM | Type-safe query building, migrations, and schema management |
| **Analytics Engine** | Prisma `$queryRaw` | Native SQL aggregation for fast analytical performance |
| **Authentication** | Custom / NextAuth / Argon2 | HTTP-only session cookies with password hashing (`argon2` / `bcrypt`) |
| **Email Service** | Resend / SendGrid | Transactional email notification API |
| **Push Service** | `web-push` (VAPID) | Native Web Push notifications targeting Service Workers (`sw.js`) |
| **Excel Processing** | XLSX / SheetJS | Server-side dry-run import and client/server export engine |
| **Deployment** | Railway | Single-container deployment with PostgreSQL plugin |

---

## Database Schema & Domain Model

The underlying relational schema is defined in `prisma/schema.prisma` (guided by `PomProc_Data_Engine_Spec.md`). Below is an overview of key domain entities:

```
                    ┌──────────────┐
                    │    Tenant    │
                    └──────┬───────┘
                           │ 1
                           │
       ┌───────────────────┼───────────────────┬───────────────────┐
       │ *                 │ *                 │ *                 │ *
┌──────┴───────┐    ┌──────┴───────┐    ┌──────┴───────┐    ┌──────┴───────┐
│     User     │    │   Category   │    │    Product   │    │    Counter   │
└──────┬───────┘    └──────┬───────┘    └──────┬───────┘    └──────────────┘
       │ 1                 │ 1                 │ 1
       │                   │ *                 │ *
       │            ┌──────┴───────┐    ┌──────┴───────────────┐
       │            │  Sub-Category│    │ ProductAssociation   │
       │            └──────────────┘    └──────────────────────┘
       │ 1
       │ *
┌──────┴───────┐ 1    * ┌──────────────┐
│    Order     ├────────┤  OrderItem   │
└──────┬───────┘        └──────────────┘
       │ 1
       ├───────────────────────┬───────────────────────┐
       │ *                     │ *                     │ *
┌──────┴───────┐        ┌──────┴───────┐        ┌──────┴───────┐
│ OrderPrepLog │        │  OrderPhoto  │        │  PushSub     │
└──────────────┘        └──────────────┘        └──────────────┘
```

### Critical Model Constraints & Notes

1. **SKU vs. Reference Number**:
   * `ref` is **not unique** (catalog items may share historical manufacturer references).
   * `sku` is the immutable unique business key (`@@unique([tenantId, sku])`) generated during seeding or import (`slug(name) + shortId`). All Excel imports match and upsert against `sku`.
2. **Precision & Money Handling**:
   * Unit prices: `Decimal(12,4)` to accommodate high-precision fractions (e.g., `€0.328` per unit).
   * Order totals & invoice balances: `Decimal(12,2)`.
3. **Amendment Provenance**:
   * Reopened/amended order items store `addedAt`, `addedByUserId`, and `addedByRole`. Baseline items maintain `null` for these fields.
4. **Photo Blobs (v1)**:
   * Stored directly in PostgreSQL byte array (`data Bytes`, `mimeType String`) under `OrderPhoto`. Prepared for an easy swap to S3 presigned URLs in Phase 2.
5. **Atomic Counter**:
   * Sequential monthly reference codes (`YYMM-XXX`, e.g., `2609-001`) are assigned exclusively on the server using atomic counter increments within DB transactions to avoid collisions.

---

## Repository Structure

```
pomproc/
├── prisma/
│   ├── schema.prisma            # Prisma domain models & indexes
│   └── seed.ts                  # Idempotent catalog & admin seeder
├── public/
│   ├── manifest.webmanifest     # PWA manifest
│   └── sw.js                    # Service worker for VAPID Web Push
├── src/
│   ├── app/                     # Next.js App Router pages & API handlers
│   │   ├── layout.tsx           # Main application root layout
│   │   ├── page.tsx             # Root dashboard / route guard
│   │   └── api/                 # REST API route handlers
│   │       ├── auth/            # Login, logout, session verification
│   │       ├── catalog/         # Catalog CRUD & Excel import endpoints
│   │       ├── orders/          # Demande lifecycle, atomic locks, prep PATCH
│   │       │   └── [id]/
│   │       │       └── photos/  # Image upload & stream handlers
│   │       ├── push/            # VAPID subscription management
│   │       └── analytics/       # Prisma raw SQL aggregation API
│   ├── components/              # UI Component Modules (ported from prototype)
│   │   ├── OrderWizard.tsx      # Step-by-step requisition interface
│   │   ├── MagasinQueue.tsx     # Live warehouse preparation queue
│   │   ├── OrderPrep.tsx        # Fulfillment, replacement, & photo capture
│   │   ├── DispatchSheet.tsx    # Printable delivery slip & box labels
│   │   ├── BillingModule.tsx    # Entity financial breakdown & XLSX export
│   │   ├── AnalyticsView.tsx    # Consumption metrics & SVG charts
│   │   ├── AdminPanel.tsx       # Catalog management, user roles, Excel import
│   │   └── ui/                  # Shared primitives (buttons, modals, badges)
│   └── lib/                     # Core Business Logic & Infrastructure
│       ├── db.ts                # Global Prisma Client instance
│       ├── auth.ts              # Argon2 password hashing & session management
│       ├── domain/              # Pure business rules (verbatim prototype math)
│       │   ├── calculations.ts  # Line total, variant closure, associations
│       │   └── reducers.ts      # Billing & analytics aggregation logic
│       ├── notifications/       # Pluggable notification channel abstraction
│       │   ├── index.ts         # Central notification event dispatcher
│       │   ├── email.ts         # Resend / SendGrid transactional driver
│       │   ├── push.ts          # Web Push VAPID payload driver
│       │   └── whatsapp.ts      # (Phase 2 Stub) Meta Cloud API driver
│       └── excel/               # Excel Parsing & Dry-Run Validation Engine
│           ├── importer.ts      # SKU matching, dry-run validator, transactional upsert
│           └── exporter.ts      # XLSX workbook builder
├── .env.example                 # Environment variables template
├── Dockerfile                   # Nixpacks / Railway container definition
├── next.config.js               # Next.js configuration
├── package.json                 # Project dependencies & scripts
├── tsconfig.json                # TypeScript compiler config
└── README.md                    # Project documentation
```

---

## Environment Variables

Copy `.env.example` to `.env` and configure the required variables:

```bash
cp .env.example .env
```

| Variable Name | Required | Default / Example | Description |
| :--- | :---: | :--- | :--- |
| `DATABASE_URL` | Yes | `postgresql://user:pass@localhost:5432/pomproc` | PostgreSQL connection string (provided by Railway) |
| `SESSION_SECRET` | Yes | `a_super_secret_random_string_min_32_chars` | Secret key for signing HTTP-only session cookies |
| `SEED_ADMIN_PASSWORD` | Yes | `AdminPass123!` | Initial password for default `admin` user during seeding |
| `EMAIL_API_KEY` | No | `re_123456789...` | Resend or SendGrid API key for email dispatch |
| `EMAIL_FROM` | No | `PomProc <no-reply@yourdomain.com>` | Sender address for automated notification emails |
| `VAPID_PUBLIC_KEY` | No | `BEl62i...` | Public VAPID key for Web Push notifications |
| `VAPID_PRIVATE_KEY` | No | `4xN_...` | Private VAPID key for Web Push signing |
| `VAPID_SUBJECT` | No | `mailto:admin@yourdomain.com` | Contact email associated with VAPID subscription |
| `NEXT_PUBLIC_VAPID_PUBLIC_KEY` | No | `BEl62i...` | Client-exposed public VAPID key for service worker registration |

*Note: You can generate VAPID keys using `npx web-push generate-vapid-keys`.*

---

## Getting Started & Local Setup

### Prerequisites

* **Node.js**: v18.0.0 or higher
* **Package Manager**: `npm` (v9+) or `pnpm`
* **PostgreSQL**: Local instance (v14+) or Docker container

### 1. Clone & Install Dependencies

```bash
git clone https://github.com/your-org/pomproc.git
cd pomproc
npm install
```

### 2. Configure Local Database

If using Docker to run PostgreSQL locally:

```bash
docker run --name pomproc-postgres -e POSTGRES_PASSWORD=postgres -e POSTGRES_DB=pomproc -p 5432:5432 -d postgres:15-alpine
```

Set your `.env`:

```env
DATABASE_URL="postgresql://postgres:postgres@localhost:5432/pomproc?schema=public"
SESSION_SECRET="development_session_secret_key_change_me_in_production"
SEED_ADMIN_PASSWORD="AdminPassword123!"
```

### 3. Apply Migrations & Seed Data

Run Prisma migrations to construct the database schema, then execute the seed script:

```bash
# Apply database schema
npx prisma migrate dev --name init

# Seed initial catalog, categories, templates, and admin user
npx prisma db seed
```

### 4. Start the Development Server

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser. Log in with:
* **Username**: `admin`
* **Password**: `AdminPassword123!` (or value configured in `SEED_ADMIN_PASSWORD`)

---

## Seeding the Database

The database seed script (`prisma/seed.ts`) is **idempotent** (safe to execute repeatedly without duplicating records). It populates the database with:

1. **Default Tenant**: Slug `xlg`.
2. **Admin User**: Username `admin` with argon2-hashed password.
3. **Categories & Subcategories**: 11 default categories with display ordering, asset-tracking rules, and variant schemas.
4. **Product Catalog**: ~187 baseline products mapped to SKUs, prices, and unit descriptions.
5. **Product Associations**: Cross-linking rules for automatic "À prévoir aussi" recommendations.
6. **Operational Templates**: Shared requisition templates (*Bureaux*, *Événementiel*, *Vitres*).

To manually re-run seeding at any time:

```bash
npx prisma db seed
```

---

## Deployment on Railway

PomProc is pre-configured for deployment on **Railway** using a single Web service and an attached Railway PostgreSQL plugin.

### Step-by-Step Deployment Instructions

1. **Push Repository to GitHub**:
   Push your code to a private or public GitHub repository.

2. **Create Railway Project**:
   * Log into [Railway.app](https://railway.app/).
   * Click **New Project** → **Deploy from GitHub repo**.
   * Select your `pomproc` repository.

3. **Add PostgreSQL Plugin**:
   * Click **+ New** in your project canvas → **Database** → **Add PostgreSQL**.
   * Railway automatically provisions the database and injects `DATABASE_URL` into your service environment variables.

4. **Configure Environment Variables**:
   In your Railway Web Service settings, add the required environment variables:
   * `SESSION_SECRET`
   * `SEED_ADMIN_PASSWORD`
   * `EMAIL_API_KEY` / `EMAIL_FROM` (optional for v1)
   * `VAPID_PUBLIC_KEY` / `VAPID_PRIVATE_KEY` / `NEXT_PUBLIC_VAPID_PUBLIC_KEY` (optional for v1)

5. **Configure Build & Start Commands**:
   Railway auto-detects Next.js via Nixpacks. Verify the build configuration:
   * **Build Command**: `npm install && npx prisma generate && npm run build`
   * **Release Command** (runs before container start): `npx prisma migrate deploy`
   * **Start Command**: `npm run start -- -p $PORT`

6. **Seed Production Database**:
   Run the seed script once via Railway CLI or one-off terminal command:
   ```bash
   railway run npx prisma db seed
   ```

7. **Verify & Smoke Test**:
   Open the generated `.up.railway.app` URL and complete the [Acceptance Testing Matrix](#-acceptance-testing--verification-matrix).

---

## Notification Engine Architecture

PomProc implements a decoupled channel dispatcher (`lib/notifications/`) to notify warehouse staff and site supervisors of critical events.

### Supported Events
* `demande.created`: Fired when a demandeur submits a new order → Targets all `magasinier` and `admin` users.
* `demande.completed`: Fired when a warehouse worker finalizes order prep → Targets the original `demandeur` user.

### Channel Interface Implementation

```typescript
// lib/notifications/index.ts
export type NotifyEvent =
  | { type: 'demande.created'; order: OrderSummary }
  | { type: 'demande.completed'; order: OrderSummary };

export interface NotificationChannel {
  name: 'email' | 'push' | 'whatsapp';
  send(event: NotifyEvent, recipients: User[]): Promise<void>;
}
```

* **In-App Alerts**: Triggered via ~10s polling hook, updating badges and playing optional notification chimes in the browser.
* **Email Channel (`email.ts`)**: Sends formatted transactional HTML emails using Resend/SendGrid.
* **Web Push Channel (`push.ts`)**: Dispatches standard encrypted VAPID web push payloads to browser service workers registered in `PushSubscription`.
* **WhatsApp Channel (`whatsapp.ts`)**: Stubbed interface ready for Phase 2 Meta Cloud API template messaging integration.

---

## Excel Catalog Import Engine

Admins can bulk-update or create catalog products and categories using `.xlsx` files based on `PomProc_database_templates.xlsx`.

### Features
* **Dry-Run Validation**: Preview modifications ("*Will create 12 products, update 4 prices, create 2 categories, 0 errors*") prior to committing database transactions.
* **SKU Upsert Match**: Matches incoming rows by `SKU` (`tenantId, sku`). Existing SKUs update prices/descriptions; new SKUs append to the catalog.
* **Two-Pass Association Resolution**: Product association SKUs are parsed in a second pass after all primary products are created, ensuring forward references resolve correctly.
* **Format Flexibility**: Accepts comma or dot decimal separators (`12,50` or `12.50`) and header text synonyms (`Ref`, `Référence`, `SKU`, `Designation`, `Prix Unit`).

---

## Acceptance Testing & Verification Matrix

Perform these smoke tests after deploying to verify system integrity:

| # | Test Scenario | Steps | Expected Outcome | Pass |
| :---: | :--- | :--- | :--- | :---: |
| **1** | **Persistence** | Create a order, restart Railway service / container. | Order persists unchanged in PostgreSQL. | [ ] |
| **2** | **Authentication** | Attempt login with invalid credentials vs valid admin. | Passwords verified via argon2 server-side; invalid logins rejected. | [ ] |
| **3** | **Live Queue Polling** | Open supervisor UI on Device A and warehouse board on Device B. Submit order on A. | Device B displays new order within ~10 seconds without page refresh. | [ ] |
| **4** | **Contention Lock** | Two warehouse workers attempt to lock the same `envoyee` order simultaneously. | First succeeds; second receives HTTP 409 conflict and board auto-refreshes. | [ ] |
| **5** | **Atomic Reference Code** | Submit two orders in rapid succession within the same minute. | Assigned reference codes are unique sequential numbers (e.g., `2609-001`, `2609-002`). | [ ] |
| **6** | **Order Amendment** | Add missing item to a `terminee` order. | Order reopens to `en_cours`; dispatch sheet groups added item under amendment date. | [ ] |
| **7** | **Notifications** | Submit order as supervisor. Finalize order as warehouse manager. | Warehouse receives email/push on creation; supervisor receives email on completion. | [ ] |
| **8** | **Billing Parity** | Toggle paid/unpaid status as Admin and export XLSX bill. | Calculated totals match prototype math; export formatted correctly. | [ ] |
| **9** | **Excel Import** | Upload `Products` sheet with modified price for 1 SKU. | Dry-run accurately reports 1 update; commit updates target product without duplicating. | [ ] |

---

## Phase 2 Roadmap

The v1 release establishes a clean foundation designed for seamless Phase 2 upgrades without requiring core architectural rewrites:

* **S3 / CDN Photo Storage**: Swap `OrderPhoto` byte array storage for S3 presigned direct uploads (`imageKey` + CDN).
* **WebSockets / SSE Realtime**: Replace the ~10s polling interval with server-sent events (`Server-Sent Events` / Socket.io) using the established `useOrderFeed()` hook abstraction.
* **WhatsApp Cloud API Integration**: Complete the `whatsapp.ts` notification driver to send Meta-approved utility templates automatically.
* **PWA Offline Drafting**: Implement IndexedDB offline draft storage and background sync service worker for site supervisors operating without cellular coverage.
* **Multi-Tenancy RLS**: Activate full multi-tenant Row-Level Security switching via tenant middleware context.

---

## License & Support

**PomProc** is proprietary software developed for facility management and site logistics teams. All rights reserved.
