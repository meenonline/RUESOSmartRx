# Pharmacy Sub-stock Management System

ระบบบริหารจัดการคลังยา Sub-stock โรงพยาบาลรือเสาะ

## Local Development Setup

### Prerequisites
- Node.js 18+
- PostgreSQL database

### Backend

```bash
cd pharmacy-substock/backend
cp .env.example .env
# Edit .env with your DATABASE_URL
npm install
npm run migrate   # Run database migration
npm run dev       # Start dev server (port 3001)
```

### Frontend

```bash
cd pharmacy-substock/frontend
cp .env.example .env
# Edit VITE_API_URL to point to your backend (e.g., http://localhost:3001)
npm install
npm run dev       # Start dev server (port 5173)
```

## Railway Deployment

1. Create a new project on [Railway](https://railway.app)
2. Add a PostgreSQL database service
3. Create two services from this repo:
   - **Backend**: Set root directory to `pharmacy-substock/backend`
   - **Frontend**: Set root directory to `pharmacy-substock/frontend`
4. Set environment variables (see below)
5. Backend will auto-deploy with `railway.toml` config
6. Run migration: connect to backend service and run `npm run migrate`

## Environment Variables

### Backend (`pharmacy-substock/backend/.env`)

| Variable | Description |
|----------|-------------|
| `DATABASE_URL` | PostgreSQL connection string |
| `PORT` | Server port (default: 3001) |
| `FRONTEND_URL` | Frontend URL for CORS |
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Messaging API token |
| `LINE_TARGET_ID` | LINE user/group ID to send notifications |
| `RESEND_API_KEY` | Resend.com API key for email reports |

### Frontend (`pharmacy-substock/frontend/.env`)

| Variable | Description |
|----------|-------------|
| `VITE_API_URL` | Backend API base URL |
| `VITE_APP_NAME` | Application display name |
| `VITE_FISCAL_YEAR` | Thai fiscal year for display |

## Database Migration

Run the migration to create all tables and views:

```bash
cd pharmacy-substock/backend
npm run migrate
```

This creates:
- `transactions` — all stock movements
- `formulary` — drug master list
- `ignore_expiry` — expiry alert suppression
- `audit_trail` — random count audit records
- `settings` — key-value configuration
- `stock_summary_lot` — view: stock by lot
- `stock_summary_nolot` — view: stock aggregated

## Features

- Dashboard with stock value, expiry alerts, and low-stock warnings
- Stock summary with lot-level and aggregated views
- Transaction history with manual entry
- Excel import/export (SheetJS, runs in browser)
- Formulary management with inline editing
- Random drug audit system
- Procurement list generation
- LINE notification integration
- Email report via Resend API
- Dark mode support
