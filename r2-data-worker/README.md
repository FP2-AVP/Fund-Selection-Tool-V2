# Fund Selection R2 Data Worker

Worker สำหรับอ่าน Base JSON และจัดการ Draft กับ Master Fund Override บน R2 แบบ private โดยตรวจ session 24 ชั่วโมงจาก KV เดียวกับ Auth Worker

ค่าเริ่มต้นอนุญาต frontend หลัก รวมถึง `http://localhost:8080` และ `http://127.0.0.1:8080` สำหรับทดสอบในเครื่อง สามารถกำหนด origin เพิ่มเติมแบบคั่นด้วย comma ผ่าน `ADDITIONAL_FRONTEND_ORIGINS`

เส้นทาง Draft:

```text
GET    /drafts?quarter=2026-Q3
GET    /drafts/{id}?quarter=2026-Q3
POST   /drafts
DELETE /drafts/{id}?quarter=2026-Q3
POST   /drafts/rebuild-index
```

ไฟล์ถูกจัดเก็บที่ `Draft/<ปี>/<Quarter>/<ชื่อไฟล์เดิม>.json`
และใช้ `Draft/index.json` สำหรับเปิดหน้ารายการอย่างรวดเร็ว

เส้นทาง Master Fund Override:

```text
GET    /master-fund-overrides?quarter=2026-Q3
POST   /master-fund-overrides
DELETE /master-fund-overrides/{key}?quarter=2026-Q3
```

ไฟล์ถูกจัดเก็บที่ `Data/<ปี>/<Quarter>/overrides/master_fund_overrides.json` โดยใช้ Master FundId เป็น key หลัก และใช้ `isin:<ISIN>` เมื่อไม่มี FundId

เส้นทาง Master Mapping ของกองทุนไทย:

```text
GET    /master-allocations?quarter=2026-Q3
POST   /master-allocations
DELETE /master-allocations/{fundKey}?quarter=2026-Q3
```

ไฟล์ถูกจัดเก็บที่ `Data/<ปี>/<Quarter>/overrides/fund_master_allocations.json` และ upsert เฉพาะ Fund Code ที่บันทึก

```bash
cd r2-data-worker
npm install
npx wrangler login
npm run deploy
```

หลัง Deploy นำ URL ไปใส่ใน `CONFIG.R2_DATA_API_URL` ที่ `js/config.override.js`

## Data For FT.com

FT data is stored as small, independently replaceable objects under `Data For FT.com/`.
The Worker exposes authenticated read routes and preserves R2 gzip metadata for yearly price shards.

```text
GET /ft/index
GET /ft/symbols/{SYMBOL_SLUG}
GET /ft/symbols/{SYMBOL_SLUG}/prices/{YYYY}
GET /ft/symbols/{SYMBOL_SLUG}/qualitative/latest
GET /ft/symbols/{SYMBOL_SLUG}/qualitative/snapshots/{YYYY-MM-DD}
```

Build the objects locally without uploading:

```bash
python scripts/sync_ft_r2.py
```

Upload them using the same R2 environment variables as `scripts/sync_r2_base.py`:

```bash
python scripts/sync_ft_r2.py --upload
```
