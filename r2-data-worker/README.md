# Fund Selection R2 Data Worker

Worker สำหรับอ่าน Base JSON และจัดการ Draft บน R2 แบบ private โดยตรวจ session 24 ชั่วโมงจาก KV เดียวกับ Auth Worker

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

```bash
cd r2-data-worker
npm install
npx wrangler login
npm run deploy
```

หลัง Deploy นำ URL ไปใส่ใน `CONFIG.R2_DATA_API_URL` ที่ `js/config.override.js`
