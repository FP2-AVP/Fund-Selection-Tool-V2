# Fund Selection R2 Data Worker

Worker สำหรับอ่าน Base JSON จาก R2 แบบ private โดยตรวจ session 24 ชั่วโมงจาก KV เดียวกับ Auth Worker

```bash
cd r2-data-worker
npm install
npx wrangler login
npm run deploy
```

หลัง Deploy นำ URL ไปใส่ใน `CONFIG.R2_DATA_API_URL` ที่ `js/config.override.js`
