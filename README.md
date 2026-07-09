# 五賀娛樂漁船 GOHO｜船班報名系統 v2 — 部署指南

Vite + React 前端，Supabase 作為資料庫（goho_kv 單表 KV 模式，與預覽版資料結構完全相同），部署於 Vercel。

## 一、Supabase 設定（約 3 分鐘）
1. 進入你的 Supabase 專案 → SQL Editor
2. 貼上 `supabase.sql` 全部內容 → Run（建立 goho_kv 表與 RLS 政策）
3. Settings → API：複製 `Project URL` 與 `anon public` 金鑰

## 二、本機測試（可略過直接部署）
```bash
npm install
cp .env.example .env      # 填入 Supabase URL / anon key / 後台密碼
npm run dev               # http://localhost:5173
```
未設定 .env 也能跑，會退回瀏覽器本機儲存（僅測試用，資料不共享）。

## 三、部署到 Vercel
### 方式 A：GitHub（建議，之後改版自動部署）
1. 將本資料夾推上 GitHub repo
2. Vercel → Add New Project → Import 該 repo（Framework 自動偵測 Vite）
3. Environment Variables 加入三個變數：
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `VITE_ADMIN_PASSWORD`（務必改掉 goho888）
4. Deploy → 取得 `https://xxx.vercel.app`

### 方式 B：CLI 直接部署
```bash
npm i -g vercel
vercel            # 依提示登入並建立專案
vercel env add VITE_SUPABASE_URL
vercel env add VITE_SUPABASE_ANON_KEY
vercel env add VITE_ADMIN_PASSWORD
vercel --prod
```

## 四、上線後檢查清單
- [ ] 開啟網址 → 系統訊息 → 註冊一筆測試客戶 → 報名一班
- [ ] 換另一台手機開啟 → 確認能看到同一筆報名（= Supabase 已生效）
- [ ] 後台登入（新密碼）→ 訂單看得到測試報名 → 匯出 Excel 正常
- [ ] Supabase Table Editor → goho_kv 表內有 `goho-erp-v2` 一筆資料
- [ ] 綁定自訂網域（Vercel → Domains，例如 booking.letsgotainan.com）

## 五、重要注意事項
- **資料模型**：MVP 採單一 KV 文件（與預覽版相同），任何寫入為整份覆蓋。同時多人報名極端情況下可能互相覆蓋；目前客流量下風險低，但正式營運穩定後建議升級為正規化資料表（trips / orders / customers 分表）+ Supabase Auth，屆時欄位可直接對應。
- **安全**：RLS 目前為公開讀寫（與預覽版信任模型一致）。後台密碼僅是前端閘門，請勿在後台存放不能外流的資料以外的機密；升級 Auth 前，匯出功能請只在自己的裝置使用。
- **備份**：Supabase Dashboard → Database → Backups 已含每日備份；也可定期在後台匯出客戶名單 Excel 留存。
