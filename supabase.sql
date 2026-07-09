-- ============ GOHO 船班報名系統 v2 — Supabase 初始化 ============
-- 在 Supabase Dashboard → SQL Editor 貼上執行一次即可

create table if not exists public.goho_kv (
  key text primary key,
  value text,
  updated_at timestamptz default now()
);

alter table public.goho_kv enable row level security;

-- 公開讀寫（MVP 信任模型，與預覽版一致；正式營運建議升級為正規化資料表 + Auth）
drop policy if exists "goho_kv_select" on public.goho_kv;
drop policy if exists "goho_kv_insert" on public.goho_kv;
drop policy if exists "goho_kv_update" on public.goho_kv;
create policy "goho_kv_select" on public.goho_kv for select using (true);
create policy "goho_kv_insert" on public.goho_kv for insert with check (true);
create policy "goho_kv_update" on public.goho_kv for update using (true);

-- 更新時間自動維護
create or replace function public.goho_kv_touch() returns trigger as $$
begin new.updated_at = now(); return new; end;
$$ language plpgsql;
drop trigger if exists goho_kv_touch on public.goho_kv;
create trigger goho_kv_touch before update on public.goho_kv
for each row execute function public.goho_kv_touch();
