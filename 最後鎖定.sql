-- ⚠️ 這段請在「新版 App 部署並測試 OK 之後」才執行（或交給 Claude 執行）。
-- 目的：把匿名寫入權限收掉，只保留公開「讀」船班/內容；個資表已在先前鎖好。

-- 1) goho_kv：公開可讀，僅登入的後台可寫
drop policy if exists goho_kv_insert on public.goho_kv;
drop policy if exists goho_kv_update on public.goho_kv;
create policy goho_kv_admin_ins on public.goho_kv for insert to authenticated with check (true);
create policy goho_kv_admin_upd on public.goho_kv for update to authenticated using (true) with check (true);
create policy goho_kv_admin_del on public.goho_kv for delete to authenticated using (true);

-- 2) goho（未使用的舊表）：移除所有匿名讀寫
drop policy if exists goho_select on public.goho;
drop policy if exists goho_insert on public.goho;
drop policy if exists goho_update on public.goho;
drop policy if exists goho_kv_select on public.goho;
drop policy if exists goho_kv_insert on public.goho;
drop policy if exists goho_kv_update on public.goho;

-- 3) 修正兩個函式的 search_path 安全提醒
alter function public.goho_kv_touch() set search_path = public;
alter function public.goho_touch() set search_path = public;
