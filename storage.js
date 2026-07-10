// GOHO 資料層
// 設計：
//  - 非個資（船班／價格／內容／庫存）：存在 Supabase goho_kv 的 blob，公開可讀、僅後台可寫
//  - 個資（客戶／訂單）：一律走 Edge Function（公眾）或已登入後台直接存取資料表（RLS 僅 authenticated）
//  - 後台登入：Supabase Auth（帳號密碼）
import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const anon = import.meta.env.VITE_SUPABASE_ANON_KEY
export const sb = url && anon ? createClient(url, anon) : null
const FN = url ? url.replace(/\/$/, '') + '/functions/v1/goho-public' : null
const LOCAL_KEYS = ['goho-lang', 'goho-quest']

if (!sb) console.warn('[GOHO] 未設定 Supabase 環境變數，個資與登入功能將無法使用')

const genCode = () => {
  const a = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'
  let s = ''; for (let i = 0; i < 6; i++) s += a[Math.floor(Math.random() * a.length)]; return s
}

/* ---------- 呼叫 Edge Function（公眾流程） ---------- */
async function callFn(payload) {
  if (!FN || !sb) throw new Error('no-backend')
  const { data: { session } } = await sb.auth.getSession()
  const token = session?.access_token || anon
  const res = await fetch(FN, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', apikey: anon, Authorization: 'Bearer ' + token },
    body: JSON.stringify(payload),
  })
  return res.json()
}

export const api = {
  list: () => callFn({ action: 'list' }),
  register: (customer) => callFn({ action: 'register', customer }),
  login: (idno, birth) => callFn({ action: 'login', idno, birth }),
  book: (p) => callFn({ action: 'book', ...p }),
  createTrip: (p) => callFn({ action: 'create_trip', ...p }),
}

/* ---------- 後台驗證（Supabase Auth） ---------- */
export const auth = {
  signIn: (email, password) => sb.auth.signInWithPassword({ email, password }),
  signOut: () => sb.auth.signOut(),
  getSession: () => sb.auth.getSession(),
  hasSession: async () => { const { data } = await sb.auth.getSession(); return !!data.session },
}

/* ---------- 欄位對應（資料表 snake_case ↔ App camelCase） ---------- */
const fromCustomerRow = (r) => ({ id: r.id, name: r.name, nickname: r.nickname, gender: r.gender, birth: r.birth, idno: r.idno, phone: r.phone, address: r.address, minors: r.minors || [], minorMode: r.minor_mode, minorCount: r.minor_count, createdAt: r.created_at })
const toCustomerRow = (c) => ({ id: c.id, name: c.name, nickname: c.nickname ?? null, gender: c.gender, birth: c.birth, idno: c.idno, phone: c.phone, address: c.address, minors: c.minors ?? [], minor_mode: c.minorMode ?? 'none', minor_count: c.minorCount ?? 0 })
const fromOrderRow = (r) => ({ id: r.id, tripId: r.trip_id, customerId: r.customer_id, name: r.name, nickname: r.nickname, phone: r.phone, role: r.role, price: r.price, rental: r.rental, rentalPrice: r.rental_price, status: r.status, paid: r.paid, isMinor: r.is_minor, bookingCode: r.booking_code, createdAt: r.created_at })
const toOrderRow = (o) => ({ id: o.id, customer_id: o.customerId ?? null, trip_id: o.tripId, name: o.name, nickname: o.nickname, phone: o.phone, role: o.role, price: o.price, rental: o.rental ?? null, rental_price: o.rentalPrice ?? 0, paid: !!o.paid, is_minor: !!o.isMinor, status: o.status, booking_code: o.bookingCode || genCode() })

/* ---------- 後台直接存取資料表（需登入；RLS 僅允許 authenticated） ---------- */
export const tables = {
  async customers() { const { data, error } = await sb.from('customers').select('*').order('created_at', { ascending: false }); if (error) throw error; return (data || []).map(fromCustomerRow) },
  async orders() { const { data, error } = await sb.from('orders').select('*').order('created_at', { ascending: false }); if (error) throw error; return (data || []).map(fromOrderRow) },
  async upsertCustomer(c) { const { error } = await sb.from('customers').upsert(toCustomerRow(c), { onConflict: 'id' }); if (error) throw error },
  async deleteCustomer(id) { const { error } = await sb.from('customers').delete().eq('id', id); if (error) throw error },
  async insertOrder(o) { const { error } = await sb.from('orders').insert(toOrderRow(o)); if (error) throw error },
  async updateOrder(id, patch) { const { error } = await sb.from('orders').update(patch).eq('id', id); if (error) throw error },
  async deleteOrder(id) { const { error } = await sb.from('orders').delete().eq('id', id); if (error) throw error },
}

/* ---------- window.storage：blob 存取（非個資） ----------
   語言／任務等個人設定存 localStorage；系統內容（船班等）存 Supabase goho_kv，僅後台（已登入）可寫。 */
window.storage = {
  async get(key) {
    if (LOCAL_KEYS.includes(key) || !sb) {
      const v = localStorage.getItem(key)
      return v === null ? null : { key, value: v }
    }
    const { data, error } = await sb.from('goho_kv').select('value').eq('key', key).maybeSingle()
    if (error) throw error
    return data ? { key, value: data.value } : null
  },
  async set(key, value) {
    if (LOCAL_KEYS.includes(key) || !sb) { localStorage.setItem(key, value); return { key, value } }
    const { error } = await sb.from('goho_kv').upsert({ key, value }, { onConflict: 'key' })
    if (error) throw error
    return { key, value }
  },
  async delete(key) {
    if (LOCAL_KEYS.includes(key) || !sb) { localStorage.removeItem(key); return { key, deleted: true } }
    const { error } = await sb.from('goho_kv').delete().eq('key', key)
    if (error) throw error
    return { key, deleted: true }
  },
}
