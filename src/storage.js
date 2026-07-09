// window.storage 相容層：預覽版 API → Supabase
// 語言偏好等個人設定存 localStorage；系統資料存 Supabase goho_kv 表
import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const anon = import.meta.env.VITE_SUPABASE_ANON_KEY
const sb = url && anon ? createClient(url, anon) : null
const LOCAL_KEYS = ['goho-lang']

if (!sb) console.warn('[GOHO] 未設定 Supabase 環境變數，改用瀏覽器本機儲存（僅供測試，資料不共享）')

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
    if (LOCAL_KEYS.includes(key) || !sb) {
      localStorage.setItem(key, value)
      return { key, value }
    }
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
  async list(prefix = '') {
    if (!sb) return { keys: Object.keys(localStorage).filter((k) => k.startsWith(prefix)) }
    const { data, error } = await sb.from('goho_kv').select('key').like('key', prefix + '%')
    if (error) throw error
    return { keys: (data || []).map((r) => r.key) }
  },
}
