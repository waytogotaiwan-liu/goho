import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const admin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false } },
);

const BLOB_KEY = "goho-erp-v2";
const S_CANCELLED = "已取消";
const S_BOOKED = "已報名";
const S_WAIT = "候補";
const R_SIGHT = "觀光";
const R_ANGLER = "釣手";
const T_OPEN = "報名中";

const ALLOWED_ORIGINS = new Set([
  "https://goho.letsgotainan.com",
  "https://goho-git-main-gohofishingboat.vercel.app",
  "http://localhost:5173",
  "http://localhost:3000",
]);

function cors(origin: string | null) {
  const allow = origin && ALLOWED_ORIGINS.has(origin) ? origin : "https://goho.letsgotainan.com";
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, content-type, apikey",
    "Vary": "Origin",
    "Content-Type": "application/json",
  };
}
const json = (b: unknown, s: number, h: Record<string, string>) => new Response(JSON.stringify(b), { status: s, headers: h });
const normBirth = (s: string) => (s || "").replace(/[民國年月日\s]/g, "-").replace(/[\/.．]/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
function code6() { const a = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"; let s = ""; for (let i = 0; i < 6; i++) s += a[Math.floor(Math.random() * a.length)]; return s; }

async function loadBlob() {
  const { data } = await admin.from("goho_kv").select("value").eq("key", BLOB_KEY).maybeSingle();
  try { return data?.value ? JSON.parse(data.value) : {}; } catch { return {}; }
}
async function saveBlob(blob: any) {
  await admin.from("goho_kv").upsert({ key: BLOB_KEY, value: JSON.stringify(blob) }, { onConflict: "key" });
}

// 防暴力嘗試：回 false = 已遭鎖定
async function throttle(key: string, limit = 10) {
  const now = Date.now();
  const { data } = await admin.from("sec_throttle").select("count, reset_at").eq("id", key).maybeSingle();
  if (data) {
    if (now < new Date(data.reset_at).getTime()) {
      if (data.count >= limit) return false;
      await admin.from("sec_throttle").update({ count: data.count + 1 }).eq("id", key);
    } else {
      await admin.from("sec_throttle").update({ count: 1, reset_at: new Date(now + 9e5).toISOString() }).eq("id", key);
    }
  } else {
    await admin.from("sec_throttle").insert({ id: key, count: 1, reset_at: new Date(now + 9e5).toISOString() });
  }
  return true;
}
const clearThrottle = (key: string) => admin.from("sec_throttle").delete().eq("id", key);

async function activeCount(tripId: string) {
  const { count } = await admin.from("orders").select("id", { count: "exact", head: true }).eq("trip_id", tripId).neq("status", S_CANCELLED);
  return count ?? 0;
}
async function verifyCustomer(idno: string, birth: string) {
  const { data: cust } = await admin.from("customers").select("*").eq("idno", idno).maybeSingle();
  if (!cust || normBirth(cust.birth) !== normBirth(birth)) return null;
  return cust;
}

Deno.serve(async (req) => {
  const headers = cors(req.headers.get("origin"));
  if (req.method === "OPTIONS") return new Response("ok", { headers });
  if (req.method !== "POST") return json({ error: "method_not_allowed" }, 405, headers);
  let body: any = {};
  try { body = await req.json(); } catch { return json({ error: "bad_json" }, 400, headers); }
  const action = String(body.action || "");

  try {
    // ===== LIST =====
    if (action === "list") {
      const blob = await loadBlob();
      const trips = Array.isArray(blob.trips) ? blob.trips : [];
      const { data: orders } = await admin.from("orders").select("trip_id, nickname, role, status");
      const rosters: Record<string, any> = {};
      for (const o of orders || []) {
        if (o.status === S_CANCELLED) continue;
        const r = (rosters[o.trip_id] ||= { count: 0, names: [] });
        r.count++;
        r.names.push({ nickname: o.nickname || R_ANGLER, role: o.role });
      }
      return json({ trips, rosters }, 200, headers);
    }

    // ===== REGISTER（首次註冊：僅建客戶） =====
    if (action === "register") {
      const c = body.customer || {};
      const miss = ["name", "gender", "birth", "idno", "address", "phone"].filter((k) => !String(c[k] ?? "").trim());
      if (miss.length) return json({ error: "missing_fields", fields: miss }, 400, headers);
      const idno = String(c.idno).trim();
      const { data: exists } = await admin.from("customers").select("id").eq("idno", idno).maybeSingle();
      if (exists) return json({ error: "duplicate_id" }, 409, headers);
      const customerId = "C" + Date.now();
      const { error } = await admin.from("customers").insert({
        id: customerId, name: c.name, nickname: c.nickname ?? null, gender: c.gender, birth: c.birth,
        idno, phone: c.phone, address: c.address, minors: c.minors ?? [], minor_mode: c.minorMode ?? "none",
        minor_count: Number(c.minorCount ?? 0) || 0,
      });
      if (error) return json({ error: "insert_failed", detail: error.message }, 500, headers);
      return json({ ok: true, customerId, customer: { ...c, id: customerId } }, 200, headers);
    }

    // ===== LOGIN（證號 + 生日） =====
    if (action === "login") {
      const idno = String(body.idno ?? "").trim();
      const birth = String(body.birth ?? "").trim();
      if (!idno || !birth) return json({ error: "missing_credentials" }, 400, headers);
      const tkey = "login:" + idno;
      if (!(await throttle(tkey))) return json({ error: "too_many_attempts" }, 429, headers);
      const cust = await verifyCustomer(idno, birth);
      if (!cust) return json({ found: false }, 200, headers);
      await clearThrottle(tkey);
      const { data: orders } = await admin.from("orders").select("*").eq("customer_id", cust.id);
      const customer = {
        id: cust.id, name: cust.name, nickname: cust.nickname, gender: cust.gender, birth: cust.birth,
        idno: cust.idno, phone: cust.phone, address: cust.address, minors: cust.minors ?? [],
        minorMode: cust.minor_mode, minorCount: cust.minor_count,
      };
      const myOrders = (orders || []).map((o) => ({
        id: o.id, tripId: o.trip_id, customerId: o.customer_id, name: o.name, nickname: o.nickname,
        phone: o.phone, role: o.role, price: o.price, rental: o.rental, rentalPrice: o.rental_price,
        status: o.status, paid: o.paid, isMinor: o.is_minor, bookingCode: o.booking_code,
      }));
      return json({ found: true, customer, orders: myOrders }, 200, headers);
    }

    // ===== BOOK（已登入客戶報名某船班） =====
    if (action === "book") {
      const idno = String(body.idno ?? "").trim();
      const birth = String(body.birth ?? "").trim();
      const tripId = String(body.tripId ?? "").trim();
      const role = body.role === R_SIGHT ? R_SIGHT : R_ANGLER;
      const rentalName = body.rental ?? null;
      const companions: string[] = Array.isArray(body.companions) ? body.companions : [];
      if (!idno || !birth || !tripId) return json({ error: "missing_fields" }, 400, headers);
      const tkey = "book:" + idno;
      if (!(await throttle(tkey, 20))) return json({ error: "too_many_attempts" }, 429, headers);

      const cust = await verifyCustomer(idno, birth);
      if (!cust) return json({ error: "auth_failed" }, 401, headers);

      const blob = await loadBlob();
      const trip = (Array.isArray(blob.trips) ? blob.trips : []).find((t: any) => t.id === tripId);
      if (!trip) return json({ error: "trip_not_found" }, 404, headers);
      if (trip.status && trip.status !== T_OPEN) return json({ error: "trip_closed", status: trip.status }, 409, headers);

      const { data: dup } = await admin.from("orders").select("booking_code").eq("trip_id", tripId).eq("customer_id", cust.id).neq("status", S_CANCELLED).maybeSingle();
      if (dup) return json({ ok: true, already: true, bookingCode: dup.booking_code }, 200, headers);

      const minors = (cust.minor_mode === "self" && Array.isArray(cust.minors)) ? cust.minors : [];
      const seatsNeeded = 1 + minors.length + companions.length;
      const cap = Number(trip.capacity ?? 10) || 10;
      const rem = cap - (await activeCount(tripId));
      const status = rem >= seatsNeeded ? S_BOOKED : S_WAIT;

      const price = Number(trip.price ?? 0) || 0;
      let rentalPrice = 0;
      if (rentalName) {
        const rr = (Array.isArray(blob.rentals) ? blob.rentals : []).find((x: any) => x.name === rentalName);
        rentalPrice = rr ? (Number(rr.price) || 0) : 0;
      }

      const now = Date.now();
      const mainCode = code6();
      const mainId = "O" + now;
      const rows: any[] = [{
        id: mainId, customer_id: cust.id, trip_id: tripId, name: cust.name, nickname: cust.nickname || cust.name,
        phone: cust.phone, role, price, rental: rentalName, rental_price: rentalPrice, paid: false,
        is_minor: false, status, booking_code: mainCode,
      }];
      minors.forEach((m: any, i: number) => rows.push({
        id: mainId + "-m" + i, customer_id: cust.id, trip_id: tripId, name: m.name, nickname: (m.name || "") + "(未成年)",
        phone: cust.phone, role: R_SIGHT, price, rental: null, rental_price: 0, paid: false, is_minor: true, status, booking_code: code6(),
      }));
      companions.forEach((n: string, i: number) => rows.push({
        id: mainId + "-c" + i, customer_id: cust.id, trip_id: tripId, name: n, nickname: n,
        phone: cust.phone, role: R_SIGHT, price, rental: null, rental_price: 0, paid: false, is_minor: false, status, booking_code: code6(),
      }));
      const { error } = await admin.from("orders").insert(rows);
      if (error) return json({ error: "book_failed", detail: error.message }, 500, headers);
      return json({ ok: true, status, bookingCode: mainCode, seats: seatsNeeded }, 200, headers);
    }

    // ===== CREATE_TRIP（已登入釣友自行開班） =====
    if (action === "create_trip") {
      const idno = String(body.idno ?? "").trim();
      const birth = String(body.birth ?? "").trim();
      const trip = body.trip || {};
      if (!idno || !birth || !trip.date) return json({ error: "missing_fields" }, 400, headers);
      const tkey = "opentrip:" + idno;
      if (!(await throttle(tkey, 10))) return json({ error: "too_many_attempts" }, 429, headers);
      const cust = await verifyCustomer(idno, birth);
      if (!cust) return json({ error: "auth_failed" }, 401, headers);

      const blob = await loadBlob();
      blob.trips = Array.isArray(blob.trips) ? blob.trips : [];
      if (blob.trips.some((x: any) => x.date === trip.date)) return json({ error: "date_taken" }, 409, headers);
      const safe = {
        id: "T" + Date.now(),
        date: String(trip.date),
        name: String(trip.name || "自訂船班").slice(0, 40),
        type: String(trip.type || "近海").slice(0, 10),
        price: Number(trip.price) || 0,
        muster: String(trip.muster || "05:00").slice(0, 10),
        rodsUp: String(trip.rodsUp || "—").slice(0, 10),
        back: String(trip.back || "—").slice(0, 20),
        timePending: !!trip.timePending,
        capacity: 10,
        targets: Array.isArray(trip.targets) ? trip.targets.slice(0, 8) : ["洽船長"],
        depth: String(trip.depth || "洽船長").slice(0, 40),
        gear: "洽船長建議",
        rigs: [],
        note: "釣友自行開班，歡迎跟報",
        status: "歡迎開班",
      };
      blob.trips.push(safe);
      blob.trips.sort((a: any, b: any) => String(a.date).localeCompare(String(b.date)));
      await saveBlob(blob);
      return json({ ok: true, trip: safe }, 200, headers);
    }

    return json({ error: "unknown_action" }, 400, headers);
  } catch (e) {
    return json({ error: "server_error", detail: String(e) }, 500, headers);
  }
});
