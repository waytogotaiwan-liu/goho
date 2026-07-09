import { useState, useEffect, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

/* ============ GOHO 五賀娛樂漁船 — 船班報名系統 & 內部資料庫 v2 ============ */
/* 配色：深海藍 #0C2D48 / 潮水青 #1EA896 / 浮球橘 #F4A259 / 沙白 #F7F3EC / 頁首 #052488（同 Logo） */
const C = {
  navy: "#0C2D48", navyDeep: "#081D30", teal: "#1EA896", tealDark: "#137A6E",
  orange: "#F4A259", sand: "#F7F3EC", yellow: "#FFD166", red: "#E76F51", gray: "#8FA3B0",
  page: "#F7F3EC", logoBlue: "#052488",
};

const CONTACT = {
  captain: "0988-098-975",
  captainLine: "https://line.me/ti/p/ayOwyVY7ka",
  lineOA: "https://lin.ee/lV6klDe",
  port: "https://maps.app.goo.gl/5T8iUM7eLcNJxhTL7",
};

const SERVICES = [
  { icon: "🐙", name: "沉船礁大物", zone: "近海", price: "NT$2,500", basePrice: 2500 },
  { icon: "🎣", name: "近海敲底小搞搞", zone: "近海", price: "NT$2,500", basePrice: 2500 },
  { icon: "🐟", name: "近海鐵板遊動丸", zone: "近海", price: "NT$2,500", basePrice: 2500 },
  { icon: "🚤", name: "中遠程鐵板遊動丸", zone: "中遠程", price: "NT$4,000", basePrice: 4000 },
  { icon: "🗺️", name: "東吉七美西淺敲底", zone: "離島", price: "NT$4,500–6,500", basePrice: 5500 },
  { icon: "🦑", name: "澎湖小卷一日", zone: "離島", price: "NT$7,000", basePrice: 7000 },
  { icon: "🐠", name: "西南花鱸赤鯮馬頭", zone: "中遠程", price: "NT$4,000", basePrice: 4000 },
  { icon: "📡", name: "海測", zone: "專案", price: "洽詢", basePrice: 0 },
  { icon: "🌅", name: "看夕陽放空潛水", zone: "休閒", price: "洽詢", basePrice: 0 },
  { icon: "🎬", name: "影片婚紗拍攝", zone: "專案", price: "洽詢", basePrice: 0 },
];

const PRICING = [
  { name: "近海船班", name_en: "Nearshore trips", name_ja: "近海便", price: "NT$2,500", hours: "約 7 小時", hours_en: "~7 hrs", hours_ja: "約7時間", note: "小搞搞／小鐵板／遊動丸／船拋／敲底小搞搞／活餌／沉船礁大物。早班 05:00 集合、12:00 起竿、約 13:00 回港；夜釣（活餌）16:00 集合、23:00 起竿、約 24:00 回港。", note_en: "Light jigging / micro-jig / Yudo / boat casting / bottom / live bait / wreck big game. Morning: muster 05:00, lines up 12:00, port ~13:00; night (live bait): 16:00 / 23:00 / ~24:00.", note_ja: "小物釣り／小型ジグ／遊動丸／キャスティング／底物／活き餌／沈船大物。朝便：5時集合・12時納竿・13時頃帰港；夜釣り（活き餌）：16時集合・23時納竿・24時頃帰港。" },
  { name: "中遠程船班", name_en: "Mid-range trips", name_ja: "中遠距離便", price: "NT$4,000", hours: "約 10 小時", hours_en: "~10 hrs", hours_ja: "約10時間", note: "鐵板遊動丸／花鱸／赤鯮／馬頭／黃雞。04:00 集合、14:00 起竿、16–17:00 回港。", note_en: "Jigging & Yudo / sea bass / crimson snapper / tilefish / yellow porgy. Muster 04:00, lines up 14:00, port 16–17:00.", note_ja: "ジグ・遊動丸／スズキ／チダイ／アマダイ／キダイ。4時集合・14時納竿・16〜17時帰港。" },
  { name: "東吉／七美／西淺敲底", name_en: "Dongji / Qimei / West-shoal bottom", name_ja: "東吉／七美／西浅場 底物", price: "NT$4,500–6,500", hours: "依航線", hours_en: "By route", hours_ja: "航路による", note: "依實際航線報價。", note_en: "Priced by actual route.", note_ja: "実際の航路により料金が異なります。" },
  { name: "澎湖小卷一日遊", name_en: "Penghu squid day trip", name_ja: "澎湖イカ日帰り便", price: "NT$7,000", hours: "一日", hours_en: "1 day", hours_ja: "1日", note: "季節限定。", note_en: "Seasonal.", note_ja: "季節限定。" },
  { name: "七美二日遊", name_en: "Qimei 2-day trip", name_ja: "七美1泊2日便", price: "NT$12,000", hours: "二日", hours_en: "2 days", hours_ja: "2日間", note: "含一晚住宿＋一頓晚餐套餐。", note_en: "Includes one night's stay + dinner set.", note_ja: "宿泊1泊＋夕食セット付き。" },
];

const RENTALS = [
  { name: "小鐵板／遊動丸 套組", name_en: "Micro-jig / Yudo rod set", name_ja: "小型ジグ／遊動丸セット", price: 800, unit: "組", unit_en: "set", unit_ja: "セット", note: "竿＋手動捲線器（耗材另計）", note_en: "Rod + manual reel (rigs extra)", note_ja: "竿＋手動リール（仕掛け別）", isSet: true, zone: "近海" },
  { name: "活餌 套組", name_en: "Live-bait set (electric reel)", name_ja: "活き餌セット（電動リール）", price: 1500, unit: "組", unit_en: "set", unit_ja: "セット", note: "竿＋電動捲線器＋竿架（耗材另計）", note_en: "Rod + electric reel + holder (rigs extra)", note_ja: "竿＋電動リール＋竿受け（仕掛け別）", isSet: true, zone: "通用" },
  { name: "遊動丸 仕掛", name_en: "Yudo rig", name_ja: "遊動丸仕掛け", price: 350, unit: "個", unit_en: "pc", unit_ja: "個", note: "加購耗材（釣組）", note_en: "Add-on rig (consumable)", note_ja: "追加仕掛け（消耗品）", isSet: false, zone: "近海" },
  { name: "小鐵板 仕掛", name_en: "Micro-jig rig", name_ja: "小型ジグ仕掛け", price: 250, unit: "個", unit_en: "pc", unit_ja: "個", note: "加購耗材（釣組）", note_en: "Add-on rig (consumable)", note_ja: "追加仕掛け（消耗品）", isSet: false, zone: "近海" },
  { name: "活餌 仕掛", name_en: "Live-bait rig", name_ja: "活き餌仕掛け", price: 300, unit: "個", unit_en: "pc", unit_ja: "個", note: "加購耗材（釣組），不含餌", note_en: "Add-on rig, bait not included", note_ja: "追加仕掛け・餌は含まれません", isSet: false, zone: "近海" },
  { name: "深場電動 套組", name_en: "Deep-water electric set", name_ja: "深場電動セット", price: 1800, unit: "組", unit_en: "set", unit_ja: "セット", note: "重竿＋大容量電動捲線器＋竿架（耗材另計）", note_en: "Heavy rod + high-capacity electric reel + holder (rigs extra)", note_ja: "強竿＋大容量電動リール＋竿受け（仕掛け別）", isSet: true, zone: "深場" },
  { name: "深場鐵板 仕掛", name_en: "Deep-water jig rig", name_ja: "深場ジグ仕掛け", price: 400, unit: "個", unit_en: "pc", unit_ja: "個", note: "加購耗材（釣組），重鉛規格", note_en: "Add-on rig, heavy-weight spec", note_ja: "追加仕掛け・重量級", isSet: false, zone: "深場" },
];

const DEFAULT_RULES = [
  { icon: "🪪", title: "一人一單，實名報關", body: "因應海巡出入港查驗，一筆報名僅限一位出海人員；管理員可代客報名。報名須提供：姓名、法定性別、出生年月日、身分證／護照／居留證號碼、地址、聯絡電話（暱稱選填，顯示於公開名單）。出海當天請攜帶證件正本供查驗。" },
  { icon: "🧭", title: "船班由釣友發起", body: "船長不主動開班。請先查詢船班行事曆，於空班日期依服務項目自行開班，或聯繫阿豪船長／LINE 小秘書。注意：出發前 24 小時內不開放開班。" },
  { icon: "👥", title: "釣手與觀光、包船優先", body: "現場可登記為「釣手」或「觀光」；釣手滿員時，觀光名額可轉為釣手。觀光乘客不提供裝備租借。包船團體優先於散客。全船乘客上限 10 位（不含船長與船員）。" },
  { icon: "⚖️", title: "低於 5 人得取消", body: "出發前一日若報名未滿 5 人，船長保留取消權利，最遲於前一日通知。候補乘客最遲於出發前 12 小時通知。" },
  { icon: "🌊", title: "天候降班機制", body: "船資 NT$3,500 以上屬中遠程班；天候不佳時先降為近海班，若近海亦不安全則取消。天候取消依船長公告辦理，不視為違約。" },
  { icon: "❌", title: "取消政策", body: "出發日（含）前 7 日內取消、未到或遲到，須支付全額船資（可自行找人替補）。7 日內因喪、病並附證明者免付。" },
  { icon: "🛟", title: "合法船舶與保險", body: "CT2 級合法娛樂漁船，總噸位逾 12 噸，領有營業執照並投保乘客保險。全程免費提供救生衣，船上設有廁所與臥艙。" },
  { icon: "🧊", title: "冰箱與活餌", body: "船上備有共用漁獲冰箱（請以束帶標記漁獲）及獨立食物冰箱；近海班乘客免自備冰塊冰箱。活餌班請自備活餌餌料。" },
];

const DEFAULT_ANNOUNCEMENT = "#更新 #五賀 2026年7-8月 船班資訊📣\n本月船長推薦：近海小鐵板遊動丸活餌、夜釣近海活餌大物、中遠程鐵板遊動丸、澎湖小卷、澎湖敲底、花鱸、黃雞班\n📢 歡迎多多開班 📢";

const DEFAULT_POSTS = [
  { id: "P1", date: "2026-07-01", title: "2026 年 7-8 月船班資訊", body: DEFAULT_ANNOUNCEMENT },
];
const DEFAULT_CONTACT_PAGE = {
  intro: "報班、開班、包船、海測、婚紗拍攝等任何需求，歡迎透過以下方式與我們聯絡：",
  items: [
    { icon: "💬", label: "官方 LINE 小秘書", value: "最快回覆，報班首選", link: CONTACT.lineOA },
    { icon: "🧑‍✈️", label: "阿豪船長 LINE", value: "船班細節、釣況討論", link: CONTACT.captainLine },
    { icon: "📞", label: "電話", value: CONTACT.captain, link: "tel:0988098975" },
    { icon: "📍", label: "出發地點：安平港", value: "點擊開啟 Google 地圖", link: CONTACT.port },
  ],
};

const DEFAULT_TRIPS = [
  { id: "T2607-10", date: "2026-07-10", name: "近海活餌夜釣", type: "近海", price: 2500, muster: "16:00", rodsUp: "23:00", back: "約 24:00", capacity: 10, targets: ["紅甘", "石斑", "白帶魚"], depth: "20–40 m", gear: "活餌組（電動捲線器）＋活餌仕掛", rigs: ["活餌仕掛 NT$300"], note: "請自備活餌餌料", status: "報名中" },
  { id: "T2607-12", date: "2026-07-12", name: "近海敲底小搞搞", type: "近海", price: 2500, muster: "05:00", rodsUp: "12:00", back: "約 13:00", capacity: 10, targets: ["黑鯛", "石斑", "紅魽"], depth: "15–30 m", gear: "假餌建議釣組裝備（僅供參考）\n釣場水深約：10-90米\n（以實際釣場為主）\n鐵板：40-120克\n遊動丸：80-200克\n船拋：18公分餌(破鋒)米諾或波趴\n捲線器需求：\n手捲（紡車）優可拋\n手捲（股式）次難拋\n補充說明：要拋18公分（破鋒假餌）最好有大約8呎長的船拋竿、配SW的捲線器\n活餌、釣組、勾子、子線、母線及捲線器搭配建議（供參考）\n＊活餌餌料需自備\n釣組：大雙培林轉環串串 或 大物天秤T型25公分粗骨（買不到我們有賣）\n鐵條重：建議半斤～1斤都備著\n🦐 蝦餌\n鉤子：OWNER 管付石班勾24-26號，建議至少15-20公分大白蝦\n子線：碳線10-20號（蝦體越大越粗）\n母線：PE 2～8\n捲線器：shimano電捲3000型 bm/md系列 或 DAIWA電捲500型，手捲、牛車輪也可\n🦀 螃蟹餌\n鉤子：OWNER管付石班勾28-30號，螃蟹餌建議含腳約15公分\n子線：50-150號碳線／80-300磅\n母線：PE 12～20\n捲線器：shimano電捲6000型 或 DAIWA電捲600-800型，手捲、牛車輪也可\n🐙 章魚餌\n鉤子：OWNER管付35-45號，勾半斤-1斤重活餌章魚\n子線：150～300磅\n母線：PE 20\n捲線器：shimano電捲6000-12000型 或 DAIWA電捲800-1800型，手捲、牛車輪也可", rigs: ["小鐵板 NT$250", "遊動丸 NT$350"], note: "新手友善班", status: "報名中" },
  { id: "T2607-15", date: "2026-07-15", name: "中遠程鐵板遊動丸", type: "中遠程", price: 4000, muster: "04:00", rodsUp: "14:00", back: "16–17:00", capacity: 10, targets: ["花鱸", "赤鯮", "馬頭", "黃雞"], depth: "60–90 m", gear: "中深場鐵板竿＋遊動丸", rigs: ["遊動丸 NT$350", "小鐵板 NT$250"], note: "天候不佳先降近海班", status: "確定出船" },
  { id: "T2607-18", date: "2026-07-18", name: "東吉敲底", type: "離島", price: 5500, muster: "03:30", rodsUp: "14:00", back: "17:00", capacity: 10, targets: ["石斑", "紅條", "紅魽"], depth: "40–80 m", gear: "敲底竿組＋重鉛", rigs: ["小鐵板 NT$250"], note: "中遠程班，天候降班機制適用", status: "報名中" },
  { id: "T2607-22", date: "2026-07-22", name: "澎湖小卷一日遊", type: "離島", price: 7000, muster: "15:00", rodsUp: "23:00", back: "翌日 02:00", capacity: 10, targets: ["小卷"], depth: "30–60 m", gear: "小卷竿＋布捲仕掛", rigs: [], note: "季節限定", status: "報名中" },
  { id: "T2607-25", date: "2026-07-25", name: "沉船礁大物", type: "近海", price: 2500, muster: "05:00", rodsUp: "12:00", back: "約 13:00", capacity: 10, targets: ["大石斑", "紅甘"], depth: "25–45 m", gear: "大物竿＋活餌組", rigs: ["活餌仕掛 NT$300"], note: "請自備活餌餌料", status: "歡迎開班" },
  { id: "T2607-29", date: "2026-07-29", name: "西南花鱸赤鯮馬頭", type: "中遠程", price: 4000, muster: "04:00", rodsUp: "14:00", back: "16–17:00", capacity: 10, targets: ["花鱸", "赤鯮", "馬頭"], depth: "70–100 m", gear: "中深場竿＋遊動丸", rigs: ["遊動丸 NT$350"], note: "", status: "歡迎開班" },
];

const DEFAULT_INVENTORY = [
  { id: "INV-1", cat: "裝備", name: "小鐵板／遊動丸 套組", qty: 6, unit: "組", low: 2 },
  { id: "INV-2", cat: "裝備", name: "活餌 套組（電捲）", qty: 4, unit: "組", low: 1 },
  { id: "INV-3", cat: "裝備", name: "救生衣", qty: 14, unit: "件", low: 12 },
  { id: "INV-4", cat: "耗材", name: "遊動丸 仕掛", qty: 28, unit: "個", low: 10 },
  { id: "INV-5", cat: "耗材", name: "小鐵板 仕掛", qty: 35, unit: "個", low: 10 },
  { id: "INV-6", cat: "耗材", name: "活餌 仕掛", qty: 22, unit: "個", low: 10 },
  { id: "INV-7", cat: "耗材", name: "漁獲標記束帶", qty: 180, unit: "條", low: 50 },
];

/* 船班狀態設定：pub=是否對外公開可報名 */
const TRIP_STATUS = {
  "歡迎開班": { color: C.teal, tone: "teal", pub: true, bookable: true },
  "報名中": { color: C.teal, tone: "teal", pub: true, bookable: true },
  "確定出船": { color: C.navy, tone: "navy", pub: true, bookable: true },
  "已完成": { color: C.gray, tone: "navy", pub: true, bookable: false },
  "人數不足取消": { color: C.red, tone: "red", pub: true, bookable: false },
  "天氣因素取消": { color: C.red, tone: "red", pub: true, bookable: false },
  "不可抗力因素取消": { color: C.red, tone: "red", pub: true, bookable: false },
};
const STATUS_LIST = Object.keys(TRIP_STATUS);

const WELCOME_MSG_1 = "您好，歡迎來到五賀娛樂漁船＊本船為合法娛樂漁船，已依政府規定取得相關營業執照及保險，並經主管機關核准合法營運。報名出海時，您可使用身分證、健保卡、駕照、護照或居留證辦理登記。＊";
const WELCOME_MSG_2 = "＊為因應漁業署及海巡署出入港查驗規定，報名時我們會請您提供出海人員資料，提供的資料包含：姓名、生日、身份證字號、地址、聯絡電話及相關資訊以利辦理報關作業。因此，為了讓您未來操作系統便利性，您是否同意提供以上資訊，並且設定您的身份證字號作為固定登入帳號、出生年月日作為固定登入密碼？";


/* ================= 多語系 i18n（中/英/日） ================= */
let LANG = "zh";
const I18N = {
zh: {
  brandSub:"台南安平出船，合法娛樂漁船，歡迎報名", loading:"⚓ 系統載入中…", gate:"請先完成上方系統訊息流程",
  sysTitle:"系統訊息", promptTitle:"提示訊息", csTitle:"真人客服",
  welcome1:"您好，歡迎來到五賀娛樂漁船＊本船為合法娛樂漁船，已依政府規定取得相關營業執照及保險，並經主管機關核准合法營運。報名出海時，您可使用身分證、健保卡、駕照、護照或居留證辦理登記。＊",
  welcome2:"＊為因應漁業署及海巡署出入港查驗規定，報名時我們會請您提供出海人員資料，提供的資料包含：姓名、生日、身份證字號、地址、聯絡電話及相關資訊以利辦理報關作業。因此，為了讓您未來操作系統便利性，您是否同意提供以上資訊，並且設定您的身份證字號作為固定登入帳號、出生年月日作為固定登入密碼？",
  agreeYes:"是，我同意提供資料並設定帳號密碼", agreeYesSub:"前往報名／登入頁", agreeNo:"否，我暫不同意", confirm:"確認",
  understandQ:"您是否已經完全了解我們的報班規定流程？", yesUnderstand:"是的我了解", backStep:"回到上一步", notUnderstand:"我不了解", contactCS:"與真人客服聯繫",
  lineOA:"加入官方 LINE 小秘書", capLine:"加阿豪船長 LINE", tel:"電話", back:"返回",
  entryTitle:"請選擇您的入口", e1t:"我是第一次來", e1d:"首次註冊：填寫出海人員資料，建立帳號", e2t:"我已經來過", e2d:"帳號登入：證號＋民國出生年月日", e3t:"船長後台登入", e3d:"管理者專用：船班、訂單、客戶、庫存",
  regTitle:"出海人員報名表單（首次註冊）", loginTitle:"船員登入", backEntry:"← 返回入口",
  phNick:"暱稱（選填，顯示於公開名單）", phName:"姓名 *", phGender:"法定性別 *", male:"男", female:"女",
  phBirth:"民國出生年月日（例：80-05-12）＝登入密碼 *", phId:"身分證／護照／居留證號碼＝登入帳號 *", phAddr:"地址 *", phPhone:"聯絡電話 *",
  minorQ:"同行是否有未滿 18 歲的未成年人？ *", has:"有", none:"無", countLbl:"人數", minorSelf:"由我代為填寫資料", minorContact:"我要主動與你們聯絡",
  minorN:"未成年人 {i}（由監護人 {g} 代填）", phMinorName:"姓名 *", phMinorGender:"性別 *", phMinorBirth:"民國生日（例：100-03-05）*", phMinorId:"證號 *",
  minorContactNote:"我們會主動與您聯絡協助未成年人報名。您也可直接聯繫 LINE 小秘書。",
  docNote:"出海當天請攜帶所有出海人員的證件正本供海巡查驗。", submitReg:"送出資料並登入",
  phLoginId:"帳號（證號）", phLoginPw:"密碼（民國出生年月日，例：80-05-12）", login:"登入",
  errRequired:"除暱稱外，所有欄位皆為必填。", errDup:"此證號已註冊，請改用「登入」。", errLogin:"帳號或密碼錯誤（帳號＝證號，密碼＝民國出生年月日）。", errMinorSel:"請選擇同行是否有未滿 18 歲的未成年人。", errMinorMode:"請選擇未成年人的處理方式。", errMinorFields:"請完整填寫每位未成年人的資料。",
  annTitle:"最新公告", annOk:"我知道了，開始導引 →",
  q_announce:"看過最新公告", qh_announce:"了解本月船長推薦班別", q_rules:"讀懂報班規定", qh_rules:"前往「規定」看完 8 條規定", q_fees:"認識費用與租借", qh_fees:"前往「費用」了解船資與裝備", q_calendar:"查一次船班行事曆", qh_calendar:"前往「船班」點開任一船班", q_book:"完成第一次報名", qh_book:"在船班詳情按下「我要報名」",
  rank0:"五賀新手", rank1:"見習船員", rank2:"合格釣手", rank3:"五賀老手",
  nextRank:"再 {n} XP 升級為「{r}」", maxRank:"已達最高階級！", taskN:"任務 {i}｜", doneLbl:"已完成", nextStep:"← 下一步", nextStepB:"下一步", questFooter:"完成全部任務即代表您已了解報班流程，可安心出海 ⚓",
  calTitle:"{y} 年 {m} 月船班", listV:"清單", monthV:"月曆", openTripBtn:"＋ 我要開班（依服務項目在空班日期開班）",
  lgAvail:"名額充足", lgFew:"即將額滿", lgFull:"已滿／取消", lgOpen:"▢ 空班可開", canOpen:"可開", cxl:"取消", remainS:"餘{n}",
  musterAt:"⏰ {t} 集合", remain:"剩 {n} 位", waitlist:"候補中",
  otTitle:"我要開班", otSub:"依規定：釣友可自行發起船班", otDate:"出發日期（限 24 小時以後）", otSvc:"服務項目",
  otErrSel:"請選擇日期與服務項目。", otErr24:"出發前 24 小時內不開放開班，請改選較晚的日期。", otErrDup:"該日期已有船班，請直接報名或改選其他日期。",
  otTime:"出發時段", tsCustom:"⏰ 自訂期望時間", customLbl:"期望集合時間", customNote:"自訂時間的船班會標記「時間待船長確認」，船長將依潮汐與航程確認或調整，並於前台更新。", pendingTag:"時間待船長確認", otErrTime:"請選擇出發時段或填寫期望集合時間。",
  otInq:"「{s}」為洽詢制服務，費用與細節需與真人確認後開班：", orCall:"或來電 {p}", otNote:"開班後狀態為「歡迎開班」，其他釣友即可跟報；未滿 5 人船長得於前一日取消。", otCreate:"建立船班並成為第一位報名者",
  rfTitle:"裝備租借費用", rfItem:"項目", rfUnit:"單位", rfFee:"費用", rfNote:"※ 備註：以上租借費用不含釣組（仕掛耗材），釣組需另行加購；活餌仕掛不含餌，請自備活餌餌料。", rfOk:"了解，繼續報名",
  backCal:"← 返回船班行事曆", backBottom:"← 返回上一頁（船班行事曆）", perAngler:"／釣手", seatsLeft:"剩餘 {n} 位", fullWait:"已滿・可候補",
  cardV:"卡片檢視", tableV:"表格檢視", muster:"集合", rodsUp:"起竿", backPort:"回港",
  lTargets:"🎯 目標魚", lDepth:"🌊 釣場水深", lGear:"🎣 裝備建議", lRigs:"🧷 建議釣組", lNote:"📌 備註", byInstr:"依現場指示",
  roster:"👥 已報名名單（公開暱稱）", rosterEmpty:"目前尚無人報名，快來當第一位！", minorTag:"・未成年",
  booked:"您已報名此船班 ✅ 可於「我的船班」查看動態與行前提醒", notBookable:"本船班狀態為「{s}」，暫不開放報名。",
  roleLbl:"報名身份", angler:"釣手", sight:"觀光", rentalLbl:"裝備租借（選填）", noRental:"自備裝備，不租借", feeLink:"查看完整租借費用表 →",
  compLbl:"攜伴觀光（選填）", compCount:"同行觀光人數", compName:"觀光同行者 {i} 姓名／暱稱 *", compNote:"觀光同行者不提供裝備租借；完整報關資料請於出海當日攜證件辦理，或由管理員代為補登。",
  sightBlock:"⚠ 觀光僅限「釣手攜伴」方式報名，無法單獨報名觀光。請改選「釣手」，並於下方「攜伴觀光」加入同行者。",
  minorSeat:"👶 系統將同時為您代填的 {n} 位未成年人一併登記座位。", bookBtn:"我要報名", waitBtn:"登記候補", policy:"一人一單；取消政策：出發前 7 日內取消須付全額",
  myTitle:"我的船班", myEmpty1:"您還沒有報名任何船班", myEmpty2:"到「船班」分頁挑一個喜歡的班別出海吧！", idLbl:"身份", paid:"✅ 已收款", unpaid:"🕐 待收款", rentShort:"租借", remTitle:"🔔 行前提醒", viewFull:"查看完整船班資訊 →", stCxl:"已取消", stWait:"候補中",
  rm1:"攜帶證件正本供海巡查驗", rm2:"提前 15 分鐘抵達安平港集合點", rm3:"暈船者請提前 30 分鐘服用暈船藥", rm4:"活餌班請自備活餌餌料", rm5:"近海班免自備冰塊冰箱，船上備有共用漁獲冰箱", rm6:"中遠程／離島班建議自備飲水與行動糧", rm7:"以束帶標記自己的漁獲",
  feeTitle:"常見船班費用（每位釣手）", rentTitle:"裝備租借費用（耗材／釣組另計）", feeFoot:"租借費用不含釣組（仕掛耗材）；活餌班請自備活餌餌料；近海班免自備冰。",
  ctTitle:"與我們聯絡", newsTitle:"最新公告", noPosts:"目前沒有公告。",
  depart:"出發地點：安平港", openMap:"開啟地圖 →",
  navQuest:"任務", navCal:"船班", navMy:"我的", navRules:"規定", navFees:"費用", navSvc:"服務", navCt:"聯絡",
  menu:"☰ 選單", mMy:"🎣 查詢我的預約", mHome:"🏠 回到首頁", mOut:"🚪 登出", adminBtn:"船長後台",
  hello:"⚓ {n}，歡迎登船", welcomeNew:"歡迎登船，{n}！", welcomeBack:"歡迎回航，{n}！",
  tSightOnly:"觀光僅限釣手攜伴報名，請以釣手身份加攜伴。", tWait:"座位不足，已為您（含同行）登記候補。", tBooked:"報名成功！＋30 XP 🎉", tBookedX:"（含 {n} 位同行）", tCreated:"船班已建立！狀態為「歡迎開班」，快報名成為第一位 🎣", tPw:"密碼錯誤",
  rzS:"近海裝備／釣組", rzD:"深場裝備／釣組", rzTagS:"近海", rzTagD:"深場", rzTagB:"通用",
  autoNote:"本班釣場水深 {d}，系統自動顯示「{z}」租借項目（以 80 米為基準）", manualNote:"⚓ 本班出租裝備由船長指定", gvAcc:"摺疊", gvTab:"分頁", gvGrid:"總覽",
  status:{"歡迎開班":"歡迎開班","報名中":"報名中","確定出船":"確定出船","已完成":"已完成","人數不足取消":"人數不足取消","天氣因素取消":"天氣因素取消","不可抗力因素取消":"不可抗力因素取消"},
  zone:{"近海":"近海","中遠程":"中遠程","離島":"離島","專案":"專案","休閒":"休閒"}, inq:"洽詢",
},
en: {
  brandSub:"Departing Anping, Tainan — licensed recreational fishing vessel. Welcome aboard!", loading:"⚓ Loading…", gate:"Please complete the system message above first",
  sysTitle:"System Message", promptTitle:"Notice", csTitle:"Contact Us",
  welcome1:"Welcome to GOHO Recreational Fishing! *This is a licensed recreational fishing vessel with government-approved permits and insurance. For boarding registration you may use a national ID, health-insurance card, driver's license, passport, or ARC.*",
  welcome2:"*Per Fisheries Agency and Coast Guard port-inspection rules, we must collect passenger information at sign-up: name, date of birth, ID number, address, and phone number for customs/port declaration. For your convenience, do you agree to provide this information and use your ID number as your login account and your date of birth as your password?",
  agreeYes:"Yes, I agree and will set up my account", agreeYesSub:"Go to registration / login", agreeNo:"No, not now", confirm:"Confirm",
  understandQ:"Have you fully understood our booking rules and process?", yesUnderstand:"Yes, I understand", backStep:"Go back", notUnderstand:"I don't understand", contactCS:"Contact a real person",
  lineOA:"Official LINE Assistant", capLine:"Captain Ahao's LINE", tel:"Phone", back:"Back",
  entryTitle:"Choose your entrance", e1t:"I'm new here", e1d:"First-time sign-up: create your account", e2t:"I've been here before", e2d:"Log in: ID number + ROC birthdate", e3t:"Captain's Admin", e3d:"Staff only: trips, orders, customers, inventory",
  regTitle:"Passenger Registration (first time)", loginTitle:"Member Login", backEntry:"← Back",
  phNick:"Nickname (optional, shown on public roster)", phName:"Full name *", phGender:"Legal gender *", male:"Male", female:"Female",
  phBirth:"ROC birthdate (e.g. 80-05-12) = password *", phId:"ID / Passport / ARC number = account *", phAddr:"Address *", phPhone:"Phone *",
  minorQ:"Any companions under 18? *", has:"Yes", none:"No", countLbl:"Count", minorSelf:"I'll fill in their details", minorContact:"I'll contact you directly",
  minorN:"Minor {i} (filled by guardian {g})", phMinorName:"Name *", phMinorGender:"Gender *", phMinorBirth:"ROC birthdate (e.g. 100-03-05) *", phMinorId:"ID number *",
  minorContactNote:"We'll reach out to help register minors. You can also message our LINE assistant directly.",
  docNote:"Bring original ID documents for all passengers for Coast Guard inspection on departure day.", submitReg:"Submit & log in",
  phLoginId:"Account (ID number)", phLoginPw:"Password (ROC birthdate, e.g. 80-05-12)", login:"Log in",
  errRequired:"All fields except nickname are required.", errDup:"This ID is already registered — please log in instead.", errLogin:"Wrong account or password (account = ID, password = ROC birthdate).", errMinorSel:"Please indicate whether minors are joining.", errMinorMode:"Please choose how to handle minor registration.", errMinorFields:"Please complete every minor's details.",
  annTitle:"Announcements", annOk:"Got it — start the guide →",
  q_announce:"Read the latest news", qh_announce:"See this month's recommended trips", q_rules:"Learn the booking rules", qh_rules:"Read all 8 rules in \u201cRules\u201d", q_fees:"Know fees & rentals", qh_fees:"Check fares and gear in \u201cFees\u201d", q_calendar:"Browse the trip calendar", qh_calendar:"Open any trip in \u201cTrips\u201d", q_book:"Make your first booking", qh_book:"Tap \u201cBook now\u201d in a trip",
  rank0:"GOHO Rookie", rank1:"Deckhand", rank2:"Qualified Angler", rank3:"GOHO Veteran",
  nextRank:"{n} XP to reach \u201c{r}\u201d", maxRank:"Max rank reached!", taskN:"Quest {i} | ", doneLbl:"Done", nextStep:"← Next", nextStepB:"Next", questFooter:"Finish all quests and you're ready to set sail ⚓",
  calTitle:"Trips · {y}/{m}", listV:"List", monthV:"Month", openTripBtn:"+ Start a trip (pick a service on an open date)",
  lgAvail:"Available", lgFew:"Almost full", lgFull:"Full / cancelled", lgOpen:"▢ Open date", canOpen:"Open", cxl:"CXL", remainS:"{n} left",
  musterAt:"⏰ Muster {t}", remain:"{n} left", waitlist:"Waitlist",
  otTitle:"Start a Trip", otSub:"Anglers may initiate trips per our rules", otDate:"Departure date (24 h+ ahead only)", otSvc:"Service",
  otErrSel:"Please choose a date and a service.", otErr24:"Trips can't be started within 24 h of departure — pick a later date.", otErrDup:"A trip already exists that day — join it or pick another date.",
  otTime:"Departure time", tsCustom:"⏰ Custom preferred time", customLbl:"Preferred muster time", customNote:"Custom-time trips are marked \u201cpending captain\u2019s confirmation\u201d \u2014 the captain will confirm or adjust based on tides and route, then update here.", pendingTag:"Time pending captain\u2019s confirmation", otErrTime:"Please pick a time slot or enter a preferred muster time.",
  otInq:"\u201c{s}\u201d is inquiry-based; please confirm details with us first:", orCall:"or call {p}", otNote:"New trips start as \u201cOpen to start\u201d so others can join; under 5 sign-ups the captain may cancel by the prior day.", otCreate:"Create trip & be the first to book",
  rfTitle:"Gear Rental Fees", rfItem:"Item", rfUnit:"Unit", rfFee:"Fee", rfNote:"* Note: rental fees do NOT include rigs (terminal tackle) — purchased separately; live-bait rigs exclude bait, please bring your own.", rfOk:"Got it, continue booking",
  backCal:"← Back to trip calendar", backBottom:"← Back to trip calendar", perAngler:"/angler", seatsLeft:"{n} seats left", fullWait:"Full · waitlist",
  cardV:"Card view", tableV:"Table view", muster:"Muster", rodsUp:"Lines up", backPort:"Return",
  lTargets:"🎯 Target fish", lDepth:"🌊 Depth", lGear:"🎣 Gear advice", lRigs:"🧷 Suggested rigs", lNote:"📌 Notes", byInstr:"Per crew instructions",
  roster:"👥 Current roster (public nicknames)", rosterEmpty:"No one yet — be the first!", minorTag:"·minor",
  booked:"You're booked ✅ See \u201cMy Trips\u201d for status & reminders", notBookable:"This trip is \u201c{s}\u201d — booking closed.",
  roleLbl:"Booking as", angler:"Angler", sight:"Sightseer", rentalLbl:"Gear rental (optional)", noRental:"Bringing my own gear", feeLink:"Full rental fee table →",
  compLbl:"Sightseeing companions (optional)", compCount:"Companions", compName:"Companion {i} name *", compNote:"Companions can't rent gear; full declaration data is completed on departure day or by our staff.",
  sightBlock:"⚠ Sightseers may only join accompanied by an angler. Choose \u201cAngler\u201d and add companions below.",
  minorSeat:"👶 Seats will also be booked for your {n} registered minor(s).", bookBtn:"Book now", waitBtn:"Join waitlist", policy:"One person per booking · cancellations within 7 days of departure owe full fare",
  myTitle:"My Trips", myEmpty1:"No bookings yet", myEmpty2:"Pick a trip you like in \u201cTrips\u201d!", idLbl:"Role", paid:"✅ Paid", unpaid:"🕐 Unpaid", rentShort:"Rental", remTitle:"🔔 Pre-departure reminders", viewFull:"Full trip details →", stCxl:"Cancelled", stWait:"Waitlisted",
  rm1:"Bring original ID for Coast Guard inspection", rm2:"Arrive at Anping Harbor 15 min early", rm3:"Take seasickness medicine 30 min before if prone", rm4:"Live-bait trips: bring your own bait feed", rm5:"Nearshore trips: no cooler needed, shared fish cooler onboard", rm6:"Mid-range/island trips: bring water and snacks", rm7:"Mark your catch with a cable tie",
  feeTitle:"Common Trip Fares (per angler)", rentTitle:"Gear Rental (rigs/consumables extra)", feeFoot:"Rentals exclude rigs; live-bait trips need your own bait; nearshore trips need no cooler.",
  ctTitle:"Contact Us", newsTitle:"Latest News", noPosts:"No announcements yet.",
  depart:"Departure: Anping Harbor", openMap:"Open map →",
  navQuest:"Guide", navCal:"Trips", navMy:"Mine", navRules:"Rules", navFees:"Fees", navSvc:"Services", navCt:"Contact",
  menu:"☰ Menu", mMy:"🎣 My bookings", mHome:"🏠 Home", mOut:"🚪 Log out", adminBtn:"Admin",
  hello:"⚓ Welcome aboard, {n}", welcomeNew:"Welcome aboard, {n}!", welcomeBack:"Welcome back, {n}!",
  tSightOnly:"Sightseers must be accompanied by an angler.", tWait:"Not enough seats — you (and companions) are waitlisted.", tBooked:"Booked! +30 XP 🎉", tBookedX:" (incl. {n} companions)", tCreated:"Trip created as \u201cOpen to start\u201d — book now to be first 🎣", tPw:"Wrong password",
  rzS:"nearshore gear/rigs", rzD:"deep-water gear/rigs", rzTagS:"Nearshore", rzTagD:"Deep", rzTagB:"All-purpose",
  autoNote:"Depth {d} — showing {z} rentals automatically (80 m threshold)", manualNote:"⚓ Rentals for this trip are specified by the captain", gvAcc:"Stack", gvTab:"Tabs", gvGrid:"Grid",
  status:{"歡迎開班":"Open to start","報名中":"Signing up","確定出船":"Confirmed","已完成":"Completed","人數不足取消":"Cancelled (low turnout)","天氣因素取消":"Cancelled (weather)","不可抗力因素取消":"Cancelled (force majeure)"},
  zone:{"近海":"Nearshore","中遠程":"Mid-range","離島":"Islands","專案":"Custom","休閒":"Leisure"}, inq:"Inquire",
},
ja: {
  brandSub:"台南・安平出港、合法遊漁船、ご予約歓迎", loading:"⚓ 読み込み中…", gate:"まず上のシステムメッセージを完了してください",
  sysTitle:"システムメッセージ", promptTitle:"お知らせ", csTitle:"有人サポート",
  welcome1:"五賀遊漁船へようこそ！＊本船は政府認可の営業許可と保険を取得した合法遊漁船です。乗船登録には身分証、健康保険証、運転免許証、パスポート、居留証がご利用いただけます。＊",
  welcome2:"＊漁業署および海巡署の出入港検査規定に基づき、お申し込み時に乗船者情報（氏名・生年月日・身分証番号・住所・電話番号）のご提供をお願いしています。今後のご利用を便利にするため、上記情報の提供に同意し、身分証番号をログインID、生年月日（民国暦）をパスワードとして設定しますか？",
  agreeYes:"はい、同意してアカウントを設定します", agreeYesSub:"登録／ログインページへ", agreeNo:"いいえ、今は同意しません", confirm:"確認",
  understandQ:"当船の予約ルールと流れを十分にご理解いただけましたか？", yesUnderstand:"はい、理解しました", backStep:"前に戻る", notUnderstand:"わかりません", contactCS:"スタッフに連絡する",
  lineOA:"公式LINEに登録", capLine:"アハオ船長のLINE", tel:"電話", back:"戻る",
  entryTitle:"入口をお選びください", e1t:"初めての方", e1d:"新規登録：乗船者情報を入力してアカウント作成", e2t:"利用したことがある方", e2d:"ログイン：証明書番号＋民国暦の生年月日", e3t:"船長管理画面", e3d:"管理者専用：便・注文・顧客・在庫",
  regTitle:"乗船者登録フォーム（初回）", loginTitle:"ログイン", backEntry:"← 入口へ戻る",
  phNick:"ニックネーム（任意・公開名簿に表示）", phName:"氏名 *", phGender:"法的性別 *", male:"男性", female:"女性",
  phBirth:"民国暦の生年月日（例：80-05-12）＝パスワード *", phId:"身分証／パスポート／居留証番号＝ID *", phAddr:"住所 *", phPhone:"電話番号 *",
  minorQ:"18歳未満の同行者はいますか？ *", has:"いる", none:"いない", countLbl:"人数", minorSelf:"私が代理で入力します", minorContact:"こちらから連絡します",
  minorN:"未成年 {i}（保護者 {g} が代理入力）", phMinorName:"氏名 *", phMinorGender:"性別 *", phMinorBirth:"民国暦生年月日（例：100-03-05）*", phMinorId:"証明書番号 *",
  minorContactNote:"未成年の方のご登録はこちらからご連絡してサポートします。公式LINEへ直接ご連絡も可能です。",
  docNote:"出港当日は全乗船者の証明書原本をご持参ください（海巡署の検査があります）。", submitReg:"送信してログイン",
  phLoginId:"ID（証明書番号）", phLoginPw:"パスワード（民国暦生年月日、例：80-05-12）", login:"ログイン",
  errRequired:"ニックネーム以外はすべて必須です。", errDup:"この番号は登録済みです。ログインをご利用ください。", errLogin:"IDまたはパスワードが違います（ID＝証明書番号、パスワード＝民国暦生年月日）。", errMinorSel:"未成年同行の有無を選択してください。", errMinorMode:"未成年の登録方法を選択してください。", errMinorFields:"未成年全員の情報を入力してください。",
  annTitle:"最新のお知らせ", annOk:"了解、ガイドを始める →",
  q_announce:"最新のお知らせを見る", qh_announce:"今月のおすすめ便をチェック", q_rules:"予約ルールを読む", qh_rules:"「規定」で8つのルールを確認", q_fees:"料金とレンタルを知る", qh_fees:"「料金」で船代と装備を確認", q_calendar:"便カレンダーを見る", qh_calendar:"「便」でどれか1つ開く", q_book:"初めての予約をする", qh_book:"便詳細で「予約する」をタップ",
  rank0:"五賀ビギナー", rank1:"見習い船員", rank2:"一人前の釣り師", rank3:"五賀ベテラン",
  nextRank:"あと {n} XP で「{r}」に昇格", maxRank:"最高ランク達成！", taskN:"クエスト {i}｜", doneLbl:"完了", nextStep:"← 次へ", nextStepB:"次へ", questFooter:"全クエスト完了＝予約の流れはバッチリ、安心して出港できます ⚓",
  calTitle:"{y}年{m}月の便", listV:"リスト", monthV:"カレンダー", openTripBtn:"＋ 便を立てる（空き日にサービスを選んで開催）",
  lgAvail:"空きあり", lgFew:"残りわずか", lgFull:"満席／中止", lgOpen:"▢ 開催可能", canOpen:"開催可", cxl:"中止", remainS:"残{n}",
  musterAt:"⏰ {t} 集合", remain:"残り {n} 名", waitlist:"キャンセル待ち",
  otTitle:"便を立てる", otSub:"規定により釣り客が便を発起できます", otDate:"出発日（24時間以降のみ）", otSvc:"サービス",
  otErrSel:"日付とサービスを選択してください。", otErr24:"出発24時間前を切った開催はできません。別の日をお選びください。", otErrDup:"その日はすでに便があります。参加するか別の日をお選びください。",
  otTime:"出発時間帯", tsCustom:"⏰ 希望時間を指定", customLbl:"希望集合時間", customNote:"時間指定の便は「船長確認待ち」となります。潮汐と航程を踏まえて船長が確定・調整し、こちらに反映されます。", pendingTag:"時間は船長確認待ち", otErrTime:"時間帯を選ぶか、希望集合時間を入力してください。",
  otInq:"「{s}」は要相談サービスです。詳細はスタッフにご確認ください：", orCall:"またはお電話 {p}", otNote:"開催後のステータスは「開催者募集中」となり、他の釣り客が参加できます。5名未満の場合、船長が前日までに中止することがあります。", otCreate:"便を作成して最初の予約者になる",
  rfTitle:"装備レンタル料金", rfItem:"項目", rfUnit:"単位", rfFee:"料金", rfNote:"※ 注意：レンタル料金に仕掛けは含まれません（別売り）。活き餌仕掛けに餌は含まれないため、餌はご持参ください。", rfOk:"了解、予約を続ける",
  backCal:"← 便カレンダーへ戻る", backBottom:"← 前のページへ戻る（便カレンダー）", perAngler:"／釣り客", seatsLeft:"残り {n} 名", fullWait:"満席・キャンセル待ち可",
  cardV:"カード表示", tableV:"表形式", muster:"集合", rodsUp:"納竿", backPort:"帰港",
  lTargets:"🎯 対象魚", lDepth:"🌊 水深", lGear:"🎣 タックル推奨", lRigs:"🧷 推奨仕掛け", lNote:"📌 備考", byInstr:"現場の指示に従う",
  roster:"👥 予約者リスト（公開ニックネーム）", rosterEmpty:"まだ予約がありません。一番乗りしましょう！", minorTag:"・未成年",
  booked:"予約済みです ✅ 「マイ便」で状況とリマインダーを確認できます", notBookable:"この便は「{s}」のため、現在予約できません。",
  roleLbl:"予約区分", angler:"釣り客", sight:"観光", rentalLbl:"装備レンタル（任意）", noRental:"持参します（レンタルなし）", feeLink:"レンタル料金表を見る →",
  compLbl:"観光同伴者（任意）", compCount:"同伴人数", compName:"同伴者 {i} の氏名 *", compNote:"同伴者はレンタル不可。正式な申告情報は当日証明書持参にて、またはスタッフが代理登録します。",
  sightBlock:"⚠ 観光は釣り客の同伴のみ可能で、単独予約はできません。「釣り客」を選び、下の「観光同伴者」で追加してください。",
  minorSeat:"👶 代理入力済みの未成年 {n} 名の座席も同時に確保されます。", bookBtn:"予約する", waitBtn:"キャンセル待ちに登録", policy:"1名1件・出発7日以内のキャンセルは全額のお支払いが必要です",
  myTitle:"マイ便", myEmpty1:"まだ予約がありません", myEmpty2:"「便」タブでお好きな便を選びましょう！", idLbl:"区分", paid:"✅ 支払済", unpaid:"🕐 未払い", rentShort:"レンタル", remTitle:"🔔 出発前リマインダー", viewFull:"便の詳細を見る →", stCxl:"キャンセル済", stWait:"キャンセル待ち",
  rm1:"証明書原本をご持参ください（海巡署検査）", rm2:"安平港の集合場所に15分前到着", rm3:"船酔いしやすい方は30分前に酔い止めを", rm4:"活き餌便は餌をご持参ください", rm5:"近海便はクーラー不要（船に共用魚クーラーあり）", rm6:"中遠距離・離島便は飲料水と軽食をご持参ください", rm7:"釣果は結束バンドで目印を",
  feeTitle:"主な船代（釣り客1名あたり）", rentTitle:"装備レンタル料金（仕掛け・消耗品別）", feeFoot:"レンタルに仕掛けは含まれません。活き餌便は餌持参、近海便は氷不要です。",
  ctTitle:"お問い合わせ", newsTitle:"最新のお知らせ", noPosts:"現在お知らせはありません。",
  depart:"出発地：安平港", openMap:"地図を開く →",
  navQuest:"ガイド", navCal:"便", navMy:"マイ", navRules:"規定", navFees:"料金", navSvc:"サービス", navCt:"連絡",
  menu:"☰ メニュー", mMy:"🎣 予約を確認", mHome:"🏠 ホームへ", mOut:"🚪 ログアウト", adminBtn:"管理画面",
  hello:"⚓ {n} さん、ようこそ", welcomeNew:"ようこそ、{n} さん！", welcomeBack:"おかえりなさい、{n} さん！",
  tSightOnly:"観光は釣り客の同伴が必要です。", tWait:"座席が足りないため（同行者含め）キャンセル待ちに登録しました。", tBooked:"予約完了！＋30 XP 🎉", tBookedX:"（同行 {n} 名含む）", tCreated:"便を作成しました！「開催者募集中」です。今すぐ予約して一番乗りに 🎣", tPw:"パスワードが違います",
  rzS:"近海用タックル", rzD:"深場用タックル", rzTagS:"近海", rzTagD:"深場", rzTagB:"共通",
  autoNote:"この便の水深は {d} のため「{z}」のレンタル品を自動表示します（基準80m）", manualNote:"⚓ この便のレンタル品は船長指定です", gvAcc:"リスト", gvTab:"タブ", gvGrid:"一覧",
  status:{"歡迎開班":"開催者募集中","報名中":"予約受付中","確定出船":"出船確定","已完成":"終了","人數不足取消":"人数不足のため中止","天氣因素取消":"天候不良のため中止","不可抗力因素取消":"不可抗力のため中止"},
  zone:{"近海":"近海","中遠程":"中遠距離","離島":"離島","專案":"特別企画","休閒":"レジャー"}, inq:"要相談",
},
};
const t = (k, vars) => {
  let str = (I18N[LANG] && I18N[LANG][k]) ?? I18N.zh[k] ?? k;
  if (typeof str === "string" && vars) for (const [key, v] of Object.entries(vars)) str = str.split("{" + key + "}").join(v);
  return str;
};
const tSt = (status) => (I18N[LANG].status && I18N[LANG].status[status]) || status;
const tZone = (z) => (I18N[LANG].zone && I18N[LANG].zone[z]) || z;
const tPrice = (p) => (p === "洽詢" ? t("inq") : p);
/* 後台多語內容：lx(物件, 欄位) → 依目前語言取 欄位_en / 欄位_ja，未填則退回中文欄位 */
const lx = (o, f) => (LANG !== "zh" && o && o[f + "_" + LANG] ? o[f + "_" + LANG] : (o ? o[f] : ""));
/* 釣場水深自動判定：取水深字串內最大數字，≤80m→近海、>80m→深場、無法判讀→null(顯示全部) */
/* 各班別建議出發時段（開班快選用） */
const TIME_PRESETS = {
  "近海": [
    { muster: "05:00", rodsUp: "12:00", back: "約 13:00" },
    { muster: "16:00", rodsUp: "23:00", back: "約 24:00" },
  ],
  "中遠程": [{ muster: "04:00", rodsUp: "14:00", back: "16–17:00" }],
  "離島": [{ muster: "03:30", rodsUp: "14:00", back: "17:00" }],
  "專案": [], "休閒": [],
};
const DEPTH_THRESHOLD = 80; // 釣場水深判定基準（米）
const depthClass = (trip) => {
  const nums = String(trip?.depth || "").match(/\d+/g);
  if (!nums || !nums.length) return null;
  return Math.max(...nums.map(Number)) > DEPTH_THRESHOLD ? "deep" : "shallow";
};
const rentalsForTrip = (trip, rentals) => {
  if (trip && trip.rentalMode === "manual") {
    const names = trip.rentalNames || [];
    return { list: rentals.filter((r) => names.includes(r.name)), mode: "manual", cls: null };
  }
  const cls = trip ? depthClass(trip) : null;
  return { list: rentals.filter((r) => rentalMatch(r, cls)), mode: cls === null ? "all" : "auto", cls };
};
const rentalMatch = (r, cls) => {
  if (cls === null) return true;
  const z = r.zone || "通用";
  if (z === "通用") return true;
  return cls === "deep" ? z === "深場" : z === "近海";
};
const lxArr = (o, f) => (LANG !== "zh" && o && o[f + "_" + LANG] && o[f + "_" + LANG].length ? o[f + "_" + LANG] : (o ? o[f] || [] : []));
const lxAnn = (a) => (typeof a === "string" ? a : (LANG !== "zh" && a[LANG]) || a.zh || "");

/* ---------- 儲存層（含舊資料合併升級） ---------- */
function seedDb() {
  return { trips: DEFAULT_TRIPS, orders: [], customers: [], inventory: DEFAULT_INVENTORY, quest: {}, announcement: { zh: DEFAULT_ANNOUNCEMENT, en: "", ja: "" }, rules: DEFAULT_RULES, posts: DEFAULT_POSTS, contactPage: DEFAULT_CONTACT_PAGE, rentals: RENTALS, pricing: PRICING };
}
async function loadStore() {
  try {
    const r = await window.storage.get("goho-erp-v2");
    if (r?.value) {
      const d = JSON.parse(r.value);
      const base = seedDb();
      const normAnn = (a) => (typeof a === "string" ? { zh: a, en: "", ja: "" } : a);
      return {
        trips: (d.trips || base.trips).map((t) => ({ status: "報名中", ...t })),
        orders: d.orders || [], customers: d.customers || [], inventory: d.inventory || base.inventory,
        quest: d.quest || {}, announcement: normAnn(d.announcement || base.announcement), rules: d.rules || base.rules,
        posts: d.posts || base.posts, contactPage: d.contactPage || base.contactPage, rentals: (d.rentals || base.rentals).map((r) => ({ zone: "近海", ...r })), pricing: d.pricing || base.pricing,
      };
    }
  } catch (e) { /* 首次使用 */ }
  return seedDb();
}
async function saveStore(db) { try { await window.storage.set("goho-erp-v2", JSON.stringify(db)); } catch (e) { console.error(e); } }

/* ---------- 小元件 ---------- */
const Btn = ({ children, onClick, kind = "primary", full, small, disabled }) => {
  const base = { primary: { background: C.orange, color: C.navyDeep }, ghost: { background: "transparent", color: C.sand, border: "1.5px solid rgba(247,243,236,.35)" }, teal: { background: C.teal, color: "#fff" }, danger: { background: C.red, color: "#fff" }, light: { background: "#fff", color: C.navy, border: "1.5px solid #0C2D4822" }, navy: { background: C.navy, color: C.sand } }[kind];
  return (
    <button onClick={onClick} disabled={disabled} style={{ ...base, opacity: disabled ? 0.45 : 1 }}
      className={`${full ? "w-full" : ""} ${small ? "px-3 py-1.5 text-sm" : "px-5 py-3"} rounded-xl font-bold tracking-wide transition-transform active:scale-95`}>
      {children}
    </button>
  );
};
const Tag = ({ children, tone = "teal" }) => {
  const bg = { teal: C.teal, orange: C.orange, navy: C.navy, red: C.red, yellow: C.yellow, gray: C.gray }[tone];
  const fg = tone === "yellow" || tone === "orange" ? C.navyDeep : "#fff";
  return <span style={{ background: bg, color: fg }} className="px-2 py-0.5 rounded-full text-xs font-bold whitespace-nowrap">{children}</span>;
};
const Wave = ({ fill }) => (
  <svg viewBox="0 0 1440 60" className="w-full block" preserveAspectRatio="none" style={{ height: 24 }}>
    <path d="M0,30 C240,60 480,0 720,30 C960,60 1200,0 1440,30 L1440,60 L0,60 Z" fill={fill || C.sand} />
  </svg>
);

/* 五賀 Logo：救生圈徽章 */
const LOGO_SRC = "data:image/jpeg;base64,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";
const Logo = ({ size = 48 }) => (
  <img src={LOGO_SRC} width={size} height={size} alt="五賀娛樂漁船" className="rounded-full" style={{ boxShadow: "0 0 0 2px #FFFFFF55" }} />
);
const LogoOld = ({ size = 46 }) => (
  <svg width={size} height={size} viewBox="0 0 100 100">
    <circle cx="50" cy="50" r="46" fill="#FFFFFF" />
    <circle cx="50" cy="50" r="46" fill="none" stroke={C.orange} strokeWidth="14" strokeDasharray="24 12.2" strokeLinecap="butt" transform="rotate(9 50 50)" />
    <circle cx="50" cy="50" r="34" fill={C.navy} />
    <path d="M22 62 C32 54, 40 70, 50 62 C60 54, 68 70, 78 62 L78 84 L22 84 Z" fill={C.teal} opacity="0.9" />
    <text x="50" y="56" textAnchor="middle" fontSize="26" fontWeight="900" fill="#FFFFFF" fontFamily="'Noto Sans TC',sans-serif">五賀</text>
  </svg>
);

/* 民國出生年月日正規化：去除民國/年/月/日字樣與空白，統一分隔為 - */
const normBirth = (s) => (s || "").replace(/[民國年月日\s]/g, "-").replace(/[\/.．]/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
function ModalShell({ children, onClose }) {
  return (
    <div className="fixed inset-0 z-50 overflow-y-auto" style={{ background: "rgba(8,29,48,.82)", backdropFilter: "blur(3px)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="min-h-full flex items-start justify-center p-3">
        <div className="w-full max-w-md rounded-2xl overflow-hidden shadow-2xl my-2" style={{ background: C.sand }} onClick={(e) => e.stopPropagation()}>{children}</div>
      </div>
    </div>
  );
}

/* ---------- 入口流程 ---------- */
function WelcomeModal({ onYes, onNo }) {
  const [agree, setAgree] = useState(null);
  return (
    <ModalShell>
      <div style={{ background: C.navy }} className="px-5 pt-5 pb-3 text-center">
        <div className="text-3xl mb-1">⚓</div>
        <h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("sysTitle")}</h2>
        <p className="text-xs mt-1" style={{ color: "#F7F3EC99" }}>GOHO 五賀娛樂漁船・安平港</p>
      </div>
      <Wave />
      <div className="px-5 pb-5">
        <p className="text-sm leading-relaxed mb-3" style={{ color: C.navyDeep }}>{t("welcome1")}</p>
        <p className="text-sm leading-relaxed mb-4" style={{ color: C.navyDeep }}>{t("welcome2")}</p>
        <div className="space-y-2 mb-4">
          <label className="flex items-start gap-3 p-3 rounded-xl cursor-pointer border-2" style={{ borderColor: agree === true ? C.teal : "#0C2D4822", background: agree === true ? "#1EA89615" : "#fff" }}>
            <input type="radio" name="consent" checked={agree === true} onChange={() => setAgree(true)} className="mt-1 accent-teal-600" />
            <span className="text-sm font-bold" style={{ color: C.navy }}>{t("agreeYes")}<span className="block font-normal text-xs mt-0.5" style={{ color: "#0C2D4899" }}>{t("agreeYesSub")}</span></span>
          </label>
          <label className="flex items-start gap-3 p-3 rounded-xl cursor-pointer border-2" style={{ borderColor: agree === false ? C.red : "#0C2D4822", background: agree === false ? "#E76F5112" : "#fff" }}>
            <input type="radio" name="consent" checked={agree === false} onChange={() => setAgree(false)} className="mt-1 accent-red-500" />
            <span className="text-sm font-bold" style={{ color: C.navy }}>{t("agreeNo")}</span>
          </label>
        </div>
        <Btn full disabled={agree === null} onClick={() => (agree ? onYes() : onNo())}>{t("confirm")}</Btn>
      </div>
    </ModalShell>
  );
}
function NoConsentModal({ onBack, onContact }) {
  return (
    <ModalShell>
      <div style={{ background: C.navy }} className="px-5 pt-5 pb-3 text-center"><div className="text-3xl mb-1">🧭</div><h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("promptTitle")}</h2></div>
      <Wave />
      <div className="px-5 pb-5">
        <p className="text-base font-bold text-center leading-relaxed mb-5" style={{ color: C.navyDeep }}>{t("understandQ")}</p>
        <div className="grid grid-cols-2 gap-3">
          <Btn kind="teal" onClick={onBack}>{t("yesUnderstand")}<span className="block text-xs font-normal opacity-90">{t("backStep")}</span></Btn>
          <Btn kind="danger" onClick={onContact}>{t("notUnderstand")}<span className="block text-xs font-normal opacity-90">{t("contactCS")}</span></Btn>
        </div>
      </div>
    </ModalShell>
  );
}
function ContactModal({ onBack }) {
  return (
    <ModalShell>
      <div style={{ background: C.navy }} className="px-5 pt-5 pb-3 text-center"><div className="text-3xl mb-1">💬</div><h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("csTitle")}</h2></div>
      <Wave />
      <div className="px-5 pb-5 space-y-3">
        <a href={CONTACT.lineOA} target="_blank" rel="noreferrer" className="block p-4 rounded-xl font-bold text-center" style={{ background: "#06C755", color: "#fff" }}>{t("lineOA")}</a>
        <a href={CONTACT.captainLine} target="_blank" rel="noreferrer" className="block p-4 rounded-xl font-bold text-center" style={{ background: C.teal, color: "#fff" }}>{t("capLine")}</a>
        <div className="text-center text-sm font-bold" style={{ color: C.navy }}>{t("tel")}：{CONTACT.captain}</div>
        <Btn kind="light" full onClick={onBack}>{t("back")}</Btn>
      </div>
    </ModalShell>
  );
}

/* 登入後第一個視窗：最新公告 */
function AnnouncementModal({ text, onClose }) {
  return (
    <ModalShell onClose={onClose}>
      <div style={{ background: `linear-gradient(150deg, ${C.navy}, ${C.tealDark})` }} className="px-5 pt-5 pb-3 text-center">
        <div className="text-3xl mb-1">📣</div><h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("annTitle")}</h2>
      </div>
      <Wave />
      <div className="px-5 pb-5">
        <div className="rounded-xl p-4 mb-4 whitespace-pre-wrap text-sm leading-relaxed font-medium" style={{ background: "#fff", border: `2px solid ${C.orange}`, color: C.navyDeep }}>{lxAnn(text)}</div>
        <Btn full onClick={onClose}>{t("annOk")}</Btn>
      </div>
    </ModalShell>
  );
}

/* ---------- 註冊 / 登入（含船長後台入口 + 未成年同行） ---------- */
function RegisterModal({ customers, onRegister, onLogin, onAdmin }) {
  const [mode, setMode] = useState(null); // null=入口選擇 | register | login
  const [f, setF] = useState({ nickname: "", name: "", gender: "", birth: "", idno: "", address: "", phone: "" });
  const [login, setLogin] = useState({ idno: "", birth: "" });
  const [hasMinor, setHasMinor] = useState(null);     // null | true | false
  const [minorMode, setMinorMode] = useState("");     // "self" 代填 | "contact" 主動聯絡
  const [minorCount, setMinorCount] = useState(1);
  const [minors, setMinors] = useState([]);
  const [err, setErr] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const inputCls = "w-full p-3 rounded-xl border-2 text-sm outline-none focus:border-teal-600";
  const inputStyle = { borderColor: "#0C2D4822", background: "#fff", color: C.navyDeep };

  const syncMinors = (n) => {
    setMinorCount(n);
    setMinors((prev) => Array.from({ length: n }, (_, i) => prev[i] || { name: "", gender: "", birth: "", idno: "" }));
  };
  const setMinorField = (i, k, v) => setMinors((prev) => prev.map((m, idx) => (idx === i ? { ...m, [k]: v } : m)));

  const submit = () => {
    if (!f.name || !f.gender || !f.birth || !f.idno || !f.address || !f.phone) { setErr(t("errRequired")); return; }
    if (customers.some((c) => c.idno === f.idno)) { setErr(t("errDup")); return; }
    if (hasMinor === null) { setErr(t("errMinorSel")); return; }
    let minorList = [];
    if (hasMinor) {
      if (!minorMode) { setErr(t("errMinorMode")); return; }
      if (minorMode === "self") {
        for (const m of minors.slice(0, minorCount)) {
          if (!m.name || !m.gender || !m.birth || !m.idno) { setErr(t("errMinorFields")); return; }
        }
        minorList = minors.slice(0, minorCount);
      }
    }
    onRegister({
      ...f, id: "C" + Date.now(), createdAt: new Date().toISOString(),
      minorMode: hasMinor ? minorMode : "none", minorCount: hasMinor ? minorCount : 0, minors: minorList,
    });
  };
  const doLogin = () => {
    const c = customers.find((x) => x.idno === login.idno && normBirth(x.birth) === normBirth(login.birth));
    if (!c) { setErr(t("errLogin")); return; }
    onLogin(c);
  };

  /* 入口選擇畫面：三種入口 */
  if (mode === null) {
    const entries = [
      { icon: "🐚", title: t("e1t"), desc: t("e1d"), action: () => setMode("register"), bg: C.teal },
      { icon: "🐟", title: t("e2t"), desc: t("e2d"), action: () => setMode("login"), bg: C.navy },
      { icon: "⚓", title: t("e3t"), desc: t("e3d"), action: onAdmin, bg: C.orange },
    ];
    return (
      <ModalShell>
        <div style={{ background: `linear-gradient(150deg, ${C.navy}, ${C.tealDark})` }} className="px-5 pt-5 pb-3 text-center">
          <div className="text-3xl mb-1">🚢</div>
          <h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("entryTitle")}</h2>
          <p className="text-xs mt-1" style={{ color: "#F7F3EC99" }}>GOHO 五賀娛樂漁船</p>
        </div>
        <Wave />
        <div className="px-5 pb-5 space-y-3">
          {entries.map((e) => (
            <button key={e.title} onClick={e.action} className="w-full text-left p-4 rounded-2xl flex items-center gap-4 transition-transform active:scale-[.98]" style={{ background: e.bg, color: e.bg === C.orange ? C.navyDeep : "#fff" }}>
              <span className="text-3xl">{e.icon}</span>
              <span>
                <span className="block font-black text-base">{e.title}</span>
                <span className="block text-xs mt-0.5 opacity-90">{e.desc}</span>
              </span>
              <span className="ml-auto text-xl font-black">→</span>
            </button>
          ))}
        </div>
      </ModalShell>
    );
  }

  return (
    <ModalShell>
      <div style={{ background: C.navy }} className="px-5 pt-5 pb-3 text-center">
        <button onClick={() => { setMode(null); setErr(""); }} className="absolute-none text-xs px-3 py-1 rounded-full font-bold float-left" style={{ background: "#F7F3EC18", color: C.sand }}>{t("backEntry")}</button>
        <div className="text-3xl mb-1">{mode === "register" ? "📝" : "🐟"}</div>
        <h2 className="font-black text-lg" style={{ color: C.yellow }}>{mode === "register" ? t("regTitle") : t("loginTitle")}</h2>
      </div>
      <Wave />
      <div className="px-5 pb-5 space-y-2.5">
        {mode === "register" ? (
          <>
            <input className={inputCls} style={inputStyle} placeholder={t("phNick")} value={f.nickname} onChange={set("nickname")} />
            <input className={inputCls} style={inputStyle} placeholder={t("phName")} value={f.name} onChange={set("name")} />
            <select className={inputCls} style={inputStyle} value={f.gender} onChange={set("gender")}><option value="">{t("phGender")}</option><option value="男">{t("male")}</option><option value="女">{t("female")}</option></select>
            <input className={inputCls} style={inputStyle} placeholder={t("phBirth")} value={f.birth} onChange={set("birth")} />
            <input className={inputCls} style={inputStyle} placeholder={t("phId")} value={f.idno} onChange={set("idno")} />
            <input className={inputCls} style={inputStyle} placeholder={t("phAddr")} value={f.address} onChange={set("address")} />
            <input className={inputCls} style={inputStyle} placeholder={t("phPhone")} value={f.phone} onChange={set("phone")} />

            {/* 未成年同行 */}
            <div className="rounded-xl p-3 border-2" style={{ borderColor: "#0C2D4822", background: "#fff" }}>
              <div className="text-sm font-bold mb-2" style={{ color: C.navy }}>{t("minorQ")}</div>
              <div className="flex gap-2 mb-2">
                {[[t("has"), true], [t("none"), false]].map(([l, v]) => (
                  <button key={l} onClick={() => { setHasMinor(v); setErr(""); }} className="flex-1 py-2 rounded-xl text-sm font-bold" style={{ background: hasMinor === v ? C.navy : "#0C2D4810", color: hasMinor === v ? C.sand : C.navy }}>{l}</button>
                ))}
              </div>
              {hasMinor && (
                <div className="space-y-2">
                  <div className="flex items-center gap-2">
                    <span className="text-sm" style={{ color: C.navy }}>{t("countLbl")}</span>
                    <select className="p-2 rounded-lg border-2 text-sm" style={inputStyle} value={minorCount} onChange={(e) => syncMinors(Number(e.target.value))}>
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </div>
                  <div className="flex gap-2">
                    {[["self", t("minorSelf")], ["contact", t("minorContact")]].map(([k, l]) => (
                      <button key={k} onClick={() => { setMinorMode(k); if (k === "self") syncMinors(minorCount); }} className="flex-1 py-2 rounded-xl text-xs font-bold" style={{ background: minorMode === k ? C.teal : "#0C2D4810", color: minorMode === k ? "#fff" : C.navy }}>{l}</button>
                    ))}
                  </div>
                  {minorMode === "self" && minors.slice(0, minorCount).map((m, i) => (
                    <div key={i} className="rounded-lg p-2.5 space-y-2" style={{ background: "#F7F3EC", border: "1px dashed #0C2D4833" }}>
                      <div className="text-xs font-bold" style={{ color: C.tealDark }}>{t("minorN", { i: i + 1, g: f.name || "—" })}</div>
                      <input className="w-full p-2 rounded-lg border text-sm" style={inputStyle} placeholder={t("phMinorName")} value={m.name} onChange={(e) => setMinorField(i, "name", e.target.value)} />
                      <div className="flex gap-2">
                        <select className="flex-1 p-2 rounded-lg border text-sm" style={inputStyle} value={m.gender} onChange={(e) => setMinorField(i, "gender", e.target.value)}><option value="">{t("phMinorGender")}</option><option value="男">{t("male")}</option><option value="女">{t("female")}</option></select>
                        <input className="flex-1 p-2 rounded-lg border text-sm" style={inputStyle} placeholder={t("phMinorBirth")} value={m.birth} onChange={(e) => setMinorField(i, "birth", e.target.value)} />
                      </div>
                      <input className="w-full p-2 rounded-lg border text-sm" style={inputStyle} placeholder={t("phMinorId")} value={m.idno} onChange={(e) => setMinorField(i, "idno", e.target.value)} />
                    </div>
                  ))}
                  {minorMode === "contact" && (
                    <div className="rounded-lg p-2.5 text-xs" style={{ background: "#FFD16622", color: C.navyDeep }}>
                      {t("minorContactNote")} <a href={CONTACT.lineOA} target="_blank" rel="noreferrer" className="underline font-bold">LINE →</a>
                    </div>
                  )}
                </div>
              )}
            </div>

            <p className="text-xs leading-relaxed" style={{ color: "#0C2D4899" }}>{t("docNote")}</p>
            {err && <p className="text-xs font-bold" style={{ color: C.red }}>{err}</p>}
            <Btn full onClick={submit}>{t("submitReg")}</Btn>
          </>
        ) : (
          <>
            <input className={inputCls} style={inputStyle} placeholder={t("phLoginId")} value={login.idno} onChange={(e) => setLogin({ ...login, idno: e.target.value })} />
            <input className={inputCls} style={inputStyle} placeholder={t("phLoginPw")} value={login.birth} onChange={(e) => setLogin({ ...login, birth: e.target.value })} />
            {err && <p className="text-xs font-bold" style={{ color: C.red }}>{err}</p>}
            <Btn full onClick={doLogin}>{t("login")}</Btn>
          </>
        )}

      </div>
    </ModalShell>
  );
}

/* ---------- 遊戲化導引 + 提示浮標 ---------- */
const QUESTS = [
  { id: "announce", icon: "📣", title: "看過最新公告", xp: 15, hint: "了解本月船長推薦班別", tab: "quest" },
  { id: "rules", icon: "📜", title: "讀懂報班規定", xp: 20, hint: "前往「規定」看完 8 條規定", tab: "rules" },
  { id: "fees", icon: "💰", title: "認識費用與租借", xp: 15, hint: "前往「費用」了解船資與裝備", tab: "fees" },
  { id: "calendar", icon: "📅", title: "查一次船班行事曆", xp: 20, hint: "前往「船班」點開任一船班", tab: "calendar" },
  { id: "book", icon: "🎣", title: "完成第一次報名", xp: 30, hint: "在船班詳情按下「我要報名」", tab: "calendar" },
];
const RANKS = [{ min: 0, key: "rank0", icon: "🐚" }, { min: 35, key: "rank1", icon: "🦀" }, { min: 70, key: "rank2", icon: "🐟" }, { min: 100, key: "rank3", icon: "🏆" }];

function nextQuest(quest) { return QUESTS.find((q) => !quest[q.id]); }

function QuestBoard({ quest, gotoTab }) {
  const xp = QUESTS.filter((q) => quest[q.id]).reduce((s, q) => s + q.xp, 0);
  const rank = [...RANKS].reverse().find((r) => xp >= r.min);
  const next = RANKS.find((r) => r.min > xp);
  const nq = nextQuest(quest);
  return (
    <div className="space-y-4">
      <div className="rounded-2xl p-5 text-center" style={{ background: `linear-gradient(160deg, ${C.navy}, ${C.tealDark})` }}>
        <div className="text-4xl">{rank.icon}</div>
        <div className="font-black text-xl mt-1" style={{ color: C.yellow }}>{t(rank.key)}</div>
        <div className="text-xs mt-1" style={{ color: "#F7F3EC99" }}>{next ? t("nextRank", { n: next.min - xp, r: t(next.key) }) : t("maxRank")}</div>
        <div className="mt-3 h-3 rounded-full overflow-hidden" style={{ background: "#081D30" }}>
          <div className="h-full rounded-full transition-all duration-700" style={{ width: `${Math.min(100, xp)}%`, background: `linear-gradient(90deg, ${C.teal}, ${C.yellow})` }} />
        </div>
        <div className="text-xs mt-1 font-bold" style={{ color: C.sand }}>{Math.min(100, xp)} / 100 XP</div>
      </div>
      <div className="space-y-2.5">
        {QUESTS.map((q, i) => {
          const done = !!quest[q.id];
          const isNext = nq && nq.id === q.id;
          return (
            <button key={q.id} onClick={() => gotoTab(q.tab)} className="w-full text-left p-4 rounded-2xl flex items-center gap-3 border-2 transition-transform active:scale-[.98] relative"
              style={{ background: done ? "#1EA89614" : "#fff", borderColor: isNext ? C.orange : done ? C.teal : "#0C2D4818", boxShadow: isNext ? "0 0 0 3px #F4A25933" : "none" }}>
              <div className="text-2xl">{done ? "✅" : q.icon}</div>
              <div className="flex-1">
                <div className="font-bold text-sm flex items-center gap-2" style={{ color: C.navy }}>
                  <span style={{ textDecoration: done ? "line-through" : "none" }}>{t("taskN", { i: i + 1 })}{t("q_" + q.id)}</span>
                  {isNext && <span className="buoy text-base" title="下一步">🛟</span>}
                </div>
                <div className="text-xs mt-0.5" style={{ color: "#0C2D4899" }}>{done ? t("doneLbl") : t("qh_" + q.id)}</div>
              </div>
              {isNext ? <Tag tone="orange">{t("nextStep")}</Tag> : <Tag tone={done ? "teal" : "gray"}>+{q.xp} XP</Tag>}
            </button>
          );
        })}
      </div>
      <p className="text-center text-xs" style={{ color: "#0C2D4877" }}>{t("questFooter")}</p>
    </div>
  );
}

/* 浮動提示浮標：固定於畫面右下，指引下一步 */
function FloatingBuoy({ quest, tab, gotoTab }) {
  const nq = nextQuest(quest);
  if (!nq || nq.tab === tab) return null;
  return (
    <button onClick={() => gotoTab(nq.tab)} className="fixed z-40 right-3 bottom-24 flex items-center gap-2 pl-2 pr-4 py-2 rounded-full shadow-xl"
      style={{ background: C.orange, color: C.navyDeep }}>
      <span className="buoy text-xl">🛟</span>
      <span className="text-xs font-black text-left leading-tight">{t("nextStepB")}<br />{t("q_" + nq.id)}</span>
    </button>
  );
}

/* ---------- 船班行事曆（狀態即時 + 空班開班） ---------- */
function remainOf(trip, orders) {
  const used = orders.filter((o) => o.tripId === trip.id && o.status !== "已取消").length;
  return Math.max(0, trip.capacity - used);
}
function hoursUntil(dateStr) { return (new Date(dateStr + "T00:00:00").getTime() - Date.now()) / 3.6e6; }

function TripCalendar({ trips, orders, onOpen, onOpenNew }) {
  const [view, setView] = useState("list");
  const now = new Date();
  const [ym, setYm] = useState({ y: now.getFullYear(), m: now.getMonth() }); // m: 0-11
  const shiftMonth = (d) => setYm(({ y, m }) => { const nm = m + d; return { y: y + Math.floor(nm / 12), m: ((nm % 12) + 12) % 12 }; });
  const days = ["日", "一", "二", "三", "四", "五", "六"];
  const first = new Date(ym.y, ym.m, 1).getDay();
  const daysInMonth = new Date(ym.y, ym.m + 1, 0).getDate();
  const byDate = Object.fromEntries(trips.map((t) => [t.date, t]));
  const cells = [...Array(first).fill(null), ...Array.from({ length: daysInMonth }, (_, i) => i + 1)];
  const cellColor = (trip, rem) => {
    const st = TRIP_STATUS[trip.status] || {};
    if (!st.bookable) return st.color || C.gray;
    return rem === 0 ? C.red : rem <= 3 ? C.orange : C.teal;
  };
  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5">
          <button onClick={() => shiftMonth(-1)} className="w-8 h-8 rounded-full font-black" style={{ background: "#0C2D4812", color: C.navy }}>‹</button>
          <h3 className="font-black text-lg" style={{ color: C.navy }}>{t("calTitle", { y: ym.y, m: ym.m + 1 })}</h3>
          <button onClick={() => shiftMonth(1)} className="w-8 h-8 rounded-full font-black" style={{ background: "#0C2D4812", color: C.navy }}>›</button>
        </div>
        <div className="flex gap-1">
          {[["list", t("listV")], ["month", t("monthV")]].map(([v, l]) => (
            <button key={v} onClick={() => setView(v)} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: view === v ? C.navy : "#0C2D4812", color: view === v ? C.sand : C.navy }}>{l}</button>
          ))}
        </div>
      </div>

      <button onClick={() => onOpenNew(null)} className="w-full rounded-2xl p-3 text-sm font-bold flex items-center justify-center gap-2 border-2 border-dashed" style={{ borderColor: C.teal, color: C.tealDark, background: "#1EA8960E" }}>
        {t("openTripBtn")}
      </button>

      {view === "month" && (
        <div className="rounded-2xl p-3" style={{ background: "#fff", border: "2px solid #0C2D4815" }}>
          <div className="grid grid-cols-7 text-center text-xs font-bold mb-1" style={{ color: "#0C2D4888" }}>{days.map((d) => <div key={d}>{d}</div>)}</div>
          <div className="grid grid-cols-7 gap-1">
            {cells.map((d, i) => {
              const key = d ? `${ym.y}-${String(ym.m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}` : null;
              const trip = key && byDate[key];
              const rem = trip ? remainOf(trip, orders) : 0;
              const openable = key && !trip && hoursUntil(key) >= 24;
              return (
                <button key={i} disabled={!trip && !openable} onClick={() => (trip ? onOpen(trip) : openable && onOpenNew(key))}
                  className="aspect-square rounded-lg text-xs font-bold flex flex-col items-center justify-center"
                  style={{ background: trip ? cellColor(trip, rem) : openable ? "#1EA89618" : d ? "#0C2D4808" : "transparent", color: trip ? "#fff" : openable ? C.tealDark : "#0C2D4855", border: openable ? `1px dashed ${C.teal}` : "none" }}>
                  {d}{trip ? <span style={{ fontSize: 9 }}>{TRIP_STATUS[trip.status]?.bookable ? t("remainS", { n: rem }) : t("cxl")}</span> : openable ? <span style={{ fontSize: 9 }}>{t("canOpen")}</span> : null}
                </button>
              );
            })}
          </div>
          <div className="flex flex-wrap gap-2 mt-2 text-xs justify-center" style={{ color: "#0C2D4899" }}>
            <span><span style={{ color: C.teal }}>●</span> {t("lgAvail")}</span><span><span style={{ color: C.orange }}>●</span> {t("lgFew")}</span><span><span style={{ color: C.red }}>●</span> {t("lgFull")}</span><span style={{ color: C.tealDark }}>{t("lgOpen")}</span>
          </div>
        </div>
      )}

      {trips.map((tp) => {
        const rem = remainOf(tp, orders);
        const st = TRIP_STATUS[tp.status] || {};
        return (
          <button key={tp.id} onClick={() => onOpen(tp)} className="w-full text-left rounded-2xl p-4 border-2 transition-transform active:scale-[.99]" style={{ background: "#fff", borderColor: st.bookable ? "#0C2D4815" : `${C.red}55` }}>
            <div className="flex items-center justify-between gap-2">
              <div className="font-black" style={{ color: C.navy }}>{tp.date.slice(5).replace("-", "/")}・{lx(tp, "name")}</div>
              <Tag tone={st.tone || "teal"}>{tSt(tp.status)}</Tag>
            </div>
            <div className="text-xs mt-1.5 flex flex-wrap gap-x-3 gap-y-1" style={{ color: "#0C2D4899" }}>
              <span>{t("musterAt", { t: tp.muster })}{tp.timePending ? " ⏳" : ""}</span><span>🎯 {lxArr(tp, "targets").join("、")}</span><span>🌊 {lx(tp, "depth")}</span>
              <span className="font-bold" style={{ color: C.tealDark }}>NT${tp.price.toLocaleString()}</span>
              {st.bookable && <Tag tone={rem === 0 ? "red" : rem <= 3 ? "orange" : "teal"}>{rem === 0 ? t("waitlist") : t("remain", { n: rem })}</Tag>}
            </div>
          </button>
        );
      })}
    </div>
  );
}

/* 釣客自行開班 */
function OpenTripModal({ presetDate, trips, onCreate, onClose }) {
  const [date, setDate] = useState(presetDate || "");
  const minDate = new Date(Date.now() + 24 * 3.6e6).toISOString().slice(0, 10);
  const [svc, setSvc] = useState("");
  const [slot, setSlot] = useState(null); // index | "custom"
  const [customTime, setCustomTime] = useState("");
  const [err, setErr] = useState("");
  const service = SERVICES.find((s) => s.name === svc);
  const presets = service ? (TIME_PRESETS[service.zone] || []) : [];
  const isInquiry = service && service.basePrice === 0;
  const create = () => {
    if (!date || !svc) { setErr(t("otErrSel")); return; }
    if (isInquiry) return;
    if (hoursUntil(date) < 24) { setErr(t("otErr24")); return; }
    if (trips.some((x) => x.date === date)) { setErr(t("otErrDup")); return; }
    if (slot === null || (slot === "custom" && !customTime)) { setErr(t("otErrTime")); return; }
    const timing = slot === "custom"
      ? { muster: customTime, rodsUp: "—", back: "—", timePending: true }
      : { ...presets[slot], timePending: false };
    onCreate(date, service, timing);
  };
  const inputStyle = { borderColor: "#0C2D4822", background: "#fff", color: C.navyDeep };
  return (
    <ModalShell onClose={onClose}>
      <div style={{ background: `linear-gradient(150deg, ${C.navy}, ${C.tealDark})` }} className="px-5 pt-5 pb-3">
        <button onClick={onClose} className="mb-2 text-xs px-3 py-1.5 rounded-full font-bold" style={{ background: "#FFFFFF22", color: "#FFFFFF" }}>{t("backCal")}</button>
        <div className="text-center"><div className="text-3xl mb-1">🚩</div><h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("otTitle")}</h2></div>
        <p className="text-xs mt-1 text-center" style={{ color: "#F7F3EC99" }}>{t("otSub")}</p>
      </div>
      <Wave />
      <div className="px-5 pb-5 space-y-3">
        <div>
          <div className="text-xs font-bold mb-1" style={{ color: C.tealDark }}>{t("otDate")}</div>
          <input type="date" min={minDate} className="w-full p-3 rounded-xl border-2 text-sm" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        {service && !isInquiry && (
          <div>
            <div className="text-xs font-bold mb-1" style={{ color: C.tealDark }}>{t("otTime")}</div>
            <div className="space-y-1.5">
              {presets.map((ps, i) => (
                <button key={i} onClick={() => setSlot(i)} className="w-full p-2.5 rounded-xl text-left text-xs font-bold border-2" style={{ borderColor: slot === i ? C.teal : "#0C2D4818", background: slot === i ? "#1EA89612" : "#fff", color: C.navy }}>
                  {t("muster")} {ps.muster}｜{t("rodsUp")} {ps.rodsUp}｜{t("backPort")} {ps.back}
                </button>
              ))}
              <button onClick={() => setSlot("custom")} className="w-full p-2.5 rounded-xl text-left text-xs font-bold border-2" style={{ borderColor: slot === "custom" ? C.orange : "#0C2D4818", background: slot === "custom" ? "#F4A25915" : "#fff", color: C.navy }}>
                {t("tsCustom")}
              </button>
              {slot === "custom" && (
                <div className="rounded-xl p-2.5 space-y-1.5" style={{ background: "#FFD16622", border: `1.5px dashed ${C.orange}` }}>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold" style={{ color: C.navy }}>{t("customLbl")}</span>
                    <input type="time" className="p-2 rounded-lg border-2 text-sm" style={{ borderColor: "#0C2D4822", background: "#fff", color: C.navy }} value={customTime} onChange={(e) => setCustomTime(e.target.value)} />
                  </div>
                  <p className="text-xs" style={{ color: C.navyDeep }}>{t("customNote")}</p>
                </div>
              )}
            </div>
          </div>
        )}
        <div>
          <div className="text-xs font-bold mb-1" style={{ color: C.tealDark }}>{t("otSvc")}</div>
          <div className="grid grid-cols-2 gap-2 max-h-52 overflow-y-auto">
            {SERVICES.map((s) => (
              <button key={s.name} onClick={() => setSvc(s.name)} className="p-2.5 rounded-xl text-left border-2 text-xs" style={{ borderColor: svc === s.name ? C.teal : "#0C2D4818", background: svc === s.name ? "#1EA89612" : "#fff" }}>
                <div className="text-lg">{s.icon}</div>
                <div className="font-bold mt-0.5" style={{ color: C.navy }}>{s.name}</div>
                <div style={{ color: "#0C2D4888" }}>{tZone(s.zone)}・{tPrice(s.price)}</div>
              </button>
            ))}
          </div>
        </div>
        {err && <p className="text-xs font-bold" style={{ color: C.red }}>{err}</p>}
        {isInquiry ? (
          <div className="rounded-xl p-3 space-y-2" style={{ background: "#FFD16622", border: `2px solid ${C.orange}` }}>
            <p className="text-sm font-bold" style={{ color: C.navyDeep }}>{t("otInq", { s: service.name })}</p>
            <a href={CONTACT.lineOA} target="_blank" rel="noreferrer" className="block p-3 rounded-xl font-bold text-center text-sm" style={{ background: "#06C755", color: "#fff" }}>💬 {t("lineOA")}</a>
            <a href={CONTACT.captainLine} target="_blank" rel="noreferrer" className="block p-3 rounded-xl font-bold text-center text-sm" style={{ background: C.teal, color: "#fff" }}>🧑\u200d✈️ {t("capLine")}</a>
            <div className="text-center text-xs font-bold" style={{ color: C.navy }}>{t("orCall", { p: CONTACT.captain })}</div>
          </div>
        ) : (
          <>
            <p className="text-xs" style={{ color: "#0C2D4877" }}>{t("otNote")}</p>
            <Btn full onClick={create}>{t("otCreate")}</Btn>
          </>
        )}
        <Btn full kind="navy" onClick={onClose}>{t("backBottom")}</Btn>
      </div>
    </ModalShell>
  );
}

/* 裝備租借費用表格提示 */
function RentalFeeModal({ rentals, trip, onClose }) {
  const rInfo = rentalsForTrip(trip, rentals);
  const dCls = rInfo.cls;
  const list = rInfo.list.filter((r) => r.isSet);
  return (
    <ModalShell onClose={onClose}>
      <div style={{ background: C.navy }} className="px-5 pt-5 pb-3 text-center"><div className="text-3xl mb-1">🎣</div><h2 className="font-black text-lg" style={{ color: C.yellow }}>{t("rfTitle")}</h2>
        {rInfo.mode === "manual" ? <p className="text-xs mt-1 font-bold" style={{ color: C.sand }}>{t("manualNote")}</p> : dCls !== null && <p className="text-xs mt-1 font-bold" style={{ color: C.sand }}>{t("autoNote", { d: lx(trip, "depth"), z: t(dCls === "deep" ? "rzD" : "rzS") })}</p>}
      </div>
      <Wave />
      <div className="px-5 pb-5">
        <div className="rounded-xl overflow-hidden border-2" style={{ borderColor: "#0C2D4815" }}>
          <div className="grid grid-cols-12 text-xs font-bold px-3 py-2" style={{ background: C.navy, color: C.sand }}>
            <div className="col-span-7">{t("rfItem")}</div><div className="col-span-2 text-center">{t("rfUnit")}</div><div className="col-span-3 text-right">{t("rfFee")}</div>
          </div>
          {list.map((r, i) => (
            <div key={r.name} className="grid grid-cols-12 px-3 py-2.5 text-xs items-center" style={{ background: i % 2 ? "#F7F3EC" : "#fff" }}>
              <div className="col-span-7"><span className="font-bold" style={{ color: C.navy }}>{lx(r, "name")}</span><Tag tone={r.zone === "深場" ? "navy" : r.zone === "通用" ? "gray" : "teal"}>{t(r.zone === "深場" ? "rzTagD" : r.zone === "通用" ? "rzTagB" : "rzTagS")}</Tag></div>
              <div className="col-span-2 text-center" style={{ color: "#0C2D4888" }}>{lx(r, "unit")}</div>
              <div className="col-span-3 text-right font-black" style={{ color: C.tealDark }}>NT${r.price}</div>
            </div>
          ))}
        </div>
        <div className="mt-3 rounded-xl p-3 text-xs font-bold" style={{ background: "#FFD16622", color: C.navyDeep }}>
          <span style={{ color: C.red }}>{t("rfNote")}</span>
        </div>
        <Btn full onClick={onClose}>{t("rfOk")}</Btn>
      </div>
    </ModalShell>
  );
}


/* 裝備建議智慧排版：管理員照「項目：內容」逐行輸入，前台自動解析為分類手風琴＋對照表
   規則：行尾為「：」→分類標題；emoji 開頭→分類標題；「標籤：內容」且標籤≤12字→對照列；其餘→補充說明 */
function parseGear(text) {
  const lines = String(text || "").split(/\n/).map((l) => l.trim()).filter(Boolean);
  const sections = [];
  let cur = null;
  const push = () => { if (cur && (cur.rows.length || cur.notes.length)) sections.push(cur); };
  const isEmojiStart = (l) => /^[\u{1F300}-\u{1FAFF}\u2600-\u27BF]/u.test(l);
  for (const line of lines) {
    const m = line.match(/^(.{1,30}?)[：:](.*)$/);
    const label = m ? m[1].trim() : null;
    const value = m ? m[2].trim() : null;
    if ((m && !value) || isEmojiStart(line) || (!m && line.length <= 22 && !cur)) {
      push(); cur = { title: line.replace(/[：:]\s*$/, ""), rows: [], notes: [] };
    } else if (m && label.length <= 12) {
      if (!cur) cur = { title: "", rows: [], notes: [] };
      cur.rows.push([label, value]);
    } else {
      if (!cur) cur = { title: "", rows: [], notes: [] };
      cur.notes.push(line);
    }
  }
  push();
  return sections;
}
function GearBlock({ text, rigs, links }) {
  const sections = parseGear(text);
  if (rigs && rigs.length) sections.push({ title: t("lRigs"), rows: [], notes: rigs });
  const structured = sections.length > 1 || (sections[0] && sections[0].rows.length >= 2);
  const [mode, setMode] = useState("acc"); // acc 摺疊 | tab 分頁 | grid 總覽
  const [tabIdx, setTabIdx] = useState(0);
  const LinkBtns = () => (links && links.length ? (
    <div className="flex flex-wrap gap-2 mt-2.5">
      {links.filter((lk) => lk.url).map((lk, i) => (
        <a key={i} href={lk.url} target="_blank" rel="noreferrer" className="px-3 py-2 rounded-xl text-xs font-black flex items-center gap-1.5" style={{ background: C.teal, color: "#fff" }}>
          🔗 {lx(lk, "label") || lk.url}
        </a>
      ))}
    </div>
  ) : null);
  if (!structured) return <div><div className="text-base font-bold leading-relaxed" style={{ color: C.navy }}>{text || "—"}</div><LinkBtns /></div>;

  const Rows = ({ sec, compact }) => (
    <>
      {sec.rows.length > 0 && (
        <div className="rounded-lg overflow-hidden" style={{ border: "1px solid #0C2D4812" }}>
          {sec.rows.map(([k, v], j) => (
            <div key={j} className={compact ? "text-xs" : "grid grid-cols-3 text-xs"} style={{ background: j % 2 ? "#fff" : "#FBFAF7" }}>
              {compact ? (
                <div className="p-1.5 leading-relaxed"><span className="font-bold" style={{ color: C.tealDark }}>{k}</span><span className="font-bold" style={{ color: C.navy }}>｜{v}</span></div>
              ) : (
                <>
                  <div className="p-2 font-bold" style={{ color: C.tealDark, background: "#0C2D4806" }}>{k}</div>
                  <div className="col-span-2 p-2 font-bold leading-relaxed" style={{ color: C.navy }}>{v}</div>
                </>
              )}
            </div>
          ))}
        </div>
      )}
      {sec.notes.map((n, j) => (
        <p key={j} className="text-xs mt-1.5 leading-relaxed" style={{ color: "#0C2D4899" }}>{n}</p>
      ))}
    </>
  );

  return (
    <div>
      <div className="flex gap-1 mb-2">
        {[["acc", t("gvAcc")], ["tab", t("gvTab")], ["grid", t("gvGrid")]].map(([m, l]) => (
          <button key={m} onClick={() => setMode(m)} className="px-2.5 py-1 rounded-full text-xs font-bold" style={{ background: mode === m ? C.navy : "#0C2D480E", color: mode === m ? C.sand : C.navy }}>{l}</button>
        ))}
      </div>

      {mode === "acc" && (
        <div className="space-y-2">
          {sections.map((sec, i) => (
            <details key={i} open={i === 0} className="rounded-xl overflow-hidden" style={{ background: "#F7F3EC", border: "1.5px solid #0C2D4815" }}>
              <summary className="px-3 py-2.5 text-sm font-black cursor-pointer flex items-center list-none" style={{ color: C.navy }}>
                {sec.title || "📋"}<span className="ml-auto" style={{ color: C.teal }}>▾</span>
              </summary>
              <div className="px-3 pb-3"><Rows sec={sec} /></div>
            </details>
          ))}
        </div>
      )}

      {mode === "tab" && (
        <div>
          <div className="flex gap-1.5 overflow-x-auto pb-2">
            {sections.map((sec, i) => (
              <button key={i} onClick={() => setTabIdx(i)} className="px-3 py-1.5 rounded-full text-xs font-black whitespace-nowrap" style={{ background: tabIdx === i ? C.teal : "#F7F3EC", color: tabIdx === i ? "#fff" : C.navy, border: "1.5px solid #0C2D4815" }}>
                {sec.title || "📋"}
              </button>
            ))}
          </div>
          <div className="rounded-xl p-3" style={{ background: "#F7F3EC", border: "1.5px solid #0C2D4815" }}>
            <Rows sec={sections[Math.min(tabIdx, sections.length - 1)]} />
          </div>
        </div>
      )}

      {mode === "grid" && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          {sections.map((sec, i) => (
            <div key={i} className="rounded-xl p-2.5" style={{ background: "#F7F3EC", border: "1.5px solid #0C2D4815" }}>
              <div className="text-xs font-black mb-1.5" style={{ color: C.navy }}>{sec.title || "📋"}</div>
              <Rows sec={sec} compact />
            </div>
          ))}
        </div>
      )}
      <LinkBtns />
    </div>
  );
}

/* ---------- 船班詳情（卡片／表格雙檢視、觀光無租借、租借跳費用表） ---------- */
function TripDetail({ trip, orders, user, rentals, onBook, onClose, onRentalInfo }) {
  const st = TRIP_STATUS[trip.status] || {};
  const rem = remainOf(trip, orders);
  const roster = orders.filter((o) => o.tripId === trip.id && o.status !== "已取消");
  const mine = user && roster.some((o) => o.customerId === user.id);
  const [detailView, setDetailView] = useState("card"); // card | table
  const [role, setRole] = useState("釣手");
  const [rental, setRental] = useState("");
  const rInfo = rentalsForTrip(trip, rentals);
  const dCls = rInfo.cls;
  const [compCount, setCompCount] = useState(0);
  const [compNames, setCompNames] = useState([]);
  const syncComp = (n) => { setCompCount(n); setCompNames((p) => Array.from({ length: n }, (_, i) => p[i] || "")); };

  const rigsArr = lxArr(trip, "rigs");
  const infoA = [
    [t("lTargets"), lxArr(trip, "targets").join("、")],
    [t("lDepth"), lx(trip, "depth")],
  ];
  const infoB = [
    [t("lNote"), lx(trip, "note") || "—"],
  ];
  const gearCard = (
    <div className="rounded-xl p-3.5" style={{ background: "#fff", border: "2px solid #0C2D4812" }}>
      <div className="text-sm font-bold mb-1.5" style={{ color: C.tealDark }}>{t("lGear")}・{t("lRigs")}</div>
      <GearBlock text={lx(trip, "gear")} rigs={rigsArr} links={trip.links || []} />
    </div>
  );
  const setRoleFn = (r) => { setRole(r); if (r === "觀光") setRental(""); };
  const chooseRental = (v) => { setRental(v); if (v) onRentalInfo(); };

  return (
    <div className="fixed inset-0 z-40 overflow-y-auto" style={{ background: "rgba(8,29,48,.8)", WebkitOverflowScrolling: "touch" }} onClick={onClose}>
      <div className="min-h-full flex items-start justify-center p-3">
      <div className="w-full max-w-md rounded-2xl overflow-hidden my-2" style={{ background: C.sand }} onClick={(e) => e.stopPropagation()}>
        <div style={{ background: `linear-gradient(150deg, ${C.navy}, ${C.tealDark})` }} className="px-5 pt-5 pb-4 flex-shrink-0">
          <button onClick={onClose} className="mb-2 text-xs px-3 py-1.5 rounded-full font-bold" style={{ background: "#F7F3EC22", color: "#FFFFFF" }}>{t("backCal")}</button>
          <div className="flex justify-between items-start">
            <div>
              <div className="text-xs font-bold" style={{ color: C.yellow }}>{trip.date}・{tZone(trip.type)}</div>
              <h3 className="font-black text-xl mt-0.5" style={{ color: "#FFFFFF" }}>{lx(trip, "name")}</h3>
            </div>
            <button onClick={onClose} className="text-3xl leading-none px-2" style={{ color: "#FFFFFF" }}>×</button>
          </div>
          <div className="flex gap-2 mt-3 flex-wrap items-center">
            <Tag tone={st.tone || "teal"}>{tSt(trip.status)}</Tag>
            {trip.timePending && <Tag tone="orange">{t("pendingTag")}</Tag>}
            {st.bookable && <Tag tone={rem === 0 ? "red" : "yellow"}>{rem === 0 ? t("fullWait") : t("seatsLeft", { n: rem })}</Tag>}
            <Tag tone="orange">NT${trip.price.toLocaleString()}{t("perAngler")}</Tag>
          </div>
        </div>

        {/* 檢視切換 */}
        <div className="flex gap-1 px-5 pt-3 flex-shrink-0">
          {[["card", t("cardV")], ["table", t("tableV")]].map(([v, l]) => (
            <button key={v} onClick={() => setDetailView(v)} className="px-3 py-1.5 rounded-full text-xs font-bold" style={{ background: detailView === v ? C.navy : "#0C2D4812", color: detailView === v ? C.sand : C.navy }}>{l}</button>
          ))}
        </div>

        <div className="p-5 space-y-3">
          <div className="grid grid-cols-3 gap-2 text-center">
            {[[t("muster"), trip.muster], [t("rodsUp"), trip.rodsUp], [t("backPort"), trip.back]].map(([k, v]) => (
              <div key={k} className="rounded-xl p-2.5" style={{ background: "#fff", border: "2px solid #0C2D4812" }}>
                <div className="text-xs" style={{ color: "#0C2D4888" }}>{k}</div>
                <div className="font-black text-base" style={{ color: C.navy }}>{v}</div>
              </div>
            ))}
          </div>

          {detailView === "card" ? (
            <div className="space-y-2.5">
              {infoA.map(([k, v]) => (
                <div key={k} className="rounded-xl p-3.5" style={{ background: "#fff", border: "2px solid #0C2D4812" }}>
                  <div className="text-sm font-bold mb-1" style={{ color: C.tealDark }}>{k}</div>
                  <div className="text-base font-bold leading-relaxed" style={{ color: C.navy }}>{v}</div>
                </div>
              ))}
              {gearCard}
              {infoB.map(([k, v]) => (
                <div key={k} className="rounded-xl p-3.5" style={{ background: "#fff", border: "2px solid #0C2D4812" }}>
                  <div className="text-sm font-bold mb-1" style={{ color: C.tealDark }}>{k}</div>
                  <div className="text-base font-bold leading-relaxed" style={{ color: C.navy }}>{v}</div>
                </div>
              ))}
            </div>
          ) : (
            <div className="space-y-2.5">
              <div className="rounded-xl overflow-hidden border-2" style={{ borderColor: "#0C2D4812" }}>
                {infoA.map(([k, v], i) => (
                  <div key={k} className="grid grid-cols-3 items-start" style={{ background: i % 2 ? "#F7F3EC" : "#fff" }}>
                    <div className="col-span-1 p-3 text-sm font-bold" style={{ color: C.tealDark, background: "#0C2D4808" }}>{k}</div>
                    <div className="col-span-2 p-3 text-base font-bold leading-relaxed" style={{ color: C.navy }}>{v}</div>
                  </div>
                ))}
              </div>
              {gearCard}
              <div className="rounded-xl overflow-hidden border-2" style={{ borderColor: "#0C2D4812" }}>
                {infoB.map(([k, v], i) => (
                  <div key={k} className="grid grid-cols-3 items-start" style={{ background: i % 2 ? "#F7F3EC" : "#fff" }}>
                    <div className="col-span-1 p-3 text-sm font-bold" style={{ color: C.tealDark, background: "#0C2D4808" }}>{k}</div>
                    <div className="col-span-2 p-3 text-base font-bold leading-relaxed" style={{ color: C.navy }}>{v}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <div className="rounded-xl p-3" style={{ background: "#fff", border: "2px solid #0C2D4812" }}>
            <div className="text-sm font-bold mb-1" style={{ color: C.tealDark }}>{t("roster")}</div>
            {roster.length === 0 ? <div className="text-sm" style={{ color: "#0C2D4877" }}>{t("rosterEmpty")}</div> : (
              <div className="flex flex-wrap gap-1.5">{roster.map((o) => <Tag key={o.id} tone={o.role === "釣手" ? "navy" : "teal"}>{o.nickname}・{o.role === "釣手" ? t("angler") : t("sight")}{o.status === "候補" ? "（" + t("waitlist") + "）" : ""}{o.isMinor ? t("minorTag") : ""}</Tag>)}</div>
            )}
          </div>

          {!st.bookable ? (
            <div className="rounded-xl p-3 text-center font-bold text-sm" style={{ background: `${C.red}18`, color: C.red }}>{t("notBookable", { s: tSt(trip.status) })}</div>
          ) : mine ? (
            <div className="rounded-xl p-3 text-center font-bold text-sm" style={{ background: "#1EA89620", color: C.tealDark }}>{t("booked")}</div>
          ) : (
            <div className="rounded-xl p-3.5 space-y-3" style={{ background: "#fff", border: `2px solid ${C.orange}` }}>
              <div>
                <div className="text-sm font-bold mb-1.5" style={{ color: C.tealDark }}>{t("roleLbl")}</div>
                <div className="flex gap-2">
                  {["釣手", "觀光"].map((r) => (
                    <button key={r} onClick={() => setRoleFn(r)} className="flex-1 py-2.5 rounded-xl text-sm font-bold" style={{ background: role === r ? C.navy : "#0C2D4810", color: role === r ? C.sand : C.navy }}>{r === "釣手" ? t("angler") : t("sight")}</button>
                  ))}
                </div>
              </div>
              {/* 觀光僅限攜伴；觀光不提供裝備租借 */}
              {role === "釣手" ? (
                <>
                  <div>
                    <div className="text-sm font-bold mb-1.5" style={{ color: C.tealDark }}>{t("rentalLbl")}</div>
                    {rInfo.mode === "manual" ? (
                      <div className="rounded-lg p-2 mb-1.5 text-xs font-bold" style={{ background: "#F4A25920", color: C.navy }}>{t("manualNote")}</div>
                    ) : dCls !== null && (
                      <div className="rounded-lg p-2 mb-1.5 text-xs font-bold" style={{ background: dCls === "deep" ? "#0C2D4812" : "#1EA89615", color: dCls === "deep" ? C.navy : C.tealDark }}>
                        {dCls === "deep" ? "🌑" : "🌊"} {t("autoNote", { d: lx(trip, "depth"), z: t(dCls === "deep" ? "rzD" : "rzS") })}
                      </div>
                    )}
                    <select value={rental} onChange={(e) => chooseRental(e.target.value)} className="w-full p-3 rounded-xl text-sm border-2" style={{ borderColor: "#0C2D4822", background: "#fff", color: C.navy }}>
                      <option value="">{t("noRental")}</option>
                      {rInfo.list.filter((r) => r.isSet).map((r) => <option key={r.name} value={r.name}>{lx(r, "name")}（NT${r.price}）</option>)}
                    </select>
                    <button onClick={onRentalInfo} className="text-xs mt-1.5 underline font-bold" style={{ color: C.tealDark }}>{t("feeLink")}</button>
                  </div>
                  <div>
                    <div className="text-sm font-bold mb-1.5" style={{ color: C.tealDark }}>{t("compLbl")}</div>
                    <div className="flex items-center gap-2">
                      <span className="text-xs" style={{ color: "#0C2D48AA" }}>{t("compCount")}</span>
                      <select value={compCount} onChange={(e) => syncComp(Number(e.target.value))} className="p-2 rounded-lg border-2 text-sm" style={{ borderColor: "#0C2D4822", background: "#fff", color: C.navy }}>
                        {[0, 1, 2, 3].map((n) => <option key={n} value={n}>{n}</option>)}
                      </select>
                    </div>
                    {compNames.slice(0, compCount).map((n, i) => (
                      <input key={i} className="w-full p-2.5 mt-2 rounded-lg border-2 text-sm" style={{ borderColor: "#0C2D4822", background: "#fff", color: C.navy }} placeholder={t("compName", { i: i + 1 })} value={n} onChange={(e) => setCompNames((p) => p.map((x, idx) => (idx === i ? e.target.value : x)))} />
                    ))}
                    {compCount > 0 && <p className="text-xs mt-1.5" style={{ color: "#0C2D48AA" }}>{t("compNote")}</p>}
                  </div>
                </>
              ) : (
                <div className="rounded-lg p-2.5 text-xs font-bold" style={{ background: "#E85D4A15", color: C.red }}>{t("sightBlock")}</div>
              )}
              {user?.minorMode === "self" && user?.minorCount > 0 && (
                <div className="rounded-lg p-2.5 text-xs font-bold" style={{ background: "#FFD16622", color: C.navyDeep }}>{t("minorSeat", { n: user.minorCount })}</div>
              )}
              <Btn full disabled={role === "觀光" || (compCount > 0 && compNames.slice(0, compCount).some((n) => !n.trim()))} onClick={() => onBook(trip, role, rental, rem, compNames.slice(0, compCount))}>{rem === 0 ? t("waitBtn") : t("bookBtn")}</Btn>
              <p className="text-xs text-center" style={{ color: "#0C2D4877" }}>{t("policy")}</p>
            </div>
          )}
          <Btn full kind="navy" onClick={onClose}>{t("backBottom")}</Btn>
        </div>
      </div>
      </div>
    </div>
  );
}

/* ---------- 我的船班（動態 + 行前提醒） ---------- */
function reminderKeysFor(trip) {
  const keys = ["rm1", "rm2", "rm3"];
  if (trip.name.includes("活餌") || (trip.note || "").includes("活餌")) keys.push("rm4");
  if (trip.type === "近海") keys.push("rm5"); else keys.push("rm6");
  keys.push("rm7");
  return keys;
}
function MyTripsTab({ trips, orders, user, onOpen }) {
  const mine = orders.filter((o) => o.customerId === user.id);
  if (mine.length === 0) return (
    <div className="text-center py-14">
      <div className="text-5xl mb-3">🎣</div>
      <p className="text-sm font-bold" style={{ color: C.navy }}>{t("myEmpty1")}</p>
      <p className="text-xs mt-1" style={{ color: "#0C2D4877" }}>{t("myEmpty2")}</p>
    </div>
  );
  return (
    <div className="space-y-3">
      <h3 className="font-black text-lg" style={{ color: C.navy }}>{t("myTitle")}</h3>
      {mine.map((o) => {
        const tr = trips.find((x) => x.id === o.tripId);
        if (!tr) return null;
        const st = TRIP_STATUS[tr.status] || {};
        return (
          <div key={o.id} className="rounded-2xl overflow-hidden border-2" style={{ background: "#fff", borderColor: st.bookable ? "#0C2D4815" : `${C.red}44` }}>
            <div className="p-4">
              <div className="flex items-center justify-between gap-2">
                <div className="font-black" style={{ color: C.navy }}>{tr.date}・{lx(tr, "name")}</div>
                <Tag tone={o.status === "已取消" ? "red" : o.status === "候補" ? "orange" : st.tone || "teal"}>{o.status === "已取消" ? t("stCxl") : o.status === "候補" ? t("stWait") : tSt(tr.status)}</Tag>
              </div>
              <div className="text-xs mt-1.5 flex flex-wrap gap-x-3 gap-y-1" style={{ color: "#0C2D4899" }}>
                <span>{t("idLbl")}：{o.role === "釣手" ? t("angler") : t("sight")}</span><span>{t("musterAt", { t: tr.muster })}{tr.timePending ? " ⏳" : ""}</span><span>💰 NT${(o.price + (o.rentalPrice || 0)).toLocaleString()}</span>
                <span>{o.paid ? t("paid") : t("unpaid")}</span>{o.rental && <span>{t("rentShort")}：{o.rental}</span>}
              </div>
            </div>
            <div className="px-4 pb-4">
              <div className="rounded-xl p-3" style={{ background: "#0C2D4808" }}>
                <div className="text-xs font-black mb-1.5" style={{ color: C.tealDark }}>{t("remTitle")}</div>
                <ul className="space-y-1">{reminderKeysFor(tr).map((k) => <li key={k} className="text-xs flex gap-1.5" style={{ color: C.navy }}><span style={{ color: C.orange }}>•</span>{t(k)}</li>)}</ul>
              </div>
              <button onClick={() => onOpen(tr)} className="text-xs mt-2 underline font-bold" style={{ color: C.tealDark }}>{t("viewFull")}</button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ---------- 規定 / 費用 / 服務 ---------- */
function RulesTab({ rules }) {
  return (
    <div className="space-y-2.5">
      {rules.map((r, i) => (
        <details key={i} className="rounded-2xl overflow-hidden border-2" style={{ background: "#fff", borderColor: "#0C2D4815" }}>
          <summary className="p-4 font-bold text-sm cursor-pointer flex items-center gap-2.5 list-none" style={{ color: C.navy }}>
            <span className="text-xl">{r.icon}</span>{lx(r, "title")}<span className="ml-auto" style={{ color: C.teal }}>▾</span>
          </summary>
          <div className="px-4 pb-4 text-sm leading-relaxed" style={{ color: "#0C2D48CC" }}>{lx(r, "body")}</div>
        </details>
      ))}
      <div className="rounded-2xl p-4 text-center text-sm font-bold" style={{ background: C.navy, color: C.sand }}>
        {t("depart")} <a href={CONTACT.port} target="_blank" rel="noreferrer" style={{ color: C.yellow }} className="underline">{t("openMap")}</a>
      </div>
    </div>
  );
}
function FeesTab({ rentals, pricing }) {
  return (
    <div className="space-y-4">
      <h3 className="font-black" style={{ color: C.navy }}>{t("feeTitle")}</h3>
      {pricing.map((p) => (
        <div key={p.name} className="rounded-2xl p-4 border-2" style={{ background: "#fff", borderColor: "#0C2D4815" }}>
          <div className="flex justify-between items-center"><div className="font-bold" style={{ color: C.navy }}>{lx(p, "name")}</div><div className="font-black" style={{ color: C.tealDark }}>{p.price}</div></div>
          <div className="text-xs mt-1" style={{ color: "#0C2D4899" }}>{lx(p, "hours")}｜{lx(p, "note")}</div>
        </div>
      ))}
      <h3 className="font-black pt-1" style={{ color: C.navy }}>{t("rentTitle")}</h3>
      <div className="rounded-2xl overflow-hidden border-2" style={{ borderColor: "#0C2D4815" }}>
        {rentals.filter((r) => r.isSet).map((r, i) => (
          <div key={r.name} className="flex justify-between items-center p-3.5 text-sm" style={{ background: i % 2 ? "#F7F3EC" : "#fff" }}>
            <div><span className="font-bold" style={{ color: C.navy }}>{lx(r, "name")} <Tag tone={r.zone === "深場" ? "navy" : r.zone === "通用" ? "gray" : "teal"}>{t(r.zone === "深場" ? "rzTagD" : r.zone === "通用" ? "rzTagB" : "rzTagS")}</Tag></span><span className="block text-xs" style={{ color: "#0C2D4888" }}>{lx(r, "note")}</span></div>
            <div className="font-black whitespace-nowrap" style={{ color: C.tealDark }}>NT${r.price}／{lx(r, "unit")}</div>
          </div>
        ))}
      </div>
      <p className="text-xs" style={{ color: "#0C2D4877" }}>{t("feeFoot")}</p>
    </div>
  );
}
function ServicesTab() {
  return (
    <div className="grid grid-cols-2 gap-2.5">
      {SERVICES.map((s) => (
        <div key={s.name} className="rounded-2xl p-4 border-2" style={{ background: "#fff", borderColor: "#0C2D4815" }}>
          <div className="text-2xl">{s.icon}</div>
          <div className="font-bold text-sm mt-1.5 leading-snug" style={{ color: C.navy }}>{s.name}</div>
          <div className="flex items-center justify-between mt-2"><Tag tone="navy">{tZone(s.zone)}</Tag><span className="text-xs font-black" style={{ color: C.tealDark }}>{tPrice(s.price)}</span></div>
        </div>
      ))}
    </div>
  );
}


/* ---------- 聯絡我們 & 最新公告 ---------- */
function ContactNewsTab({ posts, contactPage }) {
  return (
    <div className="space-y-4">
      <h3 className="font-black text-lg" style={{ color: C.navy }}>{t("ctTitle")}</h3>
      <p className="text-sm" style={{ color: "#0C2D48CC" }}>{lx(contactPage, "intro")}</p>
      <div className="space-y-2.5">
        {contactPage.items.map((it, i) => (
          <a key={i} href={it.link || undefined} target={it.link && it.link.startsWith("http") ? "_blank" : undefined} rel="noreferrer"
            className="flex items-center gap-3 p-4 rounded-2xl border-2 transition-transform active:scale-[.98]" style={{ background: "#fff", borderColor: "#0C2D4815" }}>
            <span className="text-2xl">{it.icon}</span>
            <span className="flex-1">
              <span className="block font-bold text-sm" style={{ color: C.navy }}>{lx(it, "label")}</span>
              <span className="block text-xs mt-0.5" style={{ color: "#0C2D4899" }}>{lx(it, "value")}</span>
            </span>
            {it.link && <span className="font-black" style={{ color: C.teal }}>→</span>}
          </a>
        ))}
      </div>
      <h3 className="font-black text-lg pt-2" style={{ color: C.navy }}>{t("newsTitle")}</h3>
      {posts.length === 0 && <p className="text-sm" style={{ color: "#0C2D4877" }}>{t("noPosts")}</p>}
      {[...posts].sort((a, b) => b.date.localeCompare(a.date)).map((p) => (
        <div key={p.id} className="rounded-2xl p-4 border-2" style={{ background: "#fff", borderColor: "#0C2D4815" }}>
          <div className="flex items-center justify-between gap-2">
            <div className="font-bold text-sm" style={{ color: C.navy }}>{lx(p, "title")}</div>
            <span className="text-xs" style={{ color: "#0C2D4877" }}>{p.date}</span>
          </div>
          <div className="text-sm mt-2 whitespace-pre-wrap leading-relaxed" style={{ color: "#0C2D48CC" }}>{lx(p, "body")}</div>
        </div>
      ))}
    </div>
  );
}

/* ---------- 船長後台 ---------- */
function DeleteBtn({ onConfirm, label = "刪除" }) {
  const [arm, setArm] = useState(false);
  useEffect(() => { if (arm) { const t = setTimeout(() => setArm(false), 3000); return () => clearTimeout(t); } }, [arm]);
  return arm
    ? <button onClick={onConfirm} className="px-3 py-1.5 rounded-xl text-xs font-black" style={{ background: C.red, color: "#fff" }}>確認刪除？</button>
    : <button onClick={() => setArm(true)} className="px-3 py-1.5 rounded-xl text-xs font-bold" style={{ background: "#F7F3EC15", color: C.sand }}>{label}</button>;
}

const aInput = "p-2.5 rounded-xl text-sm w-full";
const aStyle = { background: "#F7F3EC12", border: "1px solid #F7F3EC33", color: C.sand };

/* 單一船班完整編輯器（含三語內容） */
function TripEditor({ trip, orderCount, rentals, onSave, onDelete }) {
  const [open, setOpen] = useState(false);
  const [eL, setEL] = useState("zh");
  const sfx = eL === "zh" ? "" : "_" + eL;
  const langName = { zh: "中文", en: "English", ja: "日本語" }[eL];
  const join = (a) => (a || []).join("、");
  const [f, setF] = useState({
    ...trip,
    links: trip.links || [],
    rentalMode: trip.rentalMode || "auto", rentalNames: trip.rentalNames || [],
    targets: join(trip.targets), rigs: join(trip.rigs),
    targets_en: join(trip.targets_en), rigs_en: join(trip.rigs_en),
    targets_ja: join(trip.targets_ja), rigs_ja: join(trip.rigs_ja),
  });
  const split = (str) => (str ? String(str).split(/[、,，\s]+/).filter(Boolean) : []);
  const save = () => {
    onSave({
      ...trip, ...f,
      price: Number(f.price) || 0, capacity: Number(f.capacity) || 10,
      targets: split(f.targets), rigs: split(f.rigs),
      targets_en: split(f.targets_en), rigs_en: split(f.rigs_en),
      targets_ja: split(f.targets_ja), rigs_ja: split(f.rigs_ja),
    });
    setOpen(false);
  };
  const st = TRIP_STATUS[trip.status] || {};
  const L = (base, label) => (
    <input className={aInput} style={aStyle} placeholder={label + "（" + langName + "）"} value={f[base + sfx] || ""} onChange={(e) => setF({ ...f, [base + sfx]: e.target.value })} />
  );
  return (
    <div className="rounded-2xl overflow-hidden" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
      <div className="p-4 flex items-center gap-3">
        <div className="flex-1">
          <div className="font-bold text-sm" style={{ color: C.sand }}>{trip.date.slice(5)}・{trip.name}</div>
          <div className="text-xs" style={{ color: "#F7F3EC77" }}>{trip.type}｜NT${trip.price.toLocaleString()}｜已報 {orderCount}/{trip.capacity}</div>
        </div>
        <Tag tone={st.tone || "teal"}>{trip.status}</Tag>
        <button onClick={() => setOpen(!open)} className="px-3 py-1.5 rounded-xl text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>{open ? "收合" : "編輯"}</button>
      </div>
      {open && (
        <div className="px-4 pb-4 space-y-2" style={{ borderTop: "1px solid #F7F3EC1A" }}>
          <div className="pt-3 text-xs font-bold" style={{ color: C.yellow }}>船班狀態（前台即時連動）</div>
          <div className="grid grid-cols-2 gap-1.5">
            {STATUS_LIST.map((sN) => (
              <button key={sN} onClick={() => setF({ ...f, status: sN })} className="py-2 rounded-lg text-xs font-bold" style={{ background: f.status === sN ? (TRIP_STATUS[sN].color) : "#F7F3EC12", color: f.status === sN ? "#fff" : C.sand }}>{sN}</button>
            ))}
          </div>
          {f.timePending !== undefined && (
            <button onClick={() => setF({ ...f, timePending: !f.timePending })} className="w-full py-2 rounded-lg text-xs font-bold" style={{ background: f.timePending ? C.orange : C.teal, color: f.timePending ? C.navyDeep : "#fff" }}>
              {f.timePending ? "⏳ 時間待船長確認中（點擊確認時間，記得先調整集合／起竿／回港）" : "✅ 出發時間已確認"}
            </button>
          )}
          <div className="pt-1 text-xs font-bold" style={{ color: C.yellow }}>共用欄位（不分語言）</div>
          <div className="grid grid-cols-2 gap-2">
            <input className={aInput} style={aStyle} placeholder="日期" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
            <select className={aInput} style={{ ...aStyle, background: C.navy }} value={f.type} onChange={(e) => setF({ ...f, type: e.target.value })}>{["近海", "中遠程", "離島", "專案", "休閒"].map((tp) => <option key={tp}>{tp}</option>)}</select>
            <input className={aInput} style={aStyle} placeholder="船資" value={f.price} onChange={(e) => setF({ ...f, price: e.target.value })} />
            <input className={aInput} style={aStyle} placeholder="乘客上限" value={f.capacity} onChange={(e) => setF({ ...f, capacity: e.target.value })} />
            <input className={aInput} style={aStyle} placeholder="集合" value={f.muster} onChange={(e) => setF({ ...f, muster: e.target.value })} />
            <input className={aInput} style={aStyle} placeholder="起竿" value={f.rodsUp} onChange={(e) => setF({ ...f, rodsUp: e.target.value })} />
            <input className={aInput} style={aStyle} placeholder="回港" value={f.back} onChange={(e) => setF({ ...f, back: e.target.value })} />
          </div>
          <div className="flex items-center gap-2 pt-1">
            <span className="text-xs font-bold" style={{ color: C.yellow }}>🌐 內容語言</span>
            {[["zh", "中文"], ["en", "EN"], ["ja", "日"]].map(([l, lb]) => (
              <button key={l} onClick={() => setEL(l)} className="px-3 py-1.5 rounded-full text-xs font-bold" style={{ background: eL === l ? C.orange : "#F7F3EC15", color: eL === l ? C.navyDeep : C.sand }}>{lb}</button>
            ))}
            <span className="text-xs ml-auto" style={{ color: "#F7F3EC66" }}>{eL === "zh" ? "中文為基底" : "留空則前台顯示中文"}</span>
          </div>
          {L("name", "船班名稱")}
          {L("targets", "目標魚（頓號分隔）")}
          {L("depth", "水深")}
          <textarea className={aInput} style={aStyle} rows={2} placeholder={"裝備建議（" + langName + "）"} value={f["gear" + sfx] || ""} onChange={(e) => setF({ ...f, ["gear" + sfx]: e.target.value })} />
          {L("rigs", "建議釣組（頓號分隔）")}
          <div className="pt-1 text-xs font-bold" style={{ color: C.yellow }}>出租裝備顯示（前台租借選單／費用彈窗）</div>
          <div className="flex gap-1.5">
            {[["auto", "自動（依水深 80m 判定）"], ["manual", "手動指定"]].map(([m, lb]) => (
              <button key={m} onClick={() => setF({ ...f, rentalMode: m })} className="flex-1 py-2 rounded-lg text-xs font-bold" style={{ background: f.rentalMode === m ? C.orange : "#F7F3EC12", color: f.rentalMode === m ? C.navyDeep : C.sand }}>{lb}</button>
            ))}
          </div>
          {f.rentalMode === "manual" && (
            <div className="rounded-xl p-2.5 space-y-1.5" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
              {(rentals || []).map((r) => {
                const on = (f.rentalNames || []).includes(r.name);
                return (
                  <button key={r.name} onClick={() => setF({ ...f, rentalNames: on ? f.rentalNames.filter((n) => n !== r.name) : [...(f.rentalNames || []), r.name] })}
                    className="w-full flex items-center gap-2 p-2 rounded-lg text-left text-xs font-bold" style={{ background: on ? "#1EA89625" : "#F7F3EC0A", color: C.sand, border: `1px solid ${on ? C.teal : "#F7F3EC15"}` }}>
                    <span>{on ? "☑" : "☐"}</span>
                    <span className="flex-1">{r.name}</span>
                    <span style={{ color: "#F7F3EC66" }}>{r.zone || "近海"}｜NT${r.price}｜{r.isSet ? "套組" : "釣組"}</span>
                  </button>
                );
              })}
              {(f.rentalNames || []).length === 0 && <p className="text-xs" style={{ color: C.orange }}>⚠ 尚未勾選任何項目，前台租借選單將為空。</p>}
            </div>
          )}
          <textarea className={aInput} style={aStyle} rows={2} placeholder={"備註（" + langName + "）"} value={f["note" + sfx] || ""} onChange={(e) => setF({ ...f, ["note" + sfx]: e.target.value })} />
          <div className="flex items-center justify-between pt-1">
            <span className="text-xs font-bold" style={{ color: C.yellow }}>🔗 外部連結（顯示於前台裝備建議卡底部，可導向部落格／影片）</span>
            <button onClick={() => setF({ ...f, links: [...(f.links || []), { label: "", label_en: "", label_ja: "", url: "" }] })} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增</button>
          </div>
          {(f.links || []).map((lk, i) => (
            <div key={i} className="rounded-xl p-2.5 space-y-1.5" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
              <div className="flex gap-2">
                <input className={aInput} style={aStyle} placeholder={"按鈕文字（" + langName + "），例：小搞搞教學影片"} value={lk["label" + sfx] || ""} onChange={(e) => setF({ ...f, links: f.links.map((x, idx) => (idx === i ? { ...x, ["label" + sfx]: e.target.value } : x)) })} />
                <DeleteBtn onConfirm={() => setF({ ...f, links: f.links.filter((_, idx) => idx !== i) })} />
              </div>
              <input className={aInput} style={aStyle} placeholder="網址（共用，https://…）" value={lk.url || ""} onChange={(e) => setF({ ...f, links: f.links.map((x, idx) => (idx === i ? { ...x, url: e.target.value } : x)) })} />
            </div>
          ))}
          <div className="flex gap-2 pt-1">
            <Btn full small kind="teal" onClick={save}>儲存變更</Btn>
            <DeleteBtn onConfirm={() => onDelete(trip.id)} />
          </div>
        </div>
      )}
    </div>
  );
}

function AdminTrips({ trips, orders, rentals, setDb }) {
  const [f, setF] = useState({ date: "", name: "", type: "近海", price: "", muster: "", depth: "", targets: "" });
  const add = () => {
    if (!f.date || !f.name || !f.price) return;
    const t = { id: "T" + Date.now(), date: f.date, name: f.name, type: f.type, price: Number(f.price), muster: f.muster || "05:00", rodsUp: "—", back: "—", capacity: 10, targets: f.targets ? f.targets.split(/[、,，\s]+/).filter(Boolean) : ["洽船長"], depth: f.depth || "—", gear: "洽船長建議", rigs: [], note: "", status: "報名中" };
    setDb((d) => ({ ...d, trips: [...d.trips, t].sort((a, b) => a.date.localeCompare(b.date)) }));
    setF({ date: "", name: "", type: "近海", price: "", muster: "", depth: "", targets: "" });
  };
  const saveTrip = (t) => setDb((d) => ({ ...d, trips: d.trips.map((x) => (x.id === t.id ? t : x)).sort((a, b) => a.date.localeCompare(b.date)) }));
  const delTrip = (id) => setDb((d) => ({ ...d, trips: d.trips.filter((t) => t.id !== id) }));
  return (
    <div className="space-y-3">
      <div className="rounded-2xl p-4 space-y-2" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="font-bold text-sm" style={{ color: C.yellow }}>＋ 開新船班</div>
        <div className="grid grid-cols-2 gap-2">
          <input className={aInput} style={aStyle} placeholder="日期 2026-07-31" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
          <input className={aInput} style={aStyle} placeholder="船班名稱" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
          <select className={aInput} style={{ ...aStyle, background: C.navy }} value={f.type} onChange={(e) => setF({ ...f, type: e.target.value })}>{["近海", "中遠程", "離島", "專案"].map((t) => <option key={t}>{t}</option>)}</select>
          <input className={aInput} style={aStyle} placeholder="船資（數字）" value={f.price} onChange={(e) => setF({ ...f, price: e.target.value })} />
          <input className={aInput} style={aStyle} placeholder="集合 05:00" value={f.muster} onChange={(e) => setF({ ...f, muster: e.target.value })} />
          <input className={aInput} style={aStyle} placeholder="水深 20–40 m" value={f.depth} onChange={(e) => setF({ ...f, depth: e.target.value })} />
        </div>
        <input className={aInput} style={aStyle} placeholder="目標魚（頓號分隔）" value={f.targets} onChange={(e) => setF({ ...f, targets: e.target.value })} />
        <Btn full small onClick={add}>建立船班</Btn>
      </div>
      {trips.map((t) => <TripEditor key={t.id} trip={t} orderCount={orders.filter((o) => o.tripId === t.id).length} rentals={rentals} onSave={saveTrip} onDelete={delTrip} />)}
    </div>
  );
}

/* 客戶新增／編輯 */
function AdminCustomers({ customers, orders, setDb }) {
  const [editId, setEditId] = useState(null);
  const [f, setF] = useState({ nickname: "", name: "", gender: "男", birth: "", idno: "", address: "", phone: "" });
  const [adding, setAdding] = useState(false);
  const blank = { nickname: "", name: "", gender: "男", birth: "", idno: "", address: "", phone: "" };
  const startAdd = () => { setAdding(true); setEditId(null); setF(blank); };
  const startEdit = (c) => { setEditId(c.id); setAdding(false); setF({ nickname: c.nickname || "", name: c.name, gender: c.gender, birth: c.birth, idno: c.idno, address: c.address, phone: c.phone }); };
  const saveNew = () => {
    if (!f.name || !f.idno || !f.birth) return;
    setDb((d) => ({ ...d, customers: [...d.customers, { ...f, id: "C" + Date.now(), createdAt: new Date().toISOString(), minorMode: "none", minorCount: 0, minors: [] }] }));
    setAdding(false); setF(blank);
  };
  const saveEdit = () => { setDb((d) => ({ ...d, customers: d.customers.map((c) => (c.id === editId ? { ...c, ...f } : c)) })); setEditId(null); };
  const del = (id) => setDb((d) => ({ ...d, customers: d.customers.filter((c) => c.id !== id) }));
  const Fields = ({ onSave, onCancel }) => (
    <div className="rounded-2xl p-4 space-y-2 mb-2" style={{ background: "#F7F3EC10", border: `1px solid ${C.orange}` }}>
      <div className="grid grid-cols-2 gap-2">
        <input className={aInput} style={aStyle} placeholder="暱稱" value={f.nickname} onChange={(e) => setF({ ...f, nickname: e.target.value })} />
        <input className={aInput} style={aStyle} placeholder="姓名 *" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
        <select className={aInput} style={{ ...aStyle, background: C.navy }} value={f.gender} onChange={(e) => setF({ ...f, gender: e.target.value })}><option>男</option><option>女</option></select>
        <input className={aInput} style={aStyle} placeholder="民國生日（例：80-05-12）*" value={f.birth} onChange={(e) => setF({ ...f, birth: e.target.value })} />
        <input className={aInput} style={aStyle} placeholder="證號 *" value={f.idno} onChange={(e) => setF({ ...f, idno: e.target.value })} />
        <input className={aInput} style={aStyle} placeholder="電話" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} />
      </div>
      <input className={aInput} style={aStyle} placeholder="地址" value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} />
      <div className="flex gap-2"><Btn full small kind="teal" onClick={onSave}>儲存</Btn><Btn small kind="ghost" onClick={onCancel}>取消</Btn></div>
    </div>
  );
  return (
    <div className="space-y-2.5">
      {!adding && <Btn full small onClick={startAdd}>＋ 新增客戶</Btn>}
      {adding && <Fields onSave={saveNew} onCancel={() => setAdding(false)} />}
      {customers.length === 0 && !adding && <p className="text-center text-sm py-8" style={{ color: "#F7F3EC66" }}>尚無客戶資料。</p>}
      {customers.map((c) => editId === c.id ? <Fields key={c.id} onSave={saveEdit} onCancel={() => setEditId(null)} /> : (
        <div key={c.id} className="rounded-2xl p-4" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
          <div className="flex justify-between items-start gap-2">
            <div className="font-bold text-sm" style={{ color: C.sand }}>{c.name}{c.nickname ? `（${c.nickname}）` : ""}・{c.gender}</div>
            <div className="flex gap-1.5">
              <button onClick={() => startEdit(c)} className="px-3 py-1.5 rounded-xl text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>編輯</button>
              <DeleteBtn onConfirm={() => del(c.id)} />
            </div>
          </div>
          <div className="text-xs mt-1 space-y-0.5" style={{ color: "#F7F3EC88" }}>
            <div>🪪 {c.idno}｜🎂 {c.birth}</div><div>📞 {c.phone}</div><div>🏠 {c.address}</div>
            {c.minorCount > 0 && <div>👶 未成年同行：{c.minorCount} 位（{c.minorMode === "self" ? "已代填" : "待聯絡"}）</div>}
            <div>出海次數：{orders.filter((o) => o.customerId === c.id && o.status !== "已取消").length} 次</div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* 內容管理：公告 + 規定 */
function AdminContent({ db, setDb, ping }) {
  const [eLang, setELang] = useState("zh"); // 編輯語言
  const sfx = eLang === "zh" ? "" : "_" + eLang; // 欄位後綴
  const langName = { zh: "中文", en: "English", ja: "日本語" }[eLang];
  const [ann, setAnn] = useState(typeof db.announcement === "string" ? { zh: db.announcement, en: "", ja: "" } : db.announcement);
  const [rules, setRules] = useState(db.rules);
  const [posts, setPosts] = useState(db.posts);
  const [pr, setPr] = useState(db.pricing);
  const setPrItem = (i, k, v) => setPr((p) => p.map((x, idx) => (idx === i ? { ...x, [k]: v } : x)));
  const addPr = () => setPr((p) => [...p, { name: "", name_en: "", name_ja: "", price: "", hours: "", hours_en: "", hours_ja: "", note: "", note_en: "", note_ja: "" }]);
  const delPr = (i) => setPr((p) => p.filter((_, idx) => idx !== i));
  const savePr = () => { setDb((d) => ({ ...d, pricing: pr.filter((x) => x.name) })); ping && ping("✅ 船班費用已儲存"); };
  const [cp, setCp] = useState(db.contactPage);
  const saveAnn = () => { setDb((d) => ({ ...d, announcement: ann })); ping && ping("✅ 公告已儲存"); };
  const saveRules = () => { setDb((d) => ({ ...d, rules })); ping && ping("✅ 規定已儲存"); };
  const savePosts = () => { setDb((d) => ({ ...d, posts })); ping && ping("✅ 公告貼文已儲存"); };
  const saveCp = () => { setDb((d) => ({ ...d, contactPage: cp })); ping && ping("✅ 聯絡方式已儲存"); };
  const setPost = (i, k, v) => setPosts((p) => p.map((x, idx) => (idx === i ? { ...x, [k]: v } : x)));
  const addPost = () => setPosts((p) => [{ id: "P" + Date.now(), date: new Date().toISOString().slice(0, 10), title: "新公告", body: "" }, ...p]);
  const delPost = (i) => setPosts((p) => p.filter((_, idx) => idx !== i));
  const setCpItem = (i, k, v) => setCp((c) => ({ ...c, items: c.items.map((x, idx) => (idx === i ? { ...x, [k]: v } : x)) }));
  const addCpItem = () => setCp((c) => ({ ...c, items: [...c.items, { icon: "📌", label: "", value: "", link: "" }] }));
  const delCpItem = (i) => setCp((c) => ({ ...c, items: c.items.filter((_, idx) => idx !== i) }));
  const setRule = (i, k, v) => setRules((p) => p.map((r, idx) => (idx === i ? { ...r, [k]: v } : r)));
  const addRule = () => setRules((p) => [...p, { icon: "📌", title: "新規定", body: "" }]);
  const delRule = (i) => setRules((p) => p.filter((_, idx) => idx !== i));
  return (
    <div className="space-y-4">
      <div className="rounded-2xl p-3 flex items-center gap-2 sticky top-2 z-10" style={{ background: C.navyDeep, border: `1px solid ${C.orange}` }}>
        <span className="text-xs font-bold" style={{ color: C.yellow }}>🌐 編輯語言</span>
        {[["zh", "中文"], ["en", "EN"], ["ja", "日"]].map(([l, lb]) => (
          <button key={l} onClick={() => setELang(l)} className="px-3 py-1.5 rounded-full text-xs font-bold" style={{ background: eLang === l ? C.orange : "#F7F3EC15", color: eLang === l ? C.navyDeep : C.sand }}>{lb}</button>
        ))}
        <span className="text-xs ml-auto" style={{ color: "#F7F3EC77" }}>{eLang === "zh" ? "中文為必填基底" : "留空則前台顯示中文"}</span>
      </div>
      <div className="rounded-2xl p-4 space-y-2" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="font-bold text-sm" style={{ color: C.yellow }}>📣 最新公告（登入後首個彈窗）｜{langName}</div>
        <textarea className={aInput} style={aStyle} rows={5} value={ann[eLang] || ""} onChange={(e) => setAnn({ ...ann, [eLang]: e.target.value })} />
        <Btn full small kind="teal" onClick={saveAnn}>儲存公告</Btn>
      </div>
      <div className="rounded-2xl p-4 space-y-3" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="flex items-center justify-between">
          <div className="font-bold text-sm" style={{ color: C.yellow }}>💰 常見船班費用（前台「費用」頁，儲存後即時同步）</div>
          <button onClick={addPr} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增</button>
        </div>
        {pr.map((x, i) => (
          <div key={i} className="rounded-xl p-3 space-y-2" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
            <div className="flex gap-2">
              <input className={aInput} style={aStyle} placeholder={"班別名稱（" + langName + "）"} value={x["name" + sfx] || ""} onChange={(e) => setPrItem(i, "name" + sfx, e.target.value)} />
              <DeleteBtn onConfirm={() => delPr(i)} />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <input className={aInput} style={aStyle} placeholder="價格（共用，例 NT$2,500）" value={x.price} onChange={(e) => setPrItem(i, "price", e.target.value)} />
              <input className={aInput} style={aStyle} placeholder={"時數（" + langName + "）"} value={x["hours" + sfx] || ""} onChange={(e) => setPrItem(i, "hours" + sfx, e.target.value)} />
            </div>
            <textarea className={aInput} style={aStyle} rows={2} placeholder={"說明（" + langName + "）"} value={x["note" + sfx] || ""} onChange={(e) => setPrItem(i, "note" + sfx, e.target.value)} />
          </div>
        ))}
        <Btn full small kind="teal" onClick={savePr}>儲存船班費用（同步前台）</Btn>
      </div>
      <div className="rounded-2xl p-4 space-y-3" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="flex items-center justify-between">
          <div className="font-bold text-sm" style={{ color: C.yellow }}>📰 「聯絡公告」頁：公告貼文</div>
          <button onClick={addPost} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增公告</button>
        </div>
        {posts.map((po, i) => (
          <div key={po.id} className="rounded-xl p-3 space-y-2" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
            <div className="flex gap-2">
              <input className="w-32 p-2 rounded-lg text-xs" style={aStyle} placeholder="日期" value={po.date} onChange={(e) => setPost(i, "date", e.target.value)} />
              <input className={aInput} style={aStyle} placeholder={"標題（" + langName + "）"} value={po["title" + sfx] || ""} onChange={(e) => setPost(i, "title" + sfx, e.target.value)} />
              <DeleteBtn onConfirm={() => delPost(i)} />
            </div>
            <textarea className={aInput} style={aStyle} rows={3} placeholder={"內容（" + langName + "）"} value={po["body" + sfx] || ""} onChange={(e) => setPost(i, "body" + sfx, e.target.value)} />
          </div>
        ))}
        <Btn full small kind="teal" onClick={savePosts}>儲存所有公告</Btn>
      </div>
      <div className="rounded-2xl p-4 space-y-3" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="flex items-center justify-between">
          <div className="font-bold text-sm" style={{ color: C.yellow }}>💬 「聯絡公告」頁：聯絡方式</div>
          <button onClick={addCpItem} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增</button>
        </div>
        <textarea className={aInput} style={aStyle} rows={2} placeholder={"頁面說明文字（" + langName + "）"} value={cp["intro" + sfx] || ""} onChange={(e) => setCp({ ...cp, ["intro" + sfx]: e.target.value })} />
        {cp.items.map((it, i) => (
          <div key={i} className="rounded-xl p-3 space-y-2" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
            <div className="flex gap-2">
              <input className="w-14 p-2 rounded-lg text-center" style={aStyle} value={it.icon} onChange={(e) => setCpItem(i, "icon", e.target.value)} />
              <input className={aInput} style={aStyle} placeholder={"名稱（" + langName + "）"} value={it["label" + sfx] || ""} onChange={(e) => setCpItem(i, "label" + sfx, e.target.value)} />
              <DeleteBtn onConfirm={() => delCpItem(i)} />
            </div>
            <input className={aInput} style={aStyle} placeholder={"說明文字（" + langName + "）"} value={it["value" + sfx] || ""} onChange={(e) => setCpItem(i, "value" + sfx, e.target.value)} />
            <input className={aInput} style={aStyle} placeholder="連結（https:// 或 tel:，可留空）" value={it.link} onChange={(e) => setCpItem(i, "link", e.target.value)} />
          </div>
        ))}
        <Btn full small kind="teal" onClick={saveCp}>儲存聯絡方式</Btn>
      </div>
      <div className="rounded-2xl p-4 space-y-3" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="flex items-center justify-between">
          <div className="font-bold text-sm" style={{ color: C.yellow }}>📜 報班規定 / 說明</div>
          <button onClick={addRule} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增</button>
        </div>
        {rules.map((r, i) => (
          <div key={i} className="rounded-xl p-3 space-y-2" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
            <div className="flex gap-2">
              <input className="w-16 p-2 rounded-lg text-center text-lg" style={aStyle} value={r.icon} onChange={(e) => setRule(i, "icon", e.target.value)} />
              <input className={aInput} style={aStyle} placeholder={"標題（" + langName + "）"} value={r["title" + sfx] || ""} onChange={(e) => setRule(i, "title" + sfx, e.target.value)} />
              <DeleteBtn onConfirm={() => delRule(i)} />
            </div>
            <textarea className={aInput} style={aStyle} rows={2} placeholder={"內容（" + langName + "）"} value={r["body" + sfx] || ""} onChange={(e) => setRule(i, "body" + sfx, e.target.value)} />
          </div>
        ))}
        <Btn full small kind="teal" onClick={saveRules}>儲存所有規定</Btn>
      </div>
    </div>
  );
}


/* 庫存管理：三語品名、新增/編輯/刪除、數量調整 + 租借費用管理（同步前台） */
function AdminInventory({ inventory, rentals, setDb, ping }) {
  const [eL, setEL] = useState("zh");
  const sfx = eL === "zh" ? "" : "_" + eL;
  const langName = { zh: "中文", en: "English", ja: "日本語" }[eL];
  const [adding, setAdding] = useState(false);
  const [editId, setEditId] = useState(null);
  const blank = { name: "", name_en: "", name_ja: "", cat: "裝備", qty: "", unit: "組", low: "" };
  const [f, setF] = useState(blank);
  const adjInv = (id, delta) => setDb((d) => ({ ...d, inventory: d.inventory.map((i) => (i.id === id ? { ...i, qty: Math.max(0, i.qty + delta) } : i)) }));
  const del = (id) => setDb((d) => ({ ...d, inventory: d.inventory.filter((i) => i.id !== id) }));
  const saveNew = () => {
    if (!f.name) return;
    setDb((d) => ({ ...d, inventory: [...d.inventory, { id: "INV" + Date.now(), ...f, qty: Number(f.qty) || 0, unit: f.unit || "個", low: Number(f.low) || 0 }] }));
    setAdding(false); setF(blank);
  };
  const saveEdit = () => {
    setDb((d) => ({ ...d, inventory: d.inventory.map((i) => (i.id === editId ? { ...i, ...f, qty: Number(f.qty) || 0, low: Number(f.low) || 0 } : i)) }));
    setEditId(null);
  };
  const renderForm = (onSave, onCancel) => (
    <div className="rounded-2xl p-4 space-y-2 mb-2" style={{ background: "#F7F3EC10", border: `1px solid ${C.orange}` }}>
      <input className={aInput} style={aStyle} placeholder={"品項名稱（" + langName + "）" + (eL === "zh" ? " *" : "，留空顯示中文")} value={f["name" + sfx] || ""} onChange={(e) => setF({ ...f, ["name" + sfx]: e.target.value })} />
      <div className="grid grid-cols-2 gap-2">
        <select className={aInput} style={{ ...aStyle, background: C.navy }} value={f.cat} onChange={(e) => setF({ ...f, cat: e.target.value })}><option>裝備</option><option>耗材</option></select>
        <input className={aInput} style={aStyle} placeholder="數量" inputMode="numeric" value={f.qty} onChange={(e) => setF({ ...f, qty: e.target.value })} />
        <input className={aInput} style={aStyle} placeholder="單位（組/個/條…）" value={f.unit} onChange={(e) => setF({ ...f, unit: e.target.value })} />
        <input className={aInput} style={aStyle} placeholder="安全庫存" inputMode="numeric" value={f.low} onChange={(e) => setF({ ...f, low: e.target.value })} />
      </div>
      <div className="flex gap-2"><Btn full small kind="teal" onClick={onSave}>儲存</Btn><Btn small kind="ghost" onClick={onCancel}>取消</Btn></div>
    </div>
  );

  /* 租借費用編輯 */
  const [rf, setRf] = useState(rentals);
  const setRItem = (i, k, v) => setRf((p) => p.map((x, idx) => (idx === i ? { ...x, [k]: v } : x)));
  const addR = () => setRf((p) => [...p, { name: "", name_en: "", name_ja: "", price: 0, unit: "個", unit_en: "", unit_ja: "", note: "", note_en: "", note_ja: "", isSet: false, zone: "近海" }]);
  const delR = (i) => setRf((p) => p.filter((_, idx) => idx !== i));
  const saveR = () => {
    const cleaned = rf.filter((r) => r.name && String(r.name).trim()).map((r) => ({ ...r, price: Number(r.price) || 0 }));
    setDb((d) => ({ ...d, rentals: cleaned }));
    setRf(cleaned);
    if (ping) ping("✅ 租借費用已儲存，前台已同步");
  };

  return (
    <div className="space-y-3">
      <div className="rounded-2xl p-3 flex items-center gap-2" style={{ background: C.navyDeep, border: `1px solid ${C.orange}` }}>
        <span className="text-xs font-bold" style={{ color: C.yellow }}>🌐 編輯語言</span>
        {[["zh", "中文"], ["en", "EN"], ["ja", "日"]].map(([l, lb]) => (
          <button key={l} onClick={() => setEL(l)} className="px-3 py-1.5 rounded-full text-xs font-bold" style={{ background: eL === l ? C.orange : "#F7F3EC15", color: eL === l ? C.navyDeep : C.sand }}>{lb}</button>
        ))}
        <span className="text-xs ml-auto" style={{ color: "#F7F3EC77" }}>{eL === "zh" ? "中文為必填基底" : "留空則顯示中文"}</span>
      </div>

      {/* 裝備／釣組租借費用（同步前台） */}
      <div className="rounded-2xl p-4 space-y-2.5" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="flex items-center justify-between">
          <div className="font-bold text-sm" style={{ color: C.yellow }}>💰 裝備／釣組租借費用（儲存後即時同步前台）</div>
          <button onClick={addR} className="px-3 py-1 rounded-full text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>＋ 新增</button>
        </div>
        {rf.map((r, i) => (
          <div key={i} className="rounded-xl p-3 space-y-2" style={{ background: "#0C2D4830", border: "1px solid #F7F3EC1A" }}>
            <div className="flex gap-2 items-center">
              <input className={aInput} style={aStyle} placeholder={"名稱（" + langName + "）"} value={r["name" + sfx] || ""} onChange={(e) => setRItem(i, "name" + sfx, e.target.value)} />
              <DeleteBtn onConfirm={() => delR(i)} />
            </div>
            <div className="grid grid-cols-3 gap-2">
              <input className={aInput} style={aStyle} placeholder="費用 NT$" inputMode="numeric" value={r.price} onChange={(e) => setRItem(i, "price", e.target.value)} />
              <input className={aInput} style={aStyle} placeholder={"單位（" + langName + "）"} value={r["unit" + sfx] || ""} onChange={(e) => setRItem(i, "unit" + sfx, e.target.value)} />
              <button onClick={() => setRItem(i, "isSet", !r.isSet)} className="rounded-xl text-xs font-bold" style={{ background: r.isSet ? C.teal : "#F7F3EC15", color: r.isSet ? "#fff" : C.sand }}>{r.isSet ? "整組套組（可租）" : "釣組耗材（加購）"}</button>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="text-xs font-bold" style={{ color: "#F7F3EC88" }}>適用釣場</span>
              {["近海", "深場", "通用"].map((z) => (
                <button key={z} onClick={() => setRItem(i, "zone", z)} className="px-3 py-1.5 rounded-full text-xs font-bold" style={{ background: (r.zone || "近海") === z ? (z === "深場" ? C.navy : z === "通用" ? C.gray : C.teal) : "#F7F3EC15", color: (r.zone || "近海") === z ? "#fff" : C.sand }}>{z}</button>
              ))}
              <span className="text-xs ml-auto" style={{ color: "#F7F3EC55" }}>水深≤80m→近海｜{">"}80m→深場</span>
            </div>
            <input className={aInput} style={aStyle} placeholder={"備註（" + langName + "）"} value={r["note" + sfx] || ""} onChange={(e) => setRItem(i, "note" + sfx, e.target.value)} />
          </div>
        ))}
        <Btn full small kind="teal" onClick={saveR}>儲存租借費用（同步前台）</Btn>
        <p className="text-xs" style={{ color: "#F7F3EC66" }}>「整組套組」會出現在報名的租借下拉選單；「釣組耗材」僅顯示於費用表。前台會依船班釣場水深自動篩選：水深 ≤80 米顯示「近海」＋「通用」項目，{">"}80 米顯示「深場」＋「通用」項目；水深無法判讀的船班顯示全部。</p>
      </div>

      {/* 庫存品項 */}
      {!adding && <Btn full small onClick={() => { setAdding(true); setEditId(null); setF(blank); }}>＋ 新增庫存品項</Btn>}
      {adding && renderForm(saveNew, () => setAdding(false))}
      {["裝備", "耗材"].map((cat) => (
        <div key={cat}>
          <div className="font-bold text-sm mb-2" style={{ color: C.yellow }}>{cat}</div>
          {inventory.filter((i) => i.cat === cat).map((i) => editId === i.id ? <div key={i.id}>{renderForm(saveEdit, () => setEditId(null))}</div> : (
            <div key={i.id} className="rounded-2xl p-3.5 mb-2 flex items-center gap-2.5" style={{ background: "#F7F3EC10", border: `1px solid ${i.qty <= i.low ? C.orange : "#F7F3EC22"}` }}>
              <div className="flex-1">
                <div className="font-bold text-sm" style={{ color: C.sand }}>{i.name}{i.qty <= i.low && <span style={{ color: C.orange }}> ⚠</span>}</div>
                <div className="text-xs" style={{ color: "#F7F3EC77" }}>{(i.name_en || i.name_ja) ? `EN: ${i.name_en || "—"}｜日: ${i.name_ja || "—"}｜` : ""}安全庫存 {i.low} {i.unit}</div>
              </div>
              <button onClick={() => adjInv(i.id, -1)} className="w-9 h-9 rounded-xl font-black" style={{ background: "#F7F3EC15", color: C.sand }}>−</button>
              <div className="font-black w-12 text-center" style={{ color: i.qty <= i.low ? C.orange : C.teal }}>{i.qty}<span className="text-xs font-bold" style={{ color: "#F7F3EC66" }}> {i.unit}</span></div>
              <button onClick={() => adjInv(i.id, 1)} className="w-9 h-9 rounded-xl font-black" style={{ background: "#F7F3EC15", color: C.sand }}>＋</button>
              <button onClick={() => { setEditId(i.id); setAdding(false); setF({ ...blank, ...i, qty: String(i.qty), low: String(i.low) }); }} className="px-2.5 py-1.5 rounded-xl text-xs font-bold" style={{ background: C.orange, color: C.navyDeep }}>編輯</button>
              <DeleteBtn onConfirm={() => del(i.id)} />
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

/* ===== 匯出：船班報名名單 / 所有客戶名單（Excel .xlsx / Numbers 適用 .csv） ===== */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function exportRows(rows, filename, fmt) {
  if (!rows.length) return false;
  if (fmt === "xlsx") {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "名單");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename + ".xlsx");
  } else {
    const keys = Object.keys(rows[0]);
    const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const csv = "\uFEFF" + [keys.map(esc).join(","), ...rows.map((r) => keys.map((k) => esc(r[k])).join(","))].join("\r\n");
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), filename + ".csv");
  }
  return true;
}
function AdminExport({ trips, orders, customers, ping }) {
  const [tripId, setTripId] = useState(trips[0]?.id || "");
  const trip = trips.find((t) => t.id === tripId);
  const rosterRows = () => {
    if (!trip) return [];
    return orders.filter((o) => o.tripId === trip.id && o.status !== "已取消").map((o, i) => {
      const c = customers.find((x) => x.id === o.customerId) || {};
      const minor = o.isMinor && (c.minors || []).find((m) => m.name === o.name);
      return {
        "序號": i + 1, "船班日期": trip.date, "船班名稱": trip.name, "姓名": o.name, "暱稱": o.nickname,
        "身份": o.role + (o.isMinor ? "（未成年）" : o.isCompanion ? "（攜伴）" : ""),
        "法定性別": o.isCompanion && !o.isMinor ? "" : (minor?.gender || c.gender || ""),
        "民國生日": minor?.birth || (o.isCompanion && !o.isMinor ? "" : c.birth) || "",
        "證號": minor?.idno || (o.isCompanion && !o.isMinor ? "現場補登" : c.idno) || "",
        "地址": o.isCompanion ? "" : (c.address || ""), "聯絡電話": o.phone,
        "裝備租借": o.rental || "無", "金額(NT$)": (o.price || 0) + (o.rentalPrice || 0),
        "收款": o.paid ? "已收款" : "未收款", "狀態": o.status,
      };
    });
  };
  const customerRows = () => customers.map((c, i) => ({
    "序號": i + 1, "姓名": c.name, "暱稱": c.nickname || "", "法定性別": c.gender, "民國生日": c.birth,
    "證號": c.idno, "地址": c.address, "聯絡電話": c.phone,
    "未成年同行": c.minorCount ? `${c.minorCount} 位（${c.minorMode === "self" ? "已代填" : "待聯絡"}）` : "無",
    "建檔日期": (c.createdAt || "").slice(0, 10),
  }));
  const doExport = (rows, name, fmt) => {
    if (!exportRows(rows, name, fmt)) ping("沒有可匯出的資料");
    else ping("已匯出 " + name + (fmt === "xlsx" ? ".xlsx" : ".csv"));
  };
  return (
    <div className="space-y-3">
      <div className="rounded-2xl p-4 space-y-2.5" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="font-bold text-sm" style={{ color: C.yellow }}>📋 匯出船班報名客戶名單</div>
        <select className={aInput} style={{ ...aStyle, background: C.navy }} value={tripId} onChange={(e) => setTripId(e.target.value)}>
          {trips.map((t) => <option key={t.id} value={t.id}>{t.date}・{t.name}（{orders.filter((o) => o.tripId === t.id && o.status !== "已取消").length} 位）</option>)}
        </select>
        <div className="grid grid-cols-2 gap-2">
          <Btn small kind="teal" onClick={() => doExport(rosterRows(), `${trip?.date || ""}_${trip?.name || "船班"}_報名名單`, "xlsx")}>Excel 檔（.xlsx）</Btn>
          <Btn small onClick={() => doExport(rosterRows(), `${trip?.date || ""}_${trip?.name || "船班"}_報名名單`, "csv")}>Numbers 檔（.csv）</Btn>
        </div>
      </div>
      <div className="rounded-2xl p-4 space-y-2.5" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
        <div className="font-bold text-sm" style={{ color: C.yellow }}>👥 匯出所有客戶名單（{customers.length} 位）</div>
        <div className="grid grid-cols-2 gap-2">
          <Btn small kind="teal" onClick={() => doExport(customerRows(), "五賀_所有客戶名單", "xlsx")}>Excel 檔（.xlsx）</Btn>
          <Btn small onClick={() => doExport(customerRows(), "五賀_所有客戶名單", "csv")}>Numbers 檔（.csv）</Btn>
        </div>
      </div>
      <p className="text-xs" style={{ color: "#F7F3EC66" }}>※ .xlsx 與 .csv 皆可直接用 Apple Numbers 開啟；.csv 已含 BOM，中文不會亂碼。海巡報關可直接使用船班名單。</p>
    </div>
  );
}

function AdminPortal({ db, setDb, onExit }) {
  const [tab, setTab] = useState("dash");
  const [aToast, setAToast] = useState("");
  const aPing = (m) => { setAToast(m); setTimeout(() => setAToast(""), 2400); };
  const { trips, orders, customers, inventory } = db;
  const activeOrders = orders.filter((o) => o.status !== "已取消");
  const revenue = activeOrders.reduce((s, o) => s + (o.price || 0) + (o.rentalPrice || 0), 0);
  const lowStock = inventory.filter((i) => i.qty <= i.low);
  const setOrderStatus = (id, status) => setDb((d) => ({ ...d, orders: d.orders.map((o) => (o.id === id ? { ...o, status } : o)) }));
  const setPaid = (id, paid) => setDb((d) => ({ ...d, orders: d.orders.map((o) => (o.id === id ? { ...o, paid } : o)) }));
  const tabs = [["dash", "總覽"], ["orders", "訂單"], ["trips", "船班"], ["customers", "客戶"], ["inv", "庫存"], ["export", "匯出"], ["content", "內容"]];
  return (
    <div className="min-h-screen" style={{ background: C.navyDeep }}>
      <div className="max-w-2xl mx-auto p-4 pb-24">
        <div className="flex items-center justify-between mb-4">
          <div><h1 className="font-black text-xl" style={{ color: C.yellow }}>⚓ 船長後台</h1><p className="text-xs" style={{ color: "#F7F3EC77" }}>GOHO.ERP 內部資料庫</p></div>
          <Btn kind="ghost" small onClick={onExit}>返回前台</Btn>
        </div>
        <div className="flex gap-1.5 mb-4 overflow-x-auto pb-1">
          {tabs.map(([k, label]) => (
            <button key={k} onClick={() => setTab(k)} className="px-4 py-2 rounded-full text-sm font-bold whitespace-nowrap" style={{ background: tab === k ? C.orange : "#F7F3EC12", color: tab === k ? C.navyDeep : C.sand }}>{label}</button>
          ))}
        </div>

        {tab === "dash" && (
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              {[["有效訂單", activeOrders.length, "筆"], ["營收（含租借）", `NT$${revenue.toLocaleString()}`, ""], ["客戶數", customers.length, "位"], ["低庫存警示", lowStock.length, "項"]].map(([k, v, u]) => (
                <div key={k} className="rounded-2xl p-4" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
                  <div className="text-xs" style={{ color: "#F7F3EC88" }}>{k}</div>
                  <div className="font-black text-2xl mt-1" style={{ color: k === "低庫存警示" && lowStock.length ? C.orange : C.teal }}>{v}<span className="text-sm font-bold ml-1" style={{ color: "#F7F3EC66" }}>{u}</span></div>
                </div>
              ))}
            </div>
            <div className="rounded-2xl p-4" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
              <div className="font-bold text-sm mb-2" style={{ color: C.yellow }}>各船班報名狀況</div>
              {trips.map((t) => {
                const n = activeOrders.filter((o) => o.tripId === t.id).length;
                const under5 = n < 5 && TRIP_STATUS[t.status]?.bookable;
                return (
                  <div key={t.id} className="flex items-center gap-2 py-1.5 text-sm" style={{ color: C.sand }}>
                    <span className="text-xs w-12" style={{ color: "#F7F3EC77" }}>{t.date.slice(5)}</span>
                    <span className="flex-1 truncate">{t.name}</span>
                    <Tag tone={TRIP_STATUS[t.status]?.tone || "teal"}>{t.status}</Tag>
                    <span className="text-xs font-bold w-14 text-right" style={{ color: under5 ? C.orange : C.teal }}>{n}/{t.capacity}{under5 ? " ⚠" : ""}</span>
                  </div>
                );
              })}
              <p className="text-xs mt-1" style={{ color: "#F7F3EC66" }}>⚠ 未滿 5 人：依規定得於前一日通知取消</p>
            </div>
            {lowStock.length > 0 && <div className="rounded-2xl p-4" style={{ background: "#F4A25918", border: `1px solid ${C.orange}` }}><div className="font-bold text-sm" style={{ color: C.orange }}>⚠ 低庫存：{lowStock.map((i) => `${i.name}（剩 ${i.qty}${i.unit}）`).join("、")}</div></div>}
          </div>
        )}

        {tab === "orders" && (
          <div className="space-y-2.5">
            {orders.length === 0 && <p className="text-center text-sm py-8" style={{ color: "#F7F3EC66" }}>尚無訂單。前台完成報名後會即時顯示於此。</p>}
            {[...orders].reverse().map((o) => {
              const t = trips.find((x) => x.id === o.tripId);
              return (
                <div key={o.id} className="rounded-2xl p-4" style={{ background: "#F7F3EC10", border: "1px solid #F7F3EC22" }}>
                  <div className="flex justify-between items-start gap-2">
                    <div>
                      <div className="font-bold text-sm" style={{ color: C.sand }}>{o.name}（{o.nickname}）・{o.role}{o.isMinor ? "・未成年" : ""}</div>
                      <div className="text-xs mt-0.5" style={{ color: "#F7F3EC88" }}>{t?.date} {t?.name}｜📞 {o.phone}</div>
                      <div className="text-xs mt-0.5" style={{ color: "#F7F3EC88" }}>船資 NT${o.price.toLocaleString()}{o.rental ? `＋租借 ${o.rental} NT$${o.rentalPrice}` : ""}</div>
                    </div>
                    <Tag tone={o.status === "已取消" ? "red" : o.status === "候補" ? "orange" : "teal"}>{o.status}</Tag>
                  </div>
                  <div className="flex gap-2 mt-3 flex-wrap">
                    <Btn kind={o.paid ? "teal" : "ghost"} small onClick={() => setPaid(o.id, !o.paid)}>{o.paid ? "✓ 已收款" : "標記收款"}</Btn>
                    {o.status === "候補" && <Btn small onClick={() => setOrderStatus(o.id, "已報名")}>轉正取</Btn>}
                    {o.status !== "已取消" && <Btn kind="danger" small onClick={() => setOrderStatus(o.id, "已取消")}>取消</Btn>}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {tab === "trips" && <AdminTrips trips={trips} orders={activeOrders} rentals={db.rentals} setDb={setDb} />}
        {tab === "customers" && <AdminCustomers customers={customers} orders={orders} setDb={setDb} />}
        {tab === "content" && <AdminContent db={db} setDb={setDb} ping={aPing} />}

        {tab === "inv" && <AdminInventory inventory={inventory} rentals={db.rentals} setDb={setDb} ping={aPing} />}
        {aToast && <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[60] px-5 py-3 rounded-2xl font-bold text-sm shadow-xl" style={{ background: C.sand, color: C.navy, border: `2px solid ${C.teal}` }}>{aToast}</div>}
        {tab === "export" && <AdminExport trips={trips} orders={orders} customers={customers} ping={aPing} />}
      </div>
    </div>
  );
}

/* ---------- 主程式 ---------- */
export default function GohoSystem() {
  const [db, setDbState] = useState(null);
  const [stage, setStage] = useState("welcome"); // welcome | noConsent | contact | register | app
  const [user, setUser] = useState(null);
  const [tab, setTab] = useState("quest");
  const [openTrip, setOpenTrip] = useState(null);
  const [openNew, setOpenNew] = useState(null); // null=關閉；""或日期字串=開啟
  const [showRentalFee, setShowRentalFee] = useState(false);
  const [showAnnounce, setShowAnnounce] = useState(false);
  const [admin, setAdmin] = useState(false);
  const [adminAsk, setAdminAsk] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [lang, setLang] = useState("zh");
  LANG = lang;
  useEffect(() => { window.storage.get("goho-lang").then((r) => r?.value && setLang(r.value)).catch(() => {}); }, []);
  const pickLang = (l) => { setLang(l); window.storage.set("goho-lang", l).catch(() => {}); };
  const [pw, setPw] = useState("");
  const [toast, setToast] = useState("");

  useEffect(() => { loadStore().then(setDbState); }, []);
  const setDb = useCallback((fn) => { setDbState((prev) => { const next = typeof fn === "function" ? fn(prev) : fn; saveStore(next); return next; }); }, []);
  const ping = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2600); };
  const questDone = (id) => setDb((d) => ({ ...d, quest: { ...d.quest, [id]: true } }));

  if (!db) return <div className="min-h-screen flex items-center justify-center font-bold" style={{ background: C.navyDeep, color: C.sand }}>{t("loading")}</div>;

  const gotoTab = (t) => { setTab(t); if (t === "rules") questDone("rules"); if (t === "fees") questDone("fees"); };
  const afterLogin = (c, msg) => { setUser(c); setStage("app"); setShowAnnounce(true); ping(msg); };
  const handleRegister = (c) => { setDb((d) => ({ ...d, customers: [...d.customers, c] })); afterLogin(c, t("welcomeNew", { n: c.nickname || c.name })); };

  const dbRef = db;
  const handleBook = (trip, role, rental, rem, companions = []) => {
    if (role === "觀光") { ping(t("tSightOnly")); return; }
    const r = (dbRef.rentals || RENTALS).find((x) => x.name === rental);
    const now = Date.now();
    const rentalPrice = r ? r.price : 0;
    const minors = user.minorMode === "self" ? (user.minors || []) : [];
    const seatsNeeded = 1 + minors.length + companions.length;
    const status = rem >= seatsNeeded ? "已報名" : "候補";
    const mainOrder = { id: "O" + now, tripId: trip.id, customerId: user.id, name: user.name, nickname: user.nickname || user.name, phone: user.phone, role, price: trip.price, rental: rental || null, rentalPrice, status, paid: false, isMinor: false, createdAt: new Date().toISOString() };
    const minorOrders = minors.map((m, i) => ({ id: "O" + now + "-m" + i, tripId: trip.id, customerId: user.id, name: m.name, nickname: m.name + "(未成年)", phone: user.phone, role: "觀光", price: trip.price, rental: null, rentalPrice: 0, status, paid: false, isMinor: true, isCompanion: true, companionOf: "O" + now, createdAt: new Date().toISOString() }));
    const compOrders = companions.map((n, i) => ({ id: "O" + now + "-c" + i, tripId: trip.id, customerId: user.id, name: n, nickname: n, phone: user.phone, role: "觀光", price: trip.price, rental: null, rentalPrice: 0, status, paid: false, isMinor: false, isCompanion: true, companionOf: "O" + now, createdAt: new Date().toISOString() }));
    setDb((d) => ({ ...d, orders: [...d.orders, mainOrder, ...minorOrders, ...compOrders], quest: { ...d.quest, book: true } }));
    setOpenTrip(null);
    const extra = minors.length + companions.length;
    ping(status === "候補" ? t("tWait") : t("tBooked") + (extra ? t("tBookedX", { n: extra }) : ""));
  };

  const handleCreateTrip = (date, service, timing) => {
    const price = service?.basePrice || 0;
    const tm = timing || { muster: "05:00", rodsUp: "—", back: "—", timePending: false };
    const t = { id: "T" + Date.now(), date, name: service?.name || "自訂船班", type: service?.zone || "近海", price, muster: tm.muster, rodsUp: tm.rodsUp, back: tm.back, timePending: !!tm.timePending, capacity: 10, targets: ["洽船長"], depth: "洽船長", gear: "洽船長建議", rigs: [], note: "釣友自行開班，歡迎跟報", status: "歡迎開班" };
    setDb((d) => ({ ...d, trips: [...d.trips, t].sort((a, b) => a.date.localeCompare(b.date)), quest: { ...d.quest, calendar: true } }));
    setOpenNew(null);
    setTimeout(() => setOpenTrip(t), 200);
    ping(t("tCreated"));
  };

  if (admin) return <AdminPortal db={db} setDb={setDb} onExit={() => setAdmin(false)} />;

  const tabs = [["quest", "🧭", t("navQuest")], ["calendar", "📅", t("navCal")], ["mytrips", "🎣", t("navMy")], ["rules", "📜", t("navRules")], ["fees", "💰", t("navFees")], ["services", "🛎️", t("navSvc")], ["contact", "📣", t("navCt")]];

  return (
    <div className="min-h-screen" style={{ background: C.sand, fontFamily: "'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif" }}>
      <style>{`@keyframes bob{0%,100%{transform:translateY(0)}50%{transform:translateY(-5px)}}.buoy{display:inline-block;animation:bob 1.4s ease-in-out infinite}@media(prefers-reduced-motion:reduce){.buoy{animation:none}}`}</style>

      {stage !== "app" && (
        <div className="fixed top-3 right-3 z-[70] flex rounded-full overflow-hidden shadow-lg" style={{ border: "1px solid #FFFFFF55", background: "#081D30CC", backdropFilter: "blur(4px)" }}>
          {[["zh", "中"], ["en", "EN"], ["ja", "日"]].map(([l, lb]) => (
            <button key={l} onClick={() => pickLang(l)} className="px-3 py-1.5 text-xs font-bold" style={{ background: lang === l ? C.yellow : "transparent", color: lang === l ? C.navyDeep : "#FFFFFF" }}>{lb}</button>
          ))}
        </div>
      )}
      {stage === "welcome" && <WelcomeModal onYes={() => setStage("register")} onNo={() => setStage("noConsent")} />}
      {stage === "noConsent" && <NoConsentModal onBack={() => setStage("welcome")} onContact={() => setStage("contact")} />}
      {stage === "contact" && <ContactModal onBack={() => setStage("noConsent")} />}
      {stage === "register" && <RegisterModal customers={db.customers} onRegister={handleRegister} onLogin={(c) => afterLogin(c, t("welcomeBack", { n: c.nickname || c.name }))} onAdmin={() => setAdminAsk(true)} />}
      {showAnnounce && stage === "app" && <AnnouncementModal text={db.announcement} onClose={() => { setShowAnnounce(false); questDone("announce"); }} />}

      <header style={{ background: `linear-gradient(160deg, ${C.logoBlue} 0%, #0A2A6B 55%, ${C.navy} 100%)` }}>
        <div className="max-w-2xl mx-auto px-4 pt-5 pb-2 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Logo />
            <div><h1 className="font-black text-2xl tracking-wide" style={{ color: "#FFFFFF" }}>五賀娛樂漁船</h1><p className="text-xs mt-0.5 font-bold" style={{ color: "#FFFFFF" }}>{t("brandSub")}</p></div>
          </div>
          <div className="flex items-center gap-2">
          <div className="flex rounded-full overflow-hidden" style={{ border: "1px solid #FFFFFF44" }}>
            {[["zh", "中"], ["en", "EN"], ["ja", "日"]].map(([l, lb]) => (
              <button key={l} onClick={() => pickLang(l)} className="px-2.5 py-1 text-xs font-bold" style={{ background: lang === l ? C.yellow : "transparent", color: lang === l ? C.navyDeep : "#FFFFFF" }}>{lb}</button>
            ))}
          </div>
          {stage === "app" ? (
            <div className="relative">
              <button onClick={() => setMenuOpen((v) => !v)} className="text-xs px-3 py-1.5 rounded-full font-bold" style={{ background: "#F7F3EC22", color: "#FFFFFF" }}>{t("menu")}</button>
              {menuOpen && (
                <div className="absolute right-0 mt-2 w-44 rounded-2xl overflow-hidden shadow-2xl z-50" style={{ background: "#fff", border: "2px solid #0C2D4815" }}>
                  {[[t("mMy"), () => { gotoTab("mytrips"); setMenuOpen(false); }], [t("mHome"), () => { gotoTab("quest"); setMenuOpen(false); }], [t("mOut"), () => { setMenuOpen(false); setUser(null); setTab("quest"); setStage("welcome"); }]].map(([l, fn]) => (
                    <button key={l} onClick={fn} className="w-full text-left px-4 py-3 text-sm font-bold" style={{ color: C.navy, borderBottom: "1px solid #0C2D480E" }}>{l}</button>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <button onClick={() => setAdminAsk(true)} className="text-xs px-3 py-1.5 rounded-full font-bold" style={{ background: "#F7F3EC22", color: "#FFFFFF" }}>{t("adminBtn")}</button>
          )}
          </div>
        </div>
        <Wave fill={C.sand} />
      </header>

      {stage === "app" ? (
        <main className="max-w-2xl mx-auto px-4 pt-3 pb-28">
          <div className="text-sm font-bold mb-3" style={{ color: C.tealDark }}>{t("hello", { n: user?.nickname || user?.name })}</div>
          {tab === "quest" && <QuestBoard quest={db.quest} gotoTab={gotoTab} />}
          {tab === "calendar" && <TripCalendar trips={db.trips} orders={db.orders} onOpen={(t) => { setOpenTrip(t); questDone("calendar"); }} onOpenNew={(d) => setOpenNew(d ?? "")} />}
          {tab === "mytrips" && <MyTripsTab trips={db.trips} orders={db.orders} user={user} onOpen={(t) => setOpenTrip(t)} />}
          {tab === "rules" && <RulesTab rules={db.rules} />}
          {tab === "fees" && <FeesTab rentals={db.rentals} pricing={db.pricing} />}
          {tab === "services" && <ServicesTab />}
          {tab === "contact" && <ContactNewsTab posts={db.posts} contactPage={db.contactPage} />}
          <footer className="text-center pt-8 pb-2">
            <div className="text-sm font-bold" style={{ color: C.navy }}>五賀娛樂漁船 GOHO</div>
            <div className="text-xs mt-1" style={{ color: "#0C2D4877" }}>五探哥漁村有限公司</div>
          </footer>
        </main>
      ) : (
        <div className="max-w-2xl mx-auto px-4 py-16 text-center text-sm" style={{ color: "#0C2D4877" }}>{t("gate")}</div>
      )}

      {stage === "app" && (
        <>
          <FloatingBuoy quest={db.quest} tab={tab} gotoTab={gotoTab} />
          <nav className="fixed bottom-0 left-0 right-0" style={{ background: C.navy, boxShadow: "0 -4px 20px rgba(8,29,48,.3)" }}>
            <div className="max-w-2xl mx-auto grid grid-cols-7">
              {tabs.map(([k, icon, label]) => (
                <button key={k} onClick={() => gotoTab(k)} className="py-2.5 flex flex-col items-center gap-0.5">
                  <span className="text-lg">{icon}</span><span className="text-xs font-bold" style={{ color: tab === k ? C.yellow : "#F7F3EC88" }}>{label}</span>
                </button>
              ))}
            </div>
          </nav>
        </>
      )}

      {openTrip && <TripDetail trip={openTrip} orders={db.orders} user={user} rentals={db.rentals} onBook={handleBook} onClose={() => setOpenTrip(null)} onRentalInfo={() => setShowRentalFee(true)} />}
      {openNew !== null && <OpenTripModal presetDate={openNew || ""} trips={db.trips} onCreate={handleCreateTrip} onClose={() => setOpenNew(null)} />}
      {showRentalFee && <RentalFeeModal rentals={db.rentals} trip={openTrip} onClose={() => setShowRentalFee(false)} />}

      {adminAsk && (
        <ModalShell onClose={() => { setAdminAsk(false); setPw(""); }}>
          <div className="p-5 space-y-3">
            <h3 className="font-black text-center" style={{ color: C.navy }}>⚓ 船長後台登入</h3>
            <input type="password" className="w-full p-3 rounded-xl border-2 text-sm" style={{ borderColor: "#0C2D4822" }} placeholder="後台密碼" value={pw} onChange={(e) => setPw(e.target.value)} />
            <div className="grid grid-cols-2 gap-2">
              <Btn kind="light" onClick={() => { setAdminAsk(false); setPw(""); }}>取消</Btn>
              <Btn onClick={() => { if (pw === (import.meta.env?.VITE_ADMIN_PASSWORD || "goho888")) { setAdmin(true); setAdminAsk(false); setPw(""); } else ping(t("tPw")); }}>{t("login")}</Btn>
            </div>
            <p className="text-xs text-center" style={{ color: "#0C2D4866" }}>密碼由環境變數 VITE_ADMIN_PASSWORD 設定</p>
          </div>
        </ModalShell>
      )}

      {toast && <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[60] px-5 py-3 rounded-2xl font-bold text-sm shadow-xl text-center" style={{ background: C.navyDeep, color: C.yellow, border: `2px solid ${C.teal}` }}>{toast}</div>}
    </div>
  );
}
