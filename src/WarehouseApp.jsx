import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard,
  Package,
  Boxes,
  ArrowLeftRight,
  MapPin,
  History,
  Plus,
  X,
  AlertTriangle,
  Search,
  ChevronLeft,
  Check,
  RefreshCcw,
  Printer,
  PackagePlus,
  Users,
  Trash2,
  Settings,
  Pencil,
  Download,
  UserPlus,
  BookOpen,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* ثابت‌ها و توابع کمکی                                                */
/* ------------------------------------------------------------------ */

const WAREHOUSES = [
  { id: "w1", label: "انبار ۱" },
  { id: "w2", label: "انبار ۲" },
];
const GRADES = ["درجه ۱", "درجه ۲", "درجه ۳", "درجه ۴", "درجه ۵", "شیرینگ", "فله", "4G"];
const REASONS = ["شکستگی", "مفقودی", "سایر"];
const SIZES = ["6060", "8080", "100100", "120120", "60120", "80160", "40100", "40120", "3090", "3060", "4040", "3030", "6060-14", "6060-20"];
const SURFACES = ["مات", "براق", "نانو پولیش"];
const MATERIALS = ["پرسلان", "خاک سفید", "خاک قرمز"];
const CUSTOM = "__custom__";

const whLabel = (id) => WAREHOUSES.find((w) => w.id === id)?.label || id;
const uid = (p = "id") => `${p}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
const fmt = (n) => (Number(n) || 0).toLocaleString("fa-IR", { maximumFractionDigits: 2 });

function fmtDate(iso) {
  try {
    return new Intl.DateTimeFormat("fa-IR-u-ca-persian", { dateStyle: "short", timeStyle: "short" }).format(new Date(iso));
  } catch (e) {
    return new Date(iso).toLocaleString("fa-IR");
  }
}
function locLabel(loc) {
  if (!loc || (!loc.aisle && !loc.row && !loc.cell)) return "—";
  return `${loc.aisle || "-"} / ${loc.row || "-"} / ${loc.cell || "-"}`;
}
const isCartonItem = (item) => !!(item && Number(item.sqmPerBox) > 0);
function qtyDisplay(item, qty, boxCount) {
  const unit = item?.unit || "";
  if (isCartonItem(item) && boxCount != null) return `${fmt(boxCount)} کارتن (${fmt(qty)} ${unit})`;
  return `${fmt(qty)} ${unit}`;
}
const numProps = { type: "number", inputMode: "decimal", step: "0.01", min: "0" };

/* ------------------------------------------------------------------ */
/* داده‌های اولیه                                                       */
/* ------------------------------------------------------------------ */

import { createClient } from "@supabase/supabase-js";

/* ------------------------------------------------------------------ */
/* اتصال به دیتابیس Supabase                                           */
/* ------------------------------------------------------------------ */

const SUPABASE_URL = "https://rqyczuwislkdnnahbwfg.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_6mG4JD4r3kfkQzhKISArPw_gwd0K04w";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/* تبدیل نام فیلدها بین camelCase (جاوااسکریپت) و snake_case (دیتابیس) */
const FIELD_DB_TO_JS = { user_name: "user" };
const FIELD_JS_TO_DB = { user: "user_name" };
function toSnake(s) { return s.replace(/[A-Z]/g, (m) => "_" + m.toLowerCase()); }
function toCamel(s) { return s.replace(/_([a-z0-9])/g, (_, c) => c.toUpperCase()); }
function rowToCamel(row) {
  if (!row) return row;
  const out = {};
  for (const k in row) out[FIELD_DB_TO_JS[k] || toCamel(k)] = row[k];
  return out;
}
function camelToRow(obj) {
  const out = {};
  for (const k in obj) {
    if (obj[k] === undefined) continue;
    out[FIELD_JS_TO_DB[k] || toSnake(k)] = obj[k];
  }
  return out;
}
async function fetchAll(table, orderCol) {
  let q = supabase.from(table).select("*");
  if (orderCol) q = q.order(orderCol, { ascending: false });
  const { data, error } = await q;
  if (error) { console.error("خطا در دریافت", table, error); return []; }
  return (data || []).map(rowToCamel);
}
async function fetchOrdersWithLines() {
  const [orderRows, lineRows] = await Promise.all([
    supabase.from("orders").select("*").order("date", { ascending: false }),
    supabase.from("order_lines").select("*"),
  ]);
  const orders = (orderRows.data || []).map(rowToCamel);
  const lines = (lineRows.data || []).map(rowToCamel);
  return orders.map((o) => ({ ...o, lines: lines.filter((l) => l.orderId === o.id) }));
}
/* همگام‌سازی یک رکورد موجودی با دیتابیس: درج/به‌روزرسانی/حذف بر اساس وضعیت نهایی آن در nextStock */
async function syncStockRecord(nextStock, id, isNew) {
  const rec = nextStock.find((s) => s.id === id);
  if (rec) {
    if (isNew) {
      const { error } = await supabase.from("stock").insert(camelToRow(rec));
      if (error) console.error("خطا در ثبت موجودی", error);
    } else {
      const { error } = await supabase.from("stock").update(camelToRow({ qty: rec.qty, boxCount: rec.boxCount, reserved: rec.reserved, discrepancy: rec.discrepancy })).eq("id", id);
      if (error) console.error("خطا در به‌روزرسانی موجودی", error);
    }
  } else {
    const { error } = await supabase.from("stock").delete().eq("id", id);
    if (error) console.error("خطا در حذف موجودی", error);
  }
}
/* اطمینان از وجود پروفایل کارمند برای کاربر واردشده؛ در اولین ورود خودکار ساخته می‌شود */
async function ensureStaffRecord(user) {
  const { data: existing } = await supabase.from("staff").select("*").eq("auth_user_id", user.id).maybeSingle();
  if (existing) return rowToCamel(existing);
  const prefix = (user.email || "").split("@")[0];
  const nameMap = { modir: "مدیر", anbardar1: "کارگر انبار ۱", anbardar2: "کارگر انبار ۲" };
  const name = nameMap[prefix] || prefix || "کاربر جدید";
  const isManager = prefix === "modir";
  const row = { id: uid("u"), name, isManager, authUserId: user.id };
  const { data: created, error } = await supabase.from("staff").insert(camelToRow(row)).select().single();
  if (error) { console.error("خطا در ساخت پروفایل کارمند", error); return row; }
  return rowToCamel(created);
}

/* ------------------------------------------------------------------ */
/* اجزای پایه                                                          */
/* ------------------------------------------------------------------ */

function InfoTip({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span className="relative inline-block align-middle">
      <button type="button" onClick={() => setOpen((o) => !o)} className="w-4 h-4 rounded-full border border-stone-400 text-stone-400 text-[10px] font-bold inline-flex items-center justify-center hover:border-amber-700 hover:text-amber-700 leading-none">!</button>
      {open && <span className="absolute z-30 top-6 right-0 w-60 bg-stone-900 text-white text-[11px] leading-5 rounded-lg p-2.5 shadow-xl">{text}</span>}
    </span>
  );
}

function Card({ label, value, sub, tone = "default", icon: Icon, onClick }) {
  const tones = {
    default: "border-stone-200 bg-white text-stone-900",
    accent: "border-amber-700/30 bg-amber-50 text-amber-900",
    warn: "border-rose-300 bg-rose-50 text-rose-900",
    slate: "border-slate-700 bg-slate-800 text-white",
  };
  const Tag = onClick ? "button" : "div";
  return (
    <Tag onClick={onClick} className={`text-right rounded-lg border p-4 flex flex-col justify-between min-h-[92px] w-full ${tones[tone]} ${onClick ? "hover:shadow-md active:scale-[0.98] transition cursor-pointer" : ""}`}>
      <div className="flex items-center justify-between"><span className="text-xs text-stone-500">{label}</span>{Icon ? <Icon size={16} className="opacity-60" /> : null}</div>
      <div className="mt-2"><div className="text-2xl font-bold leading-none">{value}</div>{sub ? <div className="text-xs text-stone-500 mt-1">{sub}</div> : null}</div>
    </Tag>
  );
}

function SectionTitle({ children, action, info }) {
  return (
    <div className="flex items-center justify-between mb-3 gap-2 flex-wrap">
      <h2 className="text-base font-bold text-stone-900 inline-flex items-center gap-1.5">{children}{info && <InfoTip text={info} />}</h2>
      {action}
    </div>
  );
}

function BigButton({ children, onClick, tone = "primary", type = "button", disabled }) {
  const tones = {
    primary: "bg-amber-700 hover:bg-amber-800 text-white disabled:bg-stone-300",
    ghost: "bg-white border border-stone-300 text-stone-700 hover:bg-stone-50",
    danger: "bg-rose-600 hover:bg-rose-700 text-white",
  };
  return <button type={type} onClick={onClick} disabled={disabled} className={`w-full sm:w-auto px-5 py-3 rounded-lg font-bold text-sm transition-colors disabled:cursor-not-allowed ${tones[tone]}`}>{children}</button>;
}
function IconButton({ icon: Icon, onClick, title, tone = "default" }) {
  const tones = { default: "text-stone-500 hover:text-amber-800 hover:border-amber-700", danger: "text-rose-500 hover:text-rose-700 hover:border-rose-500" };
  return <button type="button" onClick={onClick} title={title} className={`p-2 rounded-lg border border-stone-300 bg-white ${tones[tone]}`}><Icon size={15} /></button>;
}
function Field({ label, children, hint }) {
  return <label className="block mb-3"><span className="block text-xs font-bold text-stone-600 mb-1">{label}</span>{children}{hint ? <span className="block text-[11px] text-stone-400 mt-1">{hint}</span> : null}</label>;
}
const inputCls = "w-full rounded-md border border-stone-300 px-3 py-2 text-sm text-stone-900 focus:outline-none focus:ring-2 focus:ring-amber-600/40 focus:border-amber-600 bg-white";

function ItemPicker({ items, value, onChange, placeholder }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const selected = items.find((i) => i.id === value);
  const q = query.trim();
  const filtered = q ? items.filter((it) => `${it.code} ${it.design} ${it.size} ${it.factory} ${it.grade}`.includes(q)) : items;
  return (
    <div className="relative">
      <input className={inputCls} placeholder={placeholder || "تایپ کنید: کد، طرح یا سایز…"} value={open ? query : selected ? `${selected.code} — ${selected.design} (${selected.size})` : ""} onFocus={() => { setOpen(true); setQuery(""); }} onChange={(e) => setQuery(e.target.value)} onBlur={() => setTimeout(() => setOpen(false), 150)} />
      {open && (
        <div className="absolute z-20 mt-1 w-full max-h-56 overflow-y-auto bg-white border border-stone-300 rounded-lg shadow-lg">
          {filtered.length === 0 ? <div className="px-3 py-2 text-xs text-stone-400">موردی یافت نشد.</div> : filtered.map((it) => (
            <button key={it.id} type="button" onMouseDown={() => { onChange(it.id); setOpen(false); setQuery(""); }} className="w-full text-right px-3 py-2 text-sm hover:bg-amber-50 border-b border-stone-100 last:border-0">
              <div className="font-bold">{it.code} — {it.design}</div><div className="text-xs text-stone-500">{it.size} · {it.grade}</div>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center bg-black/40 p-0 sm:p-4">
      <div className="bg-white w-full sm:max-w-lg rounded-t-2xl sm:rounded-2xl max-h-[92vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-stone-200 px-4 py-3 flex items-center justify-between">
          <h3 className="font-bold text-stone-900">{title}</h3>
          <button onClick={onClose} className="p-1 text-stone-500 hover:text-stone-800"><X size={20} /></button>
        </div>
        <div className="p-4">{children}</div>
      </div>
    </div>
  );
}
function Empty({ text }) { return <div className="text-center py-10 text-sm text-stone-400 border border-dashed border-stone-300 rounded-lg">{text}</div>; }
function DiscrepancyBadge() { return <span className="inline-flex items-center gap-1 text-[11px] font-bold text-rose-700 bg-rose-50 border border-rose-200 rounded px-1.5 py-0.5"><AlertTriangle size={11} /> مغایرت</span>; }
function TypeBadge({ type }) {
  const map = {
    transfer: ["انتقال", "bg-sky-50 text-sky-700 border-sky-200"],
    stockin: ["ورود کالا", "bg-emerald-50 text-emerald-700 border-emerald-200"],
    adjustment: ["اصلاح موجودی", "bg-rose-50 text-rose-700 border-rose-200"],
    delivery: ["تحویل مشتری", "bg-amber-50 text-amber-800 border-amber-200"],
  };
  const [label, cls] = map[type] || [type, "bg-stone-50 text-stone-600 border-stone-200"];
  return <span className={`text-[11px] font-bold rounded px-1.5 py-0.5 border ${cls}`}>{label}</span>;
}

function PresetOrCustomField({ label, options, value, onChange, hint }) {
  const isCustomActive = value !== "" && !options.includes(value);
  const [mode, setMode] = useState(isCustomActive ? CUSTOM : value || "");
  const selectVal = mode === CUSTOM ? CUSTOM : value;
  return (
    <Field label={label} hint={hint}>
      <select className={inputCls} value={selectVal} onChange={(e) => { const v = e.target.value; setMode(v); if (v !== CUSTOM) onChange(v); else onChange(""); }}>
        <option value="">انتخاب کنید…</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
        <option value={CUSTOM}>سایر (تایپ دستی)</option>
      </select>
      {mode === CUSTOM && <input className={inputCls + " mt-2"} placeholder="مقدار دلخواه را وارد کنید" value={value} onChange={(e) => onChange(e.target.value)} />}
    </Field>
  );
}

/* ------------------------------------------------------------------ */
/* کادر امضای لمسی                                                    */
/* ------------------------------------------------------------------ */

function SignaturePad({ onChange }) {
  const canvasRef = useRef(null);
  const drawing = useRef(false);
  const last = useRef(null);
  useEffect(() => { const ctx = canvasRef.current.getContext("2d"); ctx.lineWidth = 2.5; ctx.lineCap = "round"; ctx.strokeStyle = "#1c1917"; }, []);
  const getPos = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const t = e.touches && e.touches[0];
    const clientX = t ? t.clientX : e.clientX, clientY = t ? t.clientY : e.clientY;
    return { x: (clientX - rect.left) * (canvas.width / rect.width), y: (clientY - rect.top) * (canvas.height / rect.height) };
  };
  const start = (e) => { e.preventDefault(); drawing.current = true; last.current = getPos(e); };
  const move = (e) => { if (!drawing.current) return; e.preventDefault(); const ctx = canvasRef.current.getContext("2d"); const p = getPos(e); ctx.beginPath(); ctx.moveTo(last.current.x, last.current.y); ctx.lineTo(p.x, p.y); ctx.stroke(); last.current = p; };
  const end = () => { if (!drawing.current) return; drawing.current = false; onChange(canvasRef.current.toDataURL("image/png")); };
  const clear = () => { const c = canvasRef.current; c.getContext("2d").clearRect(0, 0, c.width, c.height); onChange(null); };
  return (
    <div>
      <canvas ref={canvasRef} width={480} height={160} className="w-full h-40 bg-stone-50 border border-dashed border-stone-300 rounded-lg touch-none" onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end} onTouchStart={start} onTouchMove={move} onTouchEnd={end} />
      <div className="flex items-center justify-between mt-1"><span className="text-[11px] text-stone-400">با انگشت یا ماوس امضا کنید</span><button type="button" onClick={clear} className="text-[11px] text-rose-600 font-bold">پاک کردن</button></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* برگه چاپ                                                            */
/* ------------------------------------------------------------------ */

function PrintSheet({ payload }) {
  if (!payload) return null;
  const today = new Intl.DateTimeFormat("fa-IR-u-ca-persian", { dateStyle: "long", timeStyle: "short" }).format(new Date());
  return (
    <div className="print-sheet p-6" dir="rtl" style={{ fontFamily: "Tahoma, Arial, sans-serif" }}>
      <div className="flex items-center justify-between border-b-2 border-stone-800 pb-3 mb-4">
        <div>
          <div className="text-lg font-extrabold">فروشگاه کاشی و سرامیک</div>
          <div className="text-xs text-stone-500">
            {payload.type === "inventory" && "گزارش موجودی انبار"}
            {payload.type === "item" && "گزارش موجودی کالا"}
            {payload.type === "delivery" && "رسید تحویل کالا"}
            {payload.type === "detail" && (payload.title || "گزارش")}
          </div>
        </div>
        <div className="text-xs text-stone-500">{today}</div>
      </div>

      {payload.type === "inventory" && (
        <table className="w-full text-sm border-collapse">
          <thead><tr className="border-b-2 border-stone-800 text-right"><th className="py-1.5">کد</th><th className="py-1.5">طرح / سایز</th><th className="py-1.5">انبار ۱</th><th className="py-1.5">انبار ۲</th><th className="py-1.5">کل</th><th className="py-1.5">آزاد</th></tr></thead>
          <tbody>{payload.rows.map((r) => (
            <tr key={r.item.id} className="border-b border-stone-300"><td className="py-1.5">{r.item.code}</td><td className="py-1.5">{r.item.design} — {r.item.size}</td><td className="py-1.5">{fmt(r.w1)}</td><td className="py-1.5">{fmt(r.w2)}</td><td className="py-1.5 font-bold">{fmt(r.total)}</td><td className="py-1.5">{fmt(r.free)}</td></tr>
          ))}</tbody>
        </table>
      )}

      {payload.type === "item" && (
        <div>
          <div className="mb-3 text-sm">
            <div className="font-bold">{payload.item.code} — {payload.item.design} ({payload.item.size})</div>
            <div className="text-stone-500">کارخانه: {payload.item.factory} · درجه: {payload.item.grade}{payload.item.material ? ` · جنس: ${payload.item.material}` : ""} · واحد: {payload.item.unit}</div>
          </div>
          <table className="w-full text-sm border-collapse">
            <thead><tr className="border-b-2 border-stone-800 text-right"><th className="py-1.5">انبار</th><th className="py-1.5">محل</th><th className="py-1.5">موجودی</th><th className="py-1.5">رزروشده</th></tr></thead>
            <tbody>{payload.records.map((r) => (<tr key={r.id} className="border-b border-stone-300"><td className="py-1.5">{whLabel(r.warehouse)}</td><td className="py-1.5">{locLabel(r)}</td><td className="py-1.5">{qtyDisplay(payload.item, r.qty, r.boxCount)}</td><td className="py-1.5">{fmt(r.reserved || 0)}</td></tr>))}</tbody>
          </table>
        </div>
      )}

      {payload.type === "delivery" && (
        <div className="text-sm">
          <div className="flex justify-between mb-3">
            <div>
              <div className="font-bold">شماره تحویل: {payload.delivery.deliveryNumber}</div>
              <div className="text-stone-500">شماره فاکتور: {payload.delivery.invoiceNumber}</div>
              <div className="text-stone-500">مشتری: {payload.delivery.customerName}</div>
            </div>
            <div className="text-stone-500 text-left"><div>{fmtDate(payload.delivery.date)}</div><div>تحویل‌دهنده: {payload.delivery.user}</div></div>
          </div>
          <table className="w-full border-collapse mb-4">
            <thead><tr className="border-b-2 border-stone-800 text-right"><th className="py-1.5">کالا</th><th className="py-1.5">مشخصات</th><th className="py-1.5">انبار / محل</th><th className="py-1.5">مقدار</th></tr></thead>
            <tbody><tr className="border-b border-stone-300"><td className="py-1.5">{payload.item.code}</td><td className="py-1.5">{payload.item.design} — {payload.item.size} — {payload.item.grade}</td><td className="py-1.5">{whLabel(payload.delivery.warehouse)} / {locLabel(payload.delivery.location)}</td><td className="py-1.5 font-bold">{qtyDisplay(payload.item, payload.delivery.qty, payload.delivery.boxCount)}</td></tr></tbody>
          </table>
          <div className="grid grid-cols-2 gap-6 mt-10">
            <div><div className="text-xs text-stone-500 mb-1">امضای تحویل‌دهنده (انبار)</div><div className="border-t border-stone-400 pt-1">{payload.delivery.user}</div></div>
            <div><div className="text-xs text-stone-500 mb-1">امضای تحویل‌گیرنده (مشتری)</div>{payload.delivery.signature ? <img src={payload.delivery.signature} alt="امضا" className="h-16" /> : <div className="h-16 border-b border-stone-400" />}</div>
          </div>
        </div>
      )}

      {payload.type === "detail" && (
        <table className="w-full text-sm border-collapse">
          <tbody>{payload.rows.map((r, i) => (<tr key={i} className="border-b border-stone-300"><td className="py-1.5">{r.label}{r.sub ? ` — ${r.sub}` : ""}</td><td className="py-1.5 font-bold text-left">{r.value}</td></tr>))}</tbody>
        </table>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* برنامه اصلی                                                          */
/* ------------------------------------------------------------------ */

const TABS = [
  { id: "dashboard", label: "داشبورد", icon: LayoutDashboard },
  { id: "items", label: "کالاها", icon: Package },
  { id: "inventory", label: "موجودی", icon: Boxes },
  { id: "stockops", label: "ورود / اصلاح", icon: PackagePlus },
  { id: "transfer", label: "انتقال", icon: ArrowLeftRight },
  { id: "orders", label: "سفارش مشتری", icon: Users },
  { id: "locations", label: "محل کالا", icon: MapPin },
  { id: "history", label: "تاریخچه", icon: History },
  { id: "settings", label: "تنظیمات", icon: Settings },
];

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = در حال بررسی، null = واردنشده

  useEffect(() => {
    let subscription;
    (async () => {
      const { data } = await supabase.auth.getSession();
      setSession(data.session);
      const sub = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
      subscription = sub.data.subscription;
    })();
    return () => { if (subscription) subscription.unsubscribe(); };
  }, []);

  if (session === undefined) {
    return <div dir="rtl" className="min-h-[400px] flex items-center justify-center text-stone-400 text-sm">در حال بررسی ورود…</div>;
  }
  if (!session) return <LoginScreen />;
  return <AuthedApp session={session} />;
}

/* ------------------------------------------------------------------ */
/* صفحه ورود                                                           */
/* ------------------------------------------------------------------ */

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setErr(""); setBusy(true);
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setBusy(false);
    if (error) setErr("ایمیل یا رمز عبور اشتباه است.");
  };

  return (
    <div dir="rtl" style={{ fontFamily: "Tahoma, 'Segoe UI', Arial, sans-serif" }} className="min-h-screen bg-stone-100 flex items-center justify-center p-4">
      <form onSubmit={submit} className="bg-white border border-stone-200 rounded-2xl shadow-sm p-6 w-full max-w-sm">
        <div className="text-center mb-5">
          <div className="text-[11px] text-amber-700 font-bold tracking-wide">سیستم مدیریت انبار کاشی و سرامیک</div>
          <div className="text-xl font-extrabold text-stone-900 mt-1">ورود به سیستم</div>
        </div>
        <Field label="ایمیل کاربری"><input className={inputCls} type="email" autoComplete="username" value={email} onChange={(e) => setEmail(e.target.value)} required /></Field>
        <Field label="رمز عبور"><input className={inputCls} type="password" autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} required /></Field>
        {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3">{err}</div>}
        <BigButton type="submit" disabled={busy}>{busy ? "در حال ورود…" : "ورود"}</BigButton>
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* برنامه اصلی (پس از ورود)                                            */
/* ------------------------------------------------------------------ */

function AuthedApp({ session }) {
  const [loading, setLoading] = useState(true);
  const [items, setItems] = useState([]);
  const [stock, setStock] = useState([]);
  const [ops, setOps] = useState([]);
  const [orders, setOrders] = useState([]);
  const [deliveries, setDeliveries] = useState([]);
  const [pendingCounts, setPendingCounts] = useState([]);
  const [staff, setStaff] = useState([]);
  const [me, setMe] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState(null);
  const [printPayload, setPrintPayload] = useState(null);

  const showToast = useCallback((text, tone = "ok") => { setToast({ text, tone }); setTimeout(() => setToast(null), 3200); }, []);
  const triggerPrint = useCallback((payload) => setPrintPayload(payload), []);

  useEffect(() => { if (!printPayload) return; const t = setTimeout(() => window.print(), 200); return () => clearTimeout(t); }, [printPayload]);
  useEffect(() => { const h = () => setPrintPayload(null); window.addEventListener("afterprint", h); return () => window.removeEventListener("afterprint", h); }, []);

  const loadAll = useCallback(async () => {
    setLoading(true);
    const meRec = await ensureStaffRecord(session.user);
    setMe(meRec);
    const [i, s, o, ord, dl, pc, st] = await Promise.all([
      fetchAll("items"), fetchAll("stock"), fetchAll("operations", "date"),
      fetchOrdersWithLines(), fetchAll("deliveries", "date"), fetchAll("pending_counts", "date"),
      fetchAll("staff"),
    ]);
    setItems(i); setStock(s); setOps(o); setOrders(ord); setDeliveries(dl); setPendingCounts(pc); setStaff(st);
    setLoading(false);
  }, [session]);

  useEffect(() => { loadAll(); }, [loadAll]);

  const currentUser = me || { name: "کاربر", isManager: false };
  const isManager = !!currentUser?.isManager;

  const itemById = useCallback((id) => items.find((it) => it.id === id), [items]);
  const stockForItem = useCallback((itemId, warehouse) => stock.filter((s) => s.itemId === itemId && (!warehouse || s.warehouse === warehouse)), [stock]);
  const itemTotals = useCallback((itemId) => {
    const recs = stock.filter((s) => s.itemId === itemId);
    const w1 = recs.filter((r) => r.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0);
    const w2 = recs.filter((r) => r.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0);
    const reserved = recs.reduce((a, r) => a + Number(r.reserved || 0), 0);
    return { w1, w2, total: w1 + w2, reserved, free: w1 + w2 - reserved };
  }, [stock]);
  const nextOpNumber = (type, prefix) => `${prefix}-${String(ops.filter((o) => o.type === type).length + 1).padStart(5, "0")}`;

  const addItem = useCallback(async (form) => {
    const item = { id: uid("itm"), code: `KJ-${1001 + items.length}`, ...form };
    const { error } = await supabase.from("items").insert(camelToRow(item));
    if (error) { showToast("خطا در ثبت کالا.", "error"); console.error(error); return item; }
    setItems((prev) => [...prev, item]);
    showToast("کالای جدید ثبت شد.");
    return item;
  }, [items, showToast]);

  const editItem = useCallback(async (itemId, patch) => {
    await supabase.from("items").update(camelToRow(patch)).eq("id", itemId);
    setItems((prev) => prev.map((it) => (it.id === itemId ? { ...it, ...patch } : it)));
    showToast("مشخصات کالا اصلاح شد.");
  }, [showToast]);

  /* ---------------- انتقال بین انبار ---------------- */
  const doTransfer = useCallback(async ({ itemId, fromWh, fromRecId, toWh, toLoc, useExistingTo, toRecId, qty, boxCount, description }) => {
    const q = Number(qty);
    const fromRec = stock.find((s) => s.id === fromRecId);
    if (!fromRec) return { ok: false, msg: "محل مبدأ معتبر نیست." };
    const available = Number(fromRec.qty || 0) - Number(fromRec.reserved || 0);
    if (!q || q <= 0) return { ok: false, msg: "مقدار انتقال باید بزرگ‌تر از صفر باشد." };
    if (q > available + 1e-6) return { ok: false, msg: `مقدار انتقالی بیشتر از موجودی قابل انتقال (${fmt(available)}) است.` };

    let nextStock = stock.map((s) => s.id === fromRec.id ? { ...s, qty: Number(s.qty) - q, boxCount: boxCount != null ? Number(s.boxCount || 0) - Number(boxCount) : s.boxCount } : s);
    let destRecId, destIsNew = false;
    if (useExistingTo && toRecId) {
      nextStock = nextStock.map((s) => s.id === toRecId ? { ...s, qty: Number(s.qty) + q, boxCount: boxCount != null ? Number(s.boxCount || 0) + Number(boxCount) : s.boxCount } : s);
      destRecId = toRecId;
    } else {
      const newRec = { id: uid("st"), itemId, warehouse: toWh, aisle: toLoc.aisle || "-", row: toLoc.row || "-", cell: toLoc.cell || "-", qty: q, boxCount: boxCount != null ? Number(boxCount) : null, reserved: 0, discrepancy: false };
      nextStock = [...nextStock, newRec];
      destRecId = newRec.id;
      destIsNew = true;
    }
    nextStock = nextStock.filter((s) => !(s.qty <= 0.001 && (s.reserved || 0) <= 0));
    const destRecAfter = nextStock.find((s) => s.id === destRecId) || { aisle: toLoc.aisle, row: toLoc.row, cell: toLoc.cell };

    const op = { id: uid("op"), type: "transfer", opNumber: nextOpNumber("transfer", "TR"), itemId, qty: q, boxCount: boxCount != null ? Number(boxCount) : null, fromWarehouse: fromWh, toWarehouse: toWh, fromLocation: { aisle: fromRec.aisle, row: fromRec.row, cell: fromRec.cell }, toLocation: { aisle: destRecAfter.aisle, row: destRecAfter.row, cell: destRecAfter.cell }, description: description || "", user: currentUser.name, date: new Date().toISOString() };

    await syncStockRecord(nextStock, fromRec.id, false);
    await syncStockRecord(nextStock, destRecId, destIsNew);
    await supabase.from("operations").insert(camelToRow(op));

    setStock(nextStock);
    setOps((prev) => [op, ...prev]);
    return { ok: true, op };
  }, [stock, ops, currentUser]);

  /* ---------------- ورود کالا به انبار ---------------- */
  const doStockIn = useCallback(async ({ itemId, warehouse, useExisting, recId, toLoc, qty, boxCount, supplier, description }) => {
    const q = Number(qty);
    if (!q || q <= 0) return { ok: false, msg: "مقدار باید بزرگ‌تر از صفر باشد." };
    let nextStock = [...stock];
    let targetId, isNew = false;
    if (useExisting && recId) {
      nextStock = nextStock.map((s) => s.id === recId ? { ...s, qty: Number(s.qty) + q, boxCount: boxCount != null ? Number(s.boxCount || 0) + Number(boxCount) : s.boxCount } : s);
      targetId = recId;
    } else {
      const rec = { id: uid("st"), itemId, warehouse, aisle: toLoc.aisle || "-", row: toLoc.row || "-", cell: toLoc.cell || "-", qty: q, boxCount: boxCount != null ? Number(boxCount) : null, reserved: 0, discrepancy: false };
      nextStock = [...nextStock, rec];
      targetId = rec.id;
      isNew = true;
    }
    const target = nextStock.find((s) => s.id === targetId);
    const op = { id: uid("op"), type: "stockin", opNumber: nextOpNumber("stockin", "IN"), itemId, qty: q, boxCount: boxCount != null ? Number(boxCount) : null, warehouse, location: { aisle: target.aisle, row: target.row, cell: target.cell }, supplier: supplier || "", description: description || "", user: currentUser.name, date: new Date().toISOString() };
    await syncStockRecord(nextStock, targetId, isNew);
    await supabase.from("operations").insert(camelToRow(op));
    setStock(nextStock);
    setOps((prev) => [op, ...prev]);
    return { ok: true, op };
  }, [stock, ops, currentUser]);

  /* ---------------- ویرایش امن ورود کالا (فقط مدیر) ---------------- */
  const editStockIn = useCallback(async (opId, patch) => {
    const op = ops.find((o) => o.id === opId && o.type === "stockin");
    if (!op) return { ok: false, msg: "عملیات یافت نشد." };
    const rec = stock.find((s) => s.itemId === op.itemId && s.warehouse === op.warehouse && s.aisle === op.location.aisle && s.row === op.location.row && s.cell === op.location.cell);
    const newQty = Number(patch.qty);
    const deltaQty = newQty - Number(op.qty);
    const deltaBox = patch.boxCount != null && op.boxCount != null ? Number(patch.boxCount) - Number(op.boxCount) : null;
    let nextStock = stock;
    if (rec) {
      if (Number(rec.qty) + deltaQty < -1e-6) return { ok: false, msg: "این اصلاح باعث منفی‌شدن موجودی می‌شود." };
      nextStock = stock.map((s) => s.id === rec.id ? { ...s, qty: Number(s.qty) + deltaQty, boxCount: deltaBox != null ? Number(s.boxCount || 0) + deltaBox : s.boxCount } : s);
    }
    const editedNote = `اصلاح توسط ${currentUser.name}: مقدار از ${fmt(op.qty)} به ${fmt(newQty)} تغییر یافت.`;
    const newBoxCount = patch.boxCount != null ? Number(patch.boxCount) : op.boxCount;
    if (rec) await syncStockRecord(nextStock, rec.id, false);
    await supabase.from("operations").update(camelToRow({ qty: newQty, boxCount: newBoxCount, editedNote })).eq("id", opId);
    setStock(nextStock);
    setOps((prev) => prev.map((o) => o.id === opId ? { ...o, qty: newQty, boxCount: newBoxCount, editedNote } : o));
    showToast("ورود کالا اصلاح شد.");
    return { ok: true };
  }, [ops, stock, currentUser, showToast]);

  /* ---------------- کسری / شکستگی ---------------- */
  const doAdjustment = useCallback(async ({ itemId, recId, qty, boxCount, reason, description }) => {
    const q = Number(qty);
    const rec = stock.find((s) => s.id === recId);
    if (!rec) return { ok: false, msg: "محل معتبر نیست." };
    if (!q || q <= 0) return { ok: false, msg: "مقدار باید بزرگ‌تر از صفر باشد." };
    if (q > Number(rec.qty) + 1e-6) return { ok: false, msg: `مقدار کسری بیشتر از موجودی ثبت‌شده (${fmt(rec.qty)}) است.` };
    let nextStock = stock.map((s) => s.id === recId ? { ...s, qty: Number(s.qty) - q, boxCount: boxCount != null ? Number(s.boxCount || 0) - Number(boxCount) : s.boxCount } : s);
    nextStock = nextStock.filter((s) => !(s.qty <= 0.001 && (s.reserved || 0) <= 0));
    const op = { id: uid("op"), type: "adjustment", opNumber: nextOpNumber("adjustment", "ADJ"), itemId, qty: q, boxCount: boxCount != null ? Number(boxCount) : null, warehouse: rec.warehouse, location: { aisle: rec.aisle, row: rec.row, cell: rec.cell }, reason, description: description || "", user: currentUser.name, date: new Date().toISOString() };
    await syncStockRecord(nextStock, recId, false);
    await supabase.from("operations").insert(camelToRow(op));
    setStock(nextStock);
    setOps((prev) => [op, ...prev]);
    return { ok: true, op };
  }, [stock, ops, currentUser]);

  /* ---------------- شمارش فیزیکی ---------------- */
  const doPhysicalCount = useCallback(async ({ itemId, recId, countedQty, countedBoxCount }) => {
    const rec = stock.find((s) => s.id === recId);
    if (!rec) return { ok: false, msg: "محل معتبر نیست." };
    const cq = Number(countedQty);
    const diff = cq - Number(rec.qty);
    if (Math.abs(diff) < 0.01) return { ok: true, matched: true };
    const nextStock = stock.map((s) => (s.id === recId ? { ...s, discrepancy: true } : s));
    const pending = { id: uid("pc"), stockRecId: recId, itemId, warehouse: rec.warehouse, location: { aisle: rec.aisle, row: rec.row, cell: rec.cell }, recordedQty: Number(rec.qty), recordedBoxCount: rec.boxCount, countedQty: cq, countedBoxCount: countedBoxCount != null ? Number(countedBoxCount) : null, diff, user: currentUser.name, date: new Date().toISOString() };
    await supabase.from("stock").update({ discrepancy: true }).eq("id", recId);
    await supabase.from("pending_counts").insert(camelToRow(pending));
    setStock(nextStock);
    setPendingCounts((prev) => [pending, ...prev]);
    return { ok: true, matched: false, pending };
  }, [stock, currentUser]);

  const resolvePendingCount = useCallback(async (pendingId, apply) => {
    const p = pendingCounts.find((x) => x.id === pendingId);
    if (!p) return;
    const rec = stock.find((s) => s.id === p.stockRecId);
    let nextStock = stock;
    if (apply) {
      nextStock = stock.map((s) => (s.id === p.stockRecId ? { ...s, qty: p.countedQty, boxCount: p.countedBoxCount != null ? p.countedBoxCount : s.boxCount, discrepancy: false } : s));
      const op = { id: uid("op"), type: "adjustment", opNumber: nextOpNumber("adjustment", "ADJ"), itemId: p.itemId, qty: Math.abs(p.diff), boxCount: p.countedBoxCount != null && p.recordedBoxCount != null ? Math.abs(p.countedBoxCount - p.recordedBoxCount) : null, warehouse: rec?.warehouse, location: rec ? { aisle: rec.aisle, row: rec.row, cell: rec.cell } : {}, reason: "اصلاح شمارش فیزیکی", description: `از ${fmt(p.recordedQty)} به ${fmt(p.countedQty)} اصلاح شد (تأیید ${currentUser.name})`, user: currentUser.name, date: new Date().toISOString() };
      if (rec) await syncStockRecord(nextStock, p.stockRecId, false);
      await supabase.from("operations").insert(camelToRow(op));
      setOps((prev) => [op, ...prev]);
    } else {
      nextStock = stock.map((s) => (s.id === p.stockRecId ? { ...s, discrepancy: false } : s));
      await supabase.from("stock").update({ discrepancy: false }).eq("id", p.stockRecId);
    }
    await supabase.from("pending_counts").delete().eq("id", pendingId);
    setStock(nextStock);
    setPendingCounts((prev) => prev.filter((x) => x.id !== pendingId));
  }, [pendingCounts, stock, ops, currentUser]);

  const toggleDiscrepancy = useCallback(async (recId) => {
    const rec = stock.find((s) => s.id === recId);
    if (!rec) return;
    const next = !rec.discrepancy;
    await supabase.from("stock").update({ discrepancy: next }).eq("id", recId);
    setStock((prev) => prev.map((s) => (s.id === recId ? { ...s, discrepancy: next } : s)));
  }, [stock]);

  /* ---------------- سفارش مشتری ---------------- */
  const createOrder = useCallback(async ({ invoiceNumber, customerName, customerPhone, lines }) => {
    if (!invoiceNumber || !customerName || lines.length === 0) return { ok: false, msg: "شماره فاکتور، نام مشتری و حداقل یک قلم کالا الزامی است." };
    const orderId = uid("ord");
    const date = new Date().toISOString();
    const order = { id: orderId, invoiceNumber, customerName, customerPhone: customerPhone || "", date, lines: lines.map((l) => ({ id: uid("ln"), orderId, itemId: l.itemId, qty: Number(l.qty), boxCount: l.boxCount != null ? Number(l.boxCount) : null, deliveredQty: 0, deliveredBoxCount: l.boxCount != null ? 0 : null })) };
    await supabase.from("orders").insert(camelToRow({ id: order.id, invoiceNumber, customerName, customerPhone: order.customerPhone, date }));
    await supabase.from("order_lines").insert(order.lines.map(camelToRow));
    setOrders((prev) => [order, ...prev]);
    return { ok: true, order };
  }, []);

  const doDelivery = useCallback(async ({ orderId, lineId, recId, qty, boxCount, signature }) => {
    const order = orders.find((o) => o.id === orderId);
    const line = order?.lines.find((l) => l.id === lineId);
    if (!order || !line) return { ok: false, msg: "سفارش یا قلم کالا معتبر نیست." };
    const remaining = Number(line.qty) - Number(line.deliveredQty);
    const q = Number(qty);
    if (!q || q <= 0) return { ok: false, msg: "مقدار تحویل باید بزرگ‌تر از صفر باشد." };
    if (q > remaining + 1e-6) return { ok: false, msg: `مقدار بیشتر از باقیمانده سفارش (${fmt(remaining)}) است.` };
    const rec = stock.find((s) => s.id === recId);
    if (!rec) return { ok: false, msg: "محل انبار معتبر نیست." };
    const available = Number(rec.qty) - Number(rec.reserved || 0);
    if (q > available + 1e-6) return { ok: false, msg: `موجودی قابل تحویل در این محل (${fmt(available)}) کافی نیست.` };

    let nextStock = stock.map((s) => s.id === recId ? { ...s, qty: Number(s.qty) - q, boxCount: boxCount != null ? Number(s.boxCount || 0) - Number(boxCount) : s.boxCount } : s);
    nextStock = nextStock.filter((s) => !(s.qty <= 0.001 && (s.reserved || 0) <= 0));
    const newDeliveredQty = Number(line.deliveredQty) + q;
    const newDeliveredBoxCount = boxCount != null ? Number(line.deliveredBoxCount || 0) + Number(boxCount) : line.deliveredBoxCount;
    const nextOrders = orders.map((o) => o.id !== orderId ? o : { ...o, lines: o.lines.map((l) => l.id !== lineId ? l : { ...l, deliveredQty: newDeliveredQty, deliveredBoxCount: newDeliveredBoxCount }) });
    const delivery = { id: uid("dl"), deliveryNumber: `DL-${String(deliveries.length + 1).padStart(5, "0")}`, orderId, lineId, itemId: line.itemId, warehouse: rec.warehouse, location: { aisle: rec.aisle, row: rec.row, cell: rec.cell }, qty: q, boxCount: boxCount != null ? Number(boxCount) : null, customerName: order.customerName, customerPhone: order.customerPhone, invoiceNumber: order.invoiceNumber, user: currentUser.name, signature: signature || null, date: new Date().toISOString() };

    await syncStockRecord(nextStock, recId, false);
    await supabase.from("order_lines").update(camelToRow({ deliveredQty: newDeliveredQty, deliveredBoxCount: newDeliveredBoxCount })).eq("id", lineId);
    await supabase.from("deliveries").insert(camelToRow(delivery));
    setStock(nextStock);
    setOrders(nextOrders);
    setDeliveries((prev) => [delivery, ...prev]);
    return { ok: true, delivery };
  }, [orders, stock, deliveries, currentUser]);

  /* ---------------- ویرایش امن تحویل مشتری (فقط مدیر) ---------------- */
  const editDelivery = useCallback(async (deliveryId, patch) => {
    const d = deliveries.find((x) => x.id === deliveryId);
    if (!d) return { ok: false, msg: "رسید یافت نشد." };
    let nextStock = stock, nextOrders = orders;
    if (patch.qty != null && Number(patch.qty) !== Number(d.qty)) {
      const deltaQty = Number(patch.qty) - Number(d.qty);
      const rec = stock.find((s) => s.itemId === d.itemId && s.warehouse === d.warehouse && s.aisle === d.location.aisle && s.row === d.location.row && s.cell === d.location.cell);
      if (rec) {
        if (Number(rec.qty) - deltaQty < -1e-6) return { ok: false, msg: "موجودی کافی برای افزایش این مقدار نیست." };
        nextStock = stock.map((s) => (s.id === rec.id ? { ...s, qty: Number(s.qty) - deltaQty } : s));
        await supabase.from("stock").update({ qty: nextStock.find((s) => s.id === rec.id).qty }).eq("id", rec.id);
      }
      nextOrders = orders.map((o) => o.id !== d.orderId ? o : { ...o, lines: o.lines.map((l) => l.id !== d.lineId ? l : { ...l, deliveredQty: Number(l.deliveredQty) + deltaQty }) });
      const newLine = nextOrders.find((o) => o.id === d.orderId)?.lines.find((l) => l.id === d.lineId);
      if (newLine) await supabase.from("order_lines").update({ delivered_qty: newLine.deliveredQty }).eq("id", d.lineId);
    }
    const editedNote = `ویرایش‌شده توسط ${currentUser.name}`;
    const patchData = { customerName: patch.customerName ?? d.customerName, invoiceNumber: patch.invoiceNumber ?? d.invoiceNumber, qty: patch.qty != null ? Number(patch.qty) : d.qty, editedNote };
    await supabase.from("deliveries").update(camelToRow(patchData)).eq("id", deliveryId);
    const nextDeliveries = deliveries.map((x) => x.id === deliveryId ? { ...x, ...patchData } : x);
    setStock(nextStock);
    setOrders(nextOrders);
    setDeliveries(nextDeliveries);
    showToast("رسید تحویل اصلاح شد.");
    return { ok: true };
  }, [deliveries, stock, orders, currentUser, showToast]);

  /* ---------------- مدیریت کارکنان ---------------- */
  const renameStaff = useCallback(async (id, name) => {
    await supabase.from("staff").update({ name }).eq("id", id);
    setStaff((prev) => prev.map((s) => s.id === id ? { ...s, name } : s));
  }, []);
  const toggleManagerFlag = useCallback(async (id, value) => {
    await supabase.from("staff").update({ is_manager: value }).eq("id", id);
    setStaff((prev) => prev.map((s) => s.id === id ? { ...s, isManager: value } : s));
  }, []);
  const removeStaffRow = useCallback(async (id) => {
    await supabase.from("staff").delete().eq("id", id);
    setStaff((prev) => prev.filter((s) => s.id !== id));
  }, []);

  /* ---------------- خروجی اکسل ---------------- */
  const exportExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(items.map((it) => ({ کد: it.code, طرح: it.design, سایز: it.size, کارخانه: it.factory, درجه: it.grade, سطح: it.surface, جنس: it.material || "", واحد: it.unit, "متراژ هر کارتن": it.sqmPerBox || "", "آستانه هشدار": it.lowThreshold || "", توضیحات: it.description || "" }))), "کالاها");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stock.map((s) => { const it = itemById(s.itemId); return { کد: it?.code, طرح: it?.design, انبار: whLabel(s.warehouse), راهرو: s.aisle, ردیف: s.row, خانه: s.cell, موجودی: s.qty, کارتن: s.boxCount ?? "", رزروشده: s.reserved || 0, مغایرت: s.discrepancy ? "دارد" : "" }; })), "موجودی");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ops.map((o) => { const it = itemById(o.itemId); return { شماره: o.opNumber, نوع: o.type, کد: it?.code, طرح: it?.design, مقدار: o.qty, انبار: whLabel(o.warehouse || o.fromWarehouse), "به انبار": o.toWarehouse ? whLabel(o.toWarehouse) : "", دلیل: o.reason || "", کاربر: o.user, تاریخ: fmtDate(o.date), توضیحات: o.description || "" }; })), "عملیات");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(orders.flatMap((o) => o.lines.map((l) => { const it = itemById(l.itemId); return { فاکتور: o.invoiceNumber, مشتری: o.customerName, تلفن: o.customerPhone || "", کد: it?.code, طرح: it?.design, سفارش: l.qty, "تحویل‌شده": l.deliveredQty, باقیمانده: Number(l.qty) - Number(l.deliveredQty) }; }))), "سفارش‌ها");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deliveries.map((d) => { const it = itemById(d.itemId); return { شماره: d.deliveryNumber, فاکتور: d.invoiceNumber, مشتری: d.customerName, کد: it?.code, طرح: it?.design, مقدار: d.qty, انبار: whLabel(d.warehouse), کاربر: d.user, تاریخ: fmtDate(d.date) }; })), "تحویل‌ها");
    XLSX.writeFile(wb, `backup-warehouse-${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast("فایل اکسل دانلود شد.");
  }, [items, stock, ops, orders, deliveries, itemById, showToast]);

  if (loading) return <div dir="rtl" className="min-h-[400px] flex items-center justify-center text-stone-400 text-sm">در حال بارگذاری اطلاعات انبار…</div>;

  return (
    <div dir="rtl" style={{ fontFamily: "Tahoma, 'Segoe UI', Arial, sans-serif" }} className="app-root bg-stone-100 min-h-screen text-stone-900">
      <style>{`
        @media print { .app-root > *:not(.print-sheet) { display: none !important; } .print-sheet { display: block !important; } }
        .print-sheet { display: none; }
      `}</style>

      <div className="no-print bg-slate-900 text-white px-4 py-3 sticky top-0 z-20">
        <div className="max-w-5xl mx-auto flex items-center justify-between gap-2">
          <div>
            <div className="text-[11px] text-amber-400 font-bold tracking-wide">سیستم مدیریت انبار کاشی و سرامیک</div>
            <div className="text-lg font-extrabold">کنترل مرکزی انبار</div>
          </div>
          <div className="flex items-center gap-2">
            <IconButton icon={RefreshCcw} title="تازه‌سازی اطلاعات" onClick={loadAll} />
            <div className="text-xs text-slate-200 bg-slate-800 border border-slate-600 rounded-lg px-2.5 py-1.5">{currentUser.name}{isManager ? " (مدیر)" : ""}</div>
            <button onClick={() => supabase.auth.signOut()} className="text-xs font-bold bg-slate-800 border border-slate-600 text-white rounded-lg px-2.5 py-1.5 hover:bg-slate-700">خروج</button>
          </div>
        </div>
      </div>

      <div className="no-print bg-white border-b border-stone-200 sticky top-[60px] z-20 overflow-x-auto">
        <div className="max-w-5xl mx-auto flex px-2">
          {TABS.map((t) => {
            const Icon = t.icon, active = tab === t.id;
            return <button key={t.id} onClick={() => setTab(t.id)} className={`flex flex-col sm:flex-row items-center gap-1 sm:gap-1.5 px-3 sm:px-4 py-2.5 text-xs sm:text-sm font-bold whitespace-nowrap border-b-2 transition-colors ${active ? "border-amber-700 text-amber-800" : "border-transparent text-stone-500 hover:text-stone-800"}`}><Icon size={16} />{t.label}</button>;
          })}
        </div>
      </div>

      <div className="no-print max-w-5xl mx-auto p-4">
        {tab === "dashboard" && <Dashboard items={items} stock={stock} ops={ops} orders={orders} itemById={itemById} onPrint={(title, rows) => triggerPrint({ type: "detail", title, rows })} />}
        {tab === "items" && (
          <ItemsPage items={items} stock={stock} ops={ops} deliveries={deliveries} itemTotals={itemTotals} onAdd={addItem} onEditItem={editItem} isManager={isManager}
            onPrintItem={(item) => triggerPrint({ type: "item", item, records: stockForItem(item.id) })}
            onPrintLedger={(title, rows) => triggerPrint({ type: "detail", title, rows })} />
        )}
        {tab === "inventory" && <InventoryPage items={items} itemTotals={itemTotals} onPrint={(rows) => triggerPrint({ type: "inventory", rows })} showToast={showToast} />}
        {tab === "stockops" && (
          <StockOpsPage items={items} stock={stock} stockForItem={stockForItem} pendingCounts={pendingCounts} itemById={itemById} onAddItem={addItem}
            doStockIn={doStockIn} doAdjustment={doAdjustment} doPhysicalCount={doPhysicalCount} resolvePendingCount={resolvePendingCount} showToast={showToast} />
        )}
        {tab === "transfer" && <TransferPage items={items} stock={stock} stockForItem={stockForItem} doTransfer={doTransfer} showToast={showToast} />}
        {tab === "orders" && (
          <OrdersPage items={items} stock={stock} orders={orders} itemById={itemById} stockForItem={stockForItem}
            createOrder={createOrder} doDelivery={doDelivery} showToast={showToast} onPrintDelivery={(delivery, item) => triggerPrint({ type: "delivery", delivery, item })} />
        )}
        {tab === "locations" && <LocationsPage items={items} stock={stock} itemById={itemById} onToggleDiscrepancy={toggleDiscrepancy} />}
        {tab === "history" && <HistoryPage ops={ops} deliveries={deliveries} itemById={itemById} items={items} isManager={isManager} editStockIn={editStockIn} editDelivery={editDelivery} showToast={showToast} />}
        {tab === "settings" && <SettingsPage staff={staff} onRename={renameStaff} onToggleManager={toggleManagerFlag} onRemoveStaff={removeStaffRow} exportExcel={exportExcel} isManager={isManager} />}
      </div>

      <PrintSheet payload={printPayload} />

      {toast && (
        <div className="no-print fixed bottom-4 inset-x-0 flex justify-center z-50 px-4">
          <div className={`px-4 py-2.5 rounded-lg text-sm font-bold shadow-lg text-white ${toast.tone === "error" ? "bg-rose-600" : "bg-stone-900"}`}>{toast.text}</div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* داشبورد                                                             */
/* ------------------------------------------------------------------ */

function DashboardDetailModal({ title, rows, onClose, onPrint }) {
  return (
    <Modal title={title} onClose={onClose}>
      <div className="flex justify-end mb-2"><IconButton icon={Printer} title="چاپ این گزارش" onClick={() => onPrint(title, rows)} /></div>
      {rows.length === 0 ? <Empty text="موردی یافت نشد." /> : (
        <div className="divide-y divide-stone-100">
          {rows.map((r, i) => (
            <div key={i} className="flex items-center justify-between py-2 text-sm gap-3">
              <div className="min-w-0"><div className="font-bold truncate">{r.label}</div>{r.sub && <div className="text-xs text-stone-500">{r.sub}</div>}</div>
              <div className="font-bold text-amber-800 whitespace-nowrap">{r.value}</div>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function Dashboard({ items, stock, ops, orders, itemById, onPrint }) {
  const [filter, setFilter] = useState("all");
  const [detail, setDetail] = useState(null);
  const scoped = useMemo(() => (filter === "all" ? stock : stock.filter((s) => s.warehouse === filter)), [stock, filter]);
  const total = scoped.reduce((a, r) => a + Number(r.qty || 0), 0);
  const w1Total = stock.filter((r) => r.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0);
  const w2Total = stock.filter((r) => r.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0);
  const reserved = scoped.reduce((a, r) => a + Number(r.reserved || 0), 0);
  const free = total - reserved;
  const discrepancyRecs = scoped.filter((r) => r.discrepancy);
  const discrepancyItems = new Set(discrepancyRecs.map((r) => r.itemId)).size;
  const noLocationList = items.filter((it) => filter === "all" ? !stock.some((s) => s.itemId === it.id) : !stock.some((s) => s.itemId === it.id && s.warehouse === filter));
  const openLines = orders.flatMap((o) => o.lines.filter((l) => Number(l.qty) - Number(l.deliveredQty) > 0.001).map((l) => ({ ...l, customerName: o.customerName, invoiceNumber: o.invoiceNumber })));
  const recentOps = ops.slice(0, 5);

  const itemScoped = (itemId) => {
    const recs = scoped.filter((s) => s.itemId === itemId);
    const w1 = recs.filter((r) => r.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0);
    const w2 = recs.filter((r) => r.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0);
    const res = recs.reduce((a, r) => a + Number(r.reserved || 0), 0);
    return { w1, w2, total: w1 + w2, reserved: res, free: w1 + w2 - res };
  };
  const globalTotal = (itemId) => { const recs = stock.filter((s) => s.itemId === itemId); return recs.reduce((a, r) => a + Number(r.qty || 0), 0); };
  const lowStockItems = items.filter((it) => it.lowThreshold && Number(it.lowThreshold) > 0 && globalTotal(it.id) <= Number(it.lowThreshold));

  const openDetail = (key) => {
    if (key === "total") setDetail({ title: "ریز موجودی کل", rows: items.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: fmt(itemScoped(it.id).total) })) });
    else if (key === "w1") setDetail({ title: "ریز موجودی انبار ۱", rows: items.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: fmt(stock.filter((s) => s.itemId === it.id && s.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0)) })) });
    else if (key === "w2") setDetail({ title: "ریز موجودی انبار ۲", rows: items.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: fmt(stock.filter((s) => s.itemId === it.id && s.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0)) })) });
    else if (key === "free") setDetail({ title: "ریز موجودی آزاد", rows: items.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: fmt(itemScoped(it.id).free) })) });
    else if (key === "reserved") setDetail({ title: "ریز موجودی رزروشده", rows: items.filter((it) => itemScoped(it.id).reserved > 0).map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: fmt(itemScoped(it.id).reserved) })) });
    else if (key === "discrepancy") setDetail({ title: "محل‌های دارای مغایرت", rows: discrepancyRecs.map((r) => { const it = itemById(r.itemId); return { label: `${it?.code} — ${it?.design}`, sub: `${whLabel(r.warehouse)} / ${locLabel(r)}`, value: qtyDisplay(it, r.qty, r.boxCount) }; }) });
    else if (key === "noLocation") setDetail({ title: "کالاهای بدون محل", rows: noLocationList.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: "—" })) });
    else if (key === "openLines") setDetail({ title: "اقلام باز مشتریان", rows: openLines.map((l) => { const it = itemById(l.itemId); return { label: `${l.customerName} — ${it?.code}`, sub: `فاکتور ${l.invoiceNumber} · ${it?.design}`, value: qtyDisplay(it, Number(l.qty) - Number(l.deliveredQty), null) }; }) });
    else if (key === "low") setDetail({ title: "کالاهای رو به اتمام", rows: lowStockItems.map((it) => ({ label: `${it.code} — ${it.design}`, sub: it.size, value: `${fmt(globalTotal(it.id))} / آستانه ${fmt(it.lowThreshold)}` })) });
  };

  return (
    <div>
      <SectionTitle info="خلاصه کلی وضعیت انبارها؛ اعداد بر اساس فیلتر انبار انتخابی محاسبه می‌شن. روی هر کارت بزنید تا ریز اطلاعاتش باز شود."
        action={<div className="flex bg-white border border-stone-300 rounded-lg overflow-hidden text-xs">
          {[{ id: "all", label: "همه انبارها" }, { id: "w1", label: "انبار ۱" }, { id: "w2", label: "انبار ۲" }].map((f) => (
            <button key={f.id} onClick={() => setFilter(f.id)} className={`px-3 py-1.5 font-bold ${filter === f.id ? "bg-amber-700 text-white" : "text-stone-600 hover:bg-stone-50"}`}>{f.label}</button>
          ))}
        </div>}>
        خلاصه وضعیت
      </SectionTitle>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
        <Card label="موجودی کل" value={fmt(total)} icon={Boxes} tone="slate" onClick={() => openDetail("total")} />
        <Card label="موجودی انبار ۱" value={fmt(w1Total)} icon={Package} onClick={() => openDetail("w1")} />
        <Card label="موجودی انبار ۲" value={fmt(w2Total)} icon={Package} onClick={() => openDetail("w2")} />
        <Card label="موجودی آزاد" value={fmt(free)} icon={Check} tone="accent" onClick={() => openDetail("free")} />
        <Card label="موجودی رزروشده" value={fmt(reserved)} icon={ChevronLeft} onClick={() => openDetail("reserved")} />
        <Card label="کالاهای دارای مغایرت" value={fmt(discrepancyItems)} icon={AlertTriangle} tone={discrepancyItems > 0 ? "warn" : "default"} onClick={() => openDetail("discrepancy")} />
        <Card label="کالاهای بدون محل" value={fmt(noLocationList.length)} icon={MapPin} tone={noLocationList.length > 0 ? "warn" : "default"} onClick={() => openDetail("noLocation")} />
        <Card label="اقلام باز مشتریان" value={fmt(openLines.length)} icon={Users} tone={openLines.length > 0 ? "warn" : "default"} onClick={() => openDetail("openLines")} />
        <Card label="کالاهای رو به اتمام" value={fmt(lowStockItems.length)} icon={AlertTriangle} tone={lowStockItems.length > 0 ? "warn" : "default"} onClick={() => openDetail("low")} />
      </div>

      {detail && <DashboardDetailModal title={detail.title} rows={detail.rows} onClose={() => setDetail(null)} onPrint={onPrint} />}

      <SectionTitle>آخرین عملیات</SectionTitle>
      {recentOps.length === 0 ? <Empty text="هنوز هیچ عملیاتی ثبت نشده است." /> : (
        <div className="bg-white border border-stone-200 rounded-lg divide-y divide-stone-100">
          {recentOps.map((op) => {
            const it = itemById(op.itemId);
            return (
              <div key={op.id} className="px-4 py-3 flex items-center justify-between text-sm">
                <div>
                  <div className="font-bold text-stone-900 flex items-center gap-2">{it ? `${it.code} — ${it.design}` : op.itemId} <TypeBadge type={op.type} /></div>
                  <div className="text-xs text-stone-500 mt-0.5">
                    {op.type === "transfer" && `${whLabel(op.fromWarehouse)} ← ${whLabel(op.toWarehouse)} · `}
                    {op.type === "adjustment" && `${op.reason} · `}
                    {qtyDisplay(it, op.qty, op.boxCount)}
                  </div>
                </div>
                <div className="text-left"><div className="text-xs text-stone-400">{op.opNumber}</div><div className="text-[11px] text-stone-400">{fmtDate(op.date)}</div></div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* کارتکس کالا (تاریخچه دقیق)                                          */
/* ------------------------------------------------------------------ */

function buildLedger(item, ops, deliveries) {
  const events = [
    ...ops.filter((o) => o.itemId === item.id).map((o) => ({ ...o, kind: "op" })),
    ...deliveries.filter((d) => d.itemId === item.id).map((d) => ({ ...d, kind: "delivery", type: "delivery", opNumber: d.deliveryNumber })),
  ].sort((a, b) => new Date(a.date) - new Date(b.date));

  let w1 = 0, w2 = 0;
  const rows = events.map((e) => {
    let dW1 = 0, dW2 = 0, desc = "";
    if (e.type === "stockin") {
      if (e.warehouse === "w1") dW1 = e.qty; else dW2 = e.qty;
      desc = `ورود کالا${e.supplier ? ` از ${e.supplier}` : ""}`;
    } else if (e.type === "transfer") {
      if (e.fromWarehouse === "w1") dW1 -= e.qty; else dW2 -= e.qty;
      if (e.toWarehouse === "w1") dW1 += e.qty; else dW2 += e.qty;
      desc = `انتقال ${whLabel(e.fromWarehouse)} ← ${whLabel(e.toWarehouse)}`;
    } else if (e.type === "adjustment") {
      if (e.warehouse === "w1") dW1 = -e.qty; else dW2 = -e.qty;
      desc = `اصلاح موجودی (${e.reason})`;
    } else if (e.kind === "delivery") {
      if (e.warehouse === "w1") dW1 = -e.qty; else dW2 = -e.qty;
      desc = `تحویل به ${e.customerName} (فاکتور ${e.invoiceNumber})`;
    }
    w1 += dW1; w2 += dW2;
    return { ...e, dW1, dW2, balanceW1: w1, balanceW2: w2, desc };
  });
  return rows.reverse();
}

function ItemLedgerModal({ item, stock, ops, deliveries, onClose, onEditItem, isManager, onPrint }) {
  const [editing, setEditing] = useState(false);
  const recs = stock.filter((s) => s.itemId === item.id);
  const w1 = recs.filter((r) => r.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0);
  const w2 = recs.filter((r) => r.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0);
  const ledger = buildLedger(item, ops, deliveries);

  const printLedger = () => onPrint(`کارتکس ${item.code} — ${item.design}`, ledger.map((r) => ({
    label: `${fmtDate(r.date)} — ${r.desc}`, sub: r.opNumber,
    value: `${r.dW1 ? `w1:${r.dW1 > 0 ? "+" : ""}${fmt(r.dW1)} ` : ""}${r.dW2 ? `w2:${r.dW2 > 0 ? "+" : ""}${fmt(r.dW2)}` : ""}`.trim() || "—",
  })));

  return (
    <Modal title={`کارتکس کالا — ${item.code}`} onClose={onClose}>
      {editing ? (
        <EditItemForm item={item} onCancel={() => setEditing(false)} onSave={async (patch) => { await onEditItem(item.id, patch); setEditing(false); }} />
      ) : (
        <>
          <div className="mb-3">
            <div className="font-bold text-sm">{item.design} — {item.size} — {item.grade}</div>
            <div className="text-xs text-stone-500">{item.factory}{item.material ? ` · ${item.material}` : ""} · {item.surface}</div>
          </div>
          <div className="grid grid-cols-3 gap-2 mb-4 text-center">
            <div className="bg-stone-50 rounded py-2"><div className="text-[10px] text-stone-400">انبار ۱</div><div className="font-bold">{fmt(w1)}</div></div>
            <div className="bg-stone-50 rounded py-2"><div className="text-[10px] text-stone-400">انبار ۲</div><div className="font-bold">{fmt(w2)}</div></div>
            <div className="bg-amber-50 rounded py-2"><div className="text-[10px] text-amber-700">کل</div><div className="font-bold text-amber-800">{fmt(w1 + w2)}</div></div>
          </div>
          <div className="flex gap-2 mb-4">
            {isManager && <button onClick={() => setEditing(true)} className="text-xs font-bold border border-stone-300 rounded-lg px-3 py-1.5 inline-flex items-center gap-1"><Pencil size={12} /> ویرایش مشخصات</button>}
            <button onClick={printLedger} className="text-xs font-bold border border-stone-300 rounded-lg px-3 py-1.5 inline-flex items-center gap-1"><Printer size={12} /> چاپ کارتکس</button>
          </div>
          <div className="text-xs font-bold text-stone-600 mb-2">تاریخچه دقیق (جدیدترین بالا)</div>
          {ledger.length === 0 ? <Empty text="هنوز اتفاقی برای این کالا ثبت نشده است." /> : (
            <div className="space-y-1.5 max-h-72 overflow-y-auto">
              {ledger.map((r, i) => (
                <div key={i} className="border border-stone-100 rounded-lg p-2 text-xs">
                  <div className="flex items-center justify-between mb-0.5"><TypeBadge type={r.type} /><span className="text-stone-400">{fmtDate(r.date)}</span></div>
                  <div className="text-stone-700">{r.desc}</div>
                  <div className="flex items-center justify-between mt-1 text-stone-500">
                    <span>{r.dW1 ? `انبار۱: ${r.dW1 > 0 ? "+" : ""}${fmt(r.dW1)} (مانده ${fmt(r.balanceW1)})` : ""}</span>
                    <span>{r.dW2 ? `انبار۲: ${r.dW2 > 0 ? "+" : ""}${fmt(r.dW2)} (مانده ${fmt(r.balanceW2)})` : ""}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </Modal>
  );
}

function EditItemForm({ item, onSave, onCancel }) {
  const [form, setForm] = useState({ ...item });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const setVal = (k) => (v) => setForm((f) => ({ ...f, [k]: v }));
  return (
    <div>
      <Field label="کارخانه"><input className={inputCls} value={form.factory} onChange={set("factory")} /></Field>
      <Field label="طرح"><input className={inputCls} value={form.design} onChange={set("design")} /></Field>
      <PresetOrCustomField label="سایز" options={SIZES} value={form.size} onChange={setVal("size")} />
      <Field label="درجه"><select className={inputCls} value={form.grade} onChange={set("grade")}>{GRADES.map((g) => <option key={g}>{g}</option>)}</select></Field>
      <PresetOrCustomField label="نوع سطح" options={SURFACES} value={form.surface} onChange={setVal("surface")} />
      <PresetOrCustomField label="جنس" options={MATERIALS} value={form.material || ""} onChange={setVal("material")} />
      <Field label="متراژ هر کارتن (اختیاری)"><input {...numProps} className={inputCls} value={form.sqmPerBox || ""} onChange={set("sqmPerBox")} /></Field>
      <Field label="آستانه هشدار کمبود (اختیاری)"><input {...numProps} className={inputCls} value={form.lowThreshold || ""} onChange={set("lowThreshold")} /></Field>
      <Field label="توضیحات"><textarea className={inputCls} rows={2} value={form.description || ""} onChange={set("description")} /></Field>
      <div className="flex gap-2"><BigButton onClick={() => onSave(form)}>ذخیره اصلاحات</BigButton><BigButton tone="ghost" onClick={onCancel}>انصراف</BigButton></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه کالاها                                                         */
/* ------------------------------------------------------------------ */

function ItemsPage({ items, stock, ops, deliveries, itemTotals, onAdd, onEditItem, isManager, onPrintItem, onPrintLedger }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const [ledgerItem, setLedgerItem] = useState(null);
  const filtered = items.filter((it) => `${it.code} ${it.factory} ${it.design} ${it.size} ${it.grade} ${it.surface}`.includes(q));

  return (
    <div>
      <SectionTitle info="تعریف مشخصات کالاها. برای دیدن تاریخچه دقیق و موجودی هر کالا، روی کارتش بزنید. موجودی از اینجا وارد نمی‌شه — برای ورود موجودی به تب «ورود / اصلاح» برید."
        action={<BigButton onClick={() => setOpen(true)}><span className="inline-flex items-center gap-1"><Plus size={16} /> کالای جدید</span></BigButton>}>
        فهرست کالاها
      </SectionTitle>

      <div className="relative mb-3">
        <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400" />
        <input className={inputCls + " pr-9"} placeholder="جستجو در کد، کارخانه، طرح، سایز…" value={q} onChange={(e) => setQ(e.target.value)} />
      </div>

      {filtered.length === 0 ? <Empty text="کالایی یافت نشد." /> : (
        <div className="grid gap-3 sm:grid-cols-2">
          {filtered.map((it) => {
            const t = itemTotals(it.id);
            return (
              <div key={it.id} onClick={() => setLedgerItem(it)} className="bg-white border border-stone-200 rounded-lg p-4 cursor-pointer hover:border-amber-400 transition">
                <div className="flex items-start justify-between">
                  <div><div className="text-xs text-amber-800 font-bold">{it.code}</div><div className="font-bold text-stone-900">{it.design} — {it.size}</div></div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-[11px] bg-stone-100 text-stone-600 rounded px-2 py-0.5">{it.grade}</span>
                    <span onClick={(e) => e.stopPropagation()}><IconButton icon={Printer} title="چاپ گزارش این کالا" onClick={() => onPrintItem(it)} /></span>
                  </div>
                </div>
                <div className="text-xs text-stone-500 mt-2 space-y-0.5">
                  <div>کارخانه: {it.factory}{it.material ? ` · جنس: ${it.material}` : ""}</div>
                  <div>نوع سطح: {it.surface} · واحد: {it.unit}{isCartonItem(it) ? ` (هر کارتن ${fmt(it.sqmPerBox)} ${it.unit})` : ""}</div>
                  {it.description && <div>توضیحات: {it.description}</div>}
                </div>
                <div className="grid grid-cols-3 gap-2 mt-3 text-center">
                  <div className="bg-stone-50 rounded py-1.5"><div className="text-[10px] text-stone-400">انبار ۱</div><div className="text-sm font-bold">{fmt(t.w1)}</div></div>
                  <div className="bg-stone-50 rounded py-1.5"><div className="text-[10px] text-stone-400">انبار ۲</div><div className="text-sm font-bold">{fmt(t.w2)}</div></div>
                  <div className="bg-amber-50 rounded py-1.5"><div className="text-[10px] text-amber-700">کل</div><div className="text-sm font-bold text-amber-800">{fmt(t.total)}</div></div>
                </div>
                <div className="text-[11px] text-amber-700 mt-2 inline-flex items-center gap-1"><BookOpen size={11} /> مشاهده کارتکس کامل</div>
              </div>
            );
          })}
        </div>
      )}
      {open && <AddItemModal items={items} onClose={() => setOpen(false)} onAdd={onAdd} />}
      {ledgerItem && <ItemLedgerModal item={ledgerItem} stock={stock} ops={ops} deliveries={deliveries} onClose={() => setLedgerItem(null)} onEditItem={onEditItem} isManager={isManager} onPrint={onPrintLedger} />}
    </div>
  );
}

function AddItemModal({ items, onClose, onAdd }) {
  const [form, setForm] = useState({ factory: "", design: "", size: "", grade: GRADES[0], surface: "", material: "", unit: "متر مربع", sqmPerBox: "", lowThreshold: "", description: "" });
  const [confirmDup, setConfirmDup] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const setVal = (k) => (v) => setForm((f) => ({ ...f, [k]: v }));
  const canSave = form.factory && form.design && form.size;
  const dup = items.find((it) => it.design === form.design && it.size === form.size && it.grade === form.grade);

  useEffect(() => { setConfirmDup(false); }, [form.design, form.size, form.grade]);

  const handleSave = async () => {
    if (dup && !confirmDup) { setConfirmDup(true); return; }
    await onAdd(form);
    onClose();
  };

  return (
    <Modal title="ثبت کالای جدید" onClose={onClose}>
      <Field label="کارخانه"><input className={inputCls} value={form.factory} onChange={set("factory")} /></Field>
      <Field label="طرح"><input className={inputCls} value={form.design} onChange={set("design")} /></Field>
      <PresetOrCustomField label="سایز" options={SIZES} value={form.size} onChange={setVal("size")} />
      <Field label="درجه"><select className={inputCls} value={form.grade} onChange={set("grade")}>{GRADES.map((g) => <option key={g}>{g}</option>)}</select></Field>
      <PresetOrCustomField label="نوع سطح" options={SURFACES} value={form.surface} onChange={setVal("surface")} />
      <PresetOrCustomField label="جنس" options={MATERIALS} value={form.material} onChange={setVal("material")} />
      <Field label="واحد"><select className={inputCls} value={form.unit} onChange={set("unit")}><option>متر مربع</option><option>کارتن</option><option>عدد</option></select></Field>
      <Field label="متراژ هر کارتن (اختیاری)" hint="اگر پر شود، بعداً فقط «تعداد کارتن» وارد می‌کنید">
        <input {...numProps} className={inputCls} value={form.sqmPerBox} onChange={set("sqmPerBox")} placeholder="مثلاً 1.44" />
      </Field>
      <Field label="آستانه هشدار کمبود (اختیاری)" hint="وقتی موجودی کل به این عدد یا کمتر برسد، در داشبورد قرمز نشان داده می‌شود">
        <input {...numProps} className={inputCls} value={form.lowThreshold} onChange={set("lowThreshold")} placeholder="مثلاً 20" />
      </Field>
      <Field label="توضیحات"><textarea className={inputCls} rows={2} value={form.description} onChange={set("description")} /></Field>

      {dup && confirmDup && (
        <div className="text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2.5 mb-3">
          یک کالای مشابه از قبل هست: <b>{dup.code}</b> ({dup.design} — {dup.size} — {dup.grade}). اگه مطمئنید دوباره روی «تأیید و ثبت» بزنید.
        </div>
      )}
      <div className="flex gap-2 pt-2">
        <BigButton disabled={!canSave} onClick={handleSave}>{dup && confirmDup ? "تأیید و ثبت با وجود مشابه" : "ثبت کالا"}</BigButton>
        <BigButton tone="ghost" onClick={onClose}>انصراف</BigButton>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه موجودی — یکجا، با گروه‌بندی و مرتب‌سازی                       */
/* ------------------------------------------------------------------ */

function copyRowsToClipboard(rows, filter, showToast) {
  const head = ["کد", "طرح", "سایز", ...(filter === "all" || filter === "w1" ? ["انبار ۱"] : []), ...(filter === "all" || filter === "w2" ? ["انبار ۲"] : []), "کل", "رزروشده", "آزاد"];
  const lines = rows.map((r) => [r.item.code, r.item.design, r.item.size, ...(filter === "all" || filter === "w1" ? [r.w1] : []), ...(filter === "all" || filter === "w2" ? [r.w2] : []), r.total, r.reserved, r.free].join("\t"));
  const text = [head.join("\t"), ...lines].join("\n");
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => showToast && showToast("جدول کپی شد؛ می‌توانید در اکسل یا واتساپ Paste کنید.")).catch(() => showToast && showToast("کپی ناموفق بود.", "error"));
  } else if (showToast) showToast("کپی خودکار پشتیبانی نمی‌شود.", "error");
}

function InventoryPage({ items, itemTotals, onPrint, showToast }) {
  const [filter, setFilter] = useState("all");
  const [groupBy, setGroupBy] = useState("none");
  const [sortBy, setSortBy] = useState("none");
  const [sizeFocus, setSizeFocus] = useState("");
  const allSizes = [...new Set(items.map((it) => it.size))].sort();

  let rows = items.map((it) => ({ item: it, ...itemTotals(it.id) }));
  const valueFor = (r) => (filter === "all" ? r.total : filter === "w1" ? r.w1 : r.w2);
  if (sortBy !== "none") rows = [...rows].sort((a, b) => (sortBy === "qtyDesc" ? valueFor(b) - valueFor(a) : valueFor(a) - valueFor(b)));

  let groups = null;
  if (sizeFocus) groups = { [sizeFocus]: rows.filter((r) => r.item.size === sizeFocus) };
  else if (groupBy === "size") { groups = {}; rows.forEach((r) => { (groups[r.item.size] = groups[r.item.size] || []).push(r); }); }

  return (
    <div>
      <SectionTitle info="نمای یکجای موجودی همه کالاها در هر دو انبار. برای دیدن یک سایز به‌تنهایی، از «سایز خاص» استفاده کنید."
        action={
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex bg-white border border-stone-300 rounded-lg overflow-hidden text-xs">
              {[{ id: "all", label: "همه" }, { id: "w1", label: "انبار ۱" }, { id: "w2", label: "انبار ۲" }].map((f) => (
                <button key={f.id} onClick={() => setFilter(f.id)} className={`px-3 py-1.5 font-bold ${filter === f.id ? "bg-amber-700 text-white" : "text-stone-600 hover:bg-stone-50"}`}>{f.label}</button>
              ))}
            </div>
            <select className="text-xs border border-stone-300 rounded-lg px-2 py-1.5 bg-white" value={sizeFocus} onChange={(e) => setSizeFocus(e.target.value)}>
              <option value="">سایز خاص: همه</option>{allSizes.map((s) => <option key={s} value={s}>فقط سایز {s}</option>)}
            </select>
            {!sizeFocus && <select className="text-xs border border-stone-300 rounded-lg px-2 py-1.5 bg-white" value={groupBy} onChange={(e) => setGroupBy(e.target.value)}><option value="none">بدون گروه‌بندی</option><option value="size">گروه‌بندی بر اساس سایز</option></select>}
            <select className="text-xs border border-stone-300 rounded-lg px-2 py-1.5 bg-white" value={sortBy} onChange={(e) => setSortBy(e.target.value)}><option value="none">بدون مرتب‌سازی</option><option value="qtyDesc">بیشترین متراژ</option><option value="qtyAsc">کمترین متراژ</option></select>
            {!groups && <IconButton icon={Printer} title="چاپ گزارش موجودی" onClick={() => onPrint(rows)} />}
          </div>
        }>
        موجودی کالاها
      </SectionTitle>

      {groups ? (
        <div className="space-y-4">
          {Object.entries(groups).map(([size, groupRows]) => {
            const groupTotal = groupRows.reduce((a, r) => a + valueFor(r), 0);
            return (
              <div key={size} className="bg-white border border-stone-200 rounded-lg overflow-hidden">
                <div className="bg-stone-50 px-3 py-2 flex items-center justify-between border-b border-stone-200 flex-wrap gap-2">
                  <span className="text-sm font-bold">سایز: {size} <span className="text-xs text-amber-800 font-bold">(جمع: {fmt(groupTotal)})</span></span>
                  <div className="flex items-center gap-1.5">
                    <IconButton icon={Printer} title="چاپ / PDF این سایز" onClick={() => onPrint(groupRows)} />
                    <button onClick={() => copyRowsToClipboard(groupRows, filter, showToast)} className="text-xs font-bold border border-stone-300 rounded-lg px-2.5 py-1.5 bg-white text-stone-600 hover:text-amber-800 hover:border-amber-700">کپی جدول</button>
                  </div>
                </div>
                <InventoryTable rows={groupRows} filter={filter} />
              </div>
            );
          })}
        </div>
      ) : <div className="bg-white border border-stone-200 rounded-lg overflow-x-auto"><InventoryTable rows={rows} filter={filter} /></div>}
    </div>
  );
}

function InventoryTable({ rows, filter }) {
  return (
    <table className="w-full text-sm min-w-[620px]">
      <thead><tr className="bg-stone-50 text-stone-500 text-xs"><th className="text-right px-3 py-2 font-bold">کالا</th>{(filter === "all" || filter === "w1") && <th className="text-right px-3 py-2 font-bold">انبار ۱</th>}{(filter === "all" || filter === "w2") && <th className="text-right px-3 py-2 font-bold">انبار ۲</th>}<th className="text-right px-3 py-2 font-bold">کل</th><th className="text-right px-3 py-2 font-bold">رزروشده</th><th className="text-right px-3 py-2 font-bold">آزاد</th></tr></thead>
      <tbody className="divide-y divide-stone-100">
        {rows.map(({ item: it, w1, w2, total, reserved, free }) => (
          <tr key={it.id}>
            <td className="px-3 py-2"><div className="font-bold text-stone-900">{it.code}</div><div className="text-xs text-stone-500">{it.design} · {it.size}</div></td>
            {(filter === "all" || filter === "w1") && <td className="px-3 py-2">{fmt(w1)}</td>}
            {(filter === "all" || filter === "w2") && <td className="px-3 py-2">{fmt(w2)}</td>}
            <td className="px-3 py-2 font-bold text-amber-800">{fmt(total)}</td><td className="px-3 py-2">{fmt(reserved)}</td><td className="px-3 py-2">{fmt(free)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/* ------------------------------------------------------------------ */
/* فیلد مقدار مشترک                                                    */
/* ------------------------------------------------------------------ */

function QuantityInput({ item, boxCount, setBoxCount, plainQty, setPlainQty, maxAvailable, label }) {
  if (isCartonItem(item)) {
    const computedQty = Number(boxCount || 0) * Number(item.sqmPerBox || 0);
    const maxBoxes = item.sqmPerBox ? maxAvailable / Number(item.sqmPerBox) : null;
    return (
      <Field label={label || "تعداد کارتن"} hint={maxAvailable != null ? `هر کارتن = ${fmt(item.sqmPerBox)} ${item.unit} · حداکثر: ${fmt(maxBoxes)} کارتن (${fmt(maxAvailable)} ${item.unit}) · این مقدار = ${fmt(computedQty)} ${item.unit}` : `هر کارتن = ${fmt(item.sqmPerBox)} ${item.unit}`}>
        <input {...numProps} className={inputCls} value={boxCount} onChange={(e) => setBoxCount(e.target.value)} />
      </Field>
    );
  }
  return (
    <Field label={label || `مقدار (${item?.unit || ""})`} hint={maxAvailable != null ? `حداکثر: ${fmt(maxAvailable)} ${item?.unit || ""}` : undefined}>
      <input {...numProps} className={inputCls} value={plainQty} onChange={(e) => setPlainQty(e.target.value)} />
    </Field>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه ورود / اصلاح انبار                                             */
/* ------------------------------------------------------------------ */

function StockOpsPage({ items, stock, stockForItem, pendingCounts, itemById, onAddItem, doStockIn, doAdjustment, doPhysicalCount, resolvePendingCount, showToast }) {
  const [mode, setMode] = useState("in");
  return (
    <div>
      <SectionTitle info="از اینجا موجودی واقعی انبار دقیق می‌شود: ورود کالای تازه، ثبت شکستگی/مفقودی، یا شمارش فیزیکی توسط انباردار.">ورود / اصلاح موجودی انبار</SectionTitle>
      <div className="flex bg-white border border-stone-300 rounded-lg overflow-hidden text-xs mb-4 w-fit">
        <button onClick={() => setMode("in")} className={`px-4 py-2 font-bold ${mode === "in" ? "bg-amber-700 text-white" : "text-stone-600"}`}>ورود کالا</button>
        <button onClick={() => setMode("adjust")} className={`px-4 py-2 font-bold ${mode === "adjust" ? "bg-amber-700 text-white" : "text-stone-600"}`}>کسری / شکستگی</button>
        <button onClick={() => setMode("count")} className={`px-4 py-2 font-bold ${mode === "count" ? "bg-amber-700 text-white" : "text-stone-600"}`}>شمارش فیزیکی</button>
      </div>
      {mode === "in" && <StockInForm items={items} stockForItem={stockForItem} onAddItem={onAddItem} doStockIn={doStockIn} showToast={showToast} />}
      {mode === "adjust" && <AdjustForm items={items} stockForItem={stockForItem} onAddItem={onAddItem} doAdjustment={doAdjustment} showToast={showToast} />}
      {mode === "count" && <CountForm items={items} stockForItem={stockForItem} doPhysicalCount={doPhysicalCount} showToast={showToast} />}

      <SectionTitle info="شمارش‌های فیزیکی که عدد کارگر با سیستم فرق داشته و منتظر تأیید مدیر برای اصلاح واقعی موجودی هستند.">در انتظار تأیید مدیر</SectionTitle>
      {pendingCounts.length === 0 ? <Empty text="موردی در انتظار تأیید نیست." /> : (
        <div className="space-y-2">
          {pendingCounts.map((p) => {
            const it = itemById(p.itemId);
            return (
              <div key={p.id} className="bg-white border border-rose-200 rounded-lg p-3.5">
                <div className="font-bold text-sm mb-1">{it?.code} — {it?.design}</div>
                <div className="text-xs text-stone-500 mb-2">{whLabel(p.warehouse)} / {locLabel(p.location)} · ثبت‌شده: {qtyDisplay(it, p.recordedQty, p.recordedBoxCount)} ← شمارش‌شده: {qtyDisplay(it, p.countedQty, p.countedBoxCount)} ({p.diff > 0 ? "+" : ""}{fmt(p.diff)})</div>
                <div className="flex gap-2">
                  <button onClick={() => resolvePendingCount(p.id, true)} className="text-xs font-bold bg-amber-700 text-white rounded-lg px-3 py-1.5">تأیید و اصلاح موجودی</button>
                  <button onClick={() => resolvePendingCount(p.id, false)} className="text-xs font-bold border border-stone-300 rounded-lg px-3 py-1.5">نادیده گرفتن</button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function NewItemInline({ items, onAddItem, onCreated }) {
  const [show, setShow] = useState(false);
  return (
    <div className="mb-3">
      <button type="button" onClick={() => setShow(true)} className="text-xs font-bold text-amber-800 border border-amber-700 rounded-lg px-3 py-1.5 inline-flex items-center gap-1"><Plus size={13} /> کالا در سیستم نیست؟ ثبت کالای جدید</button>
      {show && <AddItemModal items={items} onClose={() => setShow(false)} onAdd={async (form) => { const it = await onAddItem(form); onCreated(it); return it; }} />}
    </div>
  );
}

function StockInForm({ items, stockForItem, onAddItem, doStockIn, showToast }) {
  const [itemId, setItemId] = useState("");
  const [warehouse, setWarehouse] = useState("w1");
  const [useExisting, setUseExisting] = useState(true);
  const [recId, setRecId] = useState("");
  const [aisle, setAisle] = useState(""), [row, setRow] = useState(""), [cell, setCell] = useState("");
  const [boxCount, setBoxCount] = useState(""), [plainQty, setPlainQty] = useState("");
  const [supplier, setSupplier] = useState("");
  const [description, setDescription] = useState("");
  const [err, setErr] = useState("");
  const item = items.find((i) => i.id === itemId);
  const options = itemId ? stockForItem(itemId, warehouse) : [];
  const cartonMode = isCartonItem(item);
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  const canSubmit = itemId && q > 0 && (useExisting ? recId : aisle && row && cell);

  useEffect(() => { setRecId(""); setBoxCount(""); setPlainQty(""); }, [itemId, warehouse]);

  const submit = async () => {
    setErr("");
    const res = await doStockIn({ itemId, warehouse, useExisting, recId, toLoc: { aisle, row, cell }, qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null, supplier, description });
    if (!res.ok) { setErr(res.msg); showToast(res.msg, "error"); return; }
    showToast(`ورود کالا با شماره ${res.op.opNumber} ثبت شد.`);
    setBoxCount(""); setPlainQty(""); setDescription("");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <NewItemInline items={items} onAddItem={onAddItem} onCreated={(it) => setItemId(it.id)} />
      <Field label="کالا"><ItemPicker items={items} value={itemId} onChange={setItemId} /></Field>
      <Field label="انبار مقصد"><select className={inputCls} value={warehouse} onChange={(e) => setWarehouse(e.target.value)}>{WAREHOUSES.map((w) => <option key={w.id} value={w.id}>{w.label}</option>)}</select></Field>
      <Field label="محل">
        <div className="flex gap-2 mb-2 text-xs">
          <button type="button" onClick={() => setUseExisting(true)} className={`px-3 py-1.5 rounded-lg border font-bold ${useExisting ? "bg-amber-700 text-white border-amber-700" : "border-stone-300 text-stone-600"}`}>محل موجود</button>
          <button type="button" onClick={() => setUseExisting(false)} className={`px-3 py-1.5 rounded-lg border font-bold ${!useExisting ? "bg-amber-700 text-white border-amber-700" : "border-stone-300 text-stone-600"}`}>محل جدید</button>
        </div>
        {useExisting ? (
          <select className={inputCls} value={recId} onChange={(e) => setRecId(e.target.value)} disabled={!itemId}>
            <option value="">انتخاب محل…</option>{options.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — فعلی: {fmt(r.qty)}</option>)}
          </select>
        ) : (
          <div className="grid grid-cols-3 gap-2">
            <input className={inputCls} placeholder="راهرو" value={aisle} onChange={(e) => setAisle(e.target.value)} />
            <input className={inputCls} placeholder="ردیف" value={row} onChange={(e) => setRow(e.target.value)} />
            <input className={inputCls} placeholder="خانه" value={cell} onChange={(e) => setCell(e.target.value)} />
          </div>
        )}
      </Field>
      <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={null} label="تعداد کارتن وارده" />
      <Field label="تأمین‌کننده / کارخانه (اختیاری)"><input className={inputCls} value={supplier} onChange={(e) => setSupplier(e.target.value)} /></Field>
      <Field label="توضیحات (اختیاری)"><textarea className={inputCls} rows={2} value={description} onChange={(e) => setDescription(e.target.value)} /></Field>
      {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3">{err}</div>}
      <BigButton onClick={submit} disabled={!canSubmit}>ثبت ورود کالا</BigButton>
    </div>
  );
}

function AdjustForm({ items, stockForItem, onAddItem, doAdjustment, showToast }) {
  const [itemId, setItemId] = useState("");
  const [warehouse, setWarehouse] = useState("w1");
  const [recId, setRecId] = useState("");
  const [boxCount, setBoxCount] = useState(""), [plainQty, setPlainQty] = useState("");
  const [reason, setReason] = useState(REASONS[0]);
  const [description, setDescription] = useState("");
  const [err, setErr] = useState("");
  const item = items.find((i) => i.id === itemId);
  const options = itemId ? stockForItem(itemId, warehouse) : [];
  const rec = options.find((r) => r.id === recId);
  const cartonMode = isCartonItem(item);
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  const canSubmit = itemId && recId && q > 0;

  useEffect(() => { setRecId(""); setBoxCount(""); setPlainQty(""); }, [itemId, warehouse]);

  const submit = async () => {
    setErr("");
    const res = await doAdjustment({ itemId, recId, qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null, reason, description });
    if (!res.ok) { setErr(res.msg); showToast(res.msg, "error"); return; }
    showToast(`اصلاح موجودی با شماره ${res.op.opNumber} ثبت شد.`);
    setBoxCount(""); setPlainQty(""); setDescription("");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <NewItemInline items={items} onAddItem={onAddItem} onCreated={(it) => setItemId(it.id)} />
      <Field label="کالا"><ItemPicker items={items} value={itemId} onChange={setItemId} /></Field>
      <Field label="انبار"><select className={inputCls} value={warehouse} onChange={(e) => setWarehouse(e.target.value)}>{WAREHOUSES.map((w) => <option key={w.id} value={w.id}>{w.label}</option>)}</select></Field>
      <Field label="محل">
        <select className={inputCls} value={recId} onChange={(e) => setRecId(e.target.value)} disabled={!itemId}>
          <option value="">انتخاب محل…</option>{options.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — موجودی: {fmt(r.qty)}</option>)}
        </select>
      </Field>
      {recId && <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={Number(rec?.qty || 0)} label="تعداد کارتن کسری‌شده" />}
      <Field label="دلیل"><select className={inputCls} value={reason} onChange={(e) => setReason(e.target.value)}>{REASONS.map((r) => <option key={r}>{r}</option>)}</select></Field>
      <Field label="توضیحات (اختیاری)"><textarea className={inputCls} rows={2} value={description} onChange={(e) => setDescription(e.target.value)} /></Field>
      {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3">{err}</div>}
      <BigButton tone="danger" onClick={submit} disabled={!canSubmit}>ثبت اصلاح (کسری)</BigButton>
    </div>
  );
}

function CountForm({ items, stockForItem, doPhysicalCount, showToast }) {
  const [itemId, setItemId] = useState("");
  const [warehouse, setWarehouse] = useState("w1");
  const [recId, setRecId] = useState("");
  const [boxCount, setBoxCount] = useState(""), [plainQty, setPlainQty] = useState("");
  const item = items.find((i) => i.id === itemId);
  const options = itemId ? stockForItem(itemId, warehouse) : [];
  const cartonMode = isCartonItem(item);
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  const canSubmit = itemId && recId && (boxCount !== "" || plainQty !== "");

  useEffect(() => { setRecId(""); setBoxCount(""); setPlainQty(""); }, [itemId, warehouse]);

  const submit = async () => {
    const res = await doPhysicalCount({ itemId, recId, countedQty: q, countedBoxCount: cartonMode ? Number(boxCount || 0) : null });
    if (!res.ok) { showToast(res.msg, "error"); return; }
    if (res.matched) showToast("شمارش با موجودی سیستم مطابقت داشت — مغایرتی ثبت نشد.");
    else showToast("مغایرت ثبت شد؛ منتظر تأیید مدیر برای اصلاح است.", "error");
    setBoxCount(""); setPlainQty("");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <Field label="کالا"><ItemPicker items={items} value={itemId} onChange={setItemId} /></Field>
      <Field label="انبار"><select className={inputCls} value={warehouse} onChange={(e) => setWarehouse(e.target.value)}>{WAREHOUSES.map((w) => <option key={w.id} value={w.id}>{w.label}</option>)}</select></Field>
      <Field label="محل / پالت">
        <select className={inputCls} value={recId} onChange={(e) => setRecId(e.target.value)} disabled={!itemId}>
          <option value="">انتخاب محل…</option>{options.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — عدد سیستم: {fmt(r.qty)}</option>)}
        </select>
      </Field>
      {recId && <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={null} label="تعداد شمارش‌شده (تقریبی)" />}
      <BigButton onClick={submit} disabled={!canSubmit}>ثبت شمارش</BigButton>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه انتقال بین انبارها                                             */
/* ------------------------------------------------------------------ */

function TransferPage({ items, stock, stockForItem, doTransfer, showToast }) {
  const [itemId, setItemId] = useState("");
  const [fromWh, setFromWh] = useState("w1");
  const [fromRecId, setFromRecId] = useState("");
  const [boxCount, setBoxCount] = useState(""), [plainQty, setPlainQty] = useState("");
  const [useExistingTo, setUseExistingTo] = useState(true);
  const [toRecId, setToRecId] = useState("");
  const [toAisle, setToAisle] = useState(""), [toRow, setToRow] = useState(""), [toCell, setToCell] = useState("");
  const [description, setDescription] = useState("");
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");

  const item = items.find((i) => i.id === itemId);
  const toWh = fromWh === "w1" ? "w2" : "w1";
  const fromOptions = itemId ? stockForItem(itemId, fromWh) : [];
  const toOptions = itemId ? stockForItem(itemId, toWh) : [];
  const fromRec = stock.find((s) => s.id === fromRecId);
  const available = fromRec ? Number(fromRec.qty || 0) - Number(fromRec.reserved || 0) : 0;
  const cartonMode = isCartonItem(item);
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);

  const fillTestScenario = () => {
    setItemId("itm-a"); setFromWh("w1");
    setTimeout(() => { const recs = stockForItem("itm-a", "w1"); const rec = recs.find((r) => r.aisle === "A" && r.row === "02" && r.cell === "04") || recs[0]; if (rec) setFromRecId(rec.id); }, 0);
    setPlainQty("50"); setBoxCount(""); setUseExistingTo(true); setDescription("سناریوی تست اولیه انتقال");
  };

  useEffect(() => { setFromRecId(""); setToRecId(""); setUseExistingTo(true); setBoxCount(""); setPlainQty(""); }, [itemId, fromWh]);

  const canSubmit = itemId && fromRecId && q > 0 && (useExistingTo ? toRecId : toAisle && toRow && toCell);

  const submit = async () => {
    setErr(""); setResult(null);
    const res = await doTransfer({ itemId, fromWh, fromRecId, toWh, toLoc: { aisle: toAisle, row: toRow, cell: toCell }, useExistingTo, toRecId, qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null, description });
    if (!res.ok) { setErr(res.msg); showToast(res.msg, "error"); return; }
    setResult(res.op);
    showToast(`انتقال با شماره ${res.op.opNumber} با موفقیت ثبت شد.`);
    setPlainQty(""); setBoxCount(""); setDescription("");
  };

  return (
    <div>
      <SectionTitle info="جابه‌جایی موجودی موجود بین انبار۱ و انبار۲، بدون تغییر در موجودی کل."
        action={<button onClick={fillTestScenario} className="text-xs inline-flex items-center gap-1 text-stone-500 hover:text-amber-800 border border-stone-300 rounded-lg px-3 py-1.5"><RefreshCcw size={13} /> بارگذاری سناریوی تست</button>}>
        انتقال بین انبارها
      </SectionTitle>

      <div className="bg-white border border-stone-200 rounded-lg p-4">
        <Field label="کالا"><ItemPicker items={items} value={itemId} onChange={setItemId} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="انبار مبدأ"><select className={inputCls} value={fromWh} onChange={(e) => setFromWh(e.target.value)}>{WAREHOUSES.map((w) => <option key={w.id} value={w.id}>{w.label}</option>)}</select></Field>
          <Field label="انبار مقصد" hint="خودکار"><input className={inputCls + " bg-stone-100"} value={whLabel(toWh)} disabled /></Field>
        </div>
        <Field label="محل مبدأ">
          <select className={inputCls} value={fromRecId} onChange={(e) => setFromRecId(e.target.value)} disabled={!itemId}>
            <option value="">انتخاب محل…</option>{fromOptions.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — قابل انتقال: {qtyDisplay(item, Number(r.qty) - Number(r.reserved || 0), r.boxCount != null ? Number(r.boxCount) - Number(r.reserved || 0) : null)}</option>)}
          </select>
          {itemId && fromOptions.length === 0 && <div className="text-xs text-rose-600 mt-1">این کالا در {whLabel(fromWh)} موجودی ندارد.</div>}
        </Field>
        {fromRecId && <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={available} />}
        <Field label="محل مقصد">
          <div className="flex gap-2 mb-2 text-xs">
            <button type="button" onClick={() => setUseExistingTo(true)} className={`px-3 py-1.5 rounded-lg border font-bold ${useExistingTo ? "bg-amber-700 text-white border-amber-700" : "border-stone-300 text-stone-600"}`}>محل موجود</button>
            <button type="button" onClick={() => setUseExistingTo(false)} className={`px-3 py-1.5 rounded-lg border font-bold ${!useExistingTo ? "bg-amber-700 text-white border-amber-700" : "border-stone-300 text-stone-600"}`}>محل جدید</button>
          </div>
          {useExistingTo ? (
            <select className={inputCls} value={toRecId} onChange={(e) => setToRecId(e.target.value)} disabled={!itemId}>
              <option value="">انتخاب محل در {whLabel(toWh)}…</option>{toOptions.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — فعلی: {fmt(r.qty)}</option>)}
            </select>
          ) : (
            <div className="grid grid-cols-3 gap-2">
              <input className={inputCls} placeholder="راهرو" value={toAisle} onChange={(e) => setToAisle(e.target.value)} />
              <input className={inputCls} placeholder="ردیف" value={toRow} onChange={(e) => setToRow(e.target.value)} />
              <input className={inputCls} placeholder="خانه" value={toCell} onChange={(e) => setToCell(e.target.value)} />
            </div>
          )}
        </Field>
        <Field label="توضیحات (اختیاری)"><textarea className={inputCls} rows={2} value={description} onChange={(e) => setDescription(e.target.value)} /></Field>
        {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3">{err}</div>}
        <BigButton onClick={submit} disabled={!canSubmit}>ثبت انتقال</BigButton>
      </div>

      {result && item && (
        <div className="mt-4 bg-white border border-amber-300 rounded-lg p-4">
          <div className="text-sm font-bold text-amber-800 mb-2">انتقال با موفقیت ثبت شد — {result.opNumber}</div>
          <TransferResultSummary itemId={item.id} stock={stock} itemUnit={item.unit} />
        </div>
      )}
    </div>
  );
}

function TransferResultSummary({ itemId, stock, itemUnit }) {
  const w1 = stock.filter((s) => s.itemId === itemId && s.warehouse === "w1").reduce((a, r) => a + Number(r.qty || 0), 0);
  const w2 = stock.filter((s) => s.itemId === itemId && s.warehouse === "w2").reduce((a, r) => a + Number(r.qty || 0), 0);
  return (
    <div className="grid grid-cols-3 gap-2 text-center">
      <div className="bg-stone-50 rounded py-2"><div className="text-[10px] text-stone-400">انبار ۱</div><div className="text-lg font-bold">{fmt(w1)}</div></div>
      <div className="bg-stone-50 rounded py-2"><div className="text-[10px] text-stone-400">انبار ۲</div><div className="text-lg font-bold">{fmt(w2)}</div></div>
      <div className="bg-amber-50 rounded py-2"><div className="text-[10px] text-amber-700">کل ({itemUnit})</div><div className="text-lg font-bold text-amber-800">{fmt(w1 + w2)}</div></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه سفارش و تحویل مشتری                                            */
/* ------------------------------------------------------------------ */

function OrdersPage({ items, stock, orders, itemById, stockForItem, createOrder, doDelivery, showToast, onPrintDelivery }) {
  const [mode, setMode] = useState("new");
  const [onlyOpen, setOnlyOpen] = useState(true);
  return (
    <div>
      <SectionTitle info="ثبت سفارش چندقلمی مشتری (مطابق فاکتور دفتر) و تحویل جزئی؛ سیستم خودش بدهکاری باقیمانده را حساب می‌کند.">سفارش و تحویل مشتری</SectionTitle>
      <div className="flex bg-white border border-stone-300 rounded-lg overflow-hidden text-xs mb-4 w-fit">
        <button onClick={() => setMode("new")} className={`px-4 py-2 font-bold ${mode === "new" ? "bg-amber-700 text-white" : "text-stone-600"}`}>سفارش جدید</button>
        <button onClick={() => setMode("deliver")} className={`px-4 py-2 font-bold ${mode === "deliver" ? "bg-amber-700 text-white" : "text-stone-600"}`}>ثبت تحویل</button>
      </div>
      {mode === "new" && <NewOrderForm items={items} createOrder={createOrder} showToast={showToast} />}
      {mode === "deliver" && <DeliverForm items={items} stock={stock} orders={orders} stockForItem={stockForItem} doDelivery={doDelivery} showToast={showToast} onPrintDelivery={onPrintDelivery} />}
      <SectionTitle info="سفارش‌های ثبت‌شده و میزان بدهکاری باقیمانده به هر مشتری." action={<label className="text-xs inline-flex items-center gap-1.5 text-stone-600"><input type="checkbox" checked={onlyOpen} onChange={(e) => setOnlyOpen(e.target.checked)} /> فقط سفارش‌های باز</label>}>
        سفارش‌ها و بدهکاری مشتریان
      </SectionTitle>
      <OrdersList orders={orders} itemById={itemById} onlyOpen={onlyOpen} />
    </div>
  );
}

function NewOrderForm({ items, createOrder, showToast }) {
  const [invoiceNumber, setInvoiceNumber] = useState("");
  const [customerName, setCustomerName] = useState("");
  const [customerPhone, setCustomerPhone] = useState("");
  const [lines, setLines] = useState([{ key: uid("l"), itemId: "", boxCount: "", plainQty: "" }]);
  const [err, setErr] = useState("");
  const addLine = () => setLines((ls) => [...ls, { key: uid("l"), itemId: "", boxCount: "", plainQty: "" }]);
  const removeLine = (key) => setLines((ls) => ls.filter((l) => l.key !== key));
  const updateLine = (key, patch) => setLines((ls) => ls.map((l) => (l.key === key ? { ...l, ...patch } : l)));
  const canSubmit = invoiceNumber && customerName && lines.some((l) => l.itemId && (Number(l.boxCount) > 0 || Number(l.plainQty) > 0));

  const submit = async () => {
    setErr("");
    const compiled = lines.filter((l) => l.itemId).map((l) => {
      const item = items.find((i) => i.id === l.itemId);
      const cartonMode = isCartonItem(item);
      const qty = cartonMode ? Number(l.boxCount || 0) * Number(item.sqmPerBox || 0) : Number(l.plainQty || 0);
      return { itemId: l.itemId, qty, boxCount: cartonMode ? Number(l.boxCount || 0) : null };
    }).filter((l) => l.qty > 0);
    const res = await createOrder({ invoiceNumber, customerName, customerPhone, lines: compiled });
    if (!res.ok) { setErr(res.msg); showToast(res.msg, "error"); return; }
    showToast(`سفارش برای ${customerName} ثبت شد.`);
    setInvoiceNumber(""); setCustomerName(""); setCustomerPhone(""); setLines([{ key: uid("l"), itemId: "", boxCount: "", plainQty: "" }]);
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <div className="grid grid-cols-2 gap-3">
        <Field label="شماره فاکتور (محک)"><input className={inputCls} value={invoiceNumber} onChange={(e) => setInvoiceNumber(e.target.value)} /></Field>
        <Field label="نام مشتری"><input className={inputCls} value={customerName} onChange={(e) => setCustomerName(e.target.value)} /></Field>
      </div>
      <Field label="شماره تماس (اختیاری)"><input className={inputCls} value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)} /></Field>
      <div className="text-xs font-bold text-stone-600 mb-1">اقلام سفارش</div>
      <div className="space-y-3 mb-3">
        {lines.map((l) => {
          const item = items.find((i) => i.id === l.itemId);
          const cartonMode = isCartonItem(item);
          return (
            <div key={l.key} className="border border-stone-200 rounded-lg p-3">
              <div className="flex items-center gap-2 mb-2">
                <div className="flex-1"><ItemPicker items={items} value={l.itemId} onChange={(id) => updateLine(l.key, { itemId: id })} placeholder="تایپ کنید: طرح، کد یا سایز…" /></div>
                {lines.length > 1 && <IconButton icon={Trash2} tone="danger" title="حذف ردیف" onClick={() => removeLine(l.key)} />}
              </div>
              {l.itemId && (cartonMode
                ? <input {...numProps} className={inputCls} placeholder="تعداد کارتن" value={l.boxCount} onChange={(e) => updateLine(l.key, { boxCount: e.target.value })} />
                : <input {...numProps} className={inputCls} placeholder={`مقدار (${item?.unit || ""})`} value={l.plainQty} onChange={(e) => updateLine(l.key, { plainQty: e.target.value })} />)}
            </div>
          );
        })}
      </div>
      <button onClick={addLine} className="text-xs font-bold text-amber-800 border border-amber-700 rounded-lg px-3 py-1.5 mb-4 inline-flex items-center gap-1"><Plus size={13} /> افزودن قلم کالا</button>
      {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3">{err}</div>}
      <div><BigButton onClick={submit} disabled={!canSubmit}>ثبت سفارش</BigButton></div>
    </div>
  );
}

function DeliverForm({ items, stock, orders, stockForItem, doDelivery, showToast, onPrintDelivery }) {
  const [orderId, setOrderId] = useState("");
  const [lineId, setLineId] = useState("");
  const [warehouse, setWarehouse] = useState("w1");
  const [recId, setRecId] = useState("");
  const [boxCount, setBoxCount] = useState(""), [plainQty, setPlainQty] = useState("");
  const [signature, setSignature] = useState(null);
  const [err, setErr] = useState("");

  const openOrders = orders.filter((o) => o.lines.some((l) => Number(l.qty) - Number(l.deliveredQty) > 0.001));
  const order = orders.find((o) => o.id === orderId);
  const openLines = order ? order.lines.filter((l) => Number(l.qty) - Number(l.deliveredQty) > 0.001) : [];
  const line = order?.lines.find((l) => l.id === lineId);
  const item = line ? items.find((i) => i.id === line.itemId) : null;
  const remaining = line ? Number(line.qty) - Number(line.deliveredQty) : 0;
  const options = item ? stockForItem(item.id, warehouse) : [];
  const rec = stock.find((s) => s.id === recId);
  const available = rec ? Number(rec.qty) - Number(rec.reserved || 0) : 0;
  const cartonMode = isCartonItem(item);
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  const maxAllowed = Math.min(available || 0, remaining || 0);

  useEffect(() => { setLineId(""); }, [orderId]);
  useEffect(() => { setRecId(""); setBoxCount(""); setPlainQty(""); }, [lineId, warehouse]);

  const canSubmit = orderId && lineId && recId && q > 0 && q <= maxAllowed + 1e-6;

  const submit = async () => {
    setErr("");
    const res = await doDelivery({ orderId, lineId, recId, qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null, signature });
    if (!res.ok) { setErr(res.msg); showToast(res.msg, "error"); return; }
    showToast(`تحویل ${res.delivery.deliveryNumber} ثبت شد.`);
    onPrintDelivery(res.delivery, item);
    setBoxCount(""); setPlainQty(""); setSignature(null); setRecId("");
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4">
      <Field label="سفارش">
        <select className={inputCls} value={orderId} onChange={(e) => setOrderId(e.target.value)}>
          <option value="">انتخاب سفارش…</option>{openOrders.map((o) => <option key={o.id} value={o.id}>{o.invoiceNumber} — {o.customerName}</option>)}
        </select>
        {openOrders.length === 0 && <div className="text-xs text-stone-400 mt-1">سفارش بازی وجود ندارد.</div>}
      </Field>
      {order && (
        <Field label="قلم کالا">
          <select className={inputCls} value={lineId} onChange={(e) => setLineId(e.target.value)}>
            <option value="">انتخاب قلم…</option>
            {openLines.map((l) => { const it = items.find((i) => i.id === l.itemId); return <option key={l.id} value={l.id}>{it?.code} — {it?.design} — باقیمانده: {fmt(Number(l.qty) - Number(l.deliveredQty))}</option>; })}
          </select>
        </Field>
      )}
      {line && (
        <>
          <Field label="انبار مبدأ تحویل"><select className={inputCls} value={warehouse} onChange={(e) => setWarehouse(e.target.value)}>{WAREHOUSES.map((w) => <option key={w.id} value={w.id}>{w.label}</option>)}</select></Field>
          <Field label="محل"><select className={inputCls} value={recId} onChange={(e) => setRecId(e.target.value)}><option value="">انتخاب محل…</option>{options.map((r) => <option key={r.id} value={r.id}>{locLabel(r)} — موجود: {fmt(Number(r.qty) - Number(r.reserved || 0))}</option>)}</select></Field>
          {recId && <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={maxAllowed} label="تعداد کارتن تحویلی این نوبت" />}
          <Field label="امضای تحویل‌گیرنده (لمسی)"><SignaturePad onChange={setSignature} /></Field>
        </>
      )}
      {err && <div className="text-sm text-rose-600 bg-rose-50 border border-rose-200 rounded-lg p-2 mb-3 mt-3">{err}</div>}
      <div className="mt-3"><BigButton onClick={submit} disabled={!canSubmit}>ثبت تحویل و آماده چاپ رسید</BigButton></div>
    </div>
  );
}

function OrdersList({ orders, itemById, onlyOpen }) {
  const sorted = [...orders].sort((a, b) => new Date(b.date) - new Date(a.date));
  const filtered = onlyOpen ? sorted.filter((o) => o.lines.some((l) => Number(l.qty) - Number(l.deliveredQty) > 0.001)) : sorted;
  if (filtered.length === 0) return <Empty text="سفارشی یافت نشد." />;
  return (
    <div className="space-y-3">
      {filtered.map((o) => (
        <div key={o.id} className="bg-white border border-stone-200 rounded-lg p-3.5">
          <div className="flex items-center justify-between mb-2">
            <div><span className="text-xs font-bold text-amber-800 bg-amber-50 border border-amber-200 rounded px-2 py-0.5">{o.invoiceNumber}</span> <span className="font-bold text-sm ml-1">{o.customerName}</span></div>
            <span className="text-[11px] text-stone-400">{fmtDate(o.date)}</span>
          </div>
          <table className="w-full text-xs">
            <thead><tr className="text-stone-400"><th className="text-right py-1">کالا</th><th className="text-right py-1">سفارش</th><th className="text-right py-1">تحویل‌شده</th><th className="text-right py-1">باقیمانده</th></tr></thead>
            <tbody>
              {o.lines.map((l) => {
                const it = itemById(l.itemId);
                const remaining = Number(l.qty) - Number(l.deliveredQty);
                return (
                  <tr key={l.id} className="border-t border-stone-100">
                    <td className="py-1.5">{it?.code} — {it?.design}</td><td className="py-1.5">{fmt(l.qty)}</td><td className="py-1.5">{fmt(l.deliveredQty)}</td>
                    <td className={`py-1.5 font-bold ${remaining > 0.001 ? "text-rose-600" : "text-emerald-600"}`}>{remaining > 0.001 ? fmt(remaining) : "تکمیل شده"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه محل کالا                                                       */
/* ------------------------------------------------------------------ */

function LocationsPage({ items, stock, itemById, onToggleDiscrepancy }) {
  const [filter, setFilter] = useState("all");
  const rows = filter === "all" ? stock : stock.filter((s) => s.warehouse === filter);
  return (
    <div>
      <SectionTitle info="محل دقیق هر کالا در قفسه (راهرو/ردیف/خانه) و وضعیت مغایرت هر محل." action={<div className="flex bg-white border border-stone-300 rounded-lg overflow-hidden text-xs">{[{ id: "all", label: "همه" }, { id: "w1", label: "انبار ۱" }, { id: "w2", label: "انبار ۲" }].map((f) => (<button key={f.id} onClick={() => setFilter(f.id)} className={`px-3 py-1.5 font-bold ${filter === f.id ? "bg-amber-700 text-white" : "text-stone-600 hover:bg-stone-50"}`}>{f.label}</button>))}</div>}>
        محل دقیق کالاها
      </SectionTitle>
      {rows.length === 0 ? <Empty text="رکوردی یافت نشد." /> : (
        <div className="bg-white border border-stone-200 rounded-lg overflow-x-auto">
          <table className="w-full text-sm min-w-[680px]">
            <thead><tr className="bg-stone-50 text-stone-500 text-xs"><th className="text-right px-3 py-2 font-bold">کالا</th><th className="text-right px-3 py-2 font-bold">انبار</th><th className="text-right px-3 py-2 font-bold">محل</th><th className="text-right px-3 py-2 font-bold">موجودی</th><th className="text-right px-3 py-2 font-bold">رزروشده</th><th className="text-right px-3 py-2 font-bold">وضعیت</th></tr></thead>
            <tbody className="divide-y divide-stone-100">
              {rows.map((r) => {
                const it = itemById(r.itemId);
                return (
                  <tr key={r.id}>
                    <td className="px-3 py-2"><div className="font-bold">{it?.code || r.itemId}</div><div className="text-xs text-stone-500">{it?.design}</div></td>
                    <td className="px-3 py-2">{whLabel(r.warehouse)}</td><td className="px-3 py-2 font-mono text-xs">{locLabel(r)}</td>
                    <td className="px-3 py-2">{qtyDisplay(it, r.qty, r.boxCount)}</td><td className="px-3 py-2">{fmt(r.reserved || 0)}</td>
                    <td className="px-3 py-2"><button onClick={() => onToggleDiscrepancy(r.id)} className="inline-flex items-center gap-1" title="علامت‌گذاری دستی مغایرت">{r.discrepancy ? <DiscrepancyBadge /> : <span className="text-[11px] text-stone-400 border border-stone-200 rounded px-1.5 py-0.5">بدون مغایرت</span>}</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه تاریخچه عملیات                                                 */
/* ------------------------------------------------------------------ */

function EditStockInModal({ op, item, onSave, onClose }) {
  const cartonMode = isCartonItem(item);
  const [boxCount, setBoxCount] = useState(op.boxCount != null ? String(op.boxCount) : "");
  const [plainQty, setPlainQty] = useState(cartonMode ? "" : String(op.qty));
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  return (
    <Modal title={`ویرایش ورود کالا — ${op.opNumber}`} onClose={onClose}>
      <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={null} label="مقدار اصلاح‌شده" />
      <BigButton onClick={() => onSave({ qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null })} disabled={!(q > 0)}>ذخیره اصلاح</BigButton>
    </Modal>
  );
}
function EditDeliveryModal({ delivery, item, onSave, onClose }) {
  const [customerName, setCustomerName] = useState(delivery.customerName);
  const [invoiceNumber, setInvoiceNumber] = useState(delivery.invoiceNumber);
  const cartonMode = isCartonItem(item);
  const [boxCount, setBoxCount] = useState(delivery.boxCount != null ? String(delivery.boxCount) : "");
  const [plainQty, setPlainQty] = useState(cartonMode ? "" : String(delivery.qty));
  const q = cartonMode ? Number(boxCount || 0) * Number(item.sqmPerBox || 0) : Number(plainQty || 0);
  return (
    <Modal title={`ویرایش تحویل — ${delivery.deliveryNumber}`} onClose={onClose}>
      <Field label="نام مشتری"><input className={inputCls} value={customerName} onChange={(e) => setCustomerName(e.target.value)} /></Field>
      <Field label="شماره فاکتور"><input className={inputCls} value={invoiceNumber} onChange={(e) => setInvoiceNumber(e.target.value)} /></Field>
      <QuantityInput item={item} boxCount={boxCount} setBoxCount={setBoxCount} plainQty={plainQty} setPlainQty={setPlainQty} maxAvailable={null} label="مقدار اصلاح‌شده" />
      <BigButton onClick={() => onSave({ customerName, invoiceNumber, qty: q, boxCount: cartonMode ? Number(boxCount || 0) : null })} disabled={!(q > 0)}>ذخیره اصلاح</BigButton>
    </Modal>
  );
}

function HistoryPage({ ops, deliveries, itemById, items, isManager, editStockIn, editDelivery, showToast }) {
  const [q, setQ] = useState("");
  const [editTarget, setEditTarget] = useState(null);
  const merged = [...ops.map((o) => ({ ...o, _kind: "op" })), ...deliveries.map((d) => ({ ...d, _kind: "delivery", opNumber: d.deliveryNumber }))].sort((a, b) => new Date(b.date) - new Date(a.date));
  const filtered = merged.filter((row) => { if (!q) return true; const it = itemById(row.itemId); return `${row.opNumber} ${it?.code || ""} ${it?.design || ""} ${row.user || ""} ${row.customerName || ""} ${row.description || ""}`.includes(q); });

  return (
    <div>
      <SectionTitle info="تمام عملیات‌هایی که موجودی را تغییر داده‌اند، به‌ترتیب زمان. برای مدیر امکان اصلاح ورود کالا و تحویل مشتری هم هست.">تاریخچه عملیات</SectionTitle>
      <div className="relative mb-3"><Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400" /><input className={inputCls + " pr-9"} placeholder="جستجو…" value={q} onChange={(e) => setQ(e.target.value)} /></div>

      {filtered.length === 0 ? <Empty text="عملیاتی ثبت نشده است." /> : (
        <div className="space-y-2">
          {filtered.map((row) => {
            const it = itemById(row.itemId);
            const type = row._kind === "delivery" ? "delivery" : row.type;
            const canEditRow = isManager && (type === "stockin" || type === "delivery");
            return (
              <div key={row.id} className="bg-white border border-stone-200 rounded-lg p-3.5">
                <div className="flex items-center justify-between mb-1.5">
                  <div className="flex items-center gap-2"><span className="text-xs font-bold text-amber-800 bg-amber-50 border border-amber-200 rounded px-2 py-0.5">{row.opNumber}</span><TypeBadge type={type} /></div>
                  <div className="flex items-center gap-2">
                    {canEditRow && <button onClick={() => setEditTarget({ row, type })} className="text-stone-400 hover:text-amber-800" title="ویرایش (مدیر)"><Pencil size={13} /></button>}
                    <span className="text-[11px] text-stone-400">{fmtDate(row.date)}</span>
                  </div>
                </div>
                <div className="font-bold text-stone-900 text-sm">{it ? `${it.code} — ${it.design}` : row.itemId}</div>
                {type === "transfer" && <div className="flex items-center gap-2 text-xs text-stone-600 mt-1"><span>{whLabel(row.fromWarehouse)}</span><span className="font-mono text-stone-400">{locLabel(row.fromLocation)}</span><ArrowLeftRight size={12} className="text-stone-300" /><span>{whLabel(row.toWarehouse)}</span><span className="font-mono text-stone-400">{locLabel(row.toLocation)}</span></div>}
                {type === "stockin" && <div className="text-xs text-stone-600 mt-1">به {whLabel(row.warehouse)} / {locLabel(row.location)}{row.supplier ? ` · از ${row.supplier}` : ""}</div>}
                {type === "adjustment" && <div className="text-xs text-stone-600 mt-1">{whLabel(row.warehouse)} / {locLabel(row.location)} · دلیل: {row.reason}</div>}
                {type === "delivery" && <div className="text-xs text-stone-600 mt-1">مشتری: {row.customerName} · فاکتور: {row.invoiceNumber} · از {whLabel(row.warehouse)} / {locLabel(row.location)}</div>}
                <div className="flex items-center justify-between mt-2 text-xs"><span className="text-stone-500">مقدار: <b className="text-stone-800">{qtyDisplay(it, row.qty, row.boxCount)}</b></span><span className="text-stone-400">کاربر: {row.user}</span></div>
                {row.description && <div className="text-xs text-stone-500 mt-1.5 border-t border-stone-100 pt-1.5">{row.description}</div>}
                {row.editedNote && <div className="text-xs text-amber-700 mt-1.5 border-t border-stone-100 pt-1.5">{row.editedNote}</div>}
              </div>
            );
          })}
        </div>
      )}

      {editTarget && editTarget.type === "stockin" && (
        <EditStockInModal op={editTarget.row} item={itemById(editTarget.row.itemId)} onClose={() => setEditTarget(null)} onSave={async (patch) => { const res = await editStockIn(editTarget.row.id, patch); if (!res.ok) showToast(res.msg, "error"); setEditTarget(null); }} />
      )}
      {editTarget && editTarget.type === "delivery" && (
        <EditDeliveryModal delivery={editTarget.row} item={itemById(editTarget.row.itemId)} onClose={() => setEditTarget(null)} onSave={async (patch) => { const res = await editDelivery(editTarget.row.id, patch); if (!res.ok) showToast(res.msg, "error"); setEditTarget(null); }} />
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* صفحه تنظیمات                                                        */
/* ------------------------------------------------------------------ */

function StaffNameEditor({ staff, onRename }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(staff.name);
  if (editing) {
    return (
      <div className="flex items-center gap-1.5">
        <input className={inputCls + " py-1 text-xs w-36"} value={val} onChange={(e) => setVal(e.target.value)} />
        <button onClick={() => { onRename(staff.id, val.trim() || staff.name); setEditing(false); }} className="text-emerald-600"><Check size={15} /></button>
        <button onClick={() => { setVal(staff.name); setEditing(false); }} className="text-stone-400"><X size={15} /></button>
      </div>
    );
  }
  return (
    <button onClick={() => setEditing(true)} className="font-bold inline-flex items-center gap-1.5 hover:text-amber-800">
      {staff.name} <Pencil size={11} className="opacity-50" />
    </button>
  );
}

function SettingsPage({ staff, onRename, onToggleManager, onRemoveStaff, exportExcel, isManager }) {
  return (
    <div>
      <SectionTitle info="مدیریت کارکنانی که وارد سیستم شده‌اند و دانلود نسخه پشتیبان از کل اطلاعات.">تنظیمات</SectionTitle>

      <SectionTitle info="هر کارمند با ایمیل و رمز خودش وارد می‌شود. برای افزودن کارمند جدید باید یک حساب کاربری در پنل Supabase (بخش Authentication → Users) ساخته شود؛ در اولین ورود، نام و دسترسی او خودکار اینجا اضافه می‌شود.">کارکنان</SectionTitle>
      {staff.length === 0 ? <Empty text="هنوز کارمندی وارد سیستم نشده است." /> : (
        <div className="bg-white border border-stone-200 rounded-lg divide-y divide-stone-100 mb-3">
          {staff.map((s) => (
            <div key={s.id} className="flex items-center justify-between px-4 py-2.5 text-sm gap-2">
              <StaffNameEditor staff={s} onRename={onRename} />
              <div className="flex items-center gap-2">
                {s.isManager && <span className="text-[11px] bg-amber-50 text-amber-800 border border-amber-200 rounded px-1.5 py-0.5">مدیر</span>}
                <button onClick={() => onToggleManager(s.id, !s.isManager)} className="text-[11px] font-bold text-stone-500 border border-stone-300 rounded-lg px-2 py-1">{s.isManager ? "حذف نقش مدیر" : "تبدیل به مدیر"}</button>
                <IconButton icon={Trash2} tone="danger" title="حذف از لیست (حساب ورود او حذف نمی‌شود)" onClick={() => onRemoveStaff(s.id)} />
              </div>
            </div>
          ))}
        </div>
      )}
      <div className="text-xs text-stone-400 mb-6 bg-stone-50 border border-stone-200 rounded-lg p-3">
        برای افزودن کارمند جدید: در پنل Supabase بخش Authentication → Users روی «Add user» بزنید، یک ایمیل و رمز دلخواه وارد کنید و تیک «Auto Confirm User» را فعال کنید. کارمند با همان ایمیل و رمز وارد اپلیکیشن می‌شود و خودش در همین لیست ظاهر خواهد شد.
      </div>

      <SectionTitle info="یک فایل اکسل شامل تمام کالاها، موجودی، عملیات، سفارش‌ها و تحویل‌ها دانلود می‌شود — پشتیبان دستی مستقل از سیستم.">پشتیبان‌گیری</SectionTitle>
      <BigButton onClick={exportExcel}><span className="inline-flex items-center gap-1"><Download size={16} /> دانلود اکسل کامل</span></BigButton>
    </div>
  );
}
