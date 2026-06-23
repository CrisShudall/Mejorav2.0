import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LogIn, LogOut, Plus, Search, Upload, Download, Users as UsersIcon,
  Truck, FileText, Bell, Clock, MapPin, X, MessageSquare, Trash2,
  ShieldCheck, RotateCcw, ChevronRight, AlertTriangle, CheckCircle2,
  CircleDot, XCircle, Building2, UserPlus, Sun, Moon
} from "lucide-react";

/* ============================================================
   CONFIG / CONSTANTES
   ============================================================ */

const KEYS = {
  USERS: "sgs_users_v1",
  PROVIDERS: "sgs_providers_v1",
  SINIESTROS: "sgs_siniestros_v1",
  AUTOSUST: "sgs_autosustituto_v1",
};

const ESTADOS_MX = [
  "Aguascalientes","Baja California","Baja California Sur","Campeche","Chiapas",
  "Chihuahua","Ciudad de México","Coahuila","Colima","Durango","Guanajuato",
  "Guerrero","Hidalgo","Jalisco","Estado de México","Michoacán","Morelos",
  "Nayarit","Nuevo León","Oaxaca","Puebla","Querétaro","Quintana Roo",
  "San Luis Potosí","Sinaloa","Sonora","Tabasco","Tamaulipas","Tlaxcala",
  "Veracruz","Yucatán","Zacatecas"
];

const ESTATUS_SIN = ["abierto", "asignado", "terminado", "cancelado"];
const ESTATUS_LABEL = {
  abierto: "Abierto", asignado: "Asignado", terminado: "Terminado", cancelado: "Cancelado",
};
const ESTATUS_ICON = {
  abierto: CircleDot, asignado: ShieldCheck, terminado: CheckCircle2, cancelado: XCircle,
};

function uid(p) {
  return p + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function nowISO() { return new Date().toISOString(); }
function fmtDT(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso);
  return d.toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" }) +
    " · " + d.toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" });
}
function toLocalInputValue(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtMoney(n) {
  const v = Number(n);
  if (isNaN(v)) return "—";
  return v.toLocaleString("es-MX", { style: "currency", currency: "MXN" });
}
function normKey(s) {
  return String(s || "").trim().toLowerCase();
}

const SEED_USERS = [
  { id: uid("usr"), username: "admin", password: "admin123", role: "admin", nombre: "Administrador" },
  { id: uid("usr"), username: "usuario1", password: "usuario123", role: "usuario", nombre: "Usuario Operativo" },
];

/* ============================================================
   STORAGE HELPERS
   ============================================================ */

async function loadList(key, seed) {
  try {
    const res = await window.storage.get(key, true);
    if (res && res.value) return JSON.parse(res.value);
    return seed || [];
  } catch (e) {
    if (seed) {
      try { await window.storage.set(key, JSON.stringify(seed), true); } catch (e2) {}
      return seed;
    }
    return [];
  }
}
async function saveList(key, list) {
  try {
    await window.storage.set(key, JSON.stringify(list), true);
    return true;
  } catch (e) {
    return false;
  }
}

/* ============================================================
   ALERTAS DE TIEMPO
   ============================================================ */

function siniestroAlert(rep) {
  if (rep.estatus === "terminado") return { label: "Terminado", level: "done" };
  if (rep.estatus === "cancelado") return { label: "Cancelado", level: "off" };
  if (!rep.etaArribo) return { label: "Sin ETA", level: "off" };
  const diffMin = (new Date(rep.etaArribo).getTime() - Date.now()) / 60000;
  if (diffMin < 0) return { label: `Vencido (${Math.abs(Math.round(diffMin))} min)`, level: "danger" };
  if (diffMin <= 45) return { label: `Por vencer (${Math.round(diffMin)} min)`, level: "warn" };
  return { label: "En tiempo", level: "ok" };
}
function autoSustAlert(rep) {
  if (rep.estatus === "terminado") return { label: "Terminado", level: "done" };
  if (rep.estatus === "cancelado") return { label: "Cancelado", level: "off" };
  const limite = new Date(rep.fechaCreacion).getTime() + 24 * 3600 * 1000;
  const diffMin = (limite - Date.now()) / 60000;
  if (diffMin < 0) return { label: `Vencido (${Math.abs(Math.round(diffMin / 60))} h)`, level: "danger" };
  if (diffMin <= 120) return { label: `Por vencer (${Math.round(diffMin)} min)`, level: "warn" };
  return { label: "En tiempo", level: "ok" };
}

const LEVEL_STYLE = {
  ok: { bg: "var(--ok-bg)", fg: "var(--ok-fg)" },
  warn: { bg: "var(--warn-bg)", fg: "var(--warn-fg)" },
  danger: { bg: "var(--danger-bg)", fg: "var(--danger-fg)" },
  done: { bg: "var(--done-bg)", fg: "var(--done-fg)" },
  off: { bg: "var(--off-bg)", fg: "var(--off-fg)" },
};

function AlertBadge({ alert }) {
  const s = LEVEL_STYLE[alert.level];
  return (
    <span className="badge-pill" style={{ background: s.bg, color: s.fg }}>
      {alert.level === "danger" && <AlertTriangle size={12} />}
      {alert.level === "warn" && <Clock size={12} />}
      {alert.level === "ok" && <CheckCircle2 size={12} />}
      {alert.label}
    </span>
  );
}
function EstatusBadge({ estatus }) {
  const Icon = ESTATUS_ICON[estatus] || CircleDot;
  return (
    <span className={`badge-pill estatus-${estatus}`}>
      <Icon size={12} /> {ESTATUS_LABEL[estatus] || estatus}
    </span>
  );
}

/* ============================================================
   STYLE SHEET (token system)
   ============================================================ */

function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');
      *, *::before, *::after { box-sizing: border-box; }

      /* ── Light tokens ─────────────────────────────── */
      .sgs-root {
        --canvas:       #dde1e9;
        --canvas-grad:  radial-gradient(ellipse at 30% 20%, #d4d9e8 0%, #dde1e9 60%);
        --glass-sidebar:rgba(235,238,245,0.72);
        --glass-frame:  rgba(255,255,255,0.82);
        --glass-card:   rgba(255,255,255,0.78);
        --glass-modal:  rgba(255,255,255,0.90);
        --glass-pill:   rgba(255,255,255,0.96);
        --glass-border: rgba(255,255,255,0.80);
        --glass-border2:rgba(200,205,218,0.60);
        --glass-hi:     rgba(255,255,255,0.95);
        --blur:         blur(22px) saturate(180%);
        --shadow-sidebar:0 8px 40px rgba(0,0,0,0.10), inset 0 1px 0 rgba(255,255,255,0.9);
        --shadow-frame:  0 4px 24px rgba(0,0,0,0.07), 0 1px 2px rgba(0,0,0,0.04);
        --shadow-pill:
          0 2px 8px rgba(0,0,0,0.10),
          0 0.5px 0 rgba(0,0,0,0.06),
          inset 0 1.5px 0 rgba(255,255,255,1),
          inset 0 -1px 0 rgba(0,0,0,0.04);
        --pill-active-bg: rgba(255,255,255,0.82);
        --pill-hover-bg:  rgba(255,255,255,0.50);
        --pill-base-bg:   rgba(255,255,255,0.20);
        --pill-base-border: rgba(255,255,255,0.55);
        --pill-hover-border:rgba(255,255,255,0.80);
        --shadow-card:   0 2px 12px rgba(0,0,0,0.06);
        --ink:          #111827;
        --ink-mid:      #374151;
        --ink-soft:     #6b7280;
        --ink-faint:    #9ca3af;
        --accent:       #1c1c1e;
        --accent-fg:    #ffffff;
        --accent-soft:  rgba(28,28,30,0.07);
        --line:         rgba(0,0,0,0.08);
        --line-soft:    rgba(0,0,0,0.05);
        --ok-bg:#ecfdf5;     --ok-fg:#15803d;
        --warn-bg:#fffbeb;   --warn-fg:#92400e;
        --danger-bg:#fef2f2; --danger-fg:#b91c1c;
        --done-bg:#f3f4f6;   --done-fg:#4b5563;
        --off-bg:#f9fafb;    --off-fg:#9ca3af;
        --input-bg: rgba(255,255,255,0.70);
        --card-2:   rgba(248,249,252,0.80);
      }

      /* ── Dark tokens ──────────────────────────────── */
      .sgs-root.dark {
        --canvas:       #0d0f14;
        --canvas-grad:  radial-gradient(ellipse at 30% 20%, #141824 0%, #0d0f14 65%);
        --glass-sidebar:rgba(18,22,34,0.70);
        --glass-frame:  rgba(22,27,42,0.78);
        --glass-card:   rgba(26,32,48,0.75);
        --glass-modal:  rgba(20,25,40,0.90);
        --glass-pill:   rgba(255,255,255,0.10);
        --glass-border: rgba(255,255,255,0.10);
        --glass-border2:rgba(255,255,255,0.06);
        --glass-hi:     rgba(255,255,255,0.07);
        --shadow-sidebar:0 8px 40px rgba(0,0,0,0.40), inset 0 0 0 1px rgba(255,255,255,0.05);
        --shadow-frame:  0 4px 24px rgba(0,0,0,0.35), 0 1px 2px rgba(0,0,0,0.20);
        --shadow-pill:
          0 2px 10px rgba(0,0,0,0.35),
          0 0.5px 0 rgba(0,0,0,0.25),
          inset 0 1px 0 rgba(255,255,255,0.07),
          inset 0 -1px 0 rgba(0,0,0,0.12);
        --pill-active-bg: rgba(255,255,255,0.09);
        --pill-hover-bg:  rgba(255,255,255,0.05);
        --pill-base-bg:   rgba(255,255,255,0.02);
        --pill-base-border: rgba(255,255,255,0.07);
        --pill-hover-border:rgba(255,255,255,0.11);
        --shadow-card:   0 2px 12px rgba(0,0,0,0.30);
        --ink:          #f0f2f7;
        --ink-mid:      #c8cdd8;
        --ink-soft:     #8b93a8;
        --ink-faint:    #5c6478;
        --accent:       #f0f2f7;
        --accent-fg:    #111827;
        --accent-soft:  rgba(255,255,255,0.08);
        --line:         rgba(255,255,255,0.08);
        --line-soft:    rgba(255,255,255,0.05);
        --ok-bg:rgba(21,128,61,0.18);    --ok-fg:#4ade80;
        --warn-bg:rgba(146,64,14,0.20);  --warn-fg:#fbbf24;
        --danger-bg:rgba(185,28,28,0.18);--danger-fg:#f87171;
        --done-bg:rgba(75,85,99,0.25);   --done-fg:#9ca3af;
        --off-bg:rgba(255,255,255,0.05); --off-fg:#4b5563;
        --input-bg: rgba(255,255,255,0.06);
        --card-2:   rgba(255,255,255,0.04);
        --filter-active-bg:  rgba(255,255,255,0.11);
        --filter-active-fg:  var(--ink);
        --filter-active-border: rgba(255,255,255,0.14);
        --filter-idle-fg:    var(--ink-soft);
      }

      .sgs-root {
        font-family:'Inter',ui-sans-serif,system-ui,sans-serif;
        font-size:14px; line-height:1.5; color:var(--ink);
        background:var(--canvas-grad);
        min-height:100vh;
        transition: background 0.3s, color 0.3s;
      }
      .mono { font-family:'JetBrains Mono',ui-monospace,monospace !important; }

      /* ── Liquid Glass mixin ───────────────────────── */
      .glass {
        background: var(--glass-card);
        backdrop-filter: var(--blur);
        -webkit-backdrop-filter: var(--blur);
        border: 1px solid var(--glass-border);
        box-shadow: var(--shadow-card);
      }

      /* ── Badges ───────────────────────────────────── */
      .badge-pill {
        display:inline-flex; align-items:center; gap:3px;
        padding:2px 8px; border-radius:5px;
        font-size:11px; font-weight:500; white-space:nowrap;
      }
      .estatus-abierto  { background:#eff6ff; color:#1d4ed8; }
      .estatus-asignado { background:#f5f3ff; color:#6d28d9; }
      .estatus-terminado{ background:var(--done-bg); color:var(--done-fg); }
      .estatus-cancelado{ background:#fef2f2; color:#991b1b; }
      .dark .estatus-abierto  { background:rgba(29,78,216,0.20); color:#93c5fd; }
      .dark .estatus-asignado { background:rgba(109,40,217,0.20); color:#c4b5fd; }
      .dark .estatus-terminado{ background:rgba(75,85,99,0.25); color:#9ca3af; }
      .dark .estatus-cancelado{ background:rgba(153,27,27,0.22); color:#fca5a5; }

      /* ── Buttons ──────────────────────────────────── */
      .btn {
        display:inline-flex; align-items:center; gap:6px;
        padding:7px 13px; border-radius:9px; font-size:13px; font-weight:500;
        cursor:pointer; border:1px solid var(--glass-border2);
        transition:all .15s ease; white-space:nowrap;
        backdrop-filter: var(--blur);
      }
      .btn:active { opacity:.88; transform:translateY(1px); }
      .btn-primary {
        background:var(--accent); color:var(--accent-fg);
        border-color:transparent;
        box-shadow:0 1px 3px rgba(0,0,0,0.15), inset 0 1px 0 rgba(255,255,255,0.12);
      }
      .dark .btn-primary {
        box-shadow:0 1px 4px rgba(0,0,0,0.30), inset 0 1px 0 rgba(255,255,255,0.05);
      }
      .btn-primary:hover { opacity:.90; }
      .btn-outline {
        background:var(--glass-pill); color:var(--ink-mid);
        box-shadow:var(--shadow-pill);
      }
      .btn-outline:hover { background:var(--glass-hi); }
      .btn-ghost { background:transparent; color:var(--ink-soft); border-color:transparent; }
      .btn-ghost:hover { background:var(--accent-soft); color:var(--ink); }
      .btn-danger { background:var(--danger-fg); color:#fff; border-color:transparent; }
      .btn-sm { padding:4px 9px; font-size:12px; border-radius:7px; }
      .btn:disabled { opacity:.35; cursor:not-allowed; }

      /* ── Fields ───────────────────────────────────── */
      .field-label {
        font-size:11px; font-weight:600; text-transform:uppercase;
        letter-spacing:.05em; color:var(--ink-faint); margin-bottom:4px; display:block;
      }
      .input, select.input, textarea.input {
        width:100%; border:1px solid var(--glass-border2); border-radius:8px;
        padding:8px 11px; font-size:13.5px; color:var(--ink);
        background:var(--input-bg);
        backdrop-filter:var(--blur);
        outline:none; font-family:inherit;
        transition:border-color .15s, box-shadow .15s, background .3s;
      }
      .input::placeholder { color:var(--ink-faint); }
      .input:focus {
        border-color:var(--accent);
        box-shadow:0 0 0 3px var(--accent-soft);
        background:var(--glass-pill);
      }
      select.input { appearance:auto; }

      /* ── Card ─────────────────────────────────────── */
      .card {
        background:var(--glass-card);
        backdrop-filter:var(--blur);
        -webkit-backdrop-filter:var(--blur);
        border:1px solid var(--glass-border);
        border-radius:12px;
        box-shadow:var(--shadow-card);
      }

      /* ── Sidebar nav link ─────────────────────────── */
      /* ── Nav links — Apple Liquid Glass pills ────── */
      .nav-link {
        display:flex; align-items:center; gap:9px;
        padding:9px 13px; border-radius:999px;
        color:var(--ink-soft); font-size:13.5px; font-weight:500;
        cursor:pointer; user-select:none;
        border:1px solid transparent;
        position:relative; overflow:hidden;
        /* base pill — barely visible glass */
        background: var(--pill-base-bg);
        backdrop-filter: blur(12px) saturate(160%);
        -webkit-backdrop-filter: blur(12px) saturate(160%);
        border-color: var(--pill-base-border);
        transition:
          background     0.18s ease,
          box-shadow     0.18s ease,
          border-color   0.18s ease,
          color          0.18s ease;
      }
      /* Specular top-edge highlight via ::before */
      .nav-link::before {
        content:"";
        position:absolute; inset:0;
        border-radius:inherit;
        /* Specular highlight — luz blanca suave en modo claro */
        background: linear-gradient(
          180deg,
          rgba(255,255,255,0.52) 0%,
          rgba(255,255,255,0.00) 50%
        );
        pointer-events:none;
        opacity:0;
        transition: opacity 0.18s ease;
      }
      /* Modo oscuro: reflejo mucho más tenue y cálido, no blanco puro */
      .dark .nav-link::before {
        background: linear-gradient(
          180deg,
          rgba(255,255,255,0.06) 0%,
          rgba(255,255,255,0.00) 60%
        );
      }
      .nav-link:hover {
        background: var(--pill-hover-bg);
        border-color: var(--pill-hover-border);
        color:var(--ink);
        box-shadow:
          0 1px 6px rgba(0,0,0,0.08),
          inset 0 1px 0 rgba(255,255,255,0.55);
      }
      .dark .nav-link:hover {
        box-shadow:
          0 1px 8px rgba(0,0,0,0.25),
          inset 0 1px 0 rgba(255,255,255,0.05);
      }
      .nav-link:hover::before { opacity:1; }
      .nav-link.active {
        background: var(--pill-active-bg);
        backdrop-filter: blur(24px) saturate(200%);
        -webkit-backdrop-filter: blur(24px) saturate(200%);
        border-color: var(--pill-hover-border);
        box-shadow: var(--shadow-pill);
        color:var(--ink); font-weight:600;
      }
      .dark .nav-link.active {
        backdrop-filter: blur(20px) saturate(130%);
        -webkit-backdrop-filter: blur(20px) saturate(130%);
      }
      .nav-link.active::before { opacity:1; }

      /* ── Modal ────────────────────────────────────── */
      .modal-overlay {
        position:fixed; inset:0;
        background:rgba(0,0,0,0.35);
        backdrop-filter:blur(6px);
        display:flex; align-items:center; justify-content:center;
        z-index:50; padding:20px;
      }
      .modal-panel {
        background:var(--glass-modal);
        backdrop-filter:var(--blur);
        -webkit-backdrop-filter:var(--blur);
        border:1px solid var(--glass-border);
        border-radius:18px; width:100%; max-width:640px;
        max-height:88vh; overflow-y:auto;
        box-shadow:0 24px 60px rgba(0,0,0,0.20), inset 0 1px 0 var(--glass-hi);
      }

      /* ── Table ────────────────────────────────────── */
      table.sgs-table { width:100%; border-collapse:collapse; font-size:13px; }
      table.sgs-table th {
        text-align:left; font-size:10.5px; text-transform:uppercase; letter-spacing:.06em;
        color:var(--ink-faint); padding:9px 14px;
        border-bottom:1px solid var(--line);
        background:var(--card-2); position:sticky; top:0; font-weight:600;
      }
      table.sgs-table td {
        padding:10px 14px; border-bottom:1px solid var(--line-soft);
        vertical-align:middle; color:var(--ink-mid);
        transition:background .12s;
      }
      table.sgs-table tr:last-child td { border-bottom:none; }
      table.sgs-table tbody tr:hover td { background:var(--card-2); }

      /* ── Stat card ────────────────────────────────── */
      .stat-card {
        background:var(--glass-card);
        backdrop-filter:var(--blur);
        -webkit-backdrop-filter:var(--blur);
        border:1px solid var(--glass-border);
        border-radius:12px; padding:16px 18px;
        box-shadow:var(--shadow-card);
      }

      /* ── Avatar ───────────────────────────────────── */
      .avatar-circle {
        border-radius:50%; background:var(--accent); color:var(--accent-fg);
        display:flex; align-items:center; justify-content:center;
        font-weight:600; flex-shrink:0;
        box-shadow:0 2px 6px rgba(0,0,0,0.15);
      }

      /* ── Theme transition — quirúrgica, sin * global ─ */
      .sgs-root,
      .sgs-root .card, .sgs-root .stat-card, .sgs-root .glass,
      .sgs-root .nav-link, .sgs-root .btn,
      .sgs-root .input,
      .sgs-root .badge-pill,
      .sgs-root aside,
      .sgs-root table.sgs-table td,
      .sgs-root table.sgs-table th {
        transition:
          color          0.22s ease,
          background     0.22s ease,
          border-color   0.22s ease,
          box-shadow     0.22s ease;
      }
      /* backdrop-filter no se anima — demasiado costoso */

      /* ── Filter pills ─────────────────────────────── */
      .filter-pill {
        background: transparent;
        color: var(--filter-idle-fg);
        border: 1px solid transparent;
        border-radius: 999px;
        transition: background 0.16s ease, color 0.16s ease, border-color 0.16s ease, box-shadow 0.16s ease;
      }
      .filter-pill:hover {
        background: var(--pill-hover-bg);
        color: var(--ink);
        border-color: var(--pill-base-border);
      }
      .filter-pill.active {
        background: var(--filter-active-bg);
        color: var(--filter-active-fg);
        border-color: var(--filter-active-border);
        box-shadow: 0 1px 4px rgba(0,0,0,0.12), inset 0 1px 0 rgba(255,255,255,0.08);
        font-weight: 600;
      }
      .dark .filter-pill.active {
        box-shadow: 0 1px 6px rgba(0,0,0,0.30), inset 0 1px 0 rgba(255,255,255,0.05);
      }

      ::-webkit-scrollbar { width:8px; }
      ::-webkit-scrollbar-track { background:transparent; }
      ::-webkit-scrollbar-thumb { background:var(--glass-border2); border-radius:8px; }

      /* ── Page transition animations ─────────────────── */
      @keyframes pageEnter {
        0%   { opacity:0; transform: translateY(14px) scale(0.992); }
        100% { opacity:1; transform: translateY(0px)  scale(1); }
      }
      @keyframes pageExit {
        0%   { opacity:1; transform: translateY(0px)  scale(1); }
        100% { opacity:0; transform: translateY(-10px) scale(0.994); }
      }
      .page-enter {
        animation: pageEnter 0.30s cubic-bezier(0.22, 0.61, 0.36, 1) both;
        will-change: opacity, transform;
      }
      .page-exit {
        animation: pageExit 0.18s cubic-bezier(0.55, 0, 1, 0.45) both;
        will-change: opacity, transform;
        pointer-events: none;
      }
    `}</style>
  );
}


/* ============================================================
   LOGIN
   ============================================================ */

function LoginView({ users, onLogin, onResetUsers, dark, toggleDark }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const [err, setErr] = useState("");
  const [resetDone, setResetDone] = useState(false);
  function submit() {
    // Los usuarios de prueba (SEED_USERS) siempre funcionan como respaldo,
    // sin importar el estado del almacenamiento, para que nunca te quedes sin acceso.
    const pool = [...users];
    SEED_USERS.forEach((su) => {
      if (!pool.some((x) => normKey(x.username) === normKey(su.username))) pool.push(su);
    });
    const found = pool.find(
      (x) => normKey(x.username) === normKey(u) && x.password.trim() === p.trim()
    );
    if (!found) { setErr("Usuario o contraseña incorrectos."); return; }
    setErr("");
    onLogin(found);
  }
  function reset() {
    onResetUsers();
    setResetDone(true);
    setErr("");
  }
  return (
    <div style={{ display:"flex", minHeight:"100vh", alignItems:"center", justifyContent:"center", background:"var(--canvas-grad)" }}>
      <div style={{ position:"absolute", top:18, right:22 }}>
        <button className="btn btn-ghost btn-sm" onClick={toggleDark} style={{ borderRadius:99, padding:"7px 10px" }}>
          {dark ? <Sun size={15} /> : <Moon size={15} />}
        </button>
      </div>
      <div style={{
        width:380, padding:"32px 28px",
        background:"var(--glass-modal)",
        backdropFilter:"var(--blur)", WebkitBackdropFilter:"var(--blur)",
        border:"1px solid var(--glass-border)",
        borderRadius:20,
        boxShadow:"0 32px 80px rgba(0,0,0,0.15), inset 0 1px 0 var(--glass-hi)",
      }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:4 }}>
          <div style={{ width:38, height:38, borderRadius:12, background:"var(--accent)", color:"var(--accent-fg)", display:"flex", alignItems:"center", justifyContent:"center", boxShadow:"0 4px 12px rgba(0,0,0,0.18), inset 0 1px 0 rgba(255,255,255,0.15)" }}>
            <ShieldCheck size={20} color="currentColor" />
          </div>
          <div>
            <div style={{ fontWeight:700, fontSize:17, color:"var(--ink)", letterSpacing:"-0.01em" }}>Folios &amp; Asistencia</div>
            <div style={{ fontSize:11.5, color:"var(--ink-soft)" }}>Gestión de siniestros y auto sustituto</div>
          </div>
        </div>
        <div style={{ height: 18 }} />
        <label className="field-label">Usuario</label>
        <input className="input" value={u} onChange={(e) => setU(e.target.value)} autoFocus
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }} />
        <div style={{ height: 12 }} />
        <label className="field-label">Contraseña</label>
        <input className="input" type="password" value={p} onChange={(e) => setP(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }} />
        {err && <div style={{ color: "var(--danger-fg)", fontSize: 12.5, marginTop: 10 }}>{err}</div>}
        <button type="button" className="btn btn-primary" style={{ width: "100%", justifyContent: "center", marginTop: 18 }} onClick={submit}>
          <LogIn size={15} /> Entrar
        </button>
        <div style={{ marginTop: 16, fontSize: 11.5, color: "var(--ink-soft)", lineHeight: 1.5 }}>
          Acceso de prueba: <span className="mono">admin / admin123</span> (administrador)<br />
          <span className="mono">usuario1 / usuario123</span> (operativo)
        </div>
        <div style={{ marginTop: 8, fontSize: 10.5, color: "var(--ink-soft)" }}>
          Sistema cargado · {users.length} usuario(s) guardado(s)
        </div>
        <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px solid var(--line)" }}>
          {resetDone ? (
            <div style={{ fontSize: 11.5, color: "var(--ok-fg)" }}>Usuarios de prueba restablecidos. Intenta entrar de nuevo.</div>
          ) : (
            <button type="button" className="btn btn-ghost btn-sm" onClick={reset} style={{ padding: 0, fontSize: 11.5 }}>
              <RotateCcw size={12} /> ¿No puedes entrar? Restablecer usuarios de prueba
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   COMENTARIOS
   ============================================================ */

function Comments({ comentarios, onAdd }) {
  const [txt, setTxt] = useState("");
  return (
    <div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 220, overflowY: "auto", marginBottom: 10 }}>
        {(!comentarios || comentarios.length === 0) && (
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>Sin comentarios aún.</div>
        )}
        {(comentarios || []).slice().reverse().map((c) => (
          <div key={c.id} style={{ background: "var(--card-2)", border: "1px solid var(--line)", borderRadius: 8, padding: "8px 10px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "var(--ink-soft)", marginBottom: 3 }}>
              <span style={{ fontWeight: 700 }}>{c.usuario}</span>
              <span className="mono">{fmtDT(c.fecha)}</span>
            </div>
            <div style={{ fontSize: 13 }}>{c.texto}</div>
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <input className="input" placeholder="Agregar comentario de seguimiento…" value={txt}
          onChange={(e) => setTxt(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && txt.trim()) { onAdd(txt.trim()); setTxt(""); } }} />
        <button className="btn btn-primary btn-sm" onClick={() => { if (txt.trim()) { onAdd(txt.trim()); setTxt(""); } }}>
          <MessageSquare size={13} /> Agregar
        </button>
      </div>
    </div>
  );
}

/* ============================================================
   SINIESTROS — Crear / Detalle
   ============================================================ */

function SiniestroForm({ initial, providers, onCancel, onSave }) {
  const [f, setF] = useState(initial || {
    numeroSiniestro: "", numeroFolio: "", estatus: "abierto",
    origen: "", destino: "", costo: "", proveedor: "", etaArribo: "",
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onCancel(); }}>
      <div className="modal-panel" style={{ padding: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div className="display" style={{ fontWeight: 700, fontSize: 17 }}>
            {initial ? "Editar reporte de siniestro" : "Nuevo reporte de siniestro"}
          </div>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}><X size={16} /></button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div>
            <label className="field-label">Número de siniestro</label>
            <input className="input mono" value={f.numeroSiniestro} onChange={set("numeroSiniestro")} />
          </div>
          <div>
            <label className="field-label">Número de folio contable</label>
            <input className="input mono" value={f.numeroFolio} onChange={set("numeroFolio")} />
          </div>
          <div>
            <label className="field-label">Estatus</label>
            <select className="input" value={f.estatus} onChange={set("estatus")}>
              {ESTATUS_SIN.map((s) => <option key={s} value={s}>{ESTATUS_LABEL[s]}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Proveedor asignado</label>
            <select className="input" value={f.proveedor} onChange={set("proveedor")}>
              <option value="">— Sin asignar —</option>
              {providers.map((p) => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
            </select>
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label"><MapPin size={11} style={{ display: "inline", marginRight: 3 }} />Lugar de origen</label>
            <input className="input" placeholder="Dirección de origen (puede pegarse desde Maps)" value={f.origen} onChange={set("origen")} />
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label"><MapPin size={11} style={{ display: "inline", marginRight: 3 }} />Lugar de destino</label>
            <input className="input" placeholder="Dirección de destino (puede pegarse desde Maps)" value={f.destino} onChange={set("destino")} />
          </div>
          <div>
            <label className="field-label">Costo</label>
            <input className="input" type="number" step="0.01" placeholder="$0.00" value={f.costo} onChange={set("costo")} />
          </div>
          <div>
            <label className="field-label">ETA arribo (seguimiento preventivo)</label>
            <input className="input" type="datetime-local" value={toLocalInputValue(f.etaArribo) || f.etaArribo} onChange={(e) => setF({ ...f, etaArribo: e.target.value ? new Date(e.target.value).toISOString() : "" })} />
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 20 }}>
          <button className="btn btn-outline" onClick={onCancel}>Cancelar</button>
          <button className="btn btn-primary" onClick={() => onSave(f)} disabled={!f.numeroSiniestro.trim()}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

function SiniestroDetail({ rep, providers, currentUser, onClose, onUpdate, onAddComment }) {
  const [edit, setEdit] = useState(false);
  const alert = siniestroAlert(rep);
  if (edit) {
    return <SiniestroForm initial={rep} providers={providers} onCancel={() => setEdit(false)}
      onSave={(f) => { onUpdate(f); setEdit(false); }} />;
  }
  return (
    <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal-panel" style={{ padding: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div>
            <div className="display" style={{ fontWeight: 700, fontSize: 18 }}>Siniestro <span className="mono">{rep.numeroSiniestro}</span></div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>Folio contable: <span className="mono">{rep.numeroFolio || "—"}</span></div>
          </div>
          <button className="btn btn-ghost btn-sm" onClick={onClose}><X size={16} /></button>
        </div>
        <div style={{ display: "flex", gap: 8, margin: "10px 0 16px" }}>
          <EstatusBadge estatus={rep.estatus} />
          <AlertBadge alert={alert} />
        </div>
        <div style={{ padding: 14, background: "var(--card-2)", border: "1px solid var(--line)", borderRadius: 8, marginBottom: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13 }}>
            <div><div className="field-label">Origen</div>{rep.origen || "—"}</div>
            <div><div className="field-label">Destino</div>{rep.destino || "—"}</div>
            <div><div className="field-label">Costo</div>{fmtMoney(rep.costo)}</div>
            <div><div className="field-label">Proveedor</div>{rep.proveedor || "Sin asignar"}</div>
            <div><div className="field-label">ETA arribo</div>{fmtDT(rep.etaArribo)}</div>
            <div><div className="field-label">Creado</div>{fmtDT(rep.fechaCreacion)}</div>
          </div>
        </div>
        <button className="btn btn-outline btn-sm" onClick={() => setEdit(true)} style={{ marginBottom: 16 }}>Editar datos</button>
        <div className="field-label">Seguimiento / comentarios</div>
        <Comments comentarios={rep.comentarios} onAdd={(texto) => onAddComment(rep, texto)} />
      </div>
    </div>
  );
}

/* ============================================================
   AUTO SUSTITUTO — Crear / Detalle
   ============================================================ */

function AutoSustForm({ existing, providers, onCancel, onSave, onOpenExisting }) {
  const [f, setF] = useState(existing || {
    nombre: "", telefono: "", correo: "", numeroSiniestro: "",
    estado: "", municipio: "", proveedor: "",
  });
  const [checkResult, setCheckResult] = useState(null); // {match, action}
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  function verificar(matches) {
    if (!f.numeroSiniestro.trim()) { setCheckResult(null); return; }
    const found = matches.filter((m) => normKey(m.numeroSiniestro) === normKey(f.numeroSiniestro));
    if (found.length === 0) { setCheckResult({ type: "none" }); return; }
    const activo = found.find((m) => m.estatus === "abierto" || m.estatus === "asignado");
    if (activo) { setCheckResult({ type: "activo", rep: activo }); return; }
    const cancelado = found.find((m) => m.estatus === "cancelado");
    if (cancelado) { setCheckResult({ type: "cancelado", rep: cancelado }); return; }
    setCheckResult({ type: "terminado", rep: found[0] });
  }

  return (
    <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onCancel(); }}>
      <div className="modal-panel" style={{ padding: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div className="display" style={{ fontWeight: 700, fontSize: 17 }}>
            {existing ? "Editar reporte de auto sustituto" : "Nuevo reporte de auto sustituto"}
          </div>
          <button className="btn btn-ghost btn-sm" onClick={onCancel}><X size={16} /></button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label">Número de siniestro</label>
            <div style={{ display: "flex", gap: 8 }}>
              <input className="input mono" value={f.numeroSiniestro}
                onChange={(e) => { setF({ ...f, numeroSiniestro: e.target.value }); setCheckResult(null); }} />
              {!existing && (
                <button type="button" className="btn btn-outline btn-sm" onClick={() => verificar(window.__sgs_autosust_all || [])}>
                  <Search size={13} /> Verificar
                </button>
              )}
            </div>
            {checkResult?.type === "activo" && (
              <div style={{ marginTop: 8, padding: 10, background: "var(--warn-bg)", color: "var(--warn-fg)", borderRadius: 9, fontSize: 12.5 }}>
                Ya existe un folio <b>{ESTATUS_LABEL[checkResult.rep.estatus]}</b> para este siniestro. Continúa el seguimiento ahí en lugar de crear uno nuevo.
                <div style={{ marginTop: 6 }}>
                  <button type="button" className="btn btn-primary btn-sm" onClick={() => onOpenExisting(checkResult.rep)}>Abrir folio existente</button>
                </div>
              </div>
            )}
            {checkResult?.type === "cancelado" && (
              <div style={{ marginTop: 8, padding: 10, background: "var(--off-bg)", color: "var(--off-fg)", borderRadius: 9, fontSize: 12.5 }}>
                Existe un folio <b>cancelado</b> para este siniestro. ¿Deseas reabrirlo en lugar de crear uno nuevo?
                <div style={{ marginTop: 6, display: "flex", gap: 8 }}>
                  <button type="button" className="btn btn-primary btn-sm" onClick={() => onOpenExisting(checkResult.rep, true)}>Reabrir folio</button>
                  <button type="button" className="btn btn-outline btn-sm" onClick={() => setCheckResult({ type: "ok-new" })}>Crear uno nuevo de todas formas</button>
                </div>
              </div>
            )}
            {checkResult?.type === "terminado" && (
              <div style={{ marginTop: 8, padding: 10, background: "var(--done-bg)", color: "var(--done-fg)", borderRadius: 9, fontSize: 12.5 }}>
                Ya existe un folio terminado para este siniestro. Puedes continuar y crear uno nuevo si aplica un caso distinto.
              </div>
            )}
            {checkResult?.type === "none" && (
              <div style={{ marginTop: 8, padding: 10, background: "var(--ok-bg)", color: "var(--ok-fg)", borderRadius: 9, fontSize: 12.5 }}>
                No hay folios previos para este siniestro. Puedes continuar.
              </div>
            )}
          </div>
          <div>
            <label className="field-label">Nombre</label>
            <input className="input" value={f.nombre} onChange={set("nombre")} />
          </div>
          <div>
            <label className="field-label">Teléfono</label>
            <input className="input" value={f.telefono} onChange={set("telefono")} />
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label">Correo</label>
            <input className="input" type="email" value={f.correo} onChange={set("correo")} />
          </div>
          <div>
            <label className="field-label">Estado</label>
            <select className="input" value={f.estado} onChange={set("estado")}>
              <option value="">Selecciona…</option>
              {ESTADOS_MX.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Municipio</label>
            <input className="input" value={f.municipio} onChange={set("municipio")} />
          </div>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label">Proveedor asignado</label>
            <select className="input" value={f.proveedor} onChange={set("proveedor")}>
              <option value="">— Sin asignar —</option>
              {providers.map((p) => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
            </select>
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 20 }}>
          <button className="btn btn-outline" onClick={onCancel}>Cancelar</button>
          <button className="btn btn-primary" onClick={() => onSave(f)}
            disabled={!f.nombre.trim() || !f.numeroSiniestro.trim() || checkResult?.type === "activo"}>
            Guardar
          </button>
        </div>
      </div>
    </div>
  );
}

function AutoSustDetail({ rep, providers, onClose, onUpdate, onAddComment }) {
  const [edit, setEdit] = useState(false);
  const alert = autoSustAlert(rep);
  if (edit) {
    return <AutoSustForm existing={rep} providers={providers} onCancel={() => setEdit(false)}
      onSave={(f) => { onUpdate(f); setEdit(false); }} onOpenExisting={() => setEdit(false)} />;
  }
  return (
    <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal-panel" style={{ padding: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div>
            <div className="display" style={{ fontWeight: 700, fontSize: 18 }}>{rep.nombre}</div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>Siniestro: <span className="mono">{rep.numeroSiniestro}</span></div>
          </div>
          <button className="btn btn-ghost btn-sm" onClick={onClose}><X size={16} /></button>
        </div>
        <div style={{ display: "flex", gap: 8, margin: "10px 0 16px" }}>
          <EstatusBadge estatus={rep.estatus} />
          <AlertBadge alert={alert} />
        </div>
        <div style={{ padding: 14, background: "var(--card-2)", border: "1px solid var(--line)", borderRadius: 8, marginBottom: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, fontSize: 13 }}>
            <div><div className="field-label">Teléfono</div>{rep.telefono || "—"}</div>
            <div><div className="field-label">Correo</div>{rep.correo || "—"}</div>
            <div><div className="field-label">Estado</div>{rep.estado || "—"}</div>
            <div><div className="field-label">Municipio</div>{rep.municipio || "—"}</div>
            <div><div className="field-label">Proveedor</div>{rep.proveedor || "Sin asignar"}</div>
            <div><div className="field-label">Límite de respuesta (24h)</div>{fmtDT(new Date(new Date(rep.fechaCreacion).getTime() + 24 * 3600 * 1000).toISOString())}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <button className="btn btn-outline btn-sm" onClick={() => setEdit(true)}>Editar datos</button>
        </div>
        <div className="field-label">Seguimiento / comentarios</div>
        <Comments comentarios={rep.comentarios} onAdd={(texto) => onAddComment(rep, texto)} />
      </div>
    </div>
  );
}

/* ============================================================
   VISTA: SINIESTROS
   ============================================================ */

function SiniestrosView({ siniestros, setSiniestros, providers, currentUser }) {
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("todos");
  const [q, setQ] = useState("");
  const [importSummary, setImportSummary] = useState(null);
  const fileRef = useRef();

  const filtered = useMemo(() => {
    return siniestros.filter((r) => {
      if (filter !== "todos" && r.estatus !== filter) return false;
      if (q && !(`${r.numeroSiniestro} ${r.numeroFolio} ${r.proveedor}`.toLowerCase().includes(q.toLowerCase()))) return false;
      return true;
    }).sort((a, b) => new Date(b.fechaCreacion) - new Date(a.fechaCreacion));
  }, [siniestros, filter, q]);

  const alertas = useMemo(() =>
    siniestros.filter((r) => ["danger", "warn"].includes(siniestroAlert(r).level))
      .sort((a, b) => new Date(a.etaArribo) - new Date(b.etaArribo)),
    [siniestros]);

  function addReport(f) {
    const rep = {
      id: uid("sin"), numeroSiniestro: f.numeroSiniestro.trim(), numeroFolio: f.numeroFolio.trim(),
      estatus: f.estatus, origen: f.origen, destino: f.destino, costo: f.costo, proveedor: f.proveedor,
      etaArribo: f.etaArribo, fechaCreacion: nowISO(), comentarios: [],
    };
    setSiniestros([rep, ...siniestros]);
    setShowForm(false);
  }
  function updateReport(f) {
    setSiniestros(siniestros.map((r) => (r.id === f.id ? { ...r, ...f } : r)));
    setSelected((s) => s ? { ...s, ...f } : s);
  }
  function addComment(rep, texto) {
    const c = { id: uid("cmt"), texto, fecha: nowISO(), usuario: currentUser.username };
    const updated = { ...rep, comentarios: [...(rep.comentarios || []), c] };
    setSiniestros(siniestros.map((r) => (r.id === rep.id ? updated : r)));
    setSelected(updated);
  }

  function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        let coinciden = 0, agregados = [];
        const current = [...siniestros];
        rows.forEach((row) => {
          const norm = {};
          Object.keys(row).forEach((k) => norm[normKey(k).replace(/\s+/g, "_")] = row[k]);
          const numero = String(norm.numero_siniestro || norm.siniestro || norm.no_siniestro || "").trim();
          if (!numero) return;
          const exists = current.find((r) => normKey(r.numeroSiniestro) === normKey(numero));
          if (exists) { coinciden++; return; }
          const estatusRaw = normKey(norm.estatus || norm.status || "abierto");
          const estatus = ESTATUS_SIN.includes(estatusRaw) ? estatusRaw : "abierto";
          const rep = {
            id: uid("sin"),
            numeroSiniestro: numero,
            numeroFolio: String(norm.numero_folio || norm.folio || ""),
            estatus,
            origen: String(norm.origen || ""),
            destino: String(norm.destino || ""),
            costo: norm.costo || "",
            proveedor: String(norm.proveedor || ""),
            etaArribo: norm.eta || norm.tiempo_estimado_arribo ? new Date(norm.eta || norm.tiempo_estimado_arribo).toISOString() : "",
            fechaCreacion: norm.fecha ? new Date(norm.fecha).toISOString() : nowISO(),
            comentarios: [{ id: uid("cmt"), texto: "Folio anexado desde importación de Excel.", fecha: nowISO(), usuario: currentUser.username }],
          };
          current.push(rep);
          agregados.push(rep.numeroSiniestro);
        });
        setSiniestros(current);
        setImportSummary({ coinciden, agregados });
      } catch (err) {
        setImportSummary({ error: "No se pudo leer el archivo. Verifica que sea un .xlsx válido." });
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }

  function exportBackup() {
    const wb = XLSX.utils.book_new();
    const wsData = siniestros.map((r) => ({
      numero_siniestro: r.numeroSiniestro, numero_folio: r.numeroFolio, estatus: r.estatus,
      origen: r.origen, destino: r.destino, costo: r.costo, proveedor: r.proveedor,
      eta_arribo: r.etaArribo, fecha_creacion: r.fechaCreacion,
      comentarios: (r.comentarios || []).map((c) => `[${fmtDT(c.fecha)}] ${c.usuario}: ${c.texto}`).join(" | "),
    }));
    const ws = XLSX.utils.json_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, "Siniestros");
    XLSX.writeFile(wb, `respaldo_siniestros_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div>
          <div className="display" style={{ fontWeight: 700, fontSize: 20 }}>Reportes de siniestro</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{siniestros.length} folios registrados</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImport} />
          <button className="btn btn-outline btn-sm" onClick={() => fileRef.current.click()}><Upload size={13} /> Importar Excel</button>
          <button className="btn btn-outline btn-sm" onClick={exportBackup}><Download size={13} /> Exportar respaldo</button>
          <button className="btn btn-primary btn-sm" onClick={() => setShowForm(true)}><Plus size={13} /> Nuevo reporte</button>
        </div>
      </div>

      {alertas.length > 0 && (
        <div style={{ padding: 12, marginBottom: 14, background: "var(--warn-bg)", border: "1px solid #fde68a", borderRadius: 9 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 12.5, color: "var(--warn-fg)", marginBottom: 6 }}>
            <Bell size={14} /> Seguimiento preventivo — {alertas.length} folio(s) próximos o vencidos
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {alertas.slice(0, 8).map((r) => (
              <span key={r.id} className="badge-pill" style={{ cursor: "pointer", background: "var(--card-2)", border: "1px solid var(--line)" }} onClick={() => setSelected(r)}>
                <span className="mono">{r.numeroSiniestro}</span> · {siniestroAlert(r).label}
              </span>
            ))}
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <input className="input" style={{ maxWidth: 260 }} placeholder="Buscar por siniestro, folio o proveedor…" value={q} onChange={(e) => setQ(e.target.value)} />
        {["todos", ...ESTATUS_SIN].map((s) => (
          <button key={s} className={"btn btn-sm filter-pill" + (filter === s ? " active" : "")}
            onClick={() => setFilter(s)}>
            {s === "todos" ? "Todos" : ESTATUS_LABEL[s]}
          </button>
        ))}
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <table className="sgs-table">
          <thead>
            <tr>
              <th>Siniestro</th><th>Folio</th><th>Estatus</th><th>Alerta</th><th>Origen → Destino</th><th>Costo</th><th>Proveedor</th><th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r) => (
              <tr key={r.id} style={{ cursor: "pointer" }} onClick={() => setSelected(r)}>
                <td className="mono">{r.numeroSiniestro}</td>
                <td className="mono">{r.numeroFolio || "—"}</td>
                <td><EstatusBadge estatus={r.estatus} /></td>
                <td><AlertBadge alert={siniestroAlert(r)} /></td>
                <td style={{ fontSize: 12.5 }}>{(r.origen || "—") + " → " + (r.destino || "—")}</td>
                <td>{fmtMoney(r.costo)}</td>
                <td>{r.proveedor || "—"}</td>
                <td><ChevronRight size={14} color="var(--ink-soft)" /></td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={8} style={{ textAlign: "center", padding: 24, color: "var(--ink-soft)" }}>No hay reportes que coincidan.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && <SiniestroForm providers={providers} onCancel={() => setShowForm(false)} onSave={addReport} />}
      {selected && <SiniestroDetail rep={selected} providers={providers} currentUser={currentUser}
        onClose={() => setSelected(null)} onUpdate={updateReport} onAddComment={addComment} />}

      {importSummary && (
        <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) setImportSummary(null); }}>
          <div className="modal-panel" style={{ padding: 24, maxWidth: 420 }}>
            <div className="display" style={{ fontWeight: 700, fontSize: 16, marginBottom: 10 }}>Resultado de importación</div>
            {importSummary.error ? (
              <div style={{ color: "var(--danger-fg)", fontSize: 13 }}>{importSummary.error}</div>
            ) : (
              <div style={{ fontSize: 13.5 }}>
                <p><b>{importSummary.coinciden}</b> folios ya existían y coinciden con el sistema.</p>
                <p><b>{importSummary.agregados.length}</b> folios nuevos fueron anexados desde el Excel.</p>
                {importSummary.agregados.length > 0 && (
                  <div className="mono" style={{ fontSize: 12, maxHeight: 120, overflowY: "auto", background: "var(--card-2)", border: "1px solid var(--line)", padding: 8, borderRadius: 7 }}>
                    {importSummary.agregados.join(", ")}
                  </div>
                )}
              </div>
            )}
            <button className="btn btn-primary" style={{ marginTop: 14, width: "100%", justifyContent: "center" }} onClick={() => setImportSummary(null)}>Cerrar</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   VISTA: AUTO SUSTITUTO
   ============================================================ */

function AutoSustitutoView({ autosust, setAutosust, providers, currentUser }) {
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("todos");
  const [q, setQ] = useState("");

  useEffect(() => { window.__sgs_autosust_all = autosust; }, [autosust]);

  const filtered = useMemo(() => {
    return autosust.filter((r) => {
      if (filter !== "todos" && r.estatus !== filter) return false;
      if (q && !(`${r.nombre} ${r.numeroSiniestro} ${r.proveedor}`.toLowerCase().includes(q.toLowerCase()))) return false;
      return true;
    }).sort((a, b) => new Date(b.fechaCreacion) - new Date(a.fechaCreacion));
  }, [autosust, filter, q]);

  const alertas = useMemo(() =>
    autosust.filter((r) => ["danger", "warn"].includes(autoSustAlert(r).level)),
    [autosust]);

  function addReport(f) {
    const rep = {
      id: uid("asu"), nombre: f.nombre.trim(), telefono: f.telefono, correo: f.correo,
      numeroSiniestro: f.numeroSiniestro.trim(), estado: f.estado, municipio: f.municipio,
      proveedor: f.proveedor, estatus: "abierto", fechaCreacion: nowISO(), comentarios: [],
    };
    setAutosust([rep, ...autosust]);
    setShowForm(false);
  }
  function updateReport(f) {
    setAutosust(autosust.map((r) => (r.id === f.id ? { ...r, ...f } : r)));
    setSelected((s) => s ? { ...s, ...f } : s);
  }
  function addComment(rep, texto) {
    const c = { id: uid("cmt"), texto, fecha: nowISO(), usuario: currentUser.username };
    const updated = { ...rep, comentarios: [...(rep.comentarios || []), c] };
    setAutosust(autosust.map((r) => (r.id === rep.id ? updated : r)));
    setSelected(updated);
  }
  function openExisting(rep, reopen) {
    if (reopen) {
      const c = { id: uid("cmt"), texto: "Folio reabierto a partir de un folio cancelado previo.", fecha: nowISO(), usuario: currentUser.username };
      const updated = { ...rep, estatus: "abierto", comentarios: [...(rep.comentarios || []), c] };
      setAutosust(autosust.map((r) => (r.id === rep.id ? updated : r)));
      setSelected(updated);
    } else {
      setSelected(rep);
    }
    setShowForm(false);
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div>
          <div className="display" style={{ fontWeight: 700, fontSize: 20 }}>Reportes de auto sustituto</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{autosust.length} folios · tiempo de respuesta 24 h</div>
        </div>
        <button className="btn btn-primary btn-sm" onClick={() => setShowForm(true)}><Plus size={13} /> Nuevo reporte</button>
      </div>

      {alertas.length > 0 && (
        <div style={{ padding: 12, marginBottom: 14, background: "var(--warn-bg)", border: "1px solid #fde68a", borderRadius: 9 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 12.5, color: "var(--warn-fg)", marginBottom: 6 }}>
            <Bell size={14} /> {alertas.length} folio(s) cerca o fuera del tiempo de respuesta de 24h
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {alertas.slice(0, 8).map((r) => (
              <span key={r.id} className="badge-pill" style={{ cursor: "pointer", background: "var(--card-2)", border: "1px solid var(--line)" }} onClick={() => setSelected(r)}>
                <span className="mono">{r.numeroSiniestro}</span> · {autoSustAlert(r).label}
              </span>
            ))}
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <input className="input" style={{ maxWidth: 260 }} placeholder="Buscar por nombre, siniestro o proveedor…" value={q} onChange={(e) => setQ(e.target.value)} />
        {["todos", ...ESTATUS_SIN].map((s) => (
          <button key={s} className={"btn btn-sm filter-pill" + (filter === s ? " active" : "")}
            onClick={() => setFilter(s)}>
            {s === "todos" ? "Todos" : ESTATUS_LABEL[s]}
          </button>
        ))}
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <table className="sgs-table">
          <thead>
            <tr><th>Nombre</th><th>Siniestro</th><th>Estatus</th><th>Alerta</th><th>Estado / Municipio</th><th>Proveedor</th><th></th></tr>
          </thead>
          <tbody>
            {filtered.map((r) => (
              <tr key={r.id} style={{ cursor: "pointer" }} onClick={() => setSelected(r)}>
                <td>{r.nombre}</td>
                <td className="mono">{r.numeroSiniestro}</td>
                <td><EstatusBadge estatus={r.estatus} /></td>
                <td><AlertBadge alert={autoSustAlert(r)} /></td>
                <td style={{ fontSize: 12.5 }}>{r.estado || "—"}{r.municipio ? `, ${r.municipio}` : ""}</td>
                <td>{r.proveedor || "—"}</td>
                <td><ChevronRight size={14} color="var(--ink-soft)" /></td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={7} style={{ textAlign: "center", padding: 24, color: "var(--ink-soft)" }}>No hay reportes que coincidan.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showForm && <AutoSustForm providers={providers} onCancel={() => setShowForm(false)} onSave={addReport} onOpenExisting={openExisting} />}
      {selected && <AutoSustDetail rep={selected} providers={providers} onClose={() => setSelected(null)} onUpdate={updateReport} onAddComment={addComment} />}
    </div>
  );
}

/* ============================================================
   VISTA: PROVEEDORES (admin)
   ============================================================ */

function ProveedoresView({ providers, setProviders }) {
  const [name, setName] = useState("");
  const fileRef = useRef();

  function add() {
    if (!name.trim()) return;
    if (providers.some((p) => normKey(p.nombre) === normKey(name))) { setName(""); return; }
    setProviders([...providers, { id: uid("prv"), nombre: name.trim() }]);
    setName("");
  }
  function remove(id) { setProviders(providers.filter((p) => p.id !== id)); }

  function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const current = [...providers];
        rows.forEach((row) => {
          const norm = {};
          Object.keys(row).forEach((k) => norm[normKey(k)] = row[k]);
          const nombre = String(norm.proveedor || norm.nombre || Object.values(row)[0] || "").trim();
          if (nombre && !current.some((p) => normKey(p.nombre) === normKey(nombre))) {
            current.push({ id: uid("prv"), nombre });
          }
        });
        setProviders(current);
      } catch (err) {}
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }

  return (
    <div>
      <div className="display" style={{ fontWeight: 700, fontSize: 20, marginBottom: 4 }}>Proveedores</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 16 }}>{providers.length} proveedores disponibles para asignación</div>

      <div className="card" style={{ padding: 16, marginBottom: 16, display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 220 }}>
          <label className="field-label">Agregar proveedor manualmente</label>
          <input className="input" placeholder="Nombre del proveedor" value={name}
            onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        </div>
        <button className="btn btn-primary btn-sm" onClick={add}><Plus size={13} /> Agregar</button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImport} />
        <button className="btn btn-outline btn-sm" onClick={() => fileRef.current.click()}><Upload size={13} /> Cargar desde Excel</button>
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <table className="sgs-table">
          <thead><tr><th>Proveedor</th><th></th></tr></thead>
          <tbody>
            {providers.map((p) => (
              <tr key={p.id}>
                <td><Building2 size={13} style={{ display: "inline", marginRight: 6, verticalAlign: -2 }} color="var(--ink-soft)" />{p.nombre}</td>
                <td style={{ textAlign: "right" }}>
                  <button className="btn btn-ghost btn-sm" onClick={() => remove(p.id)}><Trash2 size={13} color="var(--danger-fg)" /></button>
                </td>
              </tr>
            ))}
            {providers.length === 0 && <tr><td colSpan={2} style={{ textAlign: "center", padding: 24, color: "var(--ink-soft)" }}>Sin proveedores aún.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================================================
   VISTA: USUARIOS (admin)
   ============================================================ */

function UsuariosView({ users, setUsers, currentUser }) {
  const [f, setF] = useState({ username: "", password: "", role: "usuario", nombre: "" });

  function add() {
    if (!f.username.trim() || !f.password.trim()) return;
    if (users.some((u) => normKey(u.username) === normKey(f.username))) { alert("Ese usuario ya existe."); return; }
    setUsers([...users, { id: uid("usr"), username: f.username.trim(), password: f.password, role: f.role, nombre: f.nombre.trim() || f.username.trim() }]);
    setF({ username: "", password: "", role: "usuario", nombre: "" });
  }
  function remove(u) {
    if (u.id === currentUser.id) { alert("No puedes eliminar tu propio usuario mientras tienes sesión activa."); return; }
    const admins = users.filter((x) => x.role === "admin");
    if (u.role === "admin" && admins.length <= 1) { alert("Debe existir al menos un administrador."); return; }
    setUsers(users.filter((x) => x.id !== u.id));
  }

  return (
    <div>
      <div className="display" style={{ fontWeight: 700, fontSize: 20, marginBottom: 4 }}>Usuarios del sistema</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 16 }}>Solo los administradores pueden agregar o eliminar usuarios.</div>

      <div className="card" style={{ padding: 16, marginBottom: 16 }}>
        <div className="field-label" style={{ marginBottom: 10 }}>Agregar nuevo usuario</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <div><label className="field-label">Nombre completo</label><input className="input" value={f.nombre} onChange={(e) => setF({ ...f, nombre: e.target.value })} /></div>
          <div><label className="field-label">Usuario</label><input className="input" value={f.username} onChange={(e) => setF({ ...f, username: e.target.value })} /></div>
          <div><label className="field-label">Contraseña</label><input className="input" type="text" value={f.password} onChange={(e) => setF({ ...f, password: e.target.value })} /></div>
          <div>
            <label className="field-label">Permisos</label>
            <select className="input" value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })}>
              <option value="usuario">Usuario estándar</option>
              <option value="admin">Administrador</option>
            </select>
          </div>
          <button className="btn btn-primary btn-sm" onClick={add}><UserPlus size={13} /> Agregar</button>
        </div>
      </div>

      <div className="card" style={{ overflow: "hidden" }}>
        <table className="sgs-table">
          <thead><tr><th>Nombre</th><th>Usuario</th><th>Permisos</th><th></th></tr></thead>
          <tbody>
            {users.map((u) => (
              <tr key={u.id}>
                <td>{u.nombre || u.username}</td>
                <td className="mono">{u.username}</td>
                <td>{u.role === "admin" ? <span className="badge-pill estatus-asignado">Administrador</span> : <span className="badge-pill estatus-abierto">Usuario</span>}</td>
                <td style={{ textAlign: "right" }}>
                  <button className="btn btn-ghost btn-sm" onClick={() => remove(u)}><Trash2 size={13} color="var(--danger-fg)" /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */

function DashboardView({ siniestros, autosust }) {
  const countsS = useMemo(() => {
    const c = { abierto: 0, asignado: 0, terminado: 0, cancelado: 0 };
    siniestros.forEach((r) => c[r.estatus] !== undefined && c[r.estatus]++);
    return c;
  }, [siniestros]);
  const countsA = useMemo(() => {
    const c = { abierto: 0, asignado: 0, terminado: 0, cancelado: 0 };
    autosust.forEach((r) => c[r.estatus] !== undefined && c[r.estatus]++);
    return c;
  }, [autosust]);
  const alertasS = siniestros.filter((r) => ["danger", "warn"].includes(siniestroAlert(r).level)).length;
  const alertasA = autosust.filter((r) => ["danger", "warn"].includes(autoSustAlert(r).level)).length;

  return (
    <div>
      <div className="display" style={{ fontWeight: 700, fontSize: 20, marginBottom: 16 }}>Resumen operativo</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 22 }}>
        <div className="stat-card">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <div className="icon-chip" style={{ background: "var(--blue-bg)" }}><FileText size={16} color="var(--blue-fg)" /></div>
            <div className="field-label" style={{ margin: 0 }}>Siniestros abiertos</div>
          </div>
          <div style={{ fontSize: 28, fontWeight: 800 }}>{countsS.abierto + countsS.asignado}</div>
        </div>
        <div className="stat-card">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <div className="icon-chip" style={{ background: "var(--purple-bg)" }}><Truck size={16} color="var(--purple-fg)" /></div>
            <div className="field-label" style={{ margin: 0 }}>Auto sustituto activos</div>
          </div>
          <div style={{ fontSize: 28, fontWeight: 800 }}>{countsA.abierto + countsA.asignado}</div>
        </div>
        <div className="stat-card">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <div className="icon-chip" style={{ background: alertasS ? "var(--peach-bg)" : "var(--off-bg)" }}><Bell size={16} color={alertasS ? "var(--peach-fg)" : "var(--off-fg)"} /></div>
            <div className="field-label" style={{ margin: 0 }}>Alertas de arribo</div>
          </div>
          <div style={{ fontSize: 28, fontWeight: 800, color: alertasS ? "var(--peach-fg)" : "var(--ink)" }}>{alertasS}</div>
        </div>
        <div className="stat-card">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
            <div className="icon-chip" style={{ background: alertasA ? "var(--peach-bg)" : "var(--off-bg)" }}><Clock size={16} color={alertasA ? "var(--peach-fg)" : "var(--off-fg)"} /></div>
            <div className="field-label" style={{ margin: 0 }}>Alertas de 24h</div>
          </div>
          <div style={{ fontSize: 28, fontWeight: 800, color: alertasA ? "var(--peach-fg)" : "var(--ink)" }}>{alertasA}</div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div className="card" style={{ padding: 18 }}>
          <div className="field-label" style={{ marginBottom: 12 }}>Siniestros por estatus</div>
          {ESTATUS_SIN.map((s) => (
            <div key={s} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: "1px solid var(--line-soft)", fontSize: 13 }}>
              <EstatusBadge estatus={s} /><span className="mono">{countsS[s]}</span>
            </div>
          ))}
        </div>
        <div className="card" style={{ padding: 18 }}>
          <div className="field-label" style={{ marginBottom: 12 }}>Auto sustituto por estatus</div>
          {ESTATUS_SIN.map((s) => (
            <div key={s} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: "1px solid var(--line-soft)", fontSize: 13 }}>
              <EstatusBadge estatus={s} /><span className="mono">{countsA[s]}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   APP RAÍZ
   ============================================================ */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [dark, setDark] = useState(false);
  const [pageKey, setPageKey] = useState(0);
  const pageTransRef = useRef(null);

  const navigateTo = useCallback((nextTab) => {
    if (nextTab === tab) return;
    // Dispara exit → luego entra la nueva página
    const el = pageTransRef.current;
    if (el) {
      el.classList.remove("page-enter");
      el.classList.add("page-exit");
      setTimeout(() => {
        setTab(nextTab);
        setPageKey(k => k + 1);
      }, 170);
    } else {
      setTab(nextTab);
      setPageKey(k => k + 1);
    }
  }, [tab]);

  const [users, setUsersState] = useState([]);
  const [providers, setProvidersState] = useState([]);
  const [siniestros, setSiniestrosState] = useState([]);
  const [autosust, setAutosustState] = useState([]);

  useEffect(() => {
    (async () => {
      const [u, p, s, a] = await Promise.all([
        loadList(KEYS.USERS, SEED_USERS),
        loadList(KEYS.PROVIDERS, []),
        loadList(KEYS.SINIESTROS, []),
        loadList(KEYS.AUTOSUST, []),
      ]);
      if (!u || u.length === 0) {
        setUsersState(SEED_USERS);
        saveList(KEYS.USERS, SEED_USERS);
      } else {
        setUsersState(u);
      }
      setProvidersState(p); setSiniestrosState(s); setAutosustState(a);
      setLoading(false);
    })();
  }, []);

  function setUsers(next) { setUsersState(next); saveList(KEYS.USERS, next); }
  function setProviders(next) { setProvidersState(next); saveList(KEYS.PROVIDERS, next); }
  function setSiniestros(next) { setSiniestrosState(next); saveList(KEYS.SINIESTROS, next); }
  function setAutosust(next) { setAutosustState(next); saveList(KEYS.AUTOSUST, next); }

  if (loading) {
    return (
      <div className={"sgs-root"+(dark?" dark":"")} style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh" }}>
        <GlobalStyles />
        <div style={{ color:"var(--ink-soft)", fontSize:13 }}>Cargando sistema…</div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className={"sgs-root"+(dark?" dark":"")}>
        <GlobalStyles />
        <LoginView users={users} onLogin={setCurrentUser} onResetUsers={()=>setUsers(SEED_USERS)} dark={dark} toggleDark={()=>setDark(d=>!d)} />
      </div>
    );
  }

  const NAV_LABEL = {
    dashboard: "Resumen", siniestros: "Siniestros", autosust: "Auto sustituto",
    proveedores: "Proveedores", usuarios: "Usuarios",
  };
  const NAV = [
    { id: "dashboard", label: "Resumen", icon: ShieldCheck },
    { id: "siniestros", label: "Siniestros", icon: FileText },
    { id: "autosust", label: "Auto sustituto", icon: Truck },
    { id: "proveedores", label: "Proveedores", icon: Building2 },
    ...(currentUser.role === "admin" ? [{ id: "usuarios", label: "Usuarios", icon: UsersIcon }] : []),
  ];
  const initials = (currentUser.nombre || currentUser.username || "?").trim().slice(0, 1).toUpperCase();

  const NAV_MAIN   = NAV.filter(n => ["dashboard","siniestros","autosust"].includes(n.id));
  const NAV_CONFIG  = NAV.filter(n => ["proveedores","usuarios"].includes(n.id));

  return (
    <div className={"sgs-root"+(dark?" dark":"")} style={{ display:"flex", minHeight:"100vh" }}>
      <GlobalStyles />

      {/* ══════════ SIDEBAR — Liquid Glass ══════════ */}
      <aside style={{
        width:260, flexShrink:0, position:"sticky", top:0, height:"100vh",
        overflowY:"auto", padding:"22px 12px",
        display:"flex", flexDirection:"column",
        /* Liquid glass */
        background:"var(--glass-sidebar)",
        backdropFilter:"var(--blur)", WebkitBackdropFilter:"var(--blur)",
        borderRight:"1px solid var(--glass-border)",
        boxShadow:"var(--shadow-sidebar)",
        /* Specular highlight on right edge */
        borderTop:"none", borderBottom:"none", borderLeft:"none",
      }}>
        {/* Logo */}
        <div style={{ display:"flex", alignItems:"center", gap:10, padding:"2px 10px", marginBottom:30 }}>
          <div style={{
            width:33, height:33, borderRadius:10, background:"var(--accent)", color:"var(--accent-fg)",
            display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0,
            boxShadow:"0 4px 12px rgba(0,0,0,0.18), inset 0 1px 0 rgba(255,255,255,0.15)",
          }}>
            <ShieldCheck size={15} color="currentColor" />
          </div>
          <span style={{ fontWeight:700, fontSize:14, color:"var(--ink)", letterSpacing:"-0.01em", lineHeight:1.2 }}>
            Folios &amp; Asistencia
          </span>
        </div>

        {/* Main nav */}
        <div style={{ marginBottom:24 }}>
          <div style={{ fontSize:10.5, fontWeight:600, textTransform:"uppercase", letterSpacing:".07em", color:"var(--ink-faint)", padding:"0 11px", marginBottom:6 }}>
            Menú principal
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:2 }}>
            {NAV_MAIN.map(n => (
              <div key={n.id} className={"nav-link"+(tab===n.id?" active":"")} onClick={()=>navigateTo(n.id)}>
                <n.icon size={15} strokeWidth={tab===n.id?2.2:1.8} /> {n.label}
              </div>
            ))}
          </div>
        </div>

        {/* Config nav */}
        {NAV_CONFIG.length > 0 && (
          <div style={{ marginBottom:24 }}>
            <div style={{ fontSize:10.5, fontWeight:600, textTransform:"uppercase", letterSpacing:".07em", color:"var(--ink-faint)", padding:"0 11px", marginBottom:6 }}>
              Configuración
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
              {NAV_CONFIG.map(n => (
                <div key={n.id} className={"nav-link"+(tab===n.id?" active":"")} onClick={()=>navigateTo(n.id)}>
                  <n.icon size={15} strokeWidth={tab===n.id?2.2:1.8} /> {n.label}
                </div>
              ))}
            </div>
          </div>
        )}

        <div style={{ flex:1 }} />

        {/* Account + theme toggle */}
        <div style={{ borderTop:"1px solid var(--glass-border)", paddingTop:14 }}>
          <div style={{ display:"flex", alignItems:"center", gap:9, padding:"4px 11px 10px" }}>
            <div className="avatar-circle" style={{ width:30, height:30, fontSize:12 }}>{initials}</div>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontSize:13, fontWeight:600, color:"var(--ink)", lineHeight:1.3, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                {currentUser.nombre || currentUser.username}
              </div>
              <div style={{ fontSize:11, color:"var(--ink-faint)" }}>
                {currentUser.role==="admin"?"Administrador":"Operativo"}
              </div>
            </div>
          </div>
          <div className="nav-link" onClick={()=>setCurrentUser(null)} style={{ color:"var(--ink-soft)", marginBottom:4 }}>
            <LogOut size={14} strokeWidth={1.8} /> Cerrar sesión
          </div>
          {/* Theme toggle */}
          <div className="nav-link" onClick={()=>setDark(d=>!d)} style={{ color:"var(--ink-soft)" }}>
            {dark ? <Sun size={14} strokeWidth={1.8} /> : <Moon size={14} strokeWidth={1.8} />}
            {dark ? "Modo claro" : "Modo oscuro"}
          </div>
        </div>
      </aside>

      {/* ══════════ CONTENT ══════════ */}
      <div style={{ flex:1, display:"flex", flexDirection:"column", minWidth:0, padding:"18px 18px 18px 0" }}>
        {/* Floating glass frame */}
        <div style={{
          flex:1, borderRadius:16,
          background:"var(--glass-frame)",
          backdropFilter:"var(--blur)", WebkitBackdropFilter:"var(--blur)",
          border:"1px solid var(--glass-border)",
          boxShadow:"var(--shadow-frame)",
          display:"flex", flexDirection:"column", overflow:"hidden",
        }}>
          {/* Top bar */}
          <div style={{
            display:"flex", alignItems:"center", justifyContent:"space-between",
            padding:"13px 26px",
            borderBottom:"1px solid var(--glass-border)",
            background:"var(--glass-hi)",
          }}>
            <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:13, color:"var(--ink-faint)" }}>
              <span style={{ fontWeight:500 }}>Folios &amp; Asistencia</span>
              <ChevronRight size={13} strokeWidth={1.5} />
              <span style={{ color:"var(--ink-mid)", fontWeight:600 }}>{NAV_LABEL[tab]}</span>
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:12 }}>
              <button className="btn btn-ghost btn-sm" onClick={()=>setDark(d=>!d)} style={{ borderRadius:99, padding:"6px 8px", border:"1px solid var(--glass-border2)" }}>
                {dark ? <Sun size={14} /> : <Moon size={14} />}
              </button>
              <Bell size={16} strokeWidth={1.6} color="var(--ink-soft)" style={{ cursor:"pointer" }} />
              <div className="avatar-circle" style={{ width:28, height:28, fontSize:11.5 }}>{initials}</div>
            </div>
          </div>

          {/* Page content — animated transition wrapper */}
          <div style={{ flex:1, overflowY:"auto", position:"relative" }}>
            <div
              key={pageKey}
              ref={pageTransRef}
              className="page-enter"
              style={{ padding:"28px 28px", minHeight:"100%" }}
            >
              {tab==="dashboard"   && <DashboardView siniestros={siniestros} autosust={autosust} />}
              {tab==="siniestros"  && <SiniestrosView siniestros={siniestros} setSiniestros={setSiniestros} providers={providers} currentUser={currentUser} />}
              {tab==="autosust"    && <AutoSustitutoView autosust={autosust} setAutosust={setAutosust} providers={providers} currentUser={currentUser} />}
              {tab==="proveedores" && <ProveedoresView providers={providers} setProviders={setProviders} />}
              {tab==="usuarios" && currentUser.role==="admin" && <UsuariosView users={users} setUsers={setUsers} currentUser={currentUser} />}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
