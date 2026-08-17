import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend, CartesianGrid
} from "recharts";
import * as XLSX from "xlsx";
import {
  Building2, Users, ClipboardList, LayoutDashboard, LogOut, Plus, X, Bell,
  Download, FileSpreadsheet, FileText, Pencil, Clock, AlertTriangle, CheckCircle2,
  ChevronRight, Stamp, CalendarClock, MessageSquare, ShieldCheck
} from "lucide-react";

/* ---------------------------------------------------------------
   Paleta / tokens — estética de "expediente administrativo":
   papel de registro, tinta institucional, sellos de criticidad.
----------------------------------------------------------------*/
const TOKENS = `
  @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Source+Sans+3:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

  .unju-root{
    --paper:#EAE6DA;
    --paper-dark:#DFDACB;
    --card:#F5F2E9;
    --ink:#20303D;
    --ink-soft:#5B6B75;
    --line:#C9C2AC;
    --seal:#7A2E2E;
    --seal-dark:#5E2222;
    --gold:#A9762F;
    --alta:#A6382C;
    --alta-bg:#F1DAD3;
    --media:#B4832A;
    --media-bg:#F1E4C7;
    --baja:#3E6E4E;
    --baja-bg:#DCE7DB;
    --navy:#22303F;
    font-family:'Source Sans 3',sans-serif;
    color:var(--ink);
    background:var(--paper);
    min-height:100vh;
  }
  .unju-root .font-display{ font-family:'Fraunces',serif; }
  .unju-root .font-mono{ font-family:'IBM Plex Mono',monospace; }

  .unju-card{
    background:var(--card);
    border:1px solid var(--line);
    box-shadow: 2px 2px 0 rgba(32,48,61,0.06);
  }
  .unju-stamp{
    display:inline-flex;align-items:center;gap:4px;
    font-family:'IBM Plex Mono',monospace;
    font-size:10px;letter-spacing:.06em;text-transform:uppercase;
    padding:3px 8px;border-radius:2px;border:1.5px solid currentColor;
    transform:rotate(-1.5deg);
    font-weight:600;
  }
  .unju-tab-active{
    background:var(--navy); color:var(--paper);
  }
  .unju-divider{ border-top:1px dashed var(--line); }
  .unju-input{
    background:var(--paper);
    border:1px solid var(--line);
    color:var(--ink);
  }
  .unju-input:focus{ outline:2px solid var(--navy); outline-offset:1px; }
  .unju-scrollbar::-webkit-scrollbar{ height:6px; width:6px; }
  .unju-scrollbar::-webkit-scrollbar-thumb{ background:var(--line); border-radius:4px; }
`;

/* ---------------------------------------------------------------
   Datos semilla
----------------------------------------------------------------*/
const AREA_SEED = [
  { id: "contabilidad", nombre: "Contabilidad" },
  { id: "secretarias", nombre: "Secretarías" },
  { id: "soporte", nombre: "Soporte de Procesos e Informática" },
  { id: "presupuesto", nombre: "Presupuesto" },
  { id: "compras", nombre: "Compras" },
];

const USER_SEED = [
  { id: "u-director", nombre: "Marcela Ríos", usuario: "mrios", clave: "1234", rol: "director", areaId: null },
  { id: "u-admin", nombre: "Julián Torres", usuario: "jtorres", clave: "1234", rol: "admin", areaId: null },
  { id: "u-contable", nombre: "Estela Farfán", usuario: "efarfan", clave: "1234", rol: "area", areaId: "contabilidad" },
  { id: "u-secret", nombre: "Pablo Guzmán", usuario: "pguzman", clave: "1234", rol: "area", areaId: "secretarias" },
  { id: "u-soporte", nombre: "Lucas Vera", usuario: "lvera", clave: "1234", rol: "area", areaId: "soporte" },
  { id: "u-presupuesto", nombre: "Noelia Choque", usuario: "nchoque", clave: "1234", rol: "area", areaId: "presupuesto" },
  { id: "u-compras", nombre: "Diego Salazar", usuario: "dsalazar", clave: "1234", rol: "area", areaId: "compras" },
];

const today = new Date();
const addDays = (n) => { const d = new Date(today); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };

const TASK_SEED = [
  {
    id: "t-1", areaId: "contabilidad", titulo: "Balance trimestral de ejecución presupuestaria",
    descripcion: "Consolidar balance del 3er trimestre con partidas de nación.",
    responsable: "Estela Farfán", criticidad: "alta", estado: "en_progreso",
    fechaLimite: addDays(3), avance: 65,
    comentarios: [
      { autor: "Estela Farfán", rol: "area", fecha: addDays(-4), texto: "Falta conciliar dos partidas con Tesorería.", tipo: "comentario" },
    ],
    historial: [],
  },
  {
    id: "t-2", areaId: "contabilidad", titulo: "Análisis de costos — división de ingresos de nación",
    descripcion: "Detallar cómo se distribuyen los ingresos girados por nación entre las distintas áreas de la universidad.",
    responsable: "Estela Farfán", criticidad: "media", estado: "pendiente",
    fechaLimite: addDays(10), avance: 15,
    comentarios: [],
    historial: [
      { fechaAnterior: addDays(-2), fechaNueva: addDays(10), motivo: "Se solicitó información adicional a Presupuesto." },
    ],
  },
  {
    id: "t-3", areaId: "secretarias", titulo: "Actualización de mesa de entradas digital",
    descripcion: "Migrar registro de expedientes a la nueva planilla unificada.",
    responsable: "Pablo Guzmán", criticidad: "baja", estado: "en_progreso",
    fechaLimite: addDays(20), avance: 40, comentarios: [], historial: [],
  },
  {
    id: "t-4", areaId: "soporte", titulo: "Migración de servidores de expedientes",
    descripcion: "Pasar el repositorio de expedientes al nuevo servidor con respaldo diario.",
    responsable: "Lucas Vera", criticidad: "alta", estado: "pendiente",
    fechaLimite: addDays(-1), avance: 20,
    comentarios: [
      { autor: "Marcela Ríos", rol: "director", fecha: addDays(-1), texto: "Tarea vencida sin reporte de avance — necesito una actualización urgente.", tipo: "alerta" },
    ],
    historial: [],
  },
  {
    id: "t-5", areaId: "presupuesto", titulo: "Proyección de presupuesto 2027",
    descripcion: "Elaborar proyección preliminar por área para el ejercicio 2027.",
    responsable: "Noelia Choque", criticidad: "media", estado: "en_progreso",
    fechaLimite: addDays(15), avance: 55, comentarios: [], historial: [],
  },
  {
    id: "t-6", areaId: "compras", titulo: "Licitación de insumos de laboratorio",
    descripcion: "Pliego de licitación pública para reposición de insumos.",
    responsable: "Diego Salazar", criticidad: "alta", estado: "finalizada",
    fechaLimite: addDays(-5), avance: 100, comentarios: [], historial: [],
  },
  {
    id: "t-7", areaId: "contabilidad", titulo: "Rendición de fondos de Bienestar Universitario",
    descripcion: "Rendir ante Contaduría los fondos ejecutados del programa de becas.",
    responsable: "Estela Farfán", criticidad: "media", estado: "en_progreso",
    fechaLimite: addDays(7), avance: 80, comentarios: [], historial: [],
  },
];

const CRIT = {
  alta: { label: "Alta", color: "var(--alta)", bg: "var(--alta-bg)" },
  media: { label: "Media", color: "var(--media)", bg: "var(--media-bg)" },
  baja: { label: "Baja", color: "var(--baja)", bg: "var(--baja-bg)" },
};
const ESTADO = {
  pendiente: "Pendiente", en_progreso: "En curso", finalizada: "Finalizada",
};

const uid = (p) => `${p}-${Date.now()}-${Math.floor(Math.random() * 9999)}`;
const fmtDate = (iso) => new Date(iso + "T00:00:00").toLocaleDateString("es-AR", { day: "2-digit", month: "short", year: "numeric" });
const daysLeft = (iso) => Math.ceil((new Date(iso + "T00:00:00") - new Date(new Date().toDateString())) / 86400000);

/* ---------------------------------------------------------------
   Persistencia
----------------------------------------------------------------*/
const STORAGE_KEY = "unju-seguimiento-estado";

async function loadState() {
  try {
    const r = await window.storage?.get(STORAGE_KEY, false);
    if (r?.value) return JSON.parse(r.value);
  } catch (e) { /* no hay estado guardado todavía */ }
  return null;
}
async function saveState(state) {
  try {
    await window.storage?.set(STORAGE_KEY, JSON.stringify(state), false);
  } catch (e) { /* almacenamiento no disponible */ }
}

/* ---------------------------------------------------------------
   Sub-componentes
----------------------------------------------------------------*/
function StampBadge({ crit, size = "normal" }) {
  const c = CRIT[crit];
  return (
    <span className="unju-stamp" style={{ color: c.color, borderColor: c.color, background: c.bg, fontSize: size === "small" ? 9 : 10 }}>
      <Stamp size={size === "small" ? 10 : 11} /> {c.label}
    </span>
  );
}

function EstadoPill({ estado }) {
  const map = {
    pendiente: { color: "var(--ink-soft)", bg: "var(--paper-dark)" },
    en_progreso: { color: "var(--navy)", bg: "#DCE3E8" },
    finalizada: { color: "var(--baja)", bg: "var(--baja-bg)" },
  };
  const s = map[estado];
  return (
    <span className="font-mono" style={{ color: s.color, background: s.bg, fontSize: 10, padding: "3px 7px", borderRadius: 2, letterSpacing: ".04em" }}>
      {ESTADO[estado].toUpperCase()}
    </span>
  );
}

function ProgressBar({ value }) {
  return (
    <div style={{ background: "var(--paper-dark)", height: 6, borderRadius: 3, overflow: "hidden", border: "1px solid var(--line)" }}>
      <div style={{ width: `${value}%`, height: "100%", background: "var(--navy)" }} />
    </div>
  );
}

function Header({ user, areas, onLogout }) {
  const roleLabel = { admin: "Administrador/a", director: "Dirección", area: "Responsable de área" }[user.rol];
  const areaName = user.areaId ? areas.find((a) => a.id === user.areaId)?.nombre : null;
  return (
    <div className="unju-card" style={{ borderLeft: "5px solid var(--seal)" }}>
      <div className="flex items-center justify-between px-4 py-3">
        <div>
          <div className="font-display" style={{ fontSize: 17, fontWeight: 600 }}>Secretaría de Administración</div>
          <div className="font-mono" style={{ fontSize: 11, color: "var(--ink-soft)" }}>
            UNJU · {roleLabel}{areaName ? ` · ${areaName}` : ""} · {user.nombre}
          </div>
        </div>
        <button onClick={onLogout} className="flex items-center gap-1 px-2 py-1.5" style={{ color: "var(--ink-soft)" }}>
          <LogOut size={16} />
        </button>
      </div>
    </div>
  );
}

function BottomNav({ tabs, active, onChange }) {
  return (
    <div className="unju-card" style={{ position: "sticky", bottom: 0, borderTop: "1px solid var(--line)" }}>
      <div className="flex">
        {tabs.map((t) => (
          <button
            key={t.key}
            onClick={() => onChange(t.key)}
            className="flex-1 flex flex-col items-center gap-1 py-2.5"
            style={{ color: active === t.key ? "var(--navy)" : "var(--ink-soft)", borderTop: active === t.key ? "2px solid var(--navy)" : "2px solid transparent" }}
          >
            <t.icon size={19} strokeWidth={active === t.key ? 2.4 : 1.8} />
            <span className="font-mono" style={{ fontSize: 9.5, letterSpacing: ".03em" }}>{t.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Login
----------------------------------------------------------------*/
function Login({ users, areas, onLogin }) {
  const [usuario, setUsuario] = useState("");
  const [clave, setClave] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    const u = users.find((x) => x.usuario === usuario.trim());
    if (!u || u.clave !== clave) { setError("Usuario o contraseña incorrectos."); return; }
    setError("");
    onLogin(u);
  };

  return (
    <div className="unju-root flex items-center justify-center" style={{ minHeight: "100vh", padding: 20 }}>
      <div className="unju-card w-full" style={{ maxWidth: 380, borderTop: "5px solid var(--seal)", padding: 28 }}>
        <div className="flex items-center gap-2 mb-1" style={{ color: "var(--seal)" }}>
          <Stamp size={20} />
          <span className="font-mono" style={{ fontSize: 11, letterSpacing: ".08em" }}>UNJU · REGISTRO DE GESTIÓN</span>
        </div>
        <h1 className="font-display" style={{ fontSize: 26, fontWeight: 700, marginTop: 4, marginBottom: 2 }}>
          Secretaría de<br />Administración
        </h1>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 20 }}>Seguimiento de actividades por área</p>

        <form onSubmit={submit} className="flex flex-col gap-3">
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)", letterSpacing: ".05em" }}>USUARIO</label>
            <input className="unju-input w-full px-3 py-2 mt-1" style={{ fontSize: 14 }} value={usuario} onChange={(e) => setUsuario(e.target.value)} placeholder="p. ej. efarfan" />
          </div>
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)", letterSpacing: ".05em" }}>CONTRASEÑA</label>
            <input type="password" className="unju-input w-full px-3 py-2 mt-1" style={{ fontSize: 14 }} value={clave} onChange={(e) => setClave(e.target.value)} placeholder="••••" />
          </div>
          {error && <div style={{ color: "var(--alta)", fontSize: 12.5 }}>{error}</div>}
          <button type="submit" className="font-mono mt-1" style={{ background: "var(--navy)", color: "var(--paper)", padding: "10px 0", fontSize: 12, letterSpacing: ".05em" }}>
            INGRESAR
          </button>
        </form>

        <div className="unju-divider mt-5 pt-4">
          <div className="font-mono mb-2" style={{ fontSize: 10, color: "var(--ink-soft)", letterSpacing: ".05em" }}>ACCESOS DE PRUEBA (clave: 1234)</div>
          <div className="flex flex-col gap-1.5">
            {users.map((u) => (
              <button key={u.id} onClick={() => { setUsuario(u.usuario); setClave("1234"); }}
                className="flex items-center justify-between px-2.5 py-1.5" style={{ background: "var(--paper)", border: "1px solid var(--line)", fontSize: 12 }}>
                <span>{u.nombre}</span>
                <span className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>{u.usuario}</span>
              </button>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Detalle de tarea (modal) — permisos según rol
----------------------------------------------------------------*/
function TaskDetail({ task, area, user, onClose, onUpdate }) {
  const [comentario, setComentario] = useState("");
  const [tipoComentario, setTipoComentario] = useState("comentario");
  const [avance, setAvance] = useState(task.avance);
  const [estado, setEstado] = useState(task.estado);
  const [criticidad, setCriticidad] = useState(task.criticidad);
  const [prorroga, setProrroga] = useState({ fecha: "", motivo: "" });
  const [showProrroga, setShowProrroga] = useState(false);

  const isAdmin = user.rol === "admin";
  const isDirector = user.rol === "director";
  const isArea = user.rol === "area";
  const overdue = daysLeft(task.fechaLimite) < 0 && task.estado !== "finalizada";

  const addComentario = () => {
    if (!comentario.trim()) return;
    const nuevo = { autor: user.nombre, rol: user.rol, fecha: new Date().toISOString().slice(0, 10), texto: comentario.trim(), tipo: isDirector ? tipoComentario : "comentario" };
    onUpdate({ ...task, comentarios: [...task.comentarios, nuevo] });
    setComentario("");
  };

  const applyCriticidad = (c) => {
    setCriticidad(c);
    onUpdate({ ...task, criticidad: c });
  };

  const applyAvanceEstado = () => {
    onUpdate({ ...task, avance: Number(avance), estado });
  };

  const solicitarProrroga = () => {
    if (!prorroga.fecha || !prorroga.motivo.trim()) return;
    onUpdate({
      ...task,
      fechaLimite: prorroga.fecha,
      historial: [...task.historial, { fechaAnterior: task.fechaLimite, fechaNueva: prorroga.fecha, motivo: prorroga.motivo.trim() }],
    });
    setShowProrroga(false);
    setProrroga({ fecha: "", motivo: "" });
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(32,48,61,0.45)", zIndex: 50, display: "flex", alignItems: "flex-end" }} onClick={onClose}>
      <div className="unju-card unju-scrollbar" style={{ width: "100%", maxHeight: "88vh", overflowY: "auto", borderRadius: "10px 10px 0 0", borderBottom: "none" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between px-4 pt-4 pb-2" style={{ borderBottom: "1px solid var(--line)" }}>
          <div style={{ paddingRight: 10 }}>
            <div className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>{area?.nombre} · {fmtDate(task.fechaLimite)}</div>
            <div className="font-display" style={{ fontSize: 18, fontWeight: 600, lineHeight: 1.25 }}>{task.titulo}</div>
          </div>
          <button onClick={onClose}><X size={20} color="var(--ink-soft)" /></button>
        </div>

        <div className="px-4 py-3 flex flex-col gap-4">
          <p style={{ fontSize: 13.5, color: "var(--ink-soft)" }}>{task.descripcion}</p>

          <div className="flex flex-wrap gap-2 items-center">
            <StampBadge crit={criticidad} />
            <EstadoPill estado={estado} />
            {overdue && (
              <span className="unju-stamp" style={{ color: "var(--alta)", borderColor: "var(--alta)", background: "var(--alta-bg)" }}>
                <AlertTriangle size={11} /> Vencida
              </span>
            )}
          </div>

          <div>
            <div className="flex justify-between mb-1">
              <span className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>AVANCE</span>
              <span className="font-mono" style={{ fontSize: 10.5 }}>{avance}%</span>
            </div>
            <ProgressBar value={avance} />
          </div>

          <div className="grid grid-cols-2 gap-3 font-mono" style={{ fontSize: 11.5 }}>
            <div><div style={{ color: "var(--ink-soft)" }}>RESPONSABLE</div><div>{task.responsable}</div></div>
            <div><div style={{ color: "var(--ink-soft)" }}>FECHA LÍMITE</div><div>{fmtDate(task.fechaLimite)}</div></div>
          </div>

          {/* Cambiar criticidad: admin y director */}
          {(isAdmin || isDirector) && (
            <div>
              <div className="font-mono mb-1.5" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>CAMBIAR CRITICIDAD</div>
              <div className="flex gap-2">
                {Object.keys(CRIT).map((c) => (
                  <button key={c} onClick={() => applyCriticidad(c)}
                    className="unju-stamp" style={{
                      color: CRIT[c].color, borderColor: CRIT[c].color,
                      background: criticidad === c ? CRIT[c].bg : "transparent", opacity: criticidad === c ? 1 : 0.55,
                    }}>
                    <Stamp size={11} /> {CRIT[c].label}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Avance/estado: admin y responsable de área */}
          {(isAdmin || isArea) && (
            <div className="unju-divider pt-3">
              <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>ACTUALIZAR AVANCE</div>
              <div className="flex items-center gap-3 mb-2">
                <input type="range" min={0} max={100} value={avance} onChange={(e) => setAvance(e.target.value)} style={{ flex: 1 }} />
                <span className="font-mono" style={{ fontSize: 12, width: 34 }}>{avance}%</span>
              </div>
              <select className="unju-input px-2 py-1.5 w-full mb-2" style={{ fontSize: 13 }} value={estado} onChange={(e) => setEstado(e.target.value)}>
                {Object.entries(ESTADO).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
              <button onClick={applyAvanceEstado} className="font-mono w-full" style={{ background: "var(--navy)", color: "var(--paper)", padding: "8px 0", fontSize: 11.5 }}>
                GUARDAR AVANCE
              </button>

              {!showProrroga ? (
                <button onClick={() => setShowProrroga(true)} className="flex items-center gap-1.5 mt-3" style={{ color: "var(--gold)", fontSize: 12.5 }}>
                  <CalendarClock size={14} /> Solicitar nueva fecha límite
                </button>
              ) : (
                <div className="mt-3" style={{ background: "var(--paper)", border: "1px solid var(--line)", padding: 10 }}>
                  <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>SOLICITAR PRÓRROGA</div>
                  <input type="date" className="unju-input w-full px-2 py-1.5 mb-2" style={{ fontSize: 13 }} value={prorroga.fecha} onChange={(e) => setProrroga({ ...prorroga, fecha: e.target.value })} />
                  <textarea className="unju-input w-full px-2 py-1.5 mb-2" style={{ fontSize: 13 }} rows={2} placeholder="Motivo del pedido de nueva fecha…" value={prorroga.motivo} onChange={(e) => setProrroga({ ...prorroga, motivo: e.target.value })} />
                  <div className="flex gap-2">
                    <button onClick={solicitarProrroga} className="font-mono flex-1" style={{ background: "var(--gold)", color: "#fff", padding: "7px 0", fontSize: 11 }}>CONFIRMAR</button>
                    <button onClick={() => setShowProrroga(false)} className="font-mono" style={{ color: "var(--ink-soft)", fontSize: 11, padding: "7px 10px" }}>CANCELAR</button>
                  </div>
                </div>
              )}
            </div>
          )}

          {task.historial.length > 0 && (
            <div className="unju-divider pt-3">
              <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>HISTORIAL DE FECHAS</div>
              {task.historial.map((h, i) => (
                <div key={i} style={{ fontSize: 12.5, marginBottom: 6 }}>
                  <span className="font-mono" style={{ color: "var(--ink-soft)" }}>{fmtDate(h.fechaAnterior)} → {fmtDate(h.fechaNueva)}</span>
                  <div style={{ color: "var(--ink-soft)" }}>{h.motivo}</div>
                </div>
              ))}
            </div>
          )}

          <div className="unju-divider pt-3">
            <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>COMENTARIOS Y ALERTAS</div>
            <div className="flex flex-col gap-2 mb-3">
              {task.comentarios.length === 0 && <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>Sin comentarios todavía.</div>}
              {task.comentarios.map((c, i) => (
                <div key={i} style={{ background: c.tipo === "alerta" ? "var(--alta-bg)" : "var(--paper)", border: "1px solid var(--line)", padding: "7px 9px" }}>
                  <div className="flex items-center justify-between">
                    <span className="font-mono" style={{ fontSize: 10.5, fontWeight: 600, color: c.tipo === "alerta" ? "var(--alta)" : "var(--ink)" }}>
                      {c.tipo === "alerta" ? "⚠ ALERTA — " : ""}{c.autor}
                    </span>
                    <span className="font-mono" style={{ fontSize: 9.5, color: "var(--ink-soft)" }}>{fmtDate(c.fecha)}</span>
                  </div>
                  <div style={{ fontSize: 12.5, marginTop: 2 }}>{c.texto}</div>
                </div>
              ))}
            </div>
            <div className="flex flex-col gap-2">
              {isDirector && (
                <div className="flex gap-2">
                  <button onClick={() => setTipoComentario("comentario")} className="font-mono" style={{ fontSize: 10.5, padding: "5px 9px", background: tipoComentario === "comentario" ? "var(--navy)" : "var(--paper)", color: tipoComentario === "comentario" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>COMENTARIO</button>
                  <button onClick={() => setTipoComentario("alerta")} className="font-mono" style={{ fontSize: 10.5, padding: "5px 9px", background: tipoComentario === "alerta" ? "var(--alta)" : "var(--paper)", color: tipoComentario === "alerta" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>ALERTA</button>
                </div>
              )}
              <textarea className="unju-input w-full px-2.5 py-2" style={{ fontSize: 13 }} rows={2} placeholder="Escribir un comentario…" value={comentario} onChange={(e) => setComentario(e.target.value)} />
              <button onClick={addComentario} className="font-mono flex items-center justify-center gap-1.5" style={{ background: "var(--navy)", color: "var(--paper)", padding: "8px 0", fontSize: 11.5 }}>
                <MessageSquare size={13} /> PUBLICAR
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Formulario de tarea (crear/editar) — admin
----------------------------------------------------------------*/
function TaskForm({ areas, users, initial, onSave, onClose }) {
  const [form, setForm] = useState(initial || {
    areaId: areas[0]?.id || "", titulo: "", descripcion: "", responsable: "",
    criticidad: "media", estado: "pendiente", fechaLimite: addDays(7), avance: 0,
  });
  const areaUsers = users.filter((u) => u.rol === "area" && u.areaId === form.areaId);

  const submit = (e) => {
    e.preventDefault();
    if (!form.titulo.trim() || !form.responsable.trim() || !form.areaId) return;
    onSave(form);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(32,48,61,0.45)", zIndex: 50, display: "flex", alignItems: "flex-end" }} onClick={onClose}>
      <div className="unju-card unju-scrollbar" style={{ width: "100%", maxHeight: "88vh", overflowY: "auto", borderRadius: "10px 10px 0 0" }} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 pt-4 pb-3" style={{ borderBottom: "1px solid var(--line)" }}>
          <div className="font-display" style={{ fontSize: 17, fontWeight: 600 }}>{initial ? "Editar tarea" : "Nueva tarea"}</div>
          <button onClick={onClose}><X size={20} color="var(--ink-soft)" /></button>
        </div>
        <form onSubmit={submit} className="px-4 py-3 flex flex-col gap-3">
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>ÁREA</label>
            <select className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.areaId} onChange={(e) => setForm({ ...form, areaId: e.target.value, responsable: "" })}>
              {areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}
            </select>
          </div>
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>TÍTULO</label>
            <input className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.titulo} onChange={(e) => setForm({ ...form, titulo: e.target.value })} />
          </div>
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>DESCRIPCIÓN</label>
            <textarea rows={2} className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.descripcion} onChange={(e) => setForm({ ...form, descripcion: e.target.value })} />
          </div>
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>RESPONSABLE</label>
            <select className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.responsable} onChange={(e) => setForm({ ...form, responsable: e.target.value })}>
              <option value="">Seleccionar…</option>
              {areaUsers.map((u) => <option key={u.id} value={u.nombre}>{u.nombre}</option>)}
            </select>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>CRITICIDAD</label>
              <select className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.criticidad} onChange={(e) => setForm({ ...form, criticidad: e.target.value })}>
                {Object.entries(CRIT).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
              </select>
            </div>
            <div>
              <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>ESTADO</label>
              <select className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.estado} onChange={(e) => setForm({ ...form, estado: e.target.value })}>
                {Object.entries(ESTADO).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            </div>
          </div>
          <div>
            <label className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>FECHA LÍMITE</label>
            <input type="date" className="unju-input w-full px-2.5 py-2 mt-1" style={{ fontSize: 13 }} value={form.fechaLimite} onChange={(e) => setForm({ ...form, fechaLimite: e.target.value })} />
          </div>
          <button type="submit" className="font-mono mt-1" style={{ background: "var(--navy)", color: "var(--paper)", padding: "10px 0", fontSize: 12 }}>
            {initial ? "GUARDAR CAMBIOS" : "CREAR TAREA"}
          </button>
        </form>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Lista de tareas (reutilizable)
----------------------------------------------------------------*/
function TaskRow({ task, area, onClick }) {
  const dl = daysLeft(task.fechaLimite);
  const overdue = dl < 0 && task.estado !== "finalizada";
  return (
    <button onClick={onClick} className="unju-card w-full text-left px-3.5 py-3 flex items-center gap-3"
      style={{ borderLeft: `4px solid ${CRIT[task.criticidad].color}` }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div className="flex items-center gap-2 flex-wrap mb-1">
          <span className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>{area?.nombre}</span>
          <EstadoPill estado={task.estado} />
        </div>
        <div style={{ fontSize: 13.5, fontWeight: 600, lineHeight: 1.25 }}>{task.titulo}</div>
        <div className="flex items-center gap-2 mt-1.5">
          <div style={{ flex: 1, maxWidth: 120 }}><ProgressBar value={task.avance} /></div>
          <span className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>{task.avance}%</span>
        </div>
        <div className="flex items-center gap-2 mt-1.5">
          <StampBadge crit={task.criticidad} size="small" />
          <span className="font-mono" style={{ fontSize: 10, color: overdue ? "var(--alta)" : "var(--ink-soft)" }}>
            {overdue ? `Vencida hace ${Math.abs(dl)}d` : `Vence en ${dl}d — ${fmtDate(task.fechaLimite)}`}
          </span>
        </div>
      </div>
      <ChevronRight size={16} color="var(--ink-soft)" />
    </button>
  );
}

/* ---------------------------------------------------------------
   Resumen / gráficos (compartido admin+director)
----------------------------------------------------------------*/
function ResumenView({ tasks, areas }) {
  const critData = Object.keys(CRIT).map((c) => ({ name: CRIT[c].label, value: tasks.filter((t) => t.criticidad === c).length, color: CRIT[c].color }));
  const areaData = areas.map((a) => ({ name: a.nombre.split(" ")[0], total: tasks.filter((t) => t.areaId === a.id).length }));
  const vencidas = tasks.filter((t) => daysLeft(t.fechaLimite) < 0 && t.estado !== "finalizada").length;
  const finalizadas = tasks.filter((t) => t.estado === "finalizada").length;
  const enCurso = tasks.filter((t) => t.estado === "en_progreso").length;

  const kpis = [
    { label: "Total tareas", value: tasks.length, icon: ClipboardList },
    { label: "En curso", value: enCurso, icon: Clock },
    { label: "Vencidas", value: vencidas, icon: AlertTriangle },
    { label: "Finalizadas", value: finalizadas, icon: CheckCircle2 },
  ];

  return (
    <div className="flex flex-col gap-4 px-4 py-4">
      <div className="grid grid-cols-2 gap-3">
        {kpis.map((k) => (
          <div key={k.label} className="unju-card px-3.5 py-3">
            <k.icon size={16} color="var(--ink-soft)" />
            <div className="font-display" style={{ fontSize: 24, fontWeight: 600, marginTop: 4 }}>{k.value}</div>
            <div className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>{k.label.toUpperCase()}</div>
          </div>
        ))}
      </div>

      <div className="unju-card px-3.5 py-3">
        <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>TAREAS POR CRITICIDAD</div>
        <ResponsiveContainer width="100%" height={180}>
          <PieChart>
            <Pie data={critData} dataKey="value" nameKey="name" innerRadius={40} outerRadius={65} paddingAngle={3}>
              {critData.map((d, i) => <Cell key={i} fill={d.color} />)}
            </Pie>
            <Tooltip />
            <Legend wrapperStyle={{ fontSize: 11 }} />
          </PieChart>
        </ResponsiveContainer>
      </div>

      <div className="unju-card px-3.5 py-3">
        <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>TAREAS POR ÁREA</div>
        <ResponsiveContainer width="100%" height={190}>
          <BarChart data={areaData}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" />
            <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-20} textAnchor="end" height={50} />
            <YAxis allowDecimals={false} tick={{ fontSize: 10 }} />
            <Tooltip />
            <Bar dataKey="total" fill="var(--navy)" radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Línea de tiempo (director)
----------------------------------------------------------------*/
function TimelineView({ tasks, areas, onOpen }) {
  const sorted = [...tasks].filter(t => t.estado !== "finalizada").sort((a, b) => a.fechaLimite.localeCompare(b.fechaLimite));
  return (
    <div className="px-4 py-4">
      <div className="font-mono mb-3" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>LÍNEA DE TIEMPO — TAREAS ABIERTAS</div>
      <div style={{ position: "relative", paddingLeft: 18 }}>
        <div style={{ position: "absolute", left: 5, top: 4, bottom: 4, width: 2, background: "var(--line)" }} />
        <div className="flex flex-col gap-3">
          {sorted.map((t) => {
            const area = areas.find((a) => a.id === t.areaId);
            const overdue = daysLeft(t.fechaLimite) < 0;
            return (
              <div key={t.id} style={{ position: "relative" }}>
                <div style={{ position: "absolute", left: -18, top: 6, width: 10, height: 10, borderRadius: "50%", background: CRIT[t.criticidad].color, border: "2px solid var(--paper)" }} />
                <button onClick={() => onOpen(t)} className="unju-card w-full text-left px-3 py-2.5">
                  <div className="flex items-center justify-between">
                    <span className="font-mono" style={{ fontSize: 10.5, color: overdue ? "var(--alta)" : "var(--ink-soft)" }}>{fmtDate(t.fechaLimite)}</span>
                    <StampBadge crit={t.criticidad} size="small" />
                  </div>
                  <div style={{ fontSize: 13, fontWeight: 600, marginTop: 2 }}>{t.titulo}</div>
                  <div className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)", marginTop: 2 }}>{area?.nombre} · {t.responsable}</div>
                </button>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Exportar (admin)
----------------------------------------------------------------*/
function exportExcel(tasks, areas) {
  const rows = tasks.map((t) => ({
    Área: areas.find((a) => a.id === t.areaId)?.nombre || "",
    Tarea: t.titulo,
    Responsable: t.responsable,
    Criticidad: CRIT[t.criticidad].label,
    Estado: ESTADO[t.estado],
    "Fecha límite": fmtDate(t.fechaLimite),
    "Avance (%)": t.avance,
    "Reprogramaciones": t.historial.length,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 16 }, { wch: 40 }, { wch: 18 }, { wch: 10 }, { wch: 12 }, { wch: 13 }, { wch: 11 }, { wch: 15 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tareas");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "tareas_secretaria_administracion.xlsx"; a.click();
  URL.revokeObjectURL(url);
}

function exportWord(tasks, areas) {
  const rows = tasks.map((t) => `
    <tr>
      <td>${areas.find((a) => a.id === t.areaId)?.nombre || ""}</td>
      <td>${t.titulo}</td>
      <td>${t.responsable}</td>
      <td>${CRIT[t.criticidad].label}</td>
      <td>${ESTADO[t.estado]}</td>
      <td>${fmtDate(t.fechaLimite)}</td>
      <td>${t.avance}%</td>
    </tr>`).join("");
  const html = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset="utf-8"><title>Reporte de tareas</title></head>
    <body style="font-family:Calibri,sans-serif;">
      <h2>Secretaría de Administración — UNJU</h2>
      <p>Reporte de actividades por área — ${new Date().toLocaleDateString("es-AR")}</p>
      <table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;font-size:11px;">
        <tr style="background:#22303F;color:#fff;">
          <th>Área</th><th>Tarea</th><th>Responsable</th><th>Criticidad</th><th>Estado</th><th>Fecha límite</th><th>Avance</th>
        </tr>
        ${rows}
      </table>
    </body></html>`;
  const blob = new Blob(["\ufeff", html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "reporte_secretaria_administracion.doc"; a.click();
  URL.revokeObjectURL(url);
}

/* ---------------------------------------------------------------
   Panel Administrador
----------------------------------------------------------------*/
function AdminPanel({ tab, tasks, areas, users, setTasks, setAreas, setUsers, openTask }) {
  const [showTaskForm, setShowTaskForm] = useState(false);
  const [editTask, setEditTask] = useState(null);
  const [newArea, setNewArea] = useState("");
  const [newUser, setNewUser] = useState({ nombre: "", usuario: "", clave: "1234", rol: "area", areaId: areas[0]?.id || "" });

  const saveTask = (form) => {
    if (editTask) {
      setTasks(tasks.map((t) => (t.id === editTask.id ? { ...t, ...form } : t)));
    } else {
      setTasks([...tasks, { ...form, id: uid("t"), comentarios: [], historial: [] }]);
    }
    setShowTaskForm(false); setEditTask(null);
  };

  if (tab === "resumen") return <ResumenView tasks={tasks} areas={areas} />;

  if (tab === "tareas") return (
    <div className="px-4 py-4">
      <div className="flex items-center justify-between mb-3">
        <div className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>{tasks.length} TAREAS</div>
        <button onClick={() => { setEditTask(null); setShowTaskForm(true); }} className="font-mono flex items-center gap-1" style={{ background: "var(--navy)", color: "var(--paper)", padding: "6px 10px", fontSize: 11 }}>
          <Plus size={13} /> NUEVA TAREA
        </button>
      </div>
      <div className="flex flex-col gap-2.5">
        {tasks.map((t) => (
          <div key={t.id} className="flex items-stretch gap-2">
            <div style={{ flex: 1 }}><TaskRow task={t} area={areas.find(a=>a.id===t.areaId)} onClick={() => openTask(t)} /></div>
            <button onClick={() => { setEditTask(t); setShowTaskForm(true); }} className="unju-card px-2.5" style={{ color: "var(--ink-soft)" }}><Pencil size={14} /></button>
          </div>
        ))}
      </div>
      {showTaskForm && <TaskForm areas={areas} users={users} initial={editTask} onSave={saveTask} onClose={() => { setShowTaskForm(false); setEditTask(null); }} />}
    </div>
  );

  if (tab === "areas") return (
    <div className="px-4 py-4">
      <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>ÁREAS ({areas.length})</div>
      <div className="flex flex-col gap-2 mb-4">
        {areas.map((a) => (
          <div key={a.id} className="unju-card px-3.5 py-2.5 flex items-center justify-between">
            <span style={{ fontSize: 13.5 }}>{a.nombre}</span>
            <span className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>{tasks.filter(t => t.areaId === a.id).length} tareas</span>
          </div>
        ))}
      </div>
      <div className="unju-card px-3.5 py-3">
        <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>CREAR NUEVA ÁREA</div>
        <div className="flex gap-2">
          <input className="unju-input flex-1 px-2.5 py-2" style={{ fontSize: 13 }} placeholder="Nombre del área" value={newArea} onChange={(e) => setNewArea(e.target.value)} />
          <button onClick={() => { if (newArea.trim()) { setAreas([...areas, { id: uid("area"), nombre: newArea.trim() }]); setNewArea(""); } }} className="font-mono" style={{ background: "var(--navy)", color: "var(--paper)", padding: "0 14px", fontSize: 11 }}>CREAR</button>
        </div>
      </div>
    </div>
  );

  if (tab === "usuarios") return (
    <div className="px-4 py-4">
      <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>USUARIOS ({users.length})</div>
      <div className="flex flex-col gap-2 mb-4">
        {users.map((u) => (
          <div key={u.id} className="unju-card px-3.5 py-2.5 flex items-center justify-between">
            <div>
              <div style={{ fontSize: 13.5 }}>{u.nombre}</div>
              <div className="font-mono" style={{ fontSize: 10, color: "var(--ink-soft)" }}>
                {u.usuario} · {{ admin: "Administrador", director: "Dirección", area: "Responsable" }[u.rol]}
                {u.areaId ? ` · ${areas.find(a => a.id === u.areaId)?.nombre}` : ""}
              </div>
            </div>
          </div>
        ))}
      </div>
      <div className="unju-card px-3.5 py-3">
        <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>CREAR USUARIO</div>
        <div className="flex flex-col gap-2">
          <input className="unju-input px-2.5 py-2" style={{ fontSize: 13 }} placeholder="Nombre completo" value={newUser.nombre} onChange={(e) => setNewUser({ ...newUser, nombre: e.target.value })} />
          <div className="grid grid-cols-2 gap-2">
            <input className="unju-input px-2.5 py-2" style={{ fontSize: 13 }} placeholder="Usuario" value={newUser.usuario} onChange={(e) => setNewUser({ ...newUser, usuario: e.target.value })} />
            <input className="unju-input px-2.5 py-2" style={{ fontSize: 13 }} placeholder="Contraseña" value={newUser.clave} onChange={(e) => setNewUser({ ...newUser, clave: e.target.value })} />
          </div>
          <div className="grid grid-cols-2 gap-2">
            <select className="unju-input px-2.5 py-2" style={{ fontSize: 13 }} value={newUser.rol} onChange={(e) => setNewUser({ ...newUser, rol: e.target.value })}>
              <option value="area">Responsable de área</option>
              <option value="director">Dirección</option>
              <option value="admin">Administrador</option>
            </select>
            <select className="unju-input px-2.5 py-2" style={{ fontSize: 13 }} disabled={newUser.rol !== "area"} value={newUser.areaId} onChange={(e) => setNewUser({ ...newUser, areaId: e.target.value })}>
              {areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}
            </select>
          </div>
          <button onClick={() => {
            if (!newUser.nombre.trim() || !newUser.usuario.trim()) return;
            setUsers([...users, { ...newUser, id: uid("u"), areaId: newUser.rol === "area" ? newUser.areaId : null }]);
            setNewUser({ nombre: "", usuario: "", clave: "1234", rol: "area", areaId: areas[0]?.id || "" });
          }} className="font-mono" style={{ background: "var(--navy)", color: "var(--paper)", padding: "9px 0", fontSize: 11.5 }}>CREAR USUARIO</button>
        </div>
      </div>
    </div>
  );

  if (tab === "exportar") return (
    <div className="px-4 py-4 flex flex-col gap-3">
      <div className="font-mono mb-1" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>DESCARGAR REPORTE ({tasks.length} tareas)</div>
      <button onClick={() => exportExcel(tasks, areas)} className="unju-card flex items-center gap-3 px-4 py-3.5 text-left">
        <FileSpreadsheet size={22} color="var(--baja)" />
        <div>
          <div style={{ fontSize: 14, fontWeight: 600 }}>Exportar a Excel</div>
          <div className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>.xlsx — planilla completa de tareas</div>
        </div>
      </button>
      <button onClick={() => exportWord(tasks, areas)} className="unju-card flex items-center gap-3 px-4 py-3.5 text-left">
        <FileText size={22} color="var(--navy)" />
        <div>
          <div style={{ fontSize: 14, fontWeight: 600 }}>Exportar a Word</div>
          <div className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>.doc — reporte con formato de tabla</div>
        </div>
      </button>
    </div>
  );

  return null;
}

/* ---------------------------------------------------------------
   Panel Dirección
----------------------------------------------------------------*/
function DirectorPanel({ tab, tasks, areas, openTask }) {
  if (tab === "resumen") return <ResumenView tasks={tasks} areas={areas} />;
  if (tab === "timeline") return <TimelineView tasks={tasks} areas={areas} onOpen={openTask} />;
  if (tab === "tareas") return (
    <div className="px-4 py-4 flex flex-col gap-2.5">
      <div className="font-mono mb-1" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>{tasks.length} TAREAS · TODAS LAS ÁREAS</div>
      {tasks.map((t) => <TaskRow key={t.id} task={t} area={areas.find(a=>a.id===t.areaId)} onClick={() => openTask(t)} />)}
    </div>
  );
  return null;
}

/* ---------------------------------------------------------------
   Panel de Área (responsable)
----------------------------------------------------------------*/
function AreaPanel({ tasks, areas, user, openTask }) {
  const mine = tasks.filter((t) => t.responsable === user.nombre);
  const area = areas.find((a) => a.id === user.areaId);
  const pendientes = mine.filter(t => t.estado !== "finalizada").length;
  return (
    <div className="px-4 py-4">
      <div className="unju-card px-3.5 py-3 mb-4" style={{ borderLeft: "4px solid var(--gold)" }}>
        <div className="font-mono" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>ÁREA</div>
        <div className="font-display" style={{ fontSize: 17, fontWeight: 600 }}>{area?.nombre}</div>
        <div className="font-mono mt-1" style={{ fontSize: 11, color: "var(--ink-soft)" }}>{pendientes} tareas abiertas de {mine.length}</div>
      </div>
      <div className="font-mono mb-2" style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>MIS TAREAS</div>
      <div className="flex flex-col gap-2.5">
        {mine.length === 0 && <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>No tenés tareas asignadas por el momento.</div>}
        {mine.map((t) => <TaskRow key={t.id} task={t} area={area} onClick={() => openTask(t)} />)}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   App
----------------------------------------------------------------*/
export default function App() {
  const [ready, setReady] = useState(false);
  const [users, setUsers] = useState(USER_SEED);
  const [areas, setAreas] = useState(AREA_SEED);
  const [tasks, setTasks] = useState(TASK_SEED);
  const [user, setUser] = useState(null);
  const [tab, setTab] = useState("resumen");
  const [openTaskObj, setOpenTaskObj] = useState(null);

  useEffect(() => {
    (async () => {
      const saved = await loadState();
      if (saved) {
        setUsers(saved.users || USER_SEED);
        setAreas(saved.areas || AREA_SEED);
        setTasks(saved.tasks || TASK_SEED);
      }
      setReady(true);
    })();
  }, []);

  useEffect(() => {
    if (!ready) return;
    saveState({ users, areas, tasks });
  }, [users, areas, tasks, ready]);

  const openTask = useCallback((t) => setOpenTaskObj(t), []);
  const updateTask = useCallback((updated) => {
    setTasks((prev) => prev.map((t) => (t.id === updated.id ? updated : t)));
    setOpenTaskObj(updated);
  }, []);

  const login = (u) => {
    setUser(u);
    setTab(u.rol === "area" ? "mis-tareas" : "resumen");
  };

  if (!ready) return null;

  if (!user) {
    return (
      <div className="unju-root">
        <style>{TOKENS}</style>
        <Login users={users} areas={areas} onLogin={login} />
      </div>
    );
  }

  const tabsByRole = {
    admin: [
      { key: "resumen", label: "Resumen", icon: LayoutDashboard },
      { key: "areas", label: "Áreas", icon: Building2 },
      { key: "usuarios", label: "Usuarios", icon: Users },
      { key: "tareas", label: "Tareas", icon: ClipboardList },
      { key: "exportar", label: "Exportar", icon: Download },
    ],
    director: [
      { key: "resumen", label: "Resumen", icon: LayoutDashboard },
      { key: "timeline", label: "Línea de tiempo", icon: CalendarClock },
      { key: "tareas", label: "Tareas", icon: ClipboardList },
    ],
    area: [
      { key: "mis-tareas", label: "Mis tareas", icon: ClipboardList },
    ],
  };

  const currentTask = openTaskObj ? tasks.find((t) => t.id === openTaskObj.id) : null;

  return (
    <div className="unju-root" style={{ display: "flex", flexDirection: "column" }}>
      <style>{TOKENS}</style>
      <div style={{ maxWidth: 480, margin: "0 auto", width: "100%", display: "flex", flexDirection: "column", minHeight: "100vh" }}>
        <Header user={user} areas={areas} onLogout={() => setUser(null)} />

        <div style={{ flex: 1, overflowY: "auto" }} className="unju-scrollbar">
          {user.rol === "admin" && (
            <AdminPanel tab={tab} tasks={tasks} areas={areas} users={users} setTasks={setTasks} setAreas={setAreas} setUsers={setUsers} openTask={openTask} />
          )}
          {user.rol === "director" && (
            <DirectorPanel tab={tab} tasks={tasks} areas={areas} openTask={openTask} />
          )}
          {user.rol === "area" && (
            <AreaPanel tasks={tasks} areas={areas} user={user} openTask={openTask} />
          )}
        </div>

        <BottomNav tabs={tabsByRole[user.rol]} active={tab} onChange={setTab} />
      </div>

      {currentTask && (
        <TaskDetail
          task={currentTask}
          area={areas.find((a) => a.id === currentTask.areaId)}
          user={user}
          onClose={() => setOpenTaskObj(null)}
          onUpdate={updateTask}
        />
      )}
    </div>
  );
}
