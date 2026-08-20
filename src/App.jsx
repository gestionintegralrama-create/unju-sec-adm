import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell,
} from "recharts";
import * as XLSX from "xlsx";
import { supabase, supabaseConfigured } from "./supabaseClient";

/* ============================================================
   RECTORADO UNJU — Sistema de gestión de tareas
   Paleta: institucional, con guiño a las franjas de color de la
   Quebrada de Humahuaca como identidad visual de la línea de tiempo.
   ============================================================ */

const STORAGE_KEY = "rectorado-unju-tareas-db";

const PALETTE = {
  bg: "#EDEAE3",
  surface: "#FFFFFF",
  surfaceSoft: "#F5F3EE",
  ink: "#1B2A4A",
  inkSoft: "#5B6478",
  line: "#DBD6C9",
  primary: "#1B2A4A",
  secondary: "#3D5A80",
  accent: "#C9862F",
  alert: "#B3473F",
  positive: "#4C7A5D",
  pending: "#9B9587",
};

const STATUS = {
  pendiente: { label: "Pendiente", color: PALETTE.pending },
  en_curso: { label: "En curso", color: PALETTE.secondary },
  demorada: { label: "Demorada", color: PALETTE.alert },
  finalizada_a_tiempo: { label: "Finalizada a tiempo", color: PALETTE.positive },
  finalizada_con_demora: { label: "Finalizada con demora", color: PALETTE.accent },
};

const PRIORITY = {
  alta: { label: "Alta", color: PALETTE.alert },
  media: { label: "Media", color: PALETTE.accent },
  baja: { label: "Baja", color: PALETTE.secondary },
};

const ROLE_LABEL = { area: "Área", admin: "Administración", direccion: "Dirección" };

function uid(prefix) {
  return prefix + "_" + Math.random().toString(36).slice(2, 9);
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function fmt(dateStr) {
  if (!dateStr) return "—";
  const [y, m, d] = dateStr.split("-");
  return `${d}/${m}/${y}`;
}

function computeStatus(task) {
  const today = todayStr();
  if (task.actualEndDate) {
    return task.actualEndDate <= task.plannedEndDate
      ? "finalizada_a_tiempo"
      : "finalizada_con_demora";
  }
  if (task.plannedEndDate < today) return "demorada";
  if (task.startDate <= today) return "en_curso";
  return "pendiente";
}

/* ---------------- Datos semilla ---------------- */

const SEED_AREAS = [
  { id: "a1", name: "Contabilidad" },
  { id: "a2", name: "Secretarías" },
  { id: "a3", name: "Soporte de Procesos e Informática" },
  { id: "a4", name: "Presupuesto" },
  { id: "a5", name: "Compras" },
];

const SEED_USERS = [
  { id: "u_admin", name: "Administración General", username: "admin", role: "admin", areaId: null },
  { id: "u_dir", name: "Rectorado — Dirección", username: "direccion", role: "direccion", areaId: null },
  { id: "u1", name: "Ana Gómez", username: "agomez", role: "area", areaId: "a1" },
  { id: "u2", name: "Carlos Pérez", username: "cperez", role: "area", areaId: "a1" },
  { id: "u3", name: "Lucía Romero", username: "lromero", role: "area", areaId: "a2" },
  { id: "u4", name: "Martín Sosa", username: "msosa", role: "area", areaId: "a3" },
  { id: "u5", name: "Elena Cruz", username: "ecruz", role: "area", areaId: "a4" },
  { id: "u6", name: "Jorge Vilte", username: "jvilte", role: "area", areaId: "a5" },
];

const SEED_TASKS = [
  { id: "t1", areaId: "a1", userId: "u1", title: "Balance trimestral Q2", description: "Cierre y presentación del balance del segundo trimestre.", priority: "alta", startDate: "2026-07-01", plannedEndDate: "2026-08-10", actualEndDate: "2026-08-09", progress: 100,
    observations: [{ id: uid("o"), author: "Administración General", date: "2026-08-09", text: "Presentado en término, sin observaciones." }], history: [] },
  { id: "t2", areaId: "a1", userId: "u2", title: "Análisis de costos — ingresos de Nación", description: "Seguimiento de la división de ingresos que llegan de Nación a la universidad.", priority: "media", startDate: "2026-07-15", plannedEndDate: "2026-08-15", actualEndDate: null, progress: 70,
    observations: [{ id: uid("o"), author: "Administración General", date: "2026-08-18", text: "Falta el detalle de partidas de agosto. Regularizar a la brevedad." }], history: [] },
  { id: "t3", areaId: "a1", userId: "u1", title: "Conciliación bancaria de agosto", description: "", priority: "baja", startDate: "2026-08-10", plannedEndDate: "2026-09-05", actualEndDate: null, progress: 30, observations: [], history: [] },
  { id: "t4", areaId: "a2", userId: "u3", title: "Actualización del padrón docente", description: "", priority: "alta", startDate: "2026-08-01", plannedEndDate: "2026-08-25", actualEndDate: null, progress: 55, observations: [], history: [] },
  { id: "t5", areaId: "a2", userId: "u3", title: "Digitalización de expedientes", description: "", priority: "media", startDate: "2026-06-01", plannedEndDate: "2026-07-10", actualEndDate: "2026-08-05", progress: 100,
    observations: [], history: [{ id: uid("h"), date: "2026-07-08", field: "plannedEndDate", oldValue: "2026-06-25", newValue: "2026-07-10", motivo: "Reprogramado por falta de personal disponible para el escaneo." }] },
  { id: "t6", areaId: "a3", userId: "u4", title: "Migración del servidor de correo", description: "", priority: "alta", startDate: "2026-08-05", plannedEndDate: "2026-08-30", actualEndDate: null, progress: 40, observations: [], history: [] },
  { id: "t7", areaId: "a4", userId: "u5", title: "Proyección presupuestaria 2027", description: "", priority: "alta", startDate: "2026-08-01", plannedEndDate: "2026-09-15", actualEndDate: null, progress: 20, observations: [], history: [] },
  { id: "t8", areaId: "a4", userId: "u5", title: "Informe de ejecución de julio", description: "", priority: "media", startDate: "2026-07-20", plannedEndDate: "2026-08-05", actualEndDate: null, progress: 85,
    observations: [{ id: uid("o"), author: "Rectorado — Dirección", date: "2026-08-17", text: "Alerta: informe vencido hace más de una semana. Priorizar su cierre." }], history: [] },
  { id: "t9", areaId: "a5", userId: "u6", title: "Licitación de equipamiento informático", description: "", priority: "alta", startDate: "2026-06-15", plannedEndDate: "2026-08-01", actualEndDate: null, progress: 60,
    observations: [{ id: uid("o"), author: "Administración General", date: "2026-08-12", text: "Se solicitó a Compras el estado del expediente ante la demora." }], history: [] },
  { id: "t10", areaId: "a5", userId: "u6", title: "Renovación de contrato de limpieza", description: "", priority: "baja", startDate: "2026-08-15", plannedEndDate: "2026-09-10", actualEndDate: null, progress: 5, observations: [], history: [] },
];

/* ---------------- Conversión filas Supabase <-> objetos de la app ---------------- */

function rowToArea(r) { return { id: r.id, name: r.name }; }
function areaToRow(a) { return { id: a.id, name: a.name }; }

function rowToUser(r) { return { id: r.id, name: r.name, username: r.username, role: r.role, areaId: r.area_id }; }
function userToRow(u) { return { id: u.id, name: u.name, username: u.username, role: u.role, area_id: u.areaId }; }

function rowToTask(r) {
  return {
    id: r.id, areaId: r.area_id, userId: r.user_id, title: r.title, description: r.description || "",
    priority: r.priority, startDate: r.start_date, plannedEndDate: r.planned_end_date,
    actualEndDate: r.actual_end_date, progress: r.progress ?? 0,
    observations: r.observations || [], history: r.history || [],
  };
}
function taskToRow(t) {
  return {
    id: t.id, area_id: t.areaId, user_id: t.userId, title: t.title, description: t.description || "",
    priority: t.priority, start_date: t.startDate, planned_end_date: t.plannedEndDate,
    actual_end_date: t.actualEndDate || null, progress: t.progress ?? 0,
    observations: t.observations || [], history: t.history || [],
  };
}



function Badge({ color, children }) {
  return (
    <span className="badge" style={{ "--c": color }}>{children}</span>
  );
}

function ProgressBar({ value, color }) {
  return (
    <div className="progress-track">
      <div className="progress-fill" style={{ width: `${value}%`, background: color }} />
    </div>
  );
}

function ColorBand() {
  const colors = [PALETTE.primary, PALETTE.secondary, PALETTE.accent, PALETTE.alert, PALETTE.positive];
  return (
    <div className="color-band">
      {colors.map((c, i) => <span key={i} style={{ background: c }} />)}
    </div>
  );
}

/* ---------------- Formulario nueva tarea ---------------- */

function TaskForm({ areas, users, defaultAreaId, defaultUserId, lockAssignment, onCancel, onSubmit }) {
  const [areaId, setAreaId] = useState(defaultAreaId || areas[0]?.id || "");
  const [userId, setUserId] = useState(defaultUserId || "");
  const [title, setTitle] = useState("");
  const [description, setDescription] = useState("");
  const [priority, setPriority] = useState("media");
  const [startDate, setStartDate] = useState(todayStr());
  const [plannedEndDate, setPlannedEndDate] = useState(todayStr());

  const areaUsers = users.filter((u) => u.areaId === areaId);

  useEffect(() => {
    if (!lockAssignment && areaUsers.length && !areaUsers.find((u) => u.id === userId)) {
      setUserId(areaUsers[0].id);
    }
  }, [areaId]); // eslint-disable-line

  function submit(e) {
    e.preventDefault();
    if (!title.trim() || !userId) return;
    onSubmit({ areaId, userId, title: title.trim(), description: description.trim(), priority, startDate, plannedEndDate });
  }

  return (
    <form className="panel form" onSubmit={submit}>
      <h3>Nueva tarea</h3>
      {!lockAssignment && (
        <div className="field-row">
          <label>
            Área
            <select value={areaId} onChange={(e) => setAreaId(e.target.value)}>
              {areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </label>
          <label>
            Responsable
            <select value={userId} onChange={(e) => setUserId(e.target.value)}>
              {areaUsers.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
            </select>
          </label>
        </div>
      )}
      <label>
        Título de la tarea
        <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Ej: Informe mensual de gastos" required />
      </label>
      <label>
        Descripción (opcional)
        <textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={2} />
      </label>
      <div className="field-row">
        <label>
          Prioridad
          <select value={priority} onChange={(e) => setPriority(e.target.value)}>
            <option value="alta">Alta</option>
            <option value="media">Media</option>
            <option value="baja">Baja</option>
          </select>
        </label>
        <label>
          Fecha de inicio
          <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} required />
        </label>
        <label>
          Fecha de fin prevista
          <input type="date" value={plannedEndDate} onChange={(e) => setPlannedEndDate(e.target.value)} required />
        </label>
      </div>
      <div className="form-actions">
        <button type="button" className="btn ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn primary">Crear tarea</button>
      </div>
    </form>
  );
}

/* ---------------- Tarjeta de tarea ---------------- */

function TaskCard({ task, area, user, role, canManage, canObserve, allAreas, allUsers, onUpdateProgress, onMarkComplete, onAddObservation, onChangeDate, onChangePriority, onReassign, onDeleteTask }) {
  const [open, setOpen] = useState(false);
  const [obsText, setObsText] = useState("");
  const [newDate, setNewDate] = useState(task.plannedEndDate);
  const [dateMotivo, setDateMotivo] = useState("");
  const [newPriority, setNewPriority] = useState(task.priority);
  const [prioMotivo, setPrioMotivo] = useState("");
  const [reassignAreaId, setReassignAreaId] = useState(task.areaId);
  const [reassignUserId, setReassignUserId] = useState(task.userId);
  const [completionDate, setCompletionDate] = useState(task.actualEndDate || todayStr());
  const status = computeStatus(task);
  const isAreaOwner = role === "area";
  const reassignAreaUsers = (allUsers || []).filter((u) => u.areaId === reassignAreaId);

  return (
    <div className="task-card">
      <div className="task-head" onClick={() => setOpen((o) => !o)}>
        <div className="task-head-main">
          <span className="task-title">{task.title}</span>
          <span className="task-sub">{area?.name} · {user?.name}</span>
        </div>
        <div className="task-head-badges">
          <Badge color={PRIORITY[task.priority].color}>{PRIORITY[task.priority].label}</Badge>
          <Badge color={STATUS[status].color}>{STATUS[status].label}</Badge>
          <span className="chevron">{open ? "▲" : "▼"}</span>
        </div>
      </div>

      <div className="task-dates">
        <span>Inicio: <b>{fmt(task.startDate)}</b></span>
        <span>Fin previsto: <b>{fmt(task.plannedEndDate)}</b></span>
        <span>Fin real: <b>{fmt(task.actualEndDate)}</b></span>
      </div>
      <ProgressBar value={task.progress} color={STATUS[status].color} />

      {open && (
        <div className="task-detail">
          {task.description && <p className="task-desc">{task.description}</p>}

          {isAreaOwner && !task.actualEndDate && (
            <div className="mini-form">
              <label>
                Avance ({task.progress}%)
                <input type="range" min={0} max={100} value={task.progress}
                  onChange={(e) => onUpdateProgress(task.id, Number(e.target.value))} />
              </label>
              <label>
                Fecha real de finalización
                <input type="date" value={completionDate} max={todayStr()} min={task.startDate}
                  onChange={(e) => setCompletionDate(e.target.value)} />
              </label>
              <button className="btn small primary" onClick={() => onMarkComplete(task.id, completionDate)}>
                Marcar como finalizada
              </button>
            </div>
          )}

          {canManage && (
            <div className="mini-form two-col">
              <div>
                <label>Reprogramar fecha de fin prevista</label>
                <input type="date" value={newDate} onChange={(e) => setNewDate(e.target.value)} />
                <input placeholder="Motivo del cambio" value={dateMotivo} onChange={(e) => setDateMotivo(e.target.value)} />
                <button className="btn small" disabled={!dateMotivo.trim() || newDate === task.plannedEndDate}
                  onClick={() => { onChangeDate(task.id, newDate, dateMotivo.trim()); setDateMotivo(""); }}>
                  Guardar nueva fecha
                </button>
              </div>
              <div>
                <label>Cambiar prioridad</label>
                <select value={newPriority} onChange={(e) => setNewPriority(e.target.value)}>
                  <option value="alta">Alta</option>
                  <option value="media">Media</option>
                  <option value="baja">Baja</option>
                </select>
                <input placeholder="Motivo del cambio" value={prioMotivo} onChange={(e) => setPrioMotivo(e.target.value)} />
                <button className="btn small" disabled={!prioMotivo.trim() || newPriority === task.priority}
                  onClick={() => { onChangePriority(task.id, newPriority, prioMotivo.trim()); setPrioMotivo(""); }}>
                  Guardar prioridad
                </button>
              </div>
            </div>
          )}

          {canManage && (
            <div className="mini-form">
              <label>Reasignar tarea a otra área / persona</label>
              <div className="field-row">
                <select value={reassignAreaId} onChange={(e) => {
                  const aid = e.target.value;
                  setReassignAreaId(aid);
                  const first = (allUsers || []).find((u) => u.areaId === aid);
                  setReassignUserId(first ? first.id : "");
                }}>
                  {(allAreas || []).map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
                <select value={reassignUserId} onChange={(e) => setReassignUserId(e.target.value)}>
                  {reassignAreaUsers.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                </select>
              </div>
              <button className="btn small" disabled={reassignAreaId === task.areaId && reassignUserId === task.userId}
                onClick={() => onReassign(task.id, reassignAreaId, reassignUserId)}>
                Confirmar reasignación
              </button>
            </div>
          )}

          {canManage && onDeleteTask && (
            <div className="mini-form">
              <button className="btn small danger" onClick={() => { if (window.confirm(`¿Eliminar la tarea "${task.title}"? Esta acción no se puede deshacer.`)) onDeleteTask(task.id); }}>
                Eliminar tarea
              </button>
            </div>
          )}

          {(task.history?.length > 0) && (
            <div className="history-list">
              <h4>Historial de cambios</h4>
              {task.history.map((h) => (
                <div key={h.id} className="history-item">
                  <span className="history-date">{fmt(h.date)}</span>
                  <span>{h.field === "plannedEndDate" ? "Fecha reprogramada" : h.field === "priority" ? "Prioridad modificada" : "Reasignación"}: {h.oldValue} → {h.newValue}</span>
                  <span className="history-motivo">Motivo: {h.motivo}</span>
                </div>
              ))}
            </div>
          )}

          <div className="obs-list">
            <h4>Observaciones {task.observations?.length ? `(${task.observations.length})` : ""}</h4>
            {task.observations?.length ? task.observations.map((o) => (
              <div key={o.id} className="obs-item">
                <span className="obs-author">{o.author}</span>
                <span className="obs-date">{fmt(o.date)}</span>
                <p>{o.text}</p>
              </div>
            )) : <p className="empty-hint">Sin observaciones registradas.</p>}

            {canObserve && (
              <div className="mini-form">
                <textarea rows={2} placeholder="Escribir una observación o alerta..." value={obsText} onChange={(e) => setObsText(e.target.value)} />
                <button className="btn small" disabled={!obsText.trim()}
                  onClick={() => { onAddObservation(task.id, obsText.trim()); setObsText(""); }}>
                  Agregar observación
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Login ---------------- */

function LoginScreen({ users, areas, onLogin }) {
  const [role, setRole] = useState("area");
  const roleUsers = users.filter((u) => u.role === role);
  return (
    <div className="login-wrap">
      <div className="login-card">
        <ColorBand />
        <p className="eyebrow">Rectorado · Universidad Nacional de Jujuy</p>
        <h1>Gestión de tareas</h1>
        <p className="login-copy">Elegí tu perfil para ingresar al sistema.</p>
        <div className="role-tabs">
          {["area", "admin", "direccion"].map((r) => (
            <button key={r} className={"role-tab" + (role === r ? " active" : "")} onClick={() => setRole(r)}>
              {ROLE_LABEL[r]}
            </button>
          ))}
        </div>
        <div className="user-list">
          {roleUsers.map((u) => (
            <button key={u.id} className="user-pick" onClick={() => onLogin(u)}>
              <span className="avatar">{u.name.split(" ").map((p) => p[0]).slice(0, 2).join("")}</span>
              <span>
                <span className="user-pick-name">{u.name}</span>
                <span className="user-pick-sub">{u.areaId ? areas.find((a) => a.id === u.areaId)?.name : ROLE_LABEL[u.role]}</span>
              </span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ---------------- Vista Área ---------------- */

function AreaView({ db, currentUser, actions }) {
  const [showForm, setShowForm] = useState(false);
  const [scope, setScope] = useState("mias");
  const myArea = db.areas.find((a) => a.id === currentUser.areaId);
  const areaTasks = db.tasks.filter((t) => t.areaId === currentUser.areaId);
  const tasks = scope === "mias" ? areaTasks.filter((t) => t.userId === currentUser.id) : areaTasks;

  return (
    <div className="view">
      <div className="view-head">
        <div>
          <p className="eyebrow">{myArea?.name}</p>
          <h2>Mis tareas</h2>
        </div>
        <button className="btn primary" onClick={() => setShowForm((s) => !s)}>
          {showForm ? "Cerrar" : "+ Nueva tarea"}
        </button>
      </div>

      {showForm && (
        <TaskForm
          areas={[myArea]}
          users={db.users}
          defaultAreaId={currentUser.areaId}
          defaultUserId={currentUser.id}
          lockAssignment
          onCancel={() => setShowForm(false)}
          onSubmit={(data) => { actions.addTask({ ...data, areaId: currentUser.areaId, userId: currentUser.id }); setShowForm(false); }}
        />
      )}

      <div className="scope-tabs">
        <button className={scope === "mias" ? "active" : ""} onClick={() => setScope("mias")}>Mis tareas ({areaTasks.filter(t=>t.userId===currentUser.id).length})</button>
        <button className={scope === "area" ? "active" : ""} onClick={() => setScope("area")}>Todas las del área ({areaTasks.length})</button>
      </div>

      <div className="task-list">
        {tasks.length === 0 && <p className="empty-hint">Todavía no hay tareas cargadas acá. Usá "+ Nueva tarea" para agregar la primera.</p>}
        {tasks.map((t) => (
          <TaskCard key={t.id} task={t} area={myArea} user={db.users.find((u) => u.id === t.userId)}
            role="area" canManage={false} canObserve={false}
            onUpdateProgress={actions.updateProgress} onMarkComplete={actions.markComplete} />
        ))}
      </div>
    </div>
  );
}

/* ---------------- Vista Administración ---------------- */

/* ---------------- Fila de usuario editable ---------------- */

function UserRow({ user, areas, taskCount, onRename, onChangeArea, onDelete }) {
  const [editing, setEditing] = useState(false);
  const [name, setName] = useState(user.name);

  function saveName() {
    if (name.trim() && name.trim() !== user.name) onRename(user.id, name.trim());
    setEditing(false);
  }

  return (
    <li className="user-row">
      {editing ? (
        <span className="user-row-edit">
          <input value={name} onChange={(e) => setName(e.target.value)} autoFocus
            onKeyDown={(e) => e.key === "Enter" && saveName()} />
          <button className="btn small primary" onClick={saveName}>Guardar</button>
          <button className="btn small ghost" onClick={() => { setName(user.name); setEditing(false); }}>Cancelar</button>
        </span>
      ) : (
        <span className="user-row-view">
          <span>{user.name}</span>
          <select className="area-inline" value={user.areaId} onChange={(e) => onChangeArea(user.id, e.target.value)}>
            {areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
          </select>
          <span className="user-row-actions">
            <button className="btn small ghost" onClick={() => setEditing(true)}>Editar</button>
            <button className="btn small danger" disabled={taskCount > 0}
              title={taskCount > 0 ? `Tiene ${taskCount} tarea(s) asignada(s). Reasignalas antes de eliminar.` : "Eliminar usuario"}
              onClick={() => { if (window.confirm(`¿Eliminar a ${user.name}?`)) onDelete(user.id); }}>
              Eliminar
            </button>
          </span>
        </span>
      )}
    </li>
  );
}

function AdminView({ db, actions }) {
  const [tab, setTab] = useState("tareas");
  const [showForm, setShowForm] = useState(false);
  const [filterArea, setFilterArea] = useState("todas");
  const [newAreaName, setNewAreaName] = useState("");
  const [newUser, setNewUser] = useState({ name: "", username: "", areaId: db.areas[0]?.id || "" });

  const tasks = db.tasks.filter((t) => filterArea === "todas" || t.areaId === filterArea);

  return (
    <div className="view">
      <div className="view-head">
        <div>
          <p className="eyebrow">Administración</p>
          <h2>Gestión general</h2>
        </div>
      </div>

      <div className="scope-tabs">
        <button className={tab === "tareas" ? "active" : ""} onClick={() => setTab("tareas")}>Todas las tareas</button>
        <button className={tab === "estructura" ? "active" : ""} onClick={() => setTab("estructura")}>Áreas y usuarios</button>
      </div>

      {tab === "tareas" && (
        <>
          <div className="view-head compact">
            <select value={filterArea} onChange={(e) => setFilterArea(e.target.value)}>
              <option value="todas">Todas las áreas</option>
              {db.areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
            <button className="btn primary" onClick={() => setShowForm((s) => !s)}>
              {showForm ? "Cerrar" : "+ Asignar tarea"}
            </button>
          </div>

          {showForm && (
            <TaskForm areas={db.areas} users={db.users.filter((u) => u.role === "area")}
              onCancel={() => setShowForm(false)}
              onSubmit={(data) => { actions.addTask(data); setShowForm(false); }} />
          )}

          <div className="task-list">
            {tasks.map((t) => (
              <TaskCard key={t.id} task={t} area={db.areas.find((a) => a.id === t.areaId)} user={db.users.find((u) => u.id === t.userId)}
                role="admin" canManage canObserve
                allAreas={db.areas} allUsers={db.users.filter((u) => u.role === "area")}
                onAddObservation={actions.addObservation}
                onChangeDate={actions.changeDate}
                onChangePriority={actions.changePriority}
                onReassign={actions.reassignTask}
                onDeleteTask={actions.deleteTask} />
            ))}
          </div>
        </>
      )}

      {tab === "estructura" && (
        <div className="two-col-layout">
          <div className="panel">
            <h3>Áreas</h3>
            <ul className="simple-list">
              {db.areas.map((a) => (
                <li key={a.id}>{a.name} <span className="muted">({db.users.filter(u=>u.areaId===a.id).length} usuarios · {db.tasks.filter(t=>t.areaId===a.id).length} tareas)</span></li>
              ))}
            </ul>
            <div className="mini-form">
              <input placeholder="Nombre de la nueva área" value={newAreaName} onChange={(e) => setNewAreaName(e.target.value)} />
              <button className="btn small primary" disabled={!newAreaName.trim()}
                onClick={() => { actions.addArea(newAreaName.trim()); setNewAreaName(""); }}>
                Agregar área
              </button>
            </div>
          </div>

          <div className="panel">
            <h3>Usuarios</h3>
            <ul className="simple-list">
              {db.users.filter((u) => u.role === "area").map((u) => (
                <UserRow key={u.id} user={u} areas={db.areas}
                  taskCount={db.tasks.filter((t) => t.userId === u.id).length}
                  onRename={actions.renameUser}
                  onChangeArea={actions.changeUserArea}
                  onDelete={actions.deleteUser} />
              ))}
            </ul>
            <div className="mini-form">
              <input placeholder="Nombre y apellido" value={newUser.name} onChange={(e) => setNewUser({ ...newUser, name: e.target.value })} />
              <input placeholder="Usuario" value={newUser.username} onChange={(e) => setNewUser({ ...newUser, username: e.target.value })} />
              <select value={newUser.areaId} onChange={(e) => setNewUser({ ...newUser, areaId: e.target.value })}>
                {db.areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
              <button className="btn small primary" disabled={!newUser.name.trim() || !newUser.username.trim()}
                onClick={() => { actions.addUser(newUser); setNewUser({ name: "", username: "", areaId: db.areas[0]?.id || "" }); }}>
                Agregar usuario
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Vista Dirección ---------------- */

function TimelineBands({ db }) {
  const allDates = db.tasks.flatMap((t) => [t.startDate, t.plannedEndDate, t.actualEndDate].filter(Boolean));
  if (!allDates.length) return null;
  const min = allDates.reduce((a, b) => (a < b ? a : b));
  const maxRaw = allDates.reduce((a, b) => (a > b ? a : b));
  const toDay = (d) => Math.floor(new Date(d).getTime() / 86400000);
  const minDay = toDay(min);
  const maxDay = Math.max(toDay(maxRaw), toDay(todayStr()));
  const span = Math.max(maxDay - minDay, 1);
  const todayPct = ((toDay(todayStr()) - minDay) / span) * 100;

  return (
    <div className="panel">
      <h3>Panorama de la línea de tiempo</h3>
      <p className="muted small">Cada franja va de la fecha de inicio a la fecha de fin (real o prevista) de la tarea. El color indica su estado.</p>
      <div className="timeline">
        <div className="timeline-today" style={{ left: `${todayPct}%` }} />
        {db.areas.map((area) => {
          const areaTasks = db.tasks.filter((t) => t.areaId === area.id);
          if (!areaTasks.length) return null;
          return (
            <div className="timeline-row" key={area.id}>
              <span className="timeline-label">{area.name}</span>
              <div className="timeline-track">
                {areaTasks.map((t) => {
                  const end = t.actualEndDate || t.plannedEndDate;
                  const left = ((toDay(t.startDate) - minDay) / span) * 100;
                  const width = Math.max(((toDay(end) - toDay(t.startDate)) / span) * 100, 1.2);
                  const status = computeStatus(t);
                  return (
                    <div key={t.id} className="timeline-bar" title={t.title}
                      style={{ left: `${left}%`, width: `${width}%`, background: STATUS[status].color }} />
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
      <div className="legend">
        {Object.entries(STATUS).map(([k, v]) => (
          <span key={k}><i style={{ background: v.color }} />{v.label}</span>
        ))}
      </div>
    </div>
  );
}

function exportTasksToExcel(db) {
  const rows = db.tasks.map((t) => {
    const status = computeStatus(t);
    return {
      "Área": db.areas.find((a) => a.id === t.areaId)?.name || "",
      "Responsable": db.users.find((u) => u.id === t.userId)?.name || "",
      "Tarea": t.title,
      "Prioridad": PRIORITY[t.priority].label,
      "Estado": STATUS[status].label,
      "Fecha de inicio": fmt(t.startDate),
      "Fecha fin prevista": fmt(t.plannedEndDate),
      "Fecha fin real": fmt(t.actualEndDate),
      "Avance (%)": t.progress,
      "Observaciones": (t.observations || []).map((o) => `[${fmt(o.date)}] ${o.author}: ${o.text}`).join(" · "),
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 26 }, { wch: 20 }, { wch: 36 }, { wch: 10 }, { wch: 22 },
    { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 10 }, { wch: 50 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tareas");
  XLSX.writeFile(wb, `resumen-tareas-rectorado-unju-${todayStr()}.xlsx`);
}

function DireccionView({ db, actions }) {
  const [tab, setTab] = useState("panorama");
  const [filterArea, setFilterArea] = useState("todas");
  const [drill, setDrill] = useState(null); // { type: 'total'|'status'|'priority'|'area', value, label, color }

  const withStatus = useMemo(() => db.tasks.map((t) => ({ ...t, status: computeStatus(t) })), [db.tasks]);

  const kpis = useMemo(() => {
    const count = (s) => withStatus.filter((t) => t.status === s).length;
    return {
      total: db.tasks.length,
      enCurso: count("en_curso"),
      demoradas: count("demorada"),
      finalizadasATiempo: count("finalizada_a_tiempo"),
      finalizadasConDemora: count("finalizada_con_demora"),
      pendientes: count("pendiente"),
    };
  }, [db.tasks, withStatus]);

  const barData = db.areas.map((area) => {
    const areaTasks = withStatus.filter((t) => t.areaId === area.id);
    const row = { area: area.name.length > 14 ? area.name.slice(0, 13) + "…" : area.name };
    Object.keys(STATUS).forEach((s) => { row[s] = areaTasks.filter((t) => t.status === s).length; });
    return row;
  });

  const pieData = ["alta", "media", "baja"].map((p) => ({
    name: PRIORITY[p].label,
    value: db.tasks.filter((t) => t.priority === p).length,
    color: PRIORITY[p].color,
  }));

  const areaIndicators = db.areas.map((area) => {
    const areaTasks = withStatus.filter((t) => t.areaId === area.id);
    return {
      area,
      total: areaTasks.length,
      demoradas: areaTasks.filter((t) => t.status === "demorada").length,
      enCurso: areaTasks.filter((t) => t.status === "en_curso").length,
      finalizadas: areaTasks.filter((t) => t.status === "finalizada_a_tiempo" || t.status === "finalizada_con_demora").length,
    };
  });

  const drillTasks = useMemo(() => {
    if (!drill) return [];
    if (drill.type === "total") return withStatus;
    if (drill.type === "status") return withStatus.filter((t) => t.status === drill.value);
    if (drill.type === "priority") return withStatus.filter((t) => t.priority === drill.value);
    if (drill.type === "area") return withStatus.filter((t) => t.areaId === drill.value);
    return [];
  }, [drill, withStatus]);

  const tasks = db.tasks.filter((t) => filterArea === "todas" || t.areaId === filterArea);

  return (
    <div className="view">
      <div className="view-head">
        <div>
          <p className="eyebrow">Dirección</p>
          <h2>Panorama de gestión</h2>
        </div>
        <button className="btn ghost" onClick={() => exportTasksToExcel(db)}>⬇ Exportar a Excel</button>
      </div>

      <div className="scope-tabs">
        <button className={tab === "panorama" ? "active" : ""} onClick={() => { setTab("panorama"); setDrill(null); }}>Resumen y gráficos</button>
        <button className={tab === "tareas" ? "active" : ""} onClick={() => setTab("tareas")}>Todas las tareas</button>
      </div>

      {tab === "panorama" && drill && (
        <div className="drill-view">
          <div className="drill-head">
            <div>
              <p className="eyebrow">Detalle</p>
              <h2>{drill.label} <span className="muted">({drillTasks.length} tarea{drillTasks.length === 1 ? "" : "s"})</span></h2>
            </div>
            <button className="btn ghost" onClick={() => setDrill(null)}>← Volver al tablero general</button>
          </div>
          <div className="task-list">
            {drillTasks.length === 0 && <p className="empty-hint">No hay tareas para esta selección.</p>}
            {drillTasks.map((t) => (
              <TaskCard key={t.id} task={t} area={db.areas.find((a) => a.id === t.areaId)} user={db.users.find((u) => u.id === t.userId)}
                role="direccion" canManage={false} canObserve
                onAddObservation={actions.addObservation} />
            ))}
          </div>
        </div>
      )}

      {tab === "panorama" && !drill && (
        <>
          <div className="kpi-row">
            <button className="kpi kpi-click" onClick={() => setDrill({ type: "total", label: "Todas las tareas" })}>
              <span className="kpi-num">{kpis.total}</span><span>Tareas totales</span>
            </button>
            <button className="kpi kpi-click" style={{ "--c": PALETTE.secondary }}
              onClick={() => setDrill({ type: "status", value: "en_curso", label: "Tareas en curso" })}>
              <span className="kpi-num">{kpis.enCurso}</span><span>En curso</span>
            </button>
            <button className="kpi kpi-click" style={{ "--c": PALETTE.alert }}
              onClick={() => setDrill({ type: "status", value: "demorada", label: "Tareas demoradas" })}>
              <span className="kpi-num">{kpis.demoradas}</span><span>Demoradas</span>
            </button>
            <button className="kpi kpi-click" style={{ "--c": PALETTE.positive }}
              onClick={() => setDrill({ type: "status", value: "finalizada_a_tiempo", label: "Finalizadas a tiempo" })}>
              <span className="kpi-num">{kpis.finalizadasATiempo}</span><span>Finalizadas a tiempo</span>
            </button>
            <button className="kpi kpi-click" style={{ "--c": PALETTE.accent }}
              onClick={() => setDrill({ type: "status", value: "finalizada_con_demora", label: "Finalizadas con demora" })}>
              <span className="kpi-num">{kpis.finalizadasConDemora}</span><span>Finalizadas con demora</span>
            </button>
          </div>

          <div className="two-col-layout">
            <div className="panel">
              <h3>Tareas por área y estado</h3>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={barData} layout="vertical" margin={{ left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={PALETTE.line} horizontal={false} />
                  <XAxis type="number" allowDecimals={false} stroke={PALETTE.inkSoft} fontSize={12} />
                  <YAxis type="category" dataKey="area" stroke={PALETTE.inkSoft} fontSize={12} width={110} />
                  <Tooltip />
                  <Legend formatter={(v) => STATUS[v]?.label || v} wrapperStyle={{ fontSize: 12 }} />
                  {Object.entries(STATUS).map(([k, v]) => (
                    <Bar key={k} dataKey={k} stackId="a" fill={v.color} radius={[0, 0, 0, 0]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="panel">
              <h3>Distribución por prioridad (unificada)</h3>
              <ResponsiveContainer width="100%" height={280}>
                <PieChart>
                  <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}
                    onClick={(entry) => {
                      const p = ["alta", "media", "baja"].find((k) => PRIORITY[k].label === entry.name);
                      setDrill({ type: "priority", value: p, label: `Prioridad ${PRIORITY[p].label.toLowerCase()}` });
                    }} style={{ cursor: "pointer" }}>
                    {pieData.map((d, i) => <Cell key={i} fill={d.color} />)}
                  </Pie>
                  <Tooltip />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="panel">
            <h3>Indicadores por área</h3>
            <p className="muted small">Tocá una tarjeta para ver el detalle de las tareas de esa área.</p>
            <div className="area-indicator-grid">
              {areaIndicators.map(({ area, total, demoradas, enCurso, finalizadas }) => (
                <button key={area.id} className="area-indicator" onClick={() => setDrill({ type: "area", value: area.id, label: area.name })}>
                  <span className="area-indicator-name">{area.name}</span>
                  <span className="area-indicator-nums">
                    <span><b>{total}</b> total</span>
                    <span style={{ color: PALETTE.secondary }}><b>{enCurso}</b> en curso</span>
                    <span style={{ color: PALETTE.alert }}><b>{demoradas}</b> demoradas</span>
                    <span style={{ color: PALETTE.positive }}><b>{finalizadas}</b> finalizadas</span>
                  </span>
                </button>
              ))}
            </div>
          </div>

          <TimelineBands db={db} />
        </>
      )}

      {tab === "tareas" && (
        <>
          <div className="view-head compact">
            <select value={filterArea} onChange={(e) => setFilterArea(e.target.value)}>
              <option value="todas">Todas las áreas</option>
              {db.areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
          <div className="task-list">
            {tasks.map((t) => (
              <TaskCard key={t.id} task={t} area={db.areas.find((a) => a.id === t.areaId)} user={db.users.find((u) => u.id === t.userId)}
                role="direccion" canManage={false} canObserve
                onAddObservation={actions.addObservation} />
            ))}
          </div>
        </>
      )}
    </div>
  );
}

/* ---------------- App raíz ---------------- */

export default function App() {
  const [db, setDb] = useState(null);
  const [currentUser, setCurrentUser] = useState(null);
  const [saveState, setSaveState] = useState("saving");

  const fetchAll = useCallback(async () => {
    setSaveState("saving");
    try {
      const [areasRes, usersRes, tasksRes] = await Promise.all([
        supabase.from("areas").select("*").order("name"),
        supabase.from("users").select("*"),
        supabase.from("tasks").select("*"),
      ]);
      if (areasRes.error) throw areasRes.error;
      if (usersRes.error) throw usersRes.error;
      if (tasksRes.error) throw tasksRes.error;
      setDb({
        areas: areasRes.data.map(rowToArea),
        users: usersRes.data.map(rowToUser),
        tasks: tasksRes.data.map(rowToTask),
      });
      setSaveState("saved");
    } catch (err) {
      console.error("Error cargando datos de Supabase:", err);
      setSaveState("error");
    }
  }, []);

  // Carga inicial: desde Supabase si está configurado, o desde este navegador si no.
  useEffect(() => {
    if (!supabaseConfigured) {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        setDb(raw ? JSON.parse(raw) : { areas: SEED_AREAS, users: SEED_USERS, tasks: SEED_TASKS });
      } catch {
        setDb({ areas: SEED_AREAS, users: SEED_USERS, tasks: SEED_TASKS });
      }
      setSaveState("saved");
      return;
    }
    fetchAll();
  }, [fetchAll]);

  // Respaldo local (solo aplica cuando no hay Supabase configurado)
  useEffect(() => {
    if (!db || supabaseConfigured) return;
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch { /* noop */ }
  }, [db]);

  // Sincronización en tiempo real: si otra PC cambia datos, se vuelven a traer acá.
  useEffect(() => {
    if (!supabaseConfigured) return;
    let debounceTimer = null;
    const scheduleRefetch = () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(fetchAll, 500);
    };
    const channel = supabase
      .channel("rectorado-unju-cambios")
      .on("postgres_changes", { event: "*", schema: "public", table: "areas" }, scheduleRefetch)
      .on("postgres_changes", { event: "*", schema: "public", table: "users" }, scheduleRefetch)
      .on("postgres_changes", { event: "*", schema: "public", table: "tasks" }, scheduleRefetch)
      .subscribe();
    return () => { clearTimeout(debounceTimer); supabase.removeChannel(channel); };
  }, [fetchAll]);

  const actions = useMemo(() => {
    function persist(promise) {
      if (!promise) return;
      setSaveState("saving");
      promise.then(({ error }) => {
        if (error) { console.error(error); setSaveState("error"); }
        else setSaveState("saved");
      });
    }

    return {
      addArea(name) {
        const area = { id: uid("a"), name };
        setDb((d) => ({ ...d, areas: [...d.areas, area] }));
        if (supabaseConfigured) persist(supabase.from("areas").insert(areaToRow(area)));
      },
      addUser({ name, username, areaId }) {
        const user = { id: uid("u"), name, username, role: "area", areaId };
        setDb((d) => ({ ...d, users: [...d.users, user] }));
        if (supabaseConfigured) persist(supabase.from("users").insert(userToRow(user)));
      },
      addTask(data) {
        const task = { id: uid("t"), ...data, actualEndDate: null, progress: 0, observations: [], history: [] };
        setDb((d) => ({ ...d, tasks: [...d.tasks, task] }));
        if (supabaseConfigured) persist(supabase.from("tasks").insert(taskToRow(task)));
      },
      updateProgress(taskId, progress) {
        setDb((d) => ({ ...d, tasks: d.tasks.map((t) => (t.id === taskId ? { ...t, progress } : t)) }));
        if (supabaseConfigured) persist(supabase.from("tasks").update({ progress }).eq("id", taskId));
      },
      markComplete(taskId, completionDate) {
        const date = completionDate || todayStr();
        setDb((d) => ({ ...d, tasks: d.tasks.map((t) => (t.id === taskId ? { ...t, actualEndDate: date, progress: 100 } : t)) }));
        if (supabaseConfigured) persist(supabase.from("tasks").update({ actual_end_date: date, progress: 100 }).eq("id", taskId));
      },
      addObservation(taskId, text) {
        setDb((d) => {
          const tasks = d.tasks.map((t) => t.id === taskId
            ? { ...t, observations: [...t.observations, { id: uid("o"), author: currentUser.name, date: todayStr(), text }] }
            : t);
          if (supabaseConfigured) {
            const updated = tasks.find((t) => t.id === taskId);
            persist(supabase.from("tasks").update({ observations: updated.observations }).eq("id", taskId));
          }
          return { ...d, tasks };
        });
      },
      changeDate(taskId, newDate, motivo) {
        setDb((d) => {
          const tasks = d.tasks.map((t) => t.id === taskId
            ? { ...t, plannedEndDate: newDate, history: [...t.history, { id: uid("h"), date: todayStr(), field: "plannedEndDate", oldValue: fmt(t.plannedEndDate), newValue: fmt(newDate), motivo }] }
            : t);
          if (supabaseConfigured) {
            const updated = tasks.find((t) => t.id === taskId);
            persist(supabase.from("tasks").update({ planned_end_date: newDate, history: updated.history }).eq("id", taskId));
          }
          return { ...d, tasks };
        });
      },
      changePriority(taskId, newPriority, motivo) {
        setDb((d) => {
          const tasks = d.tasks.map((t) => t.id === taskId
            ? { ...t, priority: newPriority, history: [...t.history, { id: uid("h"), date: todayStr(), field: "priority", oldValue: PRIORITY[t.priority].label, newValue: PRIORITY[newPriority].label, motivo }] }
            : t);
          if (supabaseConfigured) {
            const updated = tasks.find((t) => t.id === taskId);
            persist(supabase.from("tasks").update({ priority: newPriority, history: updated.history }).eq("id", taskId));
          }
          return { ...d, tasks };
        });
      },
      reassignTask(taskId, newAreaId, newUserId) {
        setDb((d) => {
          const task = d.tasks.find((t) => t.id === taskId);
          const oldAreaName = d.areas.find((a) => a.id === task.areaId)?.name || "—";
          const oldUserName = d.users.find((u) => u.id === task.userId)?.name || "—";
          const newAreaName = d.areas.find((a) => a.id === newAreaId)?.name || "—";
          const newUserName = d.users.find((u) => u.id === newUserId)?.name || "—";
          const tasks = d.tasks.map((t) => t.id === taskId
            ? {
                ...t, areaId: newAreaId, userId: newUserId,
                history: [...t.history, {
                  id: uid("h"), date: todayStr(), field: "assignment",
                  oldValue: `${oldUserName} (${oldAreaName})`, newValue: `${newUserName} (${newAreaName})`,
                  motivo: "Reasignación de tarea",
                }],
              }
            : t);
          if (supabaseConfigured) {
            const updated = tasks.find((t) => t.id === taskId);
            persist(supabase.from("tasks").update({ area_id: newAreaId, user_id: newUserId, history: updated.history }).eq("id", taskId));
          }
          return { ...d, tasks };
        });
      },
      deleteTask(taskId) {
        setDb((d) => ({ ...d, tasks: d.tasks.filter((t) => t.id !== taskId) }));
        if (supabaseConfigured) persist(supabase.from("tasks").delete().eq("id", taskId));
      },
      renameUser(userId, newName) {
        setDb((d) => ({ ...d, users: d.users.map((u) => (u.id === userId ? { ...u, name: newName } : u)) }));
        if (supabaseConfigured) persist(supabase.from("users").update({ name: newName }).eq("id", userId));
      },
      changeUserArea(userId, newAreaId) {
        setDb((d) => ({ ...d, users: d.users.map((u) => (u.id === userId ? { ...u, areaId: newAreaId } : u)) }));
        if (supabaseConfigured) persist(supabase.from("users").update({ area_id: newAreaId }).eq("id", userId));
      },
      deleteUser(userId) {
        setDb((d) => {
          const hasTasks = d.tasks.some((t) => t.userId === userId);
          if (hasTasks) return d; // protección: no se borra si tiene tareas asignadas
          if (supabaseConfigured) persist(supabase.from("users").delete().eq("id", userId));
          return { ...d, users: d.users.filter((u) => u.id !== userId) };
        });
      },
    };
  }, [currentUser]);

  return (
    <div className="app-root">
      <style>{CSS}</style>
      {!db ? (
        <div className="loading">Cargando sistema...</div>
      ) : !currentUser ? (
        <LoginScreen users={db.users} areas={db.areas} onLogin={setCurrentUser} />
      ) : (
        <>
          <header className="app-header">
            <div className="app-header-main">
              <span className="app-title">Rectorado UNJU</span>
              <span className="app-subtitle">Sistema de gestión de tareas</span>
            </div>
            <div className="app-header-user">
              <span className="save-dot" data-state={saveState} title={saveState === "saving" ? "Sincronizando..." : saveState === "error" ? "No se pudo sincronizar con la base de datos" : "Sincronizado"} />
              <span>{currentUser.name} · <b>{ROLE_LABEL[currentUser.role]}</b></span>
              <button className="btn ghost small" onClick={() => setCurrentUser(null)}>Cambiar de usuario</button>
            </div>
          </header>
          <ColorBand />
          <main className="app-main">
            {currentUser.role === "area" && <AreaView db={db} currentUser={currentUser} actions={actions} />}
            {currentUser.role === "admin" && <AdminView db={db} actions={actions} />}
            {currentUser.role === "direccion" && <DireccionView db={db} actions={actions} />}
          </main>
          <p className="storage-note">
            {supabaseConfigured
              ? "Los datos se guardan en la base de datos compartida y se sincronizan automáticamente entre todas las computadoras."
              : "⚠ No hay base de datos compartida configurada: los datos se guardan solo en este navegador."}
          </p>
        </>
      )}
    </div>
  );
}

/* ---------------- Estilos ---------------- */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap');

.app-root {
  --bg: ${PALETTE.bg}; --surface: ${PALETTE.surface}; --surface-soft: ${PALETTE.surfaceSoft};
  --ink: ${PALETTE.ink}; --ink-soft: ${PALETTE.inkSoft}; --line: ${PALETTE.line};
  font-family: 'Inter', sans-serif;
  color: var(--ink);
  background: var(--bg);
  min-height: 100vh;
}
.app-root * { box-sizing: border-box; }
h1, h2, h3, h4 { font-family: 'Source Serif 4', serif; margin: 0; font-weight: 600; }
.eyebrow { font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--ink-soft); margin: 0 0 4px; }
.muted { color: var(--ink-soft); font-weight: 400; }
.small { font-size: 12.5px; }
.empty-hint { color: var(--ink-soft); font-style: italic; font-size: 14px; }

.loading { padding: 60px; text-align: center; font-family: 'Source Serif 4', serif; font-size: 20px; }

.color-band { display: flex; height: 6px; width: 100%; }
.color-band span { flex: 1; }

.login-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px; }
.login-card { background: var(--surface); border-radius: 4px; width: 100%; max-width: 440px; box-shadow: 0 12px 40px rgba(27,42,74,0.12); overflow: hidden; }
.login-card > *:not(.color-band) { padding-left: 32px; padding-right: 32px; }
.login-card h1 { font-size: 28px; margin-top: 4px; }
.login-copy { color: var(--ink-soft); margin: 8px 0 20px; font-size: 14.5px; }
.role-tabs { display: flex; gap: 6px; margin-bottom: 18px; }
.role-tab { flex: 1; padding: 9px 6px; border: 1px solid var(--line); background: var(--surface-soft); border-radius: 3px; font-family: 'Inter', sans-serif; font-size: 13px; cursor: pointer; color: var(--ink-soft); font-weight: 600; transition: all .15s; }
.role-tab.active { background: var(--ink); color: #fff; border-color: var(--ink); }
.user-list { display: flex; flex-direction: column; gap: 6px; padding-bottom: 28px; max-height: 320px; overflow-y: auto; }
.user-pick { display: flex; align-items: center; gap: 12px; padding: 10px 12px; border: 1px solid var(--line); background: var(--surface); border-radius: 3px; cursor: pointer; text-align: left; transition: border-color .15s, background .15s; }
.user-pick:hover { border-color: ${PALETTE.secondary}; background: var(--surface-soft); }
.avatar { width: 34px; height: 34px; border-radius: 50%; background: ${PALETTE.secondary}; color: #fff; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; flex-shrink: 0; }
.user-pick-name { display: block; font-weight: 600; font-size: 14px; }
.user-pick-sub { display: block; font-size: 12px; color: var(--ink-soft); }

.app-header { display: flex; justify-content: space-between; align-items: center; padding: 16px 28px; background: var(--ink); color: #fff; flex-wrap: wrap; gap: 10px; }
.app-title { font-family: 'Source Serif 4', serif; font-size: 19px; font-weight: 700; margin-right: 10px; }
.app-subtitle { font-size: 12.5px; opacity: 0.75; }
.app-header-main { display: flex; align-items: baseline; gap: 10px; }
.app-header-user { display: flex; align-items: center; gap: 10px; font-size: 13.5px; }
.save-dot { width: 8px; height: 8px; border-radius: 50%; background: ${PALETTE.positive}; display: inline-block; }
.save-dot[data-state="saving"] { background: ${PALETTE.accent}; }
.save-dot[data-state="error"] { background: ${PALETTE.alert}; }

.app-main { max-width: 1100px; margin: 0 auto; padding: 28px 20px 60px; }
.storage-note { text-align: center; font-size: 12px; color: var(--ink-soft); padding-bottom: 30px; }

.view-head { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 18px; gap: 12px; flex-wrap: wrap; }
.view-head.compact { align-items: center; margin-bottom: 14px; }
.view-head h2 { font-size: 22px; }

.scope-tabs { display: flex; gap: 4px; margin-bottom: 16px; border-bottom: 1px solid var(--line); }
.scope-tabs button { background: none; border: none; padding: 8px 4px; margin-right: 18px; font-family: 'Inter', sans-serif; font-size: 13.5px; font-weight: 600; color: var(--ink-soft); cursor: pointer; border-bottom: 2px solid transparent; }
.scope-tabs button.active { color: var(--ink); border-bottom-color: ${PALETTE.accent}; }

.btn { font-family: 'Inter', sans-serif; font-size: 13.5px; font-weight: 600; padding: 9px 16px; border-radius: 3px; border: 1px solid transparent; cursor: pointer; }
.btn.primary { background: var(--ink); color: #fff; }
.btn.primary:hover { background: ${PALETTE.secondary}; }
.btn.ghost { background: transparent; border-color: var(--line); color: var(--ink); }
.btn.small { padding: 6px 12px; font-size: 12.5px; }
.btn.danger { background: transparent; border-color: ${PALETTE.alert}; color: ${PALETTE.alert}; }
.btn.danger:hover:not(:disabled) { background: ${PALETTE.alert}; color: #fff; }
.btn:disabled { opacity: 0.4; cursor: not-allowed; }

.user-row { display: flex; }
.user-row-view, .user-row-edit { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; width: 100%; }
.user-row-actions { display: flex; gap: 6px; margin-left: auto; }
.area-inline { font-size: 12px; padding: 4px 6px; }

.panel { background: var(--surface); border: 1px solid var(--line); border-radius: 4px; padding: 18px 20px; margin-bottom: 18px; }
.panel h3 { font-size: 15.5px; margin-bottom: 10px; }
.two-col-layout { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
@media (max-width: 720px) { .two-col-layout { grid-template-columns: 1fr; } }

.form label { display: flex; flex-direction: column; gap: 4px; font-size: 12.5px; font-weight: 600; color: var(--ink-soft); margin-bottom: 12px; }
.field-row { display: flex; gap: 14px; flex-wrap: wrap; }
.field-row label { flex: 1; min-width: 140px; }
input, select, textarea { font-family: 'Inter', sans-serif; font-size: 13.5px; padding: 8px 10px; border: 1px solid var(--line); border-radius: 3px; background: var(--surface); color: var(--ink); font-weight: 400; }
textarea { resize: vertical; }
.form-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 6px; }

.simple-list { list-style: none; padding: 0; margin: 0 0 14px; display: flex; flex-direction: column; gap: 7px; font-size: 13.5px; }
.simple-list li { padding-bottom: 7px; border-bottom: 1px solid var(--surface-soft); }

.task-list { display: flex; flex-direction: column; gap: 12px; }
.task-card { background: var(--surface); border: 1px solid var(--line); border-radius: 4px; padding: 14px 18px; }
.task-head { display: flex; justify-content: space-between; align-items: center; cursor: pointer; gap: 10px; flex-wrap: wrap; }
.task-head-main { display: flex; flex-direction: column; }
.task-title { font-weight: 700; font-size: 14.5px; }
.task-sub { font-size: 12px; color: var(--ink-soft); }
.task-head-badges { display: flex; align-items: center; gap: 8px; }
.chevron { font-size: 10px; color: var(--ink-soft); }

.badge { font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; font-weight: 500; padding: 3px 8px; border-radius: 20px; color: #fff; background: var(--c); white-space: nowrap; }

.task-dates { display: flex; gap: 16px; font-size: 12px; color: var(--ink-soft); margin: 10px 0 6px; flex-wrap: wrap; }
.progress-track { height: 5px; background: var(--surface-soft); border-radius: 3px; overflow: hidden; }
.progress-fill { height: 100%; border-radius: 3px; transition: width .2s; }

.task-detail { margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--surface-soft); display: flex; flex-direction: column; gap: 14px; }
.task-desc { font-size: 13.5px; color: var(--ink-soft); margin: 0; }
.mini-form { display: flex; flex-direction: column; gap: 8px; max-width: 420px; }
.mini-form label { font-size: 12px; font-weight: 600; color: var(--ink-soft); }
.mini-form.two-col { flex-direction: row; gap: 20px; max-width: none; flex-wrap: wrap; }
.mini-form.two-col > div { flex: 1; min-width: 220px; display: flex; flex-direction: column; gap: 6px; }

.history-list, .obs-list { display: flex; flex-direction: column; gap: 8px; }
.history-list h4, .obs-list h4 { font-size: 12.5px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-soft); font-family: 'Inter', sans-serif; }
.history-item { font-size: 12.5px; background: var(--surface-soft); padding: 8px 10px; border-radius: 3px; display: flex; flex-direction: column; gap: 2px; }
.history-date, .obs-date { font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; color: var(--ink-soft); }
.history-motivo { color: var(--ink-soft); font-style: italic; }
.obs-item { font-size: 13px; background: var(--surface-soft); padding: 8px 10px; border-radius: 3px; }
.obs-author { font-weight: 700; font-size: 12px; margin-right: 8px; }
.obs-item p { margin: 4px 0 0; }

.kpi-row { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 18px; }
@media (max-width: 820px) { .kpi-row { grid-template-columns: repeat(2, 1fr); } }
.kpi { background: var(--surface); border: 1px solid var(--line); border-top: 3px solid var(--c, ${PALETTE.ink}); border-radius: 4px; padding: 14px 12px; display: flex; flex-direction: column; gap: 2px; }
.kpi-num { font-family: 'Source Serif 4', serif; font-size: 26px; font-weight: 700; }
.kpi span:last-child { font-size: 11.5px; color: var(--ink-soft); }
.kpi-click { text-align: left; cursor: pointer; font-family: 'Inter', sans-serif; transition: transform .1s, box-shadow .15s; }
.kpi-click:hover { box-shadow: 0 4px 14px rgba(27,42,74,0.12); transform: translateY(-1px); }

.area-indicator-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 10px; margin-top: 12px; }
.area-indicator { text-align: left; background: var(--surface-soft); border: 1px solid var(--line); border-radius: 4px; padding: 12px 14px; cursor: pointer; display: flex; flex-direction: column; gap: 8px; font-family: 'Inter', sans-serif; transition: border-color .15s, background .15s; }
.area-indicator:hover { border-color: ${PALETTE.secondary}; background: var(--surface); }
.area-indicator-name { font-weight: 700; font-size: 13.5px; }
.area-indicator-nums { display: flex; flex-wrap: wrap; gap: 10px; font-size: 11.5px; color: var(--ink-soft); }
.area-indicator-nums b { font-family: 'IBM Plex Mono', monospace; color: var(--ink); }

.drill-head { display: flex; justify-content: space-between; align-items: flex-end; gap: 12px; flex-wrap: wrap; margin-bottom: 18px; }
.drill-head h2 { font-size: 20px; }

.timeline { display: flex; flex-direction: column; gap: 10px; margin: 14px 0; position: relative; }
.timeline-today { position: absolute; top: 0; bottom: 20px; width: 2px; background: ${PALETTE.ink}; opacity: 0.5; }
.timeline-row { display: flex; align-items: center; gap: 10px; }
.timeline-label { width: 130px; flex-shrink: 0; font-size: 12px; font-weight: 600; }
.timeline-track { position: relative; flex: 1; height: 16px; background: var(--surface-soft); border-radius: 2px; }
.timeline-bar { position: absolute; top: 2px; height: 12px; border-radius: 2px; min-width: 4px; }
.legend { display: flex; flex-wrap: wrap; gap: 14px; margin-top: 12px; font-size: 11.5px; color: var(--ink-soft); }
.legend i { display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 5px; vertical-align: middle; }
`;
