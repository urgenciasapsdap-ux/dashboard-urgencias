import React, { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  BarChart, Bar, LineChart, Line, ComposedChart, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, ReferenceLine
} from "recharts";

// ── Supabase config ──────────────────────────────────────────────────────────
const supabaseUrl  = import.meta.env.VITE_SUPABASE_URL || "https://placeholder.supabase.co";
const supabaseKey  = import.meta.env.VITE_SUPABASE_ANON_KEY || "placeholder";
const supabase = createClient(supabaseUrl, supabaseKey);
// ─────────────────────────────────────────────────────────────────────────────

// Semana epidemiológica estándar CDC (semana comienza el domingo)
// Usado por MINSAL Chile — coincide con boletines epidemiológicos oficiales
function getEpiWeek(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T12:00:00"); // noon para evitar problemas de zona horaria

  const year = d.getFullYear();

  const jan1 = new Date(year, 0, 1);
  const jan1Day = jan1.getDay();

  const jan4 = new Date(year, 0, 4);
  const jan4Day = jan4.getDay();
  const se01Start = new Date(jan4);
  se01Start.setDate(jan4.getDate() - jan4Day);

  if (d < se01Start) {
    return getEpiWeek((year - 1) + "-12-31");
  }

  const diffMs = d - se01Start;
  const diffDays = Math.floor(diffMs / 86400000);
  const weekNum = Math.floor(diffDays / 7) + 1;

  const nextJan4 = new Date(year + 1, 0, 4);
  const nextJan4Day = nextJan4.getDay();
  const nextSe01Start = new Date(nextJan4);
  nextSe01Start.setDate(nextJan4.getDate() - nextJan4Day);
  if (d >= nextSe01Start) return `SE 01`;

  return `SE ${weekNum.toString().padStart(2, "0")}`;
}

// Establecimientos reales SSMC agrupados por comuna
const ESTABLECIMIENTOS_GROUPED = [
  {
    comuna: "── Polo Cerrillos Maipú · Cerrillos ──",
    items: [
      "SAPU Dr. Norman Voulliéme",
      "SAR Enfermera Sofía Pincheira",
    ],
  },
  {
    comuna: "── Polo Cerrillos Maipú · Maipú ──",
    items: [
      "SAPU Maipú",
      "SAPU Dra. Ana María Juricic",
      "SAR Michelle Bachelet",
      "SAPU Dr. Iván Insunza",
    ],
  },
  {
    comuna: "── Polo Santiago Estación Central · Santiago ──",
    items: [
      "SAPU Consultorio Nº1",
      "SAPU Ignacio Domeyko",
    ],
  },
  {
    comuna: "── Polo Santiago Estación Central · Estación Central ──",
    items: [
      "SAPU Padre Vicente Irarrázabal",
      "SAPU San José de Chuchunco",
    ],
  },
];

const ESTABLECIMIENTOS = ESTABLECIMIENTOS_GROUPED.flatMap(g => g.items);

const DERIVACION_DESTINOS = [
  "HEC – Hospital El Carmen de Maipú",
  "HCSBA – Hospital Clínico San Borja Arriarán",
  "HUAP – Hospital de Urgencia Asistencia Pública",
];

const EMPTY_FORM = {
  fecha: "", semana_epi: "", establecimiento: "",
  demanda_total: "", pacientes_atendidos: "", atenciones_respiratorias: "",
  abandonos: "",
  derivaciones_hec: "", derivaciones_hcsba: "", derivaciones_huap: "",
  tiempo_espera: "",
  tiene_refuerzo: false, tipo_refuerzo: "", horas_refuerzo: "",
  refuerzo_medico: false, refuerzo_enfermera: false, refuerzo_tens: false,
  refuerzo_kinesiologo: false, refuerzo_administrativo: false,
  observaciones: "",
};

const EMPTY_AMBULANCIA_ROW = { fecha: "", establecimiento: "", hora_traslado: "", tiempo_retencion: "" };
const AMBULANCIA_FILAS = 5;

// Paleta SSMC / Minsal – fondo claro
const P = {
  azul:       "#1A3A6B",  // azul SSMC institucional
  azulDark:   "#0F2347",
  azulLight:  "#EEF2F8",
  azulMid:    "#C2D0E4",
  verde:      "#1A7A4A",
  verdeLight: "#E8F5EE",
  amber:      "#B45309",
  rojo:       "#C0392B",  // rojo SSMC institucional
  rojoLight:  "#FDECEA",
  gris:       "#6B7280",
  grisMid:    "#E5E7EB",
  bg:         "#F4F6FA",
  card:       "#FFFFFF",
  border:     "#E2E8F0",
  text:       "#1A2332",
  muted:      "#6B7280",
};

const DEMO_DATA = [
  { "id": 1, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "CESFAM N°1", "demanda_total": "80", "pacientes_atendidos": "80", "atenciones_respiratorias": "21", "tiempo_espera": "15", "abandonos": "0", "derivaciones_hec": "0", "derivaciones_hcsba": "0", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 2, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Maipú", "demanda_total": "144", "pacientes_atendidos": "142", "atenciones_respiratorias": "39", "tiempo_espera": "60", "abandonos": "2", "derivaciones_hec": "6", "derivaciones_hcsba": "3", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 3, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Voullieme", "demanda_total": "104", "pacientes_atendidos": "103", "atenciones_respiratorias": "32", "tiempo_espera": "40", "abandonos": "1", "derivaciones_hec": "2", "derivaciones_hcsba": "1", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 4, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Chuchunco", "demanda_total": "151", "pacientes_atendidos": "140", "atenciones_respiratorias": "36", "tiempo_espera": "50", "abandonos": "9", "derivaciones_hec": "2", "derivaciones_hcsba": "2", "derivaciones_huap": "2", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 5, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Juricic", "demanda_total": "173", "pacientes_atendidos": "153", "atenciones_respiratorias": "66", "tiempo_espera": "97", "abandonos": "20", "derivaciones_hec": "0", "derivaciones_hcsba": "0", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 6, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Padre Vicente", "demanda_total": "114", "pacientes_atendidos": "112", "atenciones_respiratorias": "54", "tiempo_espera": "30", "abandonos": "2", "derivaciones_hec": "3", "derivaciones_hcsba": "2", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 7, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "SAR Pincheira", "demanda_total": "215", "pacientes_atendidos": "203", "atenciones_respiratorias": "57", "tiempo_espera": "54", "abandonos": "12", "derivaciones_hec": "2", "derivaciones_hcsba": "0", "derivaciones_huap": "0", "tiene_refuerzo": true, "refuerzo_medico": true, "refuerzo_enfermera": true, "tipo_refuerzo": "Médico/Enfermera", "horas_refuerzo": "12", "observaciones": "Turno reforzado por alto flujo" },
  { "id": 8, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Insunza", "demanda_total": "115", "pacientes_atendidos": "113", "atenciones_respiratorias": "35", "tiempo_espera": "45", "abandonos": "2", "derivaciones_hec": "1", "derivaciones_hcsba": "1", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 9, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "Domeyko", "demanda_total": "59", "pacientes_atendidos": "53", "atenciones_respiratorias": "19", "tiempo_espera": "150", "abandonos": "6", "derivaciones_hec": "0", "derivaciones_hcsba": "1", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "" },
  { "id": 10, "fecha": "2026-05-03", "semana_epi": "SE 18", "establecimiento": "SAR Michelle Bachelet", "demanda_total": "176", "pacientes_atendidos": "146", "atenciones_respiratorias": "47", "tiempo_espera": "270", "abandonos": "30", "derivaciones_hec": "2", "derivaciones_hcsba": "0", "derivaciones_huap": "0", "tiene_refuerzo": false, "tipo_refuerzo": "", "horas_refuerzo": "", "observaciones": "Demora por traslado complejo" }
];

const POLO_MAP_RAW = {
  "SAPU Dr. Norman Voulliéme":          "Polo Cerrillos Maipú",
  "SAR Enfermera Sofía Pincheira":       "Polo Cerrillos Maipú",
  "SAPU Maipú":                          "Polo Cerrillos Maipú",
  "SAPU Dra. Ana María Juricic":         "Polo Cerrillos Maipú",
  "SAR Michelle Bachelet":               "Polo Cerrillos Maipú",
  "SAPU Dr. Iván Insunza":               "Polo Cerrillos Maipú",
  "Voullieme":                           "Polo Cerrillos Maipú",
  "Voulliéme":                           "Polo Cerrillos Maipú",
  "SAR Pincheira":                       "Polo Cerrillos Maipú",
  "Pincheira":                           "Polo Cerrillos Maipú",
  "Maipú":                               "Polo Cerrillos Maipú",
  "Juricic":                             "Polo Cerrillos Maipú",
  "Michelle Bachelet":                   "Polo Cerrillos Maipú",
  "Insunza":                             "Polo Cerrillos Maipú",
  "SAPU Consultorio Nº1":                "Polo Santiago Estación Central",
  "SAPU Ignacio Domeyko":                "Polo Santiago Estación Central",
  "SAPU Padre Vicente Irarrázabal":      "Polo Santiago Estación Central",
  "SAPU San José de Chuchunco":          "Polo Santiago Estación Central",
  "Consultorio Nº1":                     "Polo Santiago Estación Central",
  "Consultorio N°1":                     "Polo Santiago Estación Central",
  "SAPU Consultorio N°1":                "Polo Santiago Estación Central",
  "CESFAM N°1":                          "Polo Santiago Estación Central",
  "Domeyko":                             "Polo Santiago Estación Central",
  "Ignacio Domeyko":                     "Polo Santiago Estación Central",
  "Padre Vicente":                       "Polo Santiago Estación Central",
  "Padre Vicente Irarrázabal":           "Polo Santiago Estación Central",
  "Chuchunco":                           "Polo Santiago Estación Central",
  "San José de Chuchunco":               "Polo Santiago Estación Central",
};

const getPolo = (estab) => {
  if (!estab) return null;
  if (POLO_MAP_RAW[estab]) return POLO_MAP_RAW[estab];
  for (const [key, polo] of Object.entries(POLO_MAP_RAW)) {
    if (estab.toLowerCase().includes(key.toLowerCase()) || key.toLowerCase().includes(estab.toLowerCase())) {
      return polo;
    }
  }
  return null;
};

const POLO_MAP = new Proxy(POLO_MAP_RAW, {
  get(target, prop) { return getPolo(prop) || target[prop]; }
});

export default function App() {
  const [registros, setRegistros] = useState(DEMO_DATA);
  const [registrosAmbulancias, setRegistrosAmbulancias] = useState([]);
  const [loading, setLoading] = useState(false);
  const [form, setForm] = useState(EMPTY_FORM);
  const [editId, setEditId] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [importando, setImportando] = useState(false);
  const [proyMetodo, setProyMetodo] = useState("historico");
  const [proySemanas, setProyeSemanas] = useState(4);
  const [importResultado, setImportResultado] = useState(null);

  const _hoy = new Date();
  const _hoyStr = `${_hoy.getFullYear()}-${String(_hoy.getMonth()+1).padStart(2,"0")}-${String(_hoy.getDate()).padStart(2,"0")}`;
  const [filtroSemana, setFiltroSemana] = useState(getEpiWeek(_hoyStr) || "Todas");
  const [cmpP1desde, setCmpP1desde] = useState("");
  const [cmpP1hasta, setCmpP1hasta] = useState("");
  const [cmpP2desde, setCmpP2desde] = useState("");
  const [cmpP2hasta, setCmpP2hasta] = useState("");
  const [mostrarComparador, setMostrarComparador] = useState(false);
  const [menuMovil, setMenuMovil] = useState(false);
  const [mostrarPDF, setMostrarPDF] = useState(false);
  const [filtroEstab, setFiltroEstab] = useState("Todos");
  const [filtroPolo, setFiltroPolo] = useState("Todos");
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [toast, setToast] = useState(null);

  const [formAmbulancias, setFormAmbulancias] = useState(
    Array.from({ length: AMBULANCIA_FILAS }, () => ({ ...EMPTY_AMBULANCIA_ROW }))
  );
  const [deleteConfirmAmb, setDeleteConfirmAmb] = useState(null);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3200);
  };

  const semanas = useMemo(() => ["Todas", ...new Set(registros.map(r => r.semana_epi).filter(Boolean))].sort(), [registros]);
  const semanasOpts = useMemo(() => [...new Set(registros.map(r => r.semana_epi).filter(Boolean))].sort(), [registros]);

  const calcMetricasRango = useCallback((desde, hasta) => {
    if (!desde || !hasta) return null;
    const rango = semanasOpts.filter(se => se >= desde && se <= hasta);
    const regs = registros.filter(r =>
      rango.includes(r.semana_epi) &&
      (filtroPolo === "Todos" || POLO_MAP[r.establecimiento] === filtroPolo) &&
      (filtroEstab === "Todos" || r.establecimiento === filtroEstab)
    );
    if (regs.length === 0) return null;
    const sum = (key) => regs.reduce((a, r) => a + Number(r[key] || 0), 0);
    const conEspera = regs.filter(r => r.tiempo_espera != null && r.tiempo_espera !== "");
    return {
      semanas: rango.length,
      registros: regs.length,
      demanda:   sum("demanda_total"),
      atendidos: sum("pacientes_atendidos"),
      resp:      sum("atenciones_respiratorias"),
      abandonos: sum("abandonos"),
      espera:    conEspera.length ? Math.round(conEspera.reduce((a,r) => a + Number(r.tiempo_espera), 0) / conEspera.length) : 0,
      tasaAbandono: sum("demanda_total") ? ((sum("abandonos") / sum("demanda_total")) * 100).toFixed(1) : "0.0",
      pctResp:   sum("pacientes_atendidos") ? ((sum("atenciones_respiratorias") / sum("pacientes_atendidos")) * 100).toFixed(1) : "0.0",
    };
  }, [registros, semanasOpts, filtroPolo, filtroEstab]);

  const POLOS = ["Todos", "Polo Cerrillos Maipú", "Polo Santiago Estación Central"];

  const seActualDashboard = useMemo(() => getEpiWeek(_hoyStr), [_hoyStr]);
  const ultimasSEDashboard = useMemo(() => {
    const out = [];
    for (let i = 0; i < 5; i++) {
      const d = new Date(_hoy);
      d.setDate(d.getDate() - i * 7);
      const ds = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
      out.push(getEpiWeek(ds));
    }
    return out;
  }, [_hoyStr]);

  const pendientesPorSE = useMemo(() => {
    return ultimasSEDashboard.map(se => {
      const fechasConDatos = [...new Set(
        registros.filter(r => r.semana_epi === se && r.fecha).map(r => r.fecha)
      )].sort();

      const pendientes = ESTABLECIMIENTOS.filter(e => {
        return fechasConDatos.some(fecha =>
          !registros.some(r => r.establecimiento === e && r.semana_epi === se && r.fecha === fecha)
        );
      });

      return { se, pendientes };
    });
  }, [registros, ultimasSEDashboard]);

  const totalPendientesDashboard = useMemo(
    () => pendientesPorSE.reduce((a, s) => a + s.pendientes.length, 0),
    [pendientesPorSE]
  );

  const filtrados = useMemo(() => registros.filter(r =>
    (filtroSemana === "Todas" || r.semana_epi === filtroSemana) &&
    (filtroPolo === "Todos" || POLO_MAP[r.establecimiento] === filtroPolo) &&
    (filtroEstab === "Todos" || r.establecimiento === filtroEstab)
  ), [registros, filtroSemana, filtroEstab, filtroPolo]);

  const kpis = useMemo(() => {
    const tot = filtrados.reduce((a, r) => ({
      demanda: a.demanda + Number(r.demanda_total || 0),
      atendidos: a.atendidos + Number(r.pacientes_atendidos || 0),
      respiratorias: a.respiratorias + Number(r.atenciones_respiratorias || 0),
      abandonos: a.abandonos + Number(r.abandonos || 0),
      derivaciones: a.derivaciones + Number(r.derivaciones_hec || 0) + Number(r.derivaciones_hcsba || 0) + Number(r.derivaciones_huap || 0),
    }), { demanda: 0, atendidos: 0, respiratorias: 0, abandonos: 0, derivaciones: 0 });

    const con = filtrados.filter(r => r.tiene_refuerzo);
    const sin = filtrados.filter(r => !r.tiene_refuerzo);
    const pCon = con.length ? con.reduce((a, r) => a + Number(r.pacientes_atendidos || 0), 0) / con.length : 0;
    const pSin = sin.length ? sin.reduce((a, r) => a + Number(r.pacientes_atendidos || 0), 0) / sin.length : 0;
    const conEspera = filtrados.filter(r => r.tiempo_espera !== "" && r.tiempo_espera !== undefined && r.tiempo_espera !== null);
    const promEspera = conEspera.length
      ? (conEspera.reduce((a, r) => a + Number(r.tiempo_espera || 0), 0) / conEspera.length).toFixed(0)
      : 0;

    return {
      ...tot,
      tasaAbandono: tot.demanda ? ((tot.abandonos / tot.demanda) * 100).toFixed(1) : "0.0",
      tasaResp: tot.atendidos ? ((tot.respiratorias / tot.atendidos) * 100).toFixed(1) : "0.0",
      promCon: pCon.toFixed(1), promSin: pSin.toFixed(1),
      impacto: (pCon - pSin).toFixed(1),
      promEspera,
    };
  }, [filtrados]);

  const dataXSemana = useMemo(() => {
    const map = {};
    registros.forEach(r => {
      if (!r.semana_epi) return;
      if (!map[r.semana_epi]) map[r.semana_epi] = { semana: r.semana_epi, demanda: 0, atendidos: 0, respiratorias: 0, abandonos: 0 };
      map[r.semana_epi].demanda += Number(r.demanda_total || 0);
      map[r.semana_epi].atendidos += Number(r.pacientes_atendidos || 0);
      map[r.semana_epi].respiratorias += Number(r.atenciones_respiratorias || 0);
      map[r.semana_epi].abandonos += Number(r.abandonos || 0);
    });
    return Object.values(map).sort((a, b) => a.semana.localeCompare(b.semana));
  }, [registros]);

  const dataRespAcum = useMemo(() => {
    const map = {};
    filtrados.forEach(r => {
      if (!r.semana_epi) return;
      if (!map[r.semana_epi]) map[r.semana_epi] = { semana: r.semana_epi, atendidos: 0, respiratorias: 0 };
      map[r.semana_epi].atendidos     += Number(r.pacientes_atendidos || 0);
      map[r.semana_epi].respiratorias += Number(r.atenciones_respiratorias || 0);
    });
    const rows = Object.values(map).sort((a, b) => a.semana.localeCompare(b.semana));
    let acumAt = 0, acumResp = 0;
    return rows.map(r => {
      acumAt   += r.atendidos;
      acumResp += r.respiratorias;
      return { ...r, acumAtendidos: acumAt, acumRespiratorias: acumResp };
    });
  }, [filtrados]);

  const dataDerivaciones = useMemo(() => {
    const map = { HEC: 0, HCSBA: 0, HUAP: 0 };
    filtrados.forEach(r => {
      map.HEC    += Number(r.derivaciones_hec    || 0);
      map.HCSBA  += Number(r.derivaciones_hcsba  || 0);
      map.HUAP   += Number(r.derivaciones_huap   || 0);
    });
    return Object.entries(map).filter(([, v]) => v > 0).map(([name, value]) => ({ name, value }));
  }, [filtrados]);

  const dataXDia = useMemo(() => {
    const map = {};
    filtrados.forEach(r => {
      if (!r.fecha) return;
      if (!map[r.fecha]) map[r.fecha] = { dia: r.fecha, demanda: 0, atendidos: 0, respiratorias: 0, abandonos: 0, derivaciones: 0 };
      map[r.fecha].demanda        += Number(r.demanda_total || 0);
      map[r.fecha].atendidos      += Number(r.pacientes_atendidos || 0);
      map[r.fecha].respiratorias  += Number(r.atenciones_respiratorias || 0);
      map[r.fecha].abandonos      += Number(r.abandonos || 0);
      map[r.fecha].derivaciones   += Number(r.derivaciones_hec || 0) + Number(r.derivaciones_hcsba || 0) + Number(r.derivaciones_huap || 0);
    });
    return Object.values(map).sort((a, b) => a.dia.localeCompare(b.dia));
  }, [filtrados]);

  const dataAbsorcionDemanda = useMemo(() => {
    const base = registros.filter(r =>
      (filtroPolo === "Todos" || getPolo(r.establecimiento) === filtroPolo) &&
      (filtroSemana === "Todas" || r.semana_epi === filtroSemana)
    );
    const totalRed = base.reduce((sum, r) => sum + Number(r.demanda_total || 0), 0);
    const map = {};
    base.forEach(r => {
      const establecimiento = r.establecimiento || "Sin establecimiento";
      if (!map[establecimiento]) {
        map[establecimiento] = { establecimiento, demanda: 0, atendidos: 0 };
      }
      map[establecimiento].demanda += Number(r.demanda_total || 0);
      map[establecimiento].atendidos += Number(r.pacientes_atendidos || 0);
    });
    return Object.values(map)
      .map(row => ({
        ...row,
        absorcion: totalRed ? Number(((row.demanda / totalRed) * 100).toFixed(1)) : 0,
      }))
      .sort((a, b) => b.absorcion - a.absorcion);
  }, [registros, filtroPolo, filtroSemana]);

  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setForm(prev => ({
      ...prev,
      [name]: type === "checkbox" ? checked : value,
      ...(name === "fecha" ? { semana_epi: getEpiWeek(value) } : {}),
    }));
  };

  const handleSubmit = async () => {
    if (!form.fecha || !form.establecimiento || !form.demanda_total) {
      showToast("Completa fecha, establecimiento y demanda total", "error");
      return;
    }
    const nuevo = { ...form, id: editId || Date.now(), semana_epi: form.semana_epi || getEpiWeek(form.fecha) };
    if (editId) {
      setRegistros(prev => prev.map(r => r.id === editId ? nuevo : r));
      showToast("Registro actualizado correctamente");
      setEditId(null);
    } else {
      setRegistros(prev => [nuevo, ...prev]);
      showToast("Nuevo registro guardado con éxito");
    }
    setForm(EMPTY_FORM);
    setTab("tabla");
  };

  const handleEdit = (r) => { setForm({ ...r }); setEditId(r.id); setTab("formulario"); };

  const handleDelete = (id) => {
    setRegistros(prev => prev.filter(r => r.id !== id));
    setDeleteConfirm(null);
    showToast("Registro eliminado", "warning");
  };

  const handleChangeAmbulancia = (idx, field, value) => {
    setFormAmbulancias(prev => prev.map((row, i) => i === idx ? { ...row, [field]: value } : row));
  };

  const handleSubmitAmbulancias = () => {
    const completas = formAmbulancias.filter(r => r.fecha && r.establecimiento && r.hora_traslado && r.tiempo_retencion);
    if (completas.length === 0) {
      showToast("Completa al menos una fila con fecha, centro, horario y minutos", "error");
      return;
    }
    const nuevos = completas.map(c => ({ ...c, id: Date.now() + Math.random() }));
    setRegistrosAmbulancias(prev => [...nuevos, ...prev]);
    showToast(`${completas.length} retenciones guardadas correctamente`);
    setFormAmbulancias(Array.from({ length: AMBULANCIA_FILAS }, () => ({ ...EMPTY_AMBULANCIA_ROW })));
  };

  const handleDeleteAmbulancia = (id) => {
    setRegistrosAmbulancias(prev => prev.filter(r => r.id !== id));
    setDeleteConfirmAmb(null);
    showToast("Retención eliminada", "warning");
  };

  const kpisAmbulancias = useMemo(() => {
    const n = registrosAmbulancias.length;
    const totalMin = registrosAmbulancias.reduce((a, r) => a + Number(r.tiempo_retencion || 0), 0);
    const promedio = n ? (totalMin / n).toFixed(0) : 0;
    const porEstablecimiento = {};
    registrosAmbulancias.forEach(r => {
      porEstablecimiento[r.establecimiento] = (porEstablecimiento[r.establecimiento] || 0) + 1;
    });
    const top = Object.entries(porEstablecimiento).sort((a, b) => b[1] - a[1])[0];
    return {
      total: n,
      totalMin,
      promedio,
      establecimientoTop: top ? top[0] : "—",
      establecimientoTopCount: top ? top[1] : 0,
    };
  }, [registrosAmbulancias]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(filtrados);
    XLSX.utils.book_append_sheet(wb, ws, "Registros");
    XLSX.writeFile(wb, `AtencionesUrgencia_SSMC_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Archivo Excel exportado");
  };

  const tdS = { padding: "10px 12px", whiteSpace: "nowrap", color: P.text, borderBottom: `1px solid ${P.border}` };

  return (
    <div style={{ background: P.bg, minHeight: "100vh", fontFamily: "'Inter', sans-serif", color: P.text }}>
      {toast && (
        <div style={{
          position: "fixed", top: 20, right: 20, zIndex: 9999,
          background: toast.type === "error" ? P.rojo : toast.type === "warning" ? P.amber : P.verde,
          color: "#FFF", padding: "12px 20px", borderRadius: 8, fontSize: 13, fontWeight: 600,
          boxShadow: "0 4px 20px rgba(0,0,0,0.2)"
        }}>{toast.msg}</div>
      )}

      {/* Header Institucional SSMC */}
      <div style={{ background: P.azul, borderBottom: `3px solid ${P.rojo}` }}>
        <div style={{ padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ width: 44, height: 44, borderRadius: 6, background: "#FFF", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <img
                src="https://upload.wikimedia.org/wikipedia/commons/thumb/7/78/Escudo_de_Chile.svg/240px-Escudo_de_Chile.svg.png"
                alt="Escudo Chile"
                style={{ width: 34, height: 34, objectFit: "contain" }}
                onError={e => { e.target.style.display="none"; e.target.parentNode.innerHTML="🏥"; }}
              />
            </div>
            <div>
              <div style={{ color: "#FFF", fontWeight: 700, fontSize: 16 }}>Sistema de Monitoreo de Urgencias APS</div>
              <div style={{ color: "rgba(255,255,255,0.65)", fontSize: 11, marginTop: 2 }}>Servicio de Salud Metropolitano Central · Ministerio de Salud de Chile</div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={exportExcel} style={{ background: "transparent", border: "1px solid rgba(255,255,255,0.4)", color: "#FFF", padding: "7px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600 }}>Exportar Excel</button>
            <button onClick={() => setMostrarPDF(true)} style={{ background: P.rojo, border: "none", color: "#FFF", padding: "7px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>Reporte PDF</button>
          </div>
        </div>

        {/* Navegacion de Pestanas */}
        <div style={{ display: "flex", paddingLeft: 24, background: P.azulDark, overflowX: "auto" }}>
          {[
            { id: "dashboard", label: "Resumen Exec." },
            { id: "formulario", label: "Ingresar Datos" },
            { id: "tabla", label: "Tabla Registros" },
            { id: "ambulancias", label: "Ambulancias" },
            { id: "importar", label: "Importar Excel" },
            { id: "tiempos", label: "T° Espera" },
            { id: "proyecciones", label: "Proyecciones" },
          ].map(t => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                background: "transparent",
                color: tab === t.id ? "#FFF" : "rgba(255,255,255,0.6)",
                border: "none",
                borderBottom: tab === t.id ? `3px solid ${P.rojo}` : "3px solid transparent",
                padding: "11px 18px",
                cursor: "pointer",
                fontSize: 13,
                fontWeight: tab === t.id ? 700 : 500
              }}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <main style={{ padding: "24px 28px", maxWidth: 1400, margin: "0 auto" }}>
        {/* DASHBOARD TAB */}
        {tab === "dashboard" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
            {/* Alerta de faltantes */}
            <div style={{
              background: totalPendientesDashboard ? "#FFFBEB" : P.verdeLight,
              border: `1px solid ${totalPendientesDashboard ? "#F59E0B" : P.verde}`,
              borderRadius: 10, padding: "14px 18px"
            }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: totalPendientesDashboard ? "#92400E" : P.verde }}>
                {totalPendientesDashboard ? "⚠️ Registros Pendientes en las Últimas 5 Semanas Epidemiológicas" : "✅ Red SSMC al Día"}
              </div>
              {totalPendientesDashboard > 0 && (
                <div style={{ marginTop: 8, display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {pendientesPorSE.map(({ se, pendientes }) => (
                    pendientes.length > 0 && (
                      <span key={se} style={{ background: "#FEF3C7", border: "1px solid #F59E0B", borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, color: "#92400E" }}>
                        {se}: {pendientes.join(", ")}
                      </span>
                    )
                  ))}
                </div>
              )}
            </div>

            {/* Filtros */}
            <div style={{ background: "#FFF", border: `1px solid ${P.border}`, borderRadius: 8, padding: 16, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: P.azulDark }}>FILTROS:</div>
              <select value={filtroPolo} onChange={e => setFiltroPolo(e.target.value)} style={{ padding: "7px 12px", borderRadius: 6, border: `1px solid ${P.border}`, fontSize: 13, fontWeight: 600 }}>
                {POLOS.map(o => <option key={o}>{o}</option>)}
              </select>
              <select value={filtroSemana} onChange={e => setFiltroSemana(e.target.value)} style={{ padding: "7px 12px", borderRadius: 6, border: `1px solid ${P.border}`, fontSize: 13, fontWeight: 600 }}>
                {semanas.map(o => <option key={o}>{o}</option>)}
              </select>
              <select value={filtroEstab} onChange={e => setFiltroEstab(e.target.value)} style={{ padding: "7px 12px", borderRadius: 6, border: `1px solid ${P.border}`, fontSize: 13, fontWeight: 600 }}>
                {["Todos", ...ESTABLECIMIENTOS].map(o => <option key={o}>{o}</option>)}
              </select>
              <div style={{ marginLeft: "auto", fontSize: 12, color: P.muted }}>
                Total: <b style={{ color: P.azul, fontSize: 16 }}>{filtrados.length}</b> registros
              </div>
            </div>

            {/* KPIs */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12 }}>
              {[
                { label: "Demanda Total",    val: kpis.demanda.toLocaleString("es-CL"), color: P.azul, bg: P.azulLight },
                { label: "Atendidos",         val: kpis.atendidos.toLocaleString("es-CL"), color: P.verde, bg: P.verdeLight },
                { label: "Respiratorias",     val: kpis.respiratorias.toLocaleString("es-CL"), color: P.amber, bg: "#FEF3C7" },
                { label: "Abandonos",         val: kpis.abandonos.toLocaleString("es-CL"), color: P.rojo, bg: P.rojoLight },
                { label: "Derivaciones",      val: kpis.derivaciones.toLocaleString("es-CL"), color: "#7B3FA0", bg: "#F3E8FF" },
                { label: "Tasa Abandono",     val: `${kpis.tasaAbandono}%`, color: P.rojo, bg: P.rojoLight },
                { label: "% Respiratorio",    val: `${kpis.tasaResp}%`, color: P.amber, bg: "#FEF3C7" },
                { label: "T° Espera Prom.",   val: `${kpis.promEspera} min`, color: "#0284C7", bg: "#E0F2FE" },
              ].map(k => (
                <div key={k.label} style={{ background: k.bg, border: `1px solid ${P.border}`, borderLeft: `4px solid ${k.color}`, borderRadius: 8, padding: 14 }}>
                  <div style={{ fontSize: 22, fontWeight: 800, color: k.color }}>{k.val}</div>
                  <div style={{ fontSize: 11, color: P.muted, fontWeight: 600, marginTop: 2 }}>{k.label}</div>
                </div>
              ))}
            </div>

            {/* Absorción por Establecimiento */}
            <div style={{ background: "#FFF", border: `1px solid ${P.border}`, borderRadius: 8, padding: 20 }}>
              <h3 style={{ fontSize: 15, fontWeight: 800, color: P.azulDark, marginBottom: 14 }}>Absorción de la Demanda por Establecimiento</h3>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 18 }}>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={dataAbsorcionDemanda} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                    <XAxis type="number" tickFormatter={v => `${v}%`} />
                    <YAxis type="category" dataKey="establecimiento" width={140} tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Bar dataKey="absorcion" name="Absorción (% Red)" fill={P.azul} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
                <div style={{ border: `1px solid ${P.border}`, borderRadius: 8, overflow: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: P.azulDark, color: "#FFF" }}>
                        <th style={{ padding: 8, textAlign: "left" }}>Establecimiento</th>
                        <th style={{ padding: 8, textAlign: "right" }}>Demanda</th>
                        <th style={{ padding: 8, textAlign: "right" }}>% Red</th>
                      </tr>
                    </thead>
                    <tbody>
                      {dataAbsorcionDemanda.map(r => (
                        <tr key={r.establecimiento} style={{ borderBottom: `1px solid ${P.border}` }}>
                          <td style={{ padding: 8, fontWeight: 700 }}>{r.establecimiento}</td>
                          <td style={{ padding: 8, textAlign: "right", color: P.azul, fontWeight: 800 }}>{r.demanda.toLocaleString("es-CL")}</td>
                          <td style={{ padding: 8, textAlign: "right", fontWeight: 700, color: P.verde }}>{r.absorcion.toFixed(1)}%</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Grafico Diario */}
            <div style={{ background: "#FFF", border: `1px solid ${P.border}`, borderRadius: 8, padding: 20 }}>
              <h3 style={{ fontSize: 15, fontWeight: 800, color: P.azulDark, marginBottom: 14 }}>Comportamiento Diario de Variables</h3>
              <ResponsiveContainer width="100%" height={260}>
                <LineChart data={dataXDia}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="dia" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip />
                  <Legend />
                  <Line type="monotone" dataKey="demanda" name="Demanda" stroke={P.azul} strokeWidth={2} dot={false} />
                  <Line type="monotone" dataKey="atendidos" name="Atendidos" stroke={P.verde} strokeWidth={2} dot={false} />
                  <Line type="monotone" dataKey="respiratorias" name="Respiratorias" stroke={P.amber} strokeWidth={1.8} dot={false} />
                  <Line type="monotone" dataKey="abandonos" name="Abandonos" stroke={P.rojo} strokeWidth={1.8} dot={false} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* TABLA TAB */}
        {tab === "tabla" && (
          <div style={{ background: "#FFF", border: `1px solid ${P.border}`, borderRadius: 8, padding: 20 }}>
            <h2 style={{ fontSize: 16, fontWeight: 800, color: P.azulDark, marginBottom: 14 }}>Registros Consolidados SSMC</h2>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: P.azulDark, color: "#FFF" }}>
                    <th style={{ padding: 8, textAlign: "left" }}>Fecha</th>
                    <th style={{ padding: 8, textAlign: "left" }}>Semana</th>
                    <th style={{ padding: 8, textAlign: "left" }}>Establecimiento</th>
                    <th style={{ padding: 8, textAlign: "right" }}>Demanda</th>
                    <th style={{ padding: 8, textAlign: "right" }}>Atendidos</th>
                    <th style={{ padding: 8, textAlign: "right" }}>Resp.</th>
                    <th style={{ padding: 8, textAlign: "right" }}>Abandonos</th>
                    <th style={{ padding: 8, textAlign: "center" }}>Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  {filtrados.map(r => (
                    <tr key={r.id} style={{ borderBottom: `1px solid ${P.border}` }}>
                      <td style={{ padding: 8 }}>{r.fecha}</td>
                      <td style={{ padding: 8 }}>{r.semana_epi}</td>
                      <td style={{ padding: 8, fontWeight: 700 }}>{r.establecimiento}</td>
                      <td style={{ padding: 8, textAlign: "right" }}>{r.demanda_total}</td>
                      <td style={{ padding: 8, textAlign: "right", color: P.verde, fontWeight: 700 }}>{r.pacientes_atendidos}</td>
                      <td style={{ padding: 8, textAlign: "right", color: P.amber }}>{r.atenciones_respiratorias}</td>
                      <td style={{ padding: 8, textAlign: "right", color: P.rojo }}>{r.abandonos}</td>
                      <td style={{ padding: 8, textAlign: "center" }}>
                        <button onClick={() => handleEdit(r)} style={{ marginRight: 8, border: "none", background: "transparent", color: P.azul, cursor: "pointer", fontWeight: 600 }}>Editar</button>
                        <button onClick={() => handleDelete(r.id)} style={{ border: "none", background: "transparent", color: P.rojo, cursor: "pointer", fontWeight: 600 }}>Eliminar</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* FORMULARIO TAB */}
        {tab === "formulario" && (
          <div style={{ background: "#FFF", border: `1px solid ${P.border}`, borderRadius: 8, padding: 24, maxWidth: 650, margin: "0 auto" }}>
            <h2 style={{ fontSize: 16, fontWeight: 800, color: P.azulDark, marginBottom: 16 }}>{editId ? "Editar Registro" : "Ingresar Atenciones Urgencia"}</h2>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div>
                <label style={{ fontSize: 11, fontWeight: 700, color: P.muted, textTransform: "uppercase" }}>Fecha</label>
                <input type="date" name="fecha" value={form.fecha} onChange={handleChange} style={{ width: "100%", padding: 8, borderRadius: 6, border: `1px solid ${P.border}` }} />
              </div>
              <div>
                <label style={{ fontSize: 11, fontWeight: 700, color: P.muted, textTransform: "uppercase" }}>Establecimiento</label>
                <select name="establecimiento" value={form.establecimiento} onChange={handleChange} style={{ width: "100%", padding: 8, borderRadius: 6, border: `1px solid ${P.border}` }}>
                  <option value="">-- Seleccionar Centro SSMC --</option>
                  {ESTABLECIMIENTOS_GROUPED.map(grp => (
                    <optgroup key={grp.comuna} label={grp.comuna}>
                      {grp.items.map(item => <option key={item} value={item}>{item}</option>)}
                    </optgroup>
                  ))}
                </select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                <div>
                  <label style={{ fontSize: 11, fontWeight: 700, color: P.muted }}>Demanda Total</label>
                  <input type="number" name="demanda_total" value={form.demanda_total} onChange={handleChange} style={{ width: "100%", padding: 8, borderRadius: 6, border: `1px solid ${P.border}` }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, fontWeight: 700, color: P.muted }}>Atendidos</label>
                  <input type="number" name="pacientes_atendidos" value={form.pacientes_atendidos} onChange={handleChange} style={{ width: "100%", padding: 8, borderRadius: 6, border: `1px solid ${P.border}` }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, fontWeight: 700, color: P.muted }}>Respiratorias</label>
                  <input type="number" name="atenciones_respiratorias" value={form.atenciones_respiratorias} onChange={handleChange} style={{ width: "100%", padding: 8, borderRadius: 6, border: `1px solid ${P.border}` }} />
                </div>
              </div>
              <button onClick={handleSubmit} style={{ background: P.azul, color: "#FFF", border: "none", padding: 12, borderRadius: 6, fontWeight: 700, cursor: "pointer", marginTop: 10 }}>Guardar Registro</button>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
