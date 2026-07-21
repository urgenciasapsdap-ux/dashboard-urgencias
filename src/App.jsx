import React, { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  BarChart, Bar, LineChart, Line, ComposedChart, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, ReferenceLine
} from "recharts";

// ── Supabase config ──────────────────────────────────────────────────────────
const supabaseUrl  = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey  = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = createClient(supabaseUrl, supabaseKey);
// ─────────────────────────────────────────────────────────────────────────────

// Semana epidemiológica estándar CDC (semana comienza el domingo)
// Usado por MINSAL Chile — coincide con boletines epidemiológicos oficiales
function getEpiWeek(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T12:00:00"); // noon para evitar problemas de zona horaria

  // Encontrar el primer domingo del año (inicio SE 01)
  // La SE 01 contiene el 4 de enero (CDC: la semana que contiene el 4-ene es la SE 01)
  const year = d.getFullYear();

  // Día de la semana del 1 de enero (0=dom, 1=lun, ... 6=sáb)
  const jan1 = new Date(year, 0, 1);
  const jan1Day = jan1.getDay(); // 0=domingo

  // Inicio de la SE 01: retroceder al domingo anterior al 4 de enero
  const jan4 = new Date(year, 0, 4);
  const jan4Day = jan4.getDay();
  const se01Start = new Date(jan4);
  se01Start.setDate(jan4.getDate() - jan4Day); // domingo de esa semana

  // Si la fecha es anterior al inicio de la SE 01, corresponde a la última semana del año anterior
  if (d < se01Start) {
    return getEpiWeek((year - 1) + "-12-31");
  }

  const diffMs = d - se01Start;
  const diffDays = Math.floor(diffMs / 86400000);
  const weekNum = Math.floor(diffDays / 7) + 1;

  // Verificar que no exceda las semanas del año actual (máx 52 o 53)
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

// Lista plana para compatibilidad con filtros y tabla
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

// Fila vacía para el formulario rápido de retenciones de ambulancias
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
  {
    "id": 1,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "80",
    "pacientes_atendidos": "80",
    "atenciones_respiratorias": "21",
    "tiempo_espera": "15",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 2,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "144",
    "pacientes_atendidos": "142",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "60",
    "abandonos": "2",
    "derivaciones": "9",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 3,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "104",
    "pacientes_atendidos": "103",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "40",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 4,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "151",
    "pacientes_atendidos": "140",
    "atenciones_respiratorias": "36",
    "tiempo_espera": "50",
    "abandonos": "9",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 5,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "173",
    "pacientes_atendidos": "153",
    "atenciones_respiratorias": "66",
    "tiempo_espera": "97",
    "abandonos": "20",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 6,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "114",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "54",
    "tiempo_espera": "30",
    "abandonos": "2",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 7,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "215",
    "pacientes_atendidos": "203",
    "atenciones_respiratorias": "57",
    "tiempo_espera": "54",
    "abandonos": "12",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 8,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "115",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "35",
    "tiempo_espera": "45",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 9,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "59",
    "pacientes_atendidos": "53",
    "atenciones_respiratorias": "19",
    "tiempo_espera": "150",
    "abandonos": "6",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 10,
    "fecha": "2026-05-03",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "176",
    "pacientes_atendidos": "146",
    "atenciones_respiratorias": "47",
    "tiempo_espera": "270",
    "abandonos": "30",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 11,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "83",
    "pacientes_atendidos": "83",
    "atenciones_respiratorias": "23",
    "tiempo_espera": "120",
    "abandonos": "0",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 12,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "131",
    "pacientes_atendidos": "130",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "102",
    "abandonos": "1",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 13,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "119",
    "pacientes_atendidos": "115",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "60",
    "abandonos": "4",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 14,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "125",
    "pacientes_atendidos": "121",
    "atenciones_respiratorias": "37",
    "tiempo_espera": "40",
    "abandonos": "4",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 15,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "129",
    "pacientes_atendidos": "102",
    "atenciones_respiratorias": "58",
    "tiempo_espera": "154",
    "abandonos": "27",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 16,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "123",
    "pacientes_atendidos": "116",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "100",
    "abandonos": "7",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 17,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "283",
    "pacientes_atendidos": "236",
    "atenciones_respiratorias": "68",
    "tiempo_espera": "139",
    "abandonos": "47",
    "derivaciones": "11",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 18,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "110",
    "pacientes_atendidos": "105",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "45",
    "abandonos": "5",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 19,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "81",
    "pacientes_atendidos": "63",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "180",
    "abandonos": "18",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 20,
    "fecha": "2026-05-04",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "148",
    "pacientes_atendidos": "120",
    "atenciones_respiratorias": "51",
    "tiempo_espera": "180",
    "abandonos": "28",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 21,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "103",
    "pacientes_atendidos": "102",
    "atenciones_respiratorias": "45",
    "tiempo_espera": "120",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 22,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "111",
    "pacientes_atendidos": "110",
    "atenciones_respiratorias": "56",
    "tiempo_espera": "64",
    "abandonos": "1",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 23,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "97",
    "pacientes_atendidos": "96",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "90",
    "abandonos": "1",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 24,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "128",
    "pacientes_atendidos": "123",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "60",
    "abandonos": "4",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 25,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "134",
    "pacientes_atendidos": "122",
    "atenciones_respiratorias": "65",
    "tiempo_espera": "61",
    "abandonos": "12",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 26,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "134",
    "pacientes_atendidos": "124",
    "atenciones_respiratorias": "53",
    "tiempo_espera": "180",
    "abandonos": "10",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 27,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "245",
    "pacientes_atendidos": "217",
    "atenciones_respiratorias": "76",
    "tiempo_espera": "91",
    "abandonos": "28",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 28,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "87",
    "pacientes_atendidos": "78",
    "atenciones_respiratorias": "26",
    "tiempo_espera": "45",
    "abandonos": "8",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 29,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "62",
    "pacientes_atendidos": "61",
    "atenciones_respiratorias": "19",
    "tiempo_espera": "120",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 30,
    "fecha": "2026-05-05",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "135",
    "pacientes_atendidos": "117",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "200",
    "abandonos": "18",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 31,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "63",
    "pacientes_atendidos": "63",
    "atenciones_respiratorias": "12",
    "tiempo_espera": "36",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 32,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "113",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 33,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "42",
    "pacientes_atendidos": "42",
    "atenciones_respiratorias": "12",
    "tiempo_espera": "10",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 34,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "92",
    "pacientes_atendidos": "92",
    "atenciones_respiratorias": "31",
    "tiempo_espera": "40",
    "abandonos": "0",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 35,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "39",
    "pacientes_atendidos": "5",
    "atenciones_respiratorias": "1",
    "tiempo_espera": "15",
    "abandonos": "0",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 36,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "64",
    "pacientes_atendidos": "62",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "70",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 37,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "196",
    "pacientes_atendidos": "165",
    "atenciones_respiratorias": "45",
    "tiempo_espera": "143",
    "abandonos": "31",
    "derivaciones": "15",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 38,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "57",
    "pacientes_atendidos": "56",
    "atenciones_respiratorias": "19",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 39,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "37",
    "pacientes_atendidos": "36",
    "atenciones_respiratorias": "11",
    "tiempo_espera": "60",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 40,
    "fecha": "2026-05-06",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "127",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "36",
    "tiempo_espera": "315",
    "abandonos": "25",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 41,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "81",
    "pacientes_atendidos": "79",
    "atenciones_respiratorias": "31",
    "tiempo_espera": "60",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 42,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "98",
    "pacientes_atendidos": "98",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "67",
    "abandonos": "0",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 43,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "83",
    "pacientes_atendidos": "82",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "200",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 44,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "96",
    "pacientes_atendidos": "89",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "40",
    "abandonos": "7",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 45,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "114",
    "pacientes_atendidos": "106",
    "atenciones_respiratorias": "50",
    "tiempo_espera": "94",
    "abandonos": "8",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 46,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "94",
    "pacientes_atendidos": "80",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "",
    "abandonos": "14",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 47,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "219",
    "pacientes_atendidos": "210",
    "atenciones_respiratorias": "71",
    "tiempo_espera": "41",
    "abandonos": "9",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 48,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "89",
    "pacientes_atendidos": "73",
    "atenciones_respiratorias": "24",
    "tiempo_espera": "120",
    "abandonos": "16",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 49,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "52",
    "pacientes_atendidos": "40",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "240",
    "abandonos": "22",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 50,
    "fecha": "2026-05-07",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "139",
    "pacientes_atendidos": "104",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "180",
    "abandonos": "35",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 51,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "63",
    "pacientes_atendidos": "63",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 52,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "106",
    "pacientes_atendidos": "106",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 53,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "89",
    "pacientes_atendidos": "84",
    "atenciones_respiratorias": "25",
    "tiempo_espera": "160",
    "abandonos": "5",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 54,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "100",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "16",
    "tiempo_espera": "60",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 55,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "87",
    "pacientes_atendidos": "86",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "56",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 56,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "74",
    "pacientes_atendidos": "71",
    "atenciones_respiratorias": "31",
    "tiempo_espera": "80",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 57,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "208",
    "pacientes_atendidos": "197",
    "atenciones_respiratorias": "57",
    "tiempo_espera": "39",
    "abandonos": "11",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 58,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "62",
    "pacientes_atendidos": "62",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "25",
    "abandonos": "0",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 59,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "45",
    "pacientes_atendidos": "45",
    "atenciones_respiratorias": "12",
    "tiempo_espera": "15",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 60,
    "fecha": "2026-05-08",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "118",
    "pacientes_atendidos": "125",
    "atenciones_respiratorias": "27",
    "tiempo_espera": "180",
    "abandonos": "7",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 61,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "95",
    "pacientes_atendidos": "95",
    "atenciones_respiratorias": "25",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 62,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Maipú",
    "demanda_total": "105",
    "pacientes_atendidos": "103",
    "atenciones_respiratorias": "25",
    "tiempo_espera": "35",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 63,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Voullieme",
    "demanda_total": "91",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "26",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 64,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Chuchunco",
    "demanda_total": "141",
    "pacientes_atendidos": "126",
    "atenciones_respiratorias": "54",
    "tiempo_espera": "60",
    "abandonos": "13",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 65,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Juricic",
    "demanda_total": "140",
    "pacientes_atendidos": "135",
    "atenciones_respiratorias": "74",
    "tiempo_espera": "35",
    "abandonos": "5",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 66,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Padre Vicente",
    "demanda_total": "101",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "35",
    "tiempo_espera": "45",
    "abandonos": "4",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 67,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "160",
    "pacientes_atendidos": "160",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "20",
    "abandonos": "0",
    "derivaciones": "10",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 68,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Insunza",
    "demanda_total": "106",
    "pacientes_atendidos": "106",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "20",
    "abandonos": "0",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 69,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "Domeyko",
    "demanda_total": "87",
    "pacientes_atendidos": "87",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 70,
    "fecha": "2026-05-09",
    "semana_epi": "SE 18",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "179",
    "pacientes_atendidos": "161",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "300",
    "abandonos": "18",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 71,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "66",
    "pacientes_atendidos": "66",
    "atenciones_respiratorias": "20",
    "tiempo_espera": "40",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 72,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "126",
    "pacientes_atendidos": "126",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "25",
    "abandonos": "0",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 73,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "93",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 74,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "144",
    "pacientes_atendidos": "141",
    "atenciones_respiratorias": "56",
    "tiempo_espera": "40",
    "abandonos": "3",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 75,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "180",
    "pacientes_atendidos": "158",
    "atenciones_respiratorias": "79",
    "tiempo_espera": "61",
    "abandonos": "22",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 76,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "27",
    "pacientes_atendidos": "27",
    "atenciones_respiratorias": "11",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 77,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "184",
    "pacientes_atendidos": "172",
    "atenciones_respiratorias": "48",
    "tiempo_espera": "58",
    "abandonos": "12",
    "derivaciones": "11",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 78,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "93",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "20",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 79,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "71",
    "pacientes_atendidos": "71",
    "atenciones_respiratorias": "24",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 80,
    "fecha": "2026-05-10",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "170",
    "pacientes_atendidos": "156",
    "atenciones_respiratorias": "59",
    "tiempo_espera": "",
    "abandonos": "14",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 81,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "95",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "41",
    "tiempo_espera": "120",
    "abandonos": "2",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 82,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "130",
    "pacientes_atendidos": "128",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "87",
    "abandonos": "2",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 83,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "104",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "180",
    "abandonos": "7",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 84,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "131",
    "pacientes_atendidos": "117",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "180",
    "abandonos": "14",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 85,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "132",
    "pacientes_atendidos": "99",
    "atenciones_respiratorias": "53",
    "tiempo_espera": "128",
    "abandonos": "33",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 86,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "121",
    "pacientes_atendidos": "102",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "270",
    "abandonos": "19",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 87,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "277",
    "pacientes_atendidos": "246",
    "atenciones_respiratorias": "84",
    "tiempo_espera": "119",
    "abandonos": "31",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 88,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "103",
    "pacientes_atendidos": "94",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "45",
    "abandonos": "9",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 89,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "79",
    "pacientes_atendidos": "64",
    "atenciones_respiratorias": "28",
    "tiempo_espera": "180",
    "abandonos": "15",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 90,
    "fecha": "2026-05-11",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "175",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "300",
    "abandonos": "62",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 91,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "81",
    "pacientes_atendidos": "81",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "90",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 92,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "110",
    "pacientes_atendidos": "110",
    "atenciones_respiratorias": "54",
    "tiempo_espera": "78",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 93,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "101",
    "pacientes_atendidos": "99",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 94,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "131",
    "pacientes_atendidos": "125",
    "atenciones_respiratorias": "41",
    "tiempo_espera": "60",
    "abandonos": "6",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 95,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "120",
    "pacientes_atendidos": "109",
    "atenciones_respiratorias": "66",
    "tiempo_espera": "77",
    "abandonos": "11",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 96,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "95",
    "pacientes_atendidos": "94",
    "atenciones_respiratorias": "47",
    "tiempo_espera": "40",
    "abandonos": "1",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 97,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "268",
    "pacientes_atendidos": "222",
    "atenciones_respiratorias": "67",
    "tiempo_espera": "118",
    "abandonos": "46",
    "derivaciones": "14",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 98,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "61",
    "pacientes_atendidos": "60",
    "atenciones_respiratorias": "27",
    "tiempo_espera": "60",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 99,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "61",
    "pacientes_atendidos": "57",
    "atenciones_respiratorias": "21",
    "tiempo_espera": "120",
    "abandonos": "4",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 100,
    "fecha": "2026-05-12",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "127",
    "pacientes_atendidos": "115",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "180",
    "abandonos": "12",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 101,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "93",
    "pacientes_atendidos": "91",
    "atenciones_respiratorias": "19",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 102,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "148",
    "pacientes_atendidos": "146",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "85",
    "abandonos": "2",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 103,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "74",
    "pacientes_atendidos": "74",
    "atenciones_respiratorias": "31",
    "tiempo_espera": "120",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 104,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "125",
    "pacientes_atendidos": "119",
    "atenciones_respiratorias": "45",
    "tiempo_espera": "90",
    "abandonos": "6",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 105,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "140",
    "pacientes_atendidos": "132",
    "atenciones_respiratorias": "56",
    "tiempo_espera": "63",
    "abandonos": "8",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 106,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "93",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "45",
    "tiempo_espera": "140",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 107,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "251",
    "pacientes_atendidos": "220",
    "atenciones_respiratorias": "84",
    "tiempo_espera": "85",
    "abandonos": "31",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 108,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "81",
    "pacientes_atendidos": "73",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "45",
    "abandonos": "8",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 109,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "58",
    "pacientes_atendidos": "53",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "180",
    "abandonos": "5",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 110,
    "fecha": "2026-05-13",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "115",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 111,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "86",
    "pacientes_atendidos": "83",
    "atenciones_respiratorias": "24",
    "tiempo_espera": "90",
    "abandonos": "3",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 112,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "121",
    "pacientes_atendidos": "120",
    "atenciones_respiratorias": "44",
    "tiempo_espera": "131",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 113,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "74",
    "pacientes_atendidos": "72",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "180",
    "abandonos": "2",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 114,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "129",
    "pacientes_atendidos": "117",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "180",
    "abandonos": "12",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 115,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "113",
    "pacientes_atendidos": "104",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "74",
    "abandonos": "9",
    "derivaciones": "14",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 116,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "95",
    "pacientes_atendidos": "83",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "240",
    "abandonos": "12",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 117,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "220",
    "pacientes_atendidos": "184",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "186",
    "abandonos": "36",
    "derivaciones": "13",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 118,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "84",
    "pacientes_atendidos": "83",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "60",
    "abandonos": "1",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 119,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "70",
    "pacientes_atendidos": "59",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "180",
    "abandonos": "11",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 120,
    "fecha": "2026-05-14",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "142",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "41",
    "tiempo_espera": "500",
    "abandonos": "30",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 121,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "76",
    "pacientes_atendidos": "75",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "120",
    "abandonos": "1",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 122,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "118",
    "pacientes_atendidos": "107",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "210",
    "abandonos": "11",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 123,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "71",
    "pacientes_atendidos": "69",
    "atenciones_respiratorias": "23",
    "tiempo_espera": "60",
    "abandonos": "2",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 124,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "93",
    "pacientes_atendidos": "82",
    "atenciones_respiratorias": "18",
    "tiempo_espera": "60",
    "abandonos": "11",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 125,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "108",
    "pacientes_atendidos": "103",
    "atenciones_respiratorias": "55",
    "tiempo_espera": "99",
    "abandonos": "5",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 126,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "75",
    "pacientes_atendidos": "74",
    "atenciones_respiratorias": "36",
    "tiempo_espera": "50",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 127,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "215",
    "pacientes_atendidos": "207",
    "atenciones_respiratorias": "68",
    "tiempo_espera": "67",
    "abandonos": "7",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 128,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "",
    "pacientes_atendidos": "",
    "atenciones_respiratorias": "",
    "tiempo_espera": "",
    "abandonos": "",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 129,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "62",
    "pacientes_atendidos": "51",
    "atenciones_respiratorias": "11",
    "tiempo_espera": "30",
    "abandonos": "11",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 130,
    "fecha": "2026-05-15",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "127",
    "pacientes_atendidos": "119",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "180",
    "abandonos": "8",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 131,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "91",
    "pacientes_atendidos": "91",
    "atenciones_respiratorias": "35",
    "tiempo_espera": "91",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 132,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Maipú",
    "demanda_total": "151",
    "pacientes_atendidos": "150",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "45",
    "abandonos": "1",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 133,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Voullieme",
    "demanda_total": "96",
    "pacientes_atendidos": "96",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 134,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Chuchunco",
    "demanda_total": "173",
    "pacientes_atendidos": "171",
    "atenciones_respiratorias": "63",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 135,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Juricic",
    "demanda_total": "152",
    "pacientes_atendidos": "121",
    "atenciones_respiratorias": "60",
    "tiempo_espera": "71",
    "abandonos": "31",
    "derivaciones": "7",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 136,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Padre Vicente",
    "demanda_total": "102",
    "pacientes_atendidos": "100",
    "atenciones_respiratorias": "37",
    "tiempo_espera": "30",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 137,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "200",
    "pacientes_atendidos": "195",
    "atenciones_respiratorias": "62",
    "tiempo_espera": "38",
    "abandonos": "5",
    "derivaciones": "13",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 138,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Insunza",
    "demanda_total": "110",
    "pacientes_atendidos": "109",
    "atenciones_respiratorias": "48",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 139,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "Domeyko",
    "demanda_total": "99",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "20",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 140,
    "fecha": "2026-05-16",
    "semana_epi": "SE 19",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "182",
    "pacientes_atendidos": "120",
    "atenciones_respiratorias": "62",
    "tiempo_espera": "240",
    "abandonos": "62",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 141,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "89",
    "pacientes_atendidos": "89",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 142,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "195",
    "pacientes_atendidos": "192",
    "atenciones_respiratorias": "92",
    "tiempo_espera": "55",
    "abandonos": "3",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 143,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "114",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "40",
    "abandonos": "1",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 144,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "172",
    "pacientes_atendidos": "161",
    "atenciones_respiratorias": "67",
    "tiempo_espera": "90",
    "abandonos": "11",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 145,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "210",
    "pacientes_atendidos": "201",
    "atenciones_respiratorias": "128",
    "tiempo_espera": "49",
    "abandonos": "9",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 146,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "99",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "100",
    "abandonos": "2",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 147,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "217",
    "pacientes_atendidos": "197",
    "atenciones_respiratorias": "80",
    "tiempo_espera": "51",
    "abandonos": "20",
    "derivaciones": "13",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 148,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "111",
    "pacientes_atendidos": "110",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 149,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "80",
    "pacientes_atendidos": "80",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "10",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 150,
    "fecha": "2026-05-17",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "189",
    "pacientes_atendidos": "163",
    "atenciones_respiratorias": "75",
    "tiempo_espera": "300",
    "abandonos": "26",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 151,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "86",
    "pacientes_atendidos": "82",
    "atenciones_respiratorias": "24",
    "tiempo_espera": "120",
    "abandonos": "4",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 152,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "165",
    "pacientes_atendidos": "151",
    "atenciones_respiratorias": "54",
    "tiempo_espera": "171",
    "abandonos": "14",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 153,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "110",
    "pacientes_atendidos": "109",
    "atenciones_respiratorias": "59",
    "tiempo_espera": "180",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 154,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "128",
    "pacientes_atendidos": "115",
    "atenciones_respiratorias": "35",
    "tiempo_espera": "180",
    "abandonos": "13",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 155,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "139",
    "pacientes_atendidos": "117",
    "atenciones_respiratorias": "69",
    "tiempo_espera": "82",
    "abandonos": "22",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 156,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "111",
    "pacientes_atendidos": "100",
    "atenciones_respiratorias": "61",
    "tiempo_espera": "210",
    "abandonos": "11",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 157,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "283",
    "pacientes_atendidos": "226",
    "atenciones_respiratorias": "82",
    "tiempo_espera": "124",
    "abandonos": "57",
    "derivaciones": "12",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 158,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "119",
    "pacientes_atendidos": "103",
    "atenciones_respiratorias": "55",
    "tiempo_espera": "90",
    "abandonos": "16",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 159,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "67",
    "pacientes_atendidos": "52",
    "atenciones_respiratorias": "21",
    "tiempo_espera": "180",
    "abandonos": "15",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 160,
    "fecha": "2026-05-18",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "177",
    "pacientes_atendidos": "137",
    "atenciones_respiratorias": "68",
    "tiempo_espera": "300",
    "abandonos": "40",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 161,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "97",
    "pacientes_atendidos": "94",
    "atenciones_respiratorias": "37",
    "tiempo_espera": "90",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 162,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "147",
    "pacientes_atendidos": "145",
    "atenciones_respiratorias": "69",
    "tiempo_espera": "128",
    "abandonos": "2",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 163,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "90",
    "pacientes_atendidos": "85",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "100",
    "abandonos": "5",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 164,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "141",
    "pacientes_atendidos": "138",
    "atenciones_respiratorias": "63",
    "tiempo_espera": "60",
    "abandonos": "2",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 165,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "147",
    "pacientes_atendidos": "127",
    "atenciones_respiratorias": "79",
    "tiempo_espera": "158",
    "abandonos": "20",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 166,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "106",
    "pacientes_atendidos": "99",
    "atenciones_respiratorias": "42",
    "tiempo_espera": "130",
    "abandonos": "7",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 167,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "282",
    "pacientes_atendidos": "268",
    "atenciones_respiratorias": "118",
    "tiempo_espera": "57",
    "abandonos": "14",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 168,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "121",
    "pacientes_atendidos": "96",
    "atenciones_respiratorias": "47",
    "tiempo_espera": "120",
    "abandonos": "25",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 169,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "55",
    "pacientes_atendidos": "53",
    "atenciones_respiratorias": "13",
    "tiempo_espera": "180",
    "abandonos": "2",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 170,
    "fecha": "2026-05-19",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "151",
    "pacientes_atendidos": "129",
    "atenciones_respiratorias": "44",
    "tiempo_espera": "180",
    "abandonos": "22",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 171,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "66",
    "pacientes_atendidos": "65",
    "atenciones_respiratorias": "20",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 172,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "95",
    "pacientes_atendidos": "95",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "40",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 173,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "63",
    "pacientes_atendidos": "63",
    "atenciones_respiratorias": "23",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 174,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "101",
    "pacientes_atendidos": "97",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "40",
    "abandonos": "4",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 175,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "94",
    "pacientes_atendidos": "88",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "89",
    "abandonos": "6",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 176,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "78",
    "pacientes_atendidos": "76",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 177,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "244",
    "pacientes_atendidos": "215",
    "atenciones_respiratorias": "70",
    "tiempo_espera": "24",
    "abandonos": "29",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 178,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "58",
    "pacientes_atendidos": "52",
    "atenciones_respiratorias": "18",
    "tiempo_espera": "45",
    "abandonos": "6",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 179,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "57",
    "pacientes_atendidos": "48",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "180",
    "abandonos": "9",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 180,
    "fecha": "2026-05-20",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "104",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "245",
    "abandonos": "11",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 181,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "80",
    "pacientes_atendidos": "80",
    "atenciones_respiratorias": "29",
    "tiempo_espera": "25",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 182,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "103",
    "pacientes_atendidos": "103",
    "atenciones_respiratorias": "42",
    "tiempo_espera": "28",
    "abandonos": "0",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 183,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "89",
    "pacientes_atendidos": "89",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "10",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 184,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "130",
    "pacientes_atendidos": "126",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "60",
    "abandonos": "4",
    "derivaciones": "10",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 185,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "153",
    "pacientes_atendidos": "149",
    "atenciones_respiratorias": "89",
    "tiempo_espera": "49",
    "abandonos": "4",
    "derivaciones": "10",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 186,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "92",
    "pacientes_atendidos": "77",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "70",
    "abandonos": "15",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 187,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "154",
    "pacientes_atendidos": "152",
    "atenciones_respiratorias": "66",
    "tiempo_espera": "17",
    "abandonos": "2",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 188,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "109",
    "pacientes_atendidos": "108",
    "atenciones_respiratorias": "50",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 189,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "72",
    "pacientes_atendidos": "70",
    "atenciones_respiratorias": "20",
    "tiempo_espera": "15",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 190,
    "fecha": "2026-05-21",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "138",
    "pacientes_atendidos": "114",
    "atenciones_respiratorias": "47",
    "tiempo_espera": "240",
    "abandonos": "24",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 191,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "78",
    "pacientes_atendidos": "76",
    "atenciones_respiratorias": "26",
    "tiempo_espera": "120",
    "abandonos": "2",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 192,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "103",
    "pacientes_atendidos": "101",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "45",
    "abandonos": "2",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 193,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "66",
    "pacientes_atendidos": "65",
    "atenciones_respiratorias": "23",
    "tiempo_espera": "90",
    "abandonos": "1",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 194,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "88",
    "pacientes_atendidos": "83",
    "atenciones_respiratorias": "17",
    "tiempo_espera": "90",
    "abandonos": "6",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 195,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "92",
    "pacientes_atendidos": "90",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "37",
    "abandonos": "2",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 196,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "72",
    "pacientes_atendidos": "71",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "30",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 197,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "183",
    "pacientes_atendidos": "159",
    "atenciones_respiratorias": "66",
    "tiempo_espera": "82",
    "abandonos": "24",
    "derivaciones": "7",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 198,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "78",
    "pacientes_atendidos": "76",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "25",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 199,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "58",
    "pacientes_atendidos": "57",
    "atenciones_respiratorias": "13",
    "tiempo_espera": "20",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 200,
    "fecha": "2026-05-22",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "119",
    "pacientes_atendidos": "111",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "180",
    "abandonos": "8",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 201,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "71",
    "pacientes_atendidos": "71",
    "atenciones_respiratorias": "27",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 202,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Maipú",
    "demanda_total": "141",
    "pacientes_atendidos": "139",
    "atenciones_respiratorias": "58",
    "tiempo_espera": "35",
    "abandonos": "1",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 203,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Voullieme",
    "demanda_total": "65",
    "pacientes_atendidos": "67",
    "atenciones_respiratorias": "37",
    "tiempo_espera": "20",
    "abandonos": "1",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 204,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Chuchunco",
    "demanda_total": "127",
    "pacientes_atendidos": "116",
    "atenciones_respiratorias": "47",
    "tiempo_espera": "90",
    "abandonos": "11",
    "derivaciones": "11",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 205,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Juricic",
    "demanda_total": "132",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "63",
    "tiempo_espera": "29",
    "abandonos": "20",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 206,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Padre Vicente",
    "demanda_total": "84",
    "pacientes_atendidos": "84",
    "atenciones_respiratorias": "48",
    "tiempo_espera": "60",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 207,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "164",
    "pacientes_atendidos": "162",
    "atenciones_respiratorias": "59",
    "tiempo_espera": "24",
    "abandonos": "2",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 208,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Insunza",
    "demanda_total": "",
    "pacientes_atendidos": "",
    "atenciones_respiratorias": "",
    "tiempo_espera": "",
    "abandonos": "",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 209,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "Domeyko",
    "demanda_total": "73",
    "pacientes_atendidos": "73",
    "atenciones_respiratorias": "28",
    "tiempo_espera": "20",
    "abandonos": "0",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 210,
    "fecha": "2026-05-23",
    "semana_epi": "SE 20",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "146",
    "pacientes_atendidos": "143",
    "atenciones_respiratorias": "64",
    "tiempo_espera": "140",
    "abandonos": "3",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 211,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "82",
    "pacientes_atendidos": "82",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 212,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "151",
    "pacientes_atendidos": "150",
    "atenciones_respiratorias": "90",
    "tiempo_espera": "60",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 213,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "103",
    "pacientes_atendidos": "105",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 214,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "158",
    "pacientes_atendidos": "147",
    "atenciones_respiratorias": "64",
    "tiempo_espera": "150",
    "abandonos": "11",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 215,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "126",
    "pacientes_atendidos": "160",
    "atenciones_respiratorias": "83",
    "tiempo_espera": "114",
    "abandonos": "34",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 216,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "131",
    "pacientes_atendidos": "129",
    "atenciones_respiratorias": "71",
    "tiempo_espera": "60",
    "abandonos": "3",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 217,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "193",
    "pacientes_atendidos": "174",
    "atenciones_respiratorias": "73",
    "tiempo_espera": "90",
    "abandonos": "19",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 218,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "117",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "51",
    "tiempo_espera": "30",
    "abandonos": "4",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 219,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "66",
    "pacientes_atendidos": "66",
    "atenciones_respiratorias": "32",
    "tiempo_espera": "10",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 220,
    "fecha": "2026-05-24",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "180",
    "pacientes_atendidos": "134",
    "atenciones_respiratorias": "65",
    "tiempo_espera": "480",
    "abandonos": "46",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 221,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "90",
    "pacientes_atendidos": "87",
    "atenciones_respiratorias": "41",
    "tiempo_espera": "110",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 222,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "169",
    "pacientes_atendidos": "164",
    "atenciones_respiratorias": "69",
    "tiempo_espera": "125",
    "abandonos": "5",
    "derivaciones": "7",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 223,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "128",
    "pacientes_atendidos": "118",
    "atenciones_respiratorias": "48",
    "tiempo_espera": "140",
    "abandonos": "10",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 224,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "127",
    "pacientes_atendidos": "121",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "60",
    "abandonos": "6",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 225,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "141",
    "pacientes_atendidos": "108",
    "atenciones_respiratorias": "65",
    "tiempo_espera": "108",
    "abandonos": "33",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 226,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "124",
    "pacientes_atendidos": "107",
    "atenciones_respiratorias": "56",
    "tiempo_espera": "180",
    "abandonos": "17",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 227,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "273",
    "pacientes_atendidos": "248",
    "atenciones_respiratorias": "118",
    "tiempo_espera": "109",
    "abandonos": "25",
    "derivaciones": "17",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 228,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "111",
    "pacientes_atendidos": "107",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "45",
    "abandonos": "4",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 229,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "95",
    "pacientes_atendidos": "66",
    "atenciones_respiratorias": "33",
    "tiempo_espera": "180",
    "abandonos": "29",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 230,
    "fecha": "2026-05-25",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "172",
    "pacientes_atendidos": "110",
    "atenciones_respiratorias": "40",
    "tiempo_espera": "180",
    "abandonos": "62",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 231,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "102",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "36",
    "tiempo_espera": "75",
    "abandonos": "9",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 232,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "137",
    "pacientes_atendidos": "136",
    "atenciones_respiratorias": "74",
    "tiempo_espera": "58",
    "abandonos": "1",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 233,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "105",
    "pacientes_atendidos": "94",
    "atenciones_respiratorias": "53",
    "tiempo_espera": "120",
    "abandonos": "11",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 234,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "132",
    "pacientes_atendidos": "126",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "120",
    "abandonos": "6",
    "derivaciones": "7",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 235,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "136",
    "pacientes_atendidos": "116",
    "atenciones_respiratorias": "65",
    "tiempo_espera": "168",
    "abandonos": "20",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 236,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "102",
    "pacientes_atendidos": "86",
    "atenciones_respiratorias": "46",
    "tiempo_espera": "180",
    "abandonos": "16",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 237,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "253",
    "pacientes_atendidos": "225",
    "atenciones_respiratorias": "91",
    "tiempo_espera": "119",
    "abandonos": "28",
    "derivaciones": "9",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 238,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "93",
    "pacientes_atendidos": "79",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "90",
    "abandonos": "14",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 239,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "80",
    "pacientes_atendidos": "73",
    "atenciones_respiratorias": "27",
    "tiempo_espera": "180",
    "abandonos": "7",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 240,
    "fecha": "2026-05-26",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "145",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "52",
    "tiempo_espera": "440",
    "abandonos": "32",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 241,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "87",
    "pacientes_atendidos": "84",
    "atenciones_respiratorias": "28",
    "tiempo_espera": "40",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 242,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "139",
    "pacientes_atendidos": "136",
    "atenciones_respiratorias": "67",
    "tiempo_espera": "64",
    "abandonos": "3",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 243,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "76",
    "pacientes_atendidos": "73",
    "atenciones_respiratorias": "18",
    "tiempo_espera": "60",
    "abandonos": "3",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 244,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "121",
    "pacientes_atendidos": "115",
    "atenciones_respiratorias": "48",
    "tiempo_espera": "90",
    "abandonos": "6",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 245,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "134",
    "pacientes_atendidos": "129",
    "atenciones_respiratorias": "71",
    "tiempo_espera": "27",
    "abandonos": "5",
    "derivaciones": "4",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 246,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "112",
    "pacientes_atendidos": "107",
    "atenciones_respiratorias": "63",
    "tiempo_espera": "130",
    "abandonos": "5",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 247,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "268",
    "pacientes_atendidos": "241",
    "atenciones_respiratorias": "102",
    "tiempo_espera": "61",
    "abandonos": "27",
    "derivaciones": "18",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 248,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "",
    "pacientes_atendidos": "",
    "atenciones_respiratorias": "",
    "tiempo_espera": "",
    "abandonos": "",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 249,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "61",
    "pacientes_atendidos": "59",
    "atenciones_respiratorias": "22",
    "tiempo_espera": "90",
    "abandonos": "2",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 250,
    "fecha": "2026-05-27",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "159",
    "pacientes_atendidos": "132",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "240",
    "abandonos": "27",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 251,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "86",
    "pacientes_atendidos": "80",
    "atenciones_respiratorias": "24",
    "tiempo_espera": "150",
    "abandonos": "6",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 252,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "128",
    "pacientes_atendidos": "127",
    "atenciones_respiratorias": "55",
    "tiempo_espera": "88",
    "abandonos": "1",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 253,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "84",
    "pacientes_atendidos": "84",
    "atenciones_respiratorias": "39",
    "tiempo_espera": "180",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 254,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "131",
    "pacientes_atendidos": "113",
    "atenciones_respiratorias": "41",
    "tiempo_espera": "180",
    "abandonos": "18",
    "derivaciones": "5",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 255,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "102",
    "pacientes_atendidos": "93",
    "atenciones_respiratorias": "51",
    "tiempo_espera": "47",
    "abandonos": "9",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 256,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "109",
    "pacientes_atendidos": "102",
    "atenciones_respiratorias": "66",
    "tiempo_espera": "150",
    "abandonos": "7",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 257,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "244",
    "pacientes_atendidos": "208",
    "atenciones_respiratorias": "71",
    "tiempo_espera": "166",
    "abandonos": "36",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 258,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "67",
    "pacientes_atendidos": "51",
    "atenciones_respiratorias": "19",
    "tiempo_espera": "120",
    "abandonos": "16",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 259,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "76",
    "pacientes_atendidos": "72",
    "atenciones_respiratorias": "28",
    "tiempo_espera": "120",
    "abandonos": "4",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 260,
    "fecha": "2026-05-28",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "139",
    "pacientes_atendidos": "110",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "450",
    "abandonos": "29",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 261,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "91",
    "pacientes_atendidos": "90",
    "atenciones_respiratorias": "26",
    "tiempo_espera": "195",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 262,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "140",
    "pacientes_atendidos": "137",
    "atenciones_respiratorias": "58",
    "tiempo_espera": "51",
    "abandonos": "3",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 263,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "77",
    "pacientes_atendidos": "77",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "160",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 264,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "114",
    "pacientes_atendidos": "88",
    "atenciones_respiratorias": "16",
    "tiempo_espera": "210",
    "abandonos": "26",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 265,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "91",
    "pacientes_atendidos": "87",
    "atenciones_respiratorias": "45",
    "tiempo_espera": "47",
    "abandonos": "4",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 266,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "72",
    "pacientes_atendidos": "71",
    "atenciones_respiratorias": "43",
    "tiempo_espera": "45",
    "abandonos": "1",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 267,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "235",
    "pacientes_atendidos": "204",
    "atenciones_respiratorias": "83",
    "tiempo_espera": "65",
    "abandonos": "31",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 268,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "99",
    "pacientes_atendidos": "92",
    "atenciones_respiratorias": "38",
    "tiempo_espera": "50",
    "abandonos": "7",
    "derivaciones": "2",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 269,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "55",
    "pacientes_atendidos": "55",
    "atenciones_respiratorias": "11",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 270,
    "fecha": "2026-05-29",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "120",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "49",
    "tiempo_espera": "180",
    "abandonos": "8",
    "derivaciones": "3",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 271,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "CESFAM N°1",
    "demanda_total": "111",
    "pacientes_atendidos": "108",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "75",
    "abandonos": "3",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 272,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Maipú",
    "demanda_total": "144",
    "pacientes_atendidos": "145",
    "atenciones_respiratorias": "72",
    "tiempo_espera": "40",
    "abandonos": "1",
    "derivaciones": "7",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 273,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Voullieme",
    "demanda_total": "115",
    "pacientes_atendidos": "112",
    "atenciones_respiratorias": "54",
    "tiempo_espera": "70",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 274,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Chuchunco",
    "demanda_total": "183",
    "pacientes_atendidos": "173",
    "atenciones_respiratorias": "67",
    "tiempo_espera": "120",
    "abandonos": "10",
    "derivaciones": "8",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 275,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Juricic",
    "demanda_total": "",
    "pacientes_atendidos": "",
    "atenciones_respiratorias": "",
    "tiempo_espera": "",
    "abandonos": "",
    "derivaciones": "6",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 276,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Padre Vicente",
    "demanda_total": "107",
    "pacientes_atendidos": "107",
    "atenciones_respiratorias": "56",
    "tiempo_espera": "30",
    "abandonos": "0",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 277,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Pincheira",
    "demanda_total": "198",
    "pacientes_atendidos": "171",
    "atenciones_respiratorias": "60",
    "tiempo_espera": "97",
    "abandonos": "27",
    "derivaciones": "10",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 278,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Insunza",
    "demanda_total": "",
    "pacientes_atendidos": "",
    "atenciones_respiratorias": "",
    "tiempo_espera": "",
    "abandonos": "",
    "derivaciones": "",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 279,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "Domeyko",
    "demanda_total": "79",
    "pacientes_atendidos": "79",
    "atenciones_respiratorias": "34",
    "tiempo_espera": "15",
    "abandonos": "0",
    "derivaciones": "0",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  },
  {
    "id": 280,
    "fecha": "2026-05-30",
    "semana_epi": "SE 21",
    "establecimiento": "SAR Michelle Bachelet",
    "demanda_total": "176",
    "pacientes_atendidos": "159",
    "atenciones_respiratorias": "67",
    "tiempo_espera": "240",
    "abandonos": "17",
    "derivaciones": "1",
    "destino_derivacion": "",
    "tiene_refuerzo": false,
    "tipo_refuerzo": "",
    "horas_refuerzo": "",
    "observaciones": ""
  }
];

// Mapeo establecimiento → polo
// POLO_MAP: incluye nombres completos (desde Excel) y abreviados (ingreso manual)
const POLO_MAP_RAW = {
  // Polo Cerrillos Maipú — nombres completos
  "SAPU Dr. Norman Voulliéme":          "Polo Cerrillos Maipú",
  "SAR Enfermera Sofía Pincheira":       "Polo Cerrillos Maipú",
  "SAPU Maipú":                          "Polo Cerrillos Maipú",
  "SAPU Dra. Ana María Juricic":         "Polo Cerrillos Maipú",
  "SAR Michelle Bachelet":               "Polo Cerrillos Maipú",
  "SAPU Dr. Iván Insunza":               "Polo Cerrillos Maipú",
  // Polo Cerrillos Maipú — nombres abreviados (ingreso manual)
  "Voullieme":                           "Polo Cerrillos Maipú",
  "Voulliéme":                           "Polo Cerrillos Maipú",
  "SAR Pincheira":                       "Polo Cerrillos Maipú",
  "Pincheira":                           "Polo Cerrillos Maipú",
  "Maipú":                               "Polo Cerrillos Maipú",
  "Juricic":                             "Polo Cerrillos Maipú",
  "Michelle Bachelet":                   "Polo Cerrillos Maipú",
  "Insunza":                             "Polo Cerrillos Maipú",
  // Polo Santiago Estación Central — nombres completos
  "SAPU Consultorio Nº1":                "Polo Santiago Estación Central",
  "SAPU Ignacio Domeyko":                "Polo Santiago Estación Central",
  "SAPU Padre Vicente Irarrázabal":      "Polo Santiago Estación Central",
  "SAPU San José de Chuchunco":          "Polo Santiago Estación Central",
  // Polo Santiago Estación Central — nombres abreviados (ingreso manual)
  "Consultorio Nº1":                     "Polo Santiago Estación Central",
  "Consultorio N°1":                     "Polo Santiago Estación Central",
  "SAPU Consultorio N°1":                "Polo Santiago Estación Central",
  "Domeyko":                             "Polo Santiago Estación Central",
  "Ignacio Domeyko":                     "Polo Santiago Estación Central",
  "Padre Vicente":                       "Polo Santiago Estación Central",
  "Padre Vicente Irarrázabal":           "Polo Santiago Estación Central",
  "Chuchunco":                           "Polo Santiago Estación Central",
  "San José de Chuchunco":               "Polo Santiago Estación Central",
};

// Función que busca polo por nombre exacto o por coincidencia parcial
const getPolo = (estab) => {
  if (!estab) return null;
  // Búsqueda exacta
  if (POLO_MAP_RAW[estab]) return POLO_MAP_RAW[estab];
  // Búsqueda parcial (si el nombre contiene alguna clave)
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
  const [registros, setRegistros] = useState([]);
  const [registrosAmbulancias, setRegistrosAmbulancias] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState(EMPTY_FORM);
  const [editId, setEditId] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [importando, setImportando] = useState(false);
  const [proyMetodo, setProyMetodo] = useState("historico");
  const [proySemanas, setProyeSemanas] = useState(4);
  const [importResultado, setImportResultado] = useState(null);
  // Iniciar con la SE actual seleccionada
  const _hoy = new Date();
  const _hoyStr = `${_hoy.getFullYear()}-${String(_hoy.getMonth()+1).padStart(2,"0")}-${String(_hoy.getDate()).padStart(2,"0")}`;
  const [filtroSemana, setFiltroSemana] = useState(getEpiWeek(_hoyStr));
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

  // Retenciones de Ambulancias
  const [formAmbulancias, setFormAmbulancias] = useState(
    Array.from({ length: AMBULANCIA_FILAS }, () => ({ ...EMPTY_AMBULANCIA_ROW }))
  );
  const [deleteConfirmAmb, setDeleteConfirmAmb] = useState(null);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3200);
  };

  const semanas = useMemo(() => ["Todas", ...new Set(registros.map(r => r.semana_epi))].sort(), [registros]);
  const semanasOpts = useMemo(() => [...new Set(registros.map(r => r.semana_epi).filter(Boolean))].sort(), [registros]);

  // Comparador: calcular métricas para un rango de SE
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
      tasaAbandono: sum("demanda_total") ? ((sum("abandonos") / sum("demanda_total")) * 100).toFixed(1) : 0,
      pctResp:   sum("pacientes_atendidos") ? ((sum("atenciones_respiratorias") / sum("pacientes_atendidos")) * 100).toFixed(1) : 0,
    };
  }, [registros, semanasOpts, filtroPolo, filtroEstab]);
  const estabs  = useMemo(() => ["Todos", ...new Set(registros.map(r => r.establecimiento))], [registros]);
  const POLOS   = ["Todos", "Polo Cerrillos Maipú", "Polo Santiago Estación Central"];

  // Centros con ingreso de datos pendiente — últimas 5 semanas epidemiológicas
  const seActualDashboard = useMemo(() => getEpiWeek(_hoyStr), [_hoyStr]);
  const ultimasSEDashboard = useMemo(() => {
    const out = [];
    for (let i = 0; i < 5; i++) {
      const d = new Date(_hoy);
      d.setDate(d.getDate() - i * 7);
      const ds = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
      out.push(getEpiWeek(ds));
    }
    return out; // [actual, -1, -2]
  }, [_hoyStr]);
  const pendientesPorSE = useMemo(() => {
    return ultimasSEDashboard.map(se => {
      // Fechas en que ALGÚN establecimiento ya registró datos para esta SE
      const fechasConDatos = [...new Set(
        registros.filter(r => r.semana_epi === se && r.fecha).map(r => r.fecha)
      )].sort();

      const pendientes = ESTABLECIMIENTOS.filter(e => {
        // ¿Le falta algún día en que otros centros ya reportaron en esta SE?
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
    const conEspera = filtrados.filter(r => r.tiempo_espera !== "" && r.tiempo_espera !== undefined);
    const promEspera = conEspera.length
      ? (conEspera.reduce((a, r) => a + Number(r.tiempo_espera || 0), 0) / conEspera.length).toFixed(0)
      : 0;
    return {
      ...tot,
      tasaAbandono: tot.demanda ? ((tot.abandonos / tot.demanda) * 100).toFixed(1) : 0,
      tasaResp: tot.atendidos ? ((tot.respiratorias / tot.atendidos) * 100).toFixed(1) : 0,
      promCon: pCon.toFixed(1), promSin: pSin.toFixed(1),
      impacto: (pCon - pSin).toFixed(1),
      promEspera,
    };
  }, [filtrados]);

  const dataXSemana = useMemo(() => {
    const map = {};
    registros.forEach(r => {
      if (!map[r.semana_epi]) map[r.semana_epi] = { semana: r.semana_epi, demanda: 0, atendidos: 0, respiratorias: 0, abandonos: 0 };
      map[r.semana_epi].demanda += Number(r.demanda_total || 0);
      map[r.semana_epi].atendidos += Number(r.pacientes_atendidos || 0);
      map[r.semana_epi].respiratorias += Number(r.atenciones_respiratorias || 0);
      map[r.semana_epi].abandonos += Number(r.abandonos || 0);
    });
    return Object.values(map).sort((a, b) => a.semana.localeCompare(b.semana));
  }, [registros]);

  // Gráfico: atendidos vs respiratorias por semana + acumulado
  const dataRespAcum = useMemo(() => {
    const map = {};
    filtrados.forEach(r => {
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

  // Absorción: respeta filtro de Polo y SE, pero IGNORA filtro de Establecimiento
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

  const dataRefuerzo = useMemo(() => [
    { name: "Con refuerzo", atendidos: Number(kpis.promCon) },
    { name: "Sin refuerzo", atendidos: Number(kpis.promSin) },
  ], [kpis]);

  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setForm(prev => ({
      ...prev,
      [name]: type === "checkbox" ? checked : value,
      ...(name === "fecha" ? { semana_epi: getEpiWeek(value) } : {}),
    }));
  };

  // ── Suscripciones Firestore en tiempo real ──────────────────────────────
  useEffect(() => {
    // Carga inicial de registros
    // Traer todos los registros en lotes de 1000 (límite de Supabase)
    const fetchAll = async (tabla) => {
      const pageSize = 1000;
      let all = [];
      let from = 0;
      while (true) {
        const { data, error } = await supabase
          .from(tabla)
          .select("*")
          .order("fecha", { ascending: false })
          .range(from, from + pageSize - 1);
        if (error || !data || data.length === 0) break;
        all = [...all, ...data];
        if (data.length < pageSize) break;
        from += pageSize;
      }
      return all;
    };

    const fetchData = async () => {
      try {
        const regs = await fetchAll("registros");
        setRegistros(regs);
        const ambs = await fetchAll("retenciones_ambulancias");
        setRegistrosAmbulancias(ambs);
      } catch {
        // silencioso
      } finally {
        setLoading(false);
      }
    };
    fetchData();

    // Suscripciones en tiempo real via Supabase Realtime
    const chanRegistros = supabase
      .channel("registros-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "registros" }, () => {
        fetchAll("registros").then(data => setRegistros(data));
      })
      .subscribe();

    const chanAmbulancias = supabase
      .channel("ambulancias-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "retenciones_ambulancias" }, () => {
        fetchAll("retenciones_ambulancias").then(data => setRegistrosAmbulancias(data));
      })
      .subscribe();

    return () => {
      supabase.removeChannel(chanRegistros);
      supabase.removeChannel(chanAmbulancias);
    };
  }, []);

  const handleImportExcel = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportando(true);
    setImportResultado(null);
    try {
      const buffer = await file.arrayBuffer();
      // raw:true para leer valores crudos (números seriales de fecha)
      const wb = XLSX.read(buffer, { cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // raw:true mantiene fechas como objetos Date
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      // Helper: convierte Date o string a YYYY-MM-DD
      const toISO = (v) => {
        if (!v) return null;
        if (v instanceof Date) {
          const y = v.getFullYear();
          const m = String(v.getMonth()+1).padStart(2,"0");
          const d = String(v.getDate()).padStart(2,"0");
          return `${y}-${m}-${d}`;
        }
        // A veces viene como string "dd/mm/yyyy" o "yyyy-mm-dd"
        if (typeof v === "string") {
          const partes = v.split("/");
          if (partes.length === 3) return `${partes[2]}-${partes[1].padStart(2,"0")}-${partes[0].padStart(2,"0")}`;
          if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
        }
        return null;
      };

      // Buscar la fila que tiene fechas (objetos Date o strings de fecha)
      let filaFechas = -1;
      for (let r = 0; r < Math.min(raw.length, 10); r++) {
        const tieneFecha = raw[r].some(v => v instanceof Date || (typeof v === "string" && /\d{1,4}[/-]\d{1,2}[/-]\d{2,4}/.test(v)));
        if (tieneFecha) { filaFechas = r; break; }
      }
      if (filaFechas === -1) throw new Error("No se encontraron fechas en el archivo. Verifica el formato.");

      const filaHeaders = filaFechas + 1;
      const filasDatos  = filaFechas + 2;
      const fechasFila  = raw[filaFechas] || [];

      // Mapear columna → fecha usando headers relativos (máx 8 cols por bloque)
      // Estrategia: para cada fecha, buscar campos por nombre en las siguientes cols
      // pero deteniéndose al encontrar la próxima fecha o "TOTAL DEMANDA" repetido
      const headersFila = raw[filaHeaders] || [];
      const bloques = [];
      for (let c = 1; c < fechasFila.length; c++) {
        const fecha = toISO(fechasFila[c]);
        if (!fecha) continue;

        // Buscar solo hasta la próxima fecha real (máx 8 cols)
        const hdrMap = {};
        let demandaVisto = false;
        for (let i = 0; i < 8; i++) {
          // Si hay una fecha en la siguiente columna de datos, parar
          if (i > 0 && toISO(fechasFila[c+i])) break;
          const h = String(headersFila[c+i] || "").toUpperCase().trim();
          if (!h) continue;
          // Si vemos TOTAL DEMANDA por segunda vez, es el bloque siguiente
          if (h.includes("TOTAL DEMANDA") || h === "DEMANDA") {
            if (demandaVisto) break;
            demandaVisto = true;
            hdrMap.demanda = i;
          } else if (h.includes("ATENCIONES TOT")) { hdrMap.atendidos = i; }
          else if (h.includes("RESP"))              { hdrMap.resp      = i; }
          else if (h.includes("ESPERA"))            { hdrMap.espera    = i; }
          else if (h.includes("ABANDONO"))          { hdrMap.abandonos = i; }
          else if (h.includes("HEC"))               { hdrMap.hec       = i; }
          else if (h.includes("HUAP"))              { hdrMap.huap      = i; }
          else if (h.includes("HCSBA"))             { hdrMap.hcsba     = i; }
        }
        if (hdrMap.demanda !== undefined) {
          bloques.push({ col: c, fecha, hdrMap });
        }
      }

      if (bloques.length === 0) throw new Error("No se pudieron leer las fechas. Verifica el formato del archivo.");

      const n = (fila, c, idx) => {
        if (idx === undefined) return null;
        const v = fila[c + idx];
        return v != null && !isNaN(Number(v)) ? Number(v) : null;
      };

      const registrosNuevos = [];
      for (let r = filasDatos; r < raw.length; r++) {
        const fila = raw[r];
        if (!fila) continue;
        const estab = fila[0];
        if (!estab || String(estab).toUpperCase().includes("TOTAL")) continue;

        for (const { col: c, fecha, hdrMap } of bloques) {
          const demanda   = n(fila, c, hdrMap.demanda);
          const atendidos = n(fila, c, hdrMap.atendidos);
          const resp      = n(fila, c, hdrMap.resp);
          const espera    = n(fila, c, hdrMap.espera);
          const abandonos = n(fila, c, hdrMap.abandonos);
          const hec       = n(fila, c, hdrMap.hec);
          const huap      = n(fila, c, hdrMap.huap);
          const hcsba     = n(fila, c, hdrMap.hcsba);

          if (demanda === null && atendidos === null) continue;

          registrosNuevos.push({
            fecha,
            semana_epi:               getEpiWeek(fecha),
            establecimiento:          String(estab).trim(),
            demanda_total:            demanda,
            pacientes_atendidos:      atendidos,
            atenciones_respiratorias: resp,
            tiempo_espera:            espera,
            abandonos:                abandonos,
            derivaciones_hec:         hec,
            derivaciones_huap:        huap,
            derivaciones_hcsba:       hcsba,
          });
        }
      }

      if (registrosNuevos.length === 0) {
        setImportResultado({ ok: false, msg: "No se encontraron datos válidos en el archivo." });
        return;
      }

      // Verificar duplicados en lotes de 50 fechas para no saturar Supabase
      const fechasUnicas = [...new Set(registrosNuevos.map(r => r.fecha))];
      const claveExistente = new Set();
      for (let i = 0; i < fechasUnicas.length; i += 50) {
        const lote = fechasUnicas.slice(i, i + 50);
        const { data: existentes } = await supabase
          .from("registros")
          .select("fecha, establecimiento")
          .in("fecha", lote);
        (existentes || []).forEach(r => claveExistente.add(`${r.fecha}__${r.establecimiento}`));
      }

      const soloNuevos = registrosNuevos.filter(r =>
        !claveExistente.has(`${r.fecha}__${r.establecimiento}`)
      );
      const duplicados = registrosNuevos.length - soloNuevos.length;

      if (soloNuevos.length === 0) {
        setImportResultado({ ok: false, msg: `⚠️ Todos los registros ya existen en la base de datos (${duplicados} duplicados omitidos). No se cargó nada nuevo.` });
        return;
      }

      // Insertar solo los nuevos en lotes de 100
      let insertados = 0;
      for (let i = 0; i < soloNuevos.length; i += 100) {
        const lote = soloNuevos.slice(i, i + 100);
        const { error } = await supabase.from("registros").insert(lote);
        if (error) throw error;
        insertados += lote.length;
      }

      const msg = duplicados > 0
        ? `✅ ${insertados} registros importados. ${duplicados} duplicados omitidos automáticamente.`
        : `✅ ${insertados} registros importados correctamente desde ${file.name}`;

      setImportResultado({ ok: true, msg });
      e.target.value = "";
    } catch (err) {
      setImportResultado({ ok: false, msg: `❌ Error al importar: ${err.message}` });
    } finally {
      setImportando(false);
    }
  };

  const handleSubmit = async () => {
    if (!form.fecha || !form.establecimiento || !form.demanda_total) {
      showToast("Completa los campos obligatorios: fecha, establecimiento y demanda", "error");
      return;
    }
    // Convertir campos numéricos vacíos a null para Supabase
    const limpiarForm = (f) => {
      const numericos = ["demanda_total","pacientes_atendidos","atenciones_respiratorias",
        "abandonos","derivaciones_hec","derivaciones_hcsba","derivaciones_huap",
        "tiempo_espera","horas_refuerzo"];
      const limpio = { ...f };
      numericos.forEach(k => { if (limpio[k] === "" || limpio[k] === undefined) limpio[k] = null; });
      if (limpio.tipo_refuerzo === "") limpio.tipo_refuerzo = null;
      if (limpio.observaciones === "") limpio.observaciones = null;
      if (limpio.semana_epi === "") limpio.semana_epi = null;
      return limpio;
    };
    try {
      if (editId !== null) {
        const { id: _id, created_at: _ca, ...data } = limpiarForm(form);
        const { error } = await supabase.from("registros").update(data).eq("id", editId);
        if (error) throw error;
        showToast("Registro actualizado correctamente");
        setEditId(null);
      } else {
        const { error } = await supabase.from("registros").insert(limpiarForm(form));
        if (error) throw error;
        showToast("Registro guardado correctamente");
      }
      setForm(EMPTY_FORM);
      setTab("tabla");
    } catch {
      showToast("Error al guardar en Supabase", "error");
    }
  };

  const handleEdit = (r) => { setForm({ ...r }); setEditId(r.id); setTab("formulario"); };
  const handleDelete = async (id) => {
    try {
      const { error } = await supabase.from("registros").delete().eq("id", id);
      if (error) throw error;
      setDeleteConfirm(null);
      showToast("Registro eliminado", "warning");
    } catch {
      showToast("Error al eliminar", "error");
    }
  };

  // ── Retenciones de Ambulancias ──────────────────────────────
  const handleChangeAmbulancia = (idx, field, value) => {
    setFormAmbulancias(prev => prev.map((row, i) => i === idx ? { ...row, [field]: value } : row));
  };

  const handleSubmitAmbulancias = async () => {
    const filasCompletas = formAmbulancias.filter(r =>
      r.fecha && r.establecimiento && r.hora_traslado && r.tiempo_retencion !== ""
    );
    const filasParciales = formAmbulancias.filter(r =>
      (r.fecha || r.establecimiento || r.hora_traslado || r.tiempo_retencion !== "") &&
      !(r.fecha && r.establecimiento && r.hora_traslado && r.tiempo_retencion !== "")
    );
    if (filasCompletas.length === 0) {
      showToast("Completa al menos una línea con fecha, establecimiento, horario y tiempo de retención", "error");
      return;
    }
    try {
      const { error } = await supabase.from("retenciones_ambulancias").insert(filasCompletas);
      if (error) throw error;
      showToast(
        filasParciales.length > 0
          ? `${filasCompletas.length} registro(s) guardado(s). ${filasParciales.length} línea(s) incompleta(s) se descartaron.`
          : `${filasCompletas.length} registro(s) de retención guardado(s) correctamente`
      );
      setFormAmbulancias(Array.from({ length: AMBULANCIA_FILAS }, () => ({ ...EMPTY_AMBULANCIA_ROW })));
    } catch {
      showToast("Error al guardar en Supabase", "error");
    }
  };

  const handleDeleteAmbulancia = async (id) => {
    try {
      const { error } = await supabase.from("retenciones_ambulancias").delete().eq("id", id);
      if (error) throw error;
      setDeleteConfirmAmb(null);
      showToast("Registro de retención eliminado", "warning");
    } catch {
      showToast("Error al eliminar", "error");
    }
  };

  const registrosAmbulanciasOrdenados = useMemo(
    () => [...registrosAmbulancias].sort((a, b) =>
      (b.fecha + b.hora_traslado).localeCompare(a.fecha + a.hora_traslado)
    ),
    [registrosAmbulancias]
  );

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

  const dataAmbulanciasXEstablecimiento = useMemo(() => {
    const map = {};
    registrosAmbulancias.forEach(r => {
      const k = r.establecimiento || "Sin establecimiento";
      if (!map[k]) map[k] = { establecimiento: k, retenciones: 0, totalMin: 0 };
      map[k].retenciones += 1;
      map[k].totalMin += Number(r.tiempo_retencion || 0);
    });
    return Object.values(map)
      .map(r => ({ ...r, promedioMin: Math.round(r.totalMin / r.retenciones) }))
      .sort((a, b) => b.retenciones - a.retenciones);
  }, [registrosAmbulancias]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(filtrados.map(r => ({
      Fecha: r.fecha, "Semana Epi.": r.semana_epi, Establecimiento: r.establecimiento,
      "Demanda Total": r.demanda_total, "Pac. Atendidos": r.pacientes_atendidos,
      "At. Respiratorias": r.atenciones_respiratorias, Abandonos: r.abandonos,
      "Deriv. HEC": r.derivaciones_hec, "Deriv. HCSBA": r.derivaciones_hcsba, "Deriv. HUAP": r.derivaciones_huap,
      "T° Espera Prom. (min)": r.tiempo_espera,
      "Tiene Refuerzo": r.tiene_refuerzo ? "Sí" : "No", "Tipo Refuerzo": r.tipo_refuerzo,
      "Horas Refuerzo": r.horas_refuerzo, Observaciones: r.observaciones,
    })));
    XLSX.utils.book_append_sheet(wb, ws, "Registros");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataXSemana), "Resumen por Semana");
    if (registrosAmbulancias.length > 0) {
      const wsAmb = XLSX.utils.json_to_sheet(registrosAmbulanciasOrdenados.map(r => ({
        Fecha: r.fecha, Establecimiento: r.establecimiento,
        "Horario Traslado": r.hora_traslado, "Tiempo Retención (min)": r.tiempo_retencion,
      })));
      XLSX.utils.book_append_sheet(wb, wsAmb, "Retenciones Ambulancias");
    }
    XLSX.writeFile(wb, `AtencionesUrgencia_SSMC_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Archivo Excel exportado");
  };

  const PIE_COLORS = [P.azul, P.verde, P.amber, P.rojo, "#7B3FA0", "#1A7A9A"];

  const labelS = { display: "block", fontSize: 11, color: P.muted, marginBottom: 5, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.4px" };
  const inpS = { width: "100%", background: "#F8FAFC", color: P.text, border: `1px solid ${P.border}`, borderRadius: 7, padding: "9px 12px", fontSize: 13, boxSizing: "border-box", outline: "none" };
  const tdS = { padding: "10px 12px", whiteSpace: "nowrap", color: P.text, borderBottom: `1px solid ${P.border}` };

  return (
    <div style={{ background: P.bg, minHeight: "100vh", fontFamily: "'Inter', 'Segoe UI', 'Helvetica Neue', Arial, sans-serif", color: P.text }}>

      {loading && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(255,255,255,0.92)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", zIndex: 9999 }}>
          <div style={{ fontSize: 38, marginBottom: 14 }}>🔄</div>
          <div style={{ fontSize: 16, fontWeight: 700, color: P.azul }}>Cargando datos...</div>
          <div style={{ fontSize: 12, color: P.muted, marginTop: 6 }}>Cargando datos en tiempo real</div>
        </div>
      )}
      {/* Toast */}
      {toast && (
        <div style={{
          position: "fixed", top: 20, right: 20, zIndex: 9999,
          background: toast.type === "error" ? P.rojo : toast.type === "warning" ? "#92400E" : P.verde,
          color: "#fff", padding: "12px 20px", borderRadius: 8, fontSize: 13, fontWeight: 600,
          boxShadow: "0 4px 20px rgba(0,0,0,0.2)", animation: "fadeIn 0.3s ease"
        }}>{toast.msg}</div>
      )}

      {/* Header SSMC — institucional minimalista */}
      <div style={{ background: P.azul, borderBottom: `3px solid ${P.rojo}` }}>
        {/* Banda principal */}
        <div style={{ padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            {/* Logo SSMC */}
            <div style={{
              width: 46, height: 46, borderRadius: 6, background: "#fff",
              display: "flex", alignItems: "center", justifyContent: "center",
              flexShrink: 0, padding: 3, overflow: "hidden"
            }}>
              <img
                src="https://upload.wikimedia.org/wikipedia/commons/thumb/7/78/Escudo_de_Chile.svg/240px-Escudo_de_Chile.svg.png"
                alt="Escudo Chile"
                style={{ width: 36, height: 36, objectFit: "contain" }}
                onError={e => { e.target.style.display="none"; e.target.parentNode.innerHTML="<span style='font-size:22px'>🏥</span>"; }}
              />
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 700, fontSize: 15, letterSpacing: "-0.2px", lineHeight: 1.3 }}>
                Urgencias APS · SSMC
              </div>
              <div style={{ color: "rgba(255,255,255,0.6)", fontSize: 11, marginTop: 1 }}>
                Servicio de Salud Metropolitano Central · Ministerio de Salud
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
            <button onClick={exportExcel} style={{
              background: "transparent", color: "#fff",
              border: "1.5px solid rgba(255,255,255,0.35)",
              borderRadius: 6, padding: "6px 12px", cursor: "pointer",
              fontSize: 12, fontWeight: 600, display: "flex", alignItems: "center", gap: 4,
            }}>↓ Excel</button>
            <button onClick={() => setMostrarPDF(true)} style={{
              background: P.rojo, color: "#fff", border: "none",
              borderRadius: 6, padding: "6px 12px", cursor: "pointer",
              fontSize: 12, fontWeight: 600, display: "flex", alignItems: "center", gap: 4,
            }}>↓ PDF</button>
          </div>
        </div>

        {/* Nav Tabs — desktop */}
        <div className="nav-desktop" style={{ display: "flex", paddingLeft: 16, overflowX: "auto" }}>
          {[
            { id: "dashboard",    label: "Resumen" },
            { id: "formulario",   label: "Ingresar" },
            { id: "tabla",        label: "Tabla" },
            { id: "ambulancias",  label: "Ambulancias" },
            { id: "importar",     label: "Importar" },
            { id: "tiempos",      label: "T° Espera" },
            { id: "proyecciones", label: "Proyecciones" },
          ].map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              background: "transparent",
              color: tab === t.id ? "#fff" : "rgba(255,255,255,0.5)",
              border: "none",
              borderBottom: tab === t.id ? `3px solid ${P.rojo}` : "3px solid transparent",
              padding: "9px 16px", cursor: "pointer", fontSize: 12,
              fontWeight: tab === t.id ? 700 : 400,
              transition: "all 0.15s", whiteSpace: "nowrap",
            }}>{t.label}</button>
          ))}
        </div>

        {/* Nav Tabs — móvil hamburguesa */}
        <div className="nav-mobile" style={{ display: "none" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 16px" }}>
            <span style={{ color: "rgba(255,255,255,0.8)", fontSize: 13, fontWeight: 600 }}>
              {[
                {id:"dashboard",label:"Resumen"},{id:"formulario",label:"Ingresar"},
                {id:"tabla",label:"Tabla"},{id:"ambulancias",label:"Ambulancias"},
                {id:"importar",label:"Importar"},{id:"tiempos",label:"T° Espera"},
                {id:"proyecciones",label:"Proyecciones"},
              ].find(t => t.id === tab)?.label}
            </span>
            <button onClick={() => setMenuMovil(v => !v)} style={{
              background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 6,
              color: "#fff", padding: "5px 11px", cursor: "pointer", fontSize: 18, lineHeight: 1,
            }}>{menuMovil ? "✕" : "☰"}</button>
          </div>
          {menuMovil && (
            <div style={{ background: P.azulDark, borderTop: "1px solid rgba(255,255,255,0.1)" }}>
              {[
                {id:"dashboard",label:"📊 Resumen"},{id:"formulario",label:"➕ Ingresar"},
                {id:"tabla",label:"📋 Tabla"},{id:"ambulancias",label:"🚑 Ambulancias"},
                {id:"importar",label:"📥 Importar"},{id:"tiempos",label:"⏱️ T° Espera"},
                {id:"proyecciones",label:"📈 Proyecciones"},
              ].map(t => (
                <button key={t.id} onClick={() => { setTab(t.id); setMenuMovil(false); }} style={{
                  display: "block", width: "100%", textAlign: "left",
                  background: tab === t.id ? "rgba(192,57,43,0.25)" : "transparent",
                  color: tab === t.id ? "#fff" : "rgba(255,255,255,0.65)",
                  border: "none", borderLeft: tab === t.id ? `3px solid ${P.rojo}` : "3px solid transparent",
                  padding: "12px 20px", cursor: "pointer", fontSize: 13,
                  fontWeight: tab === t.id ? 700 : 400,
                }}>{t.label}</button>
              ))}
            </div>
          )}
        </div>
      </div>

      <div style={{ padding: "24px 28px" }}>

        {/* ── DASHBOARD ─────────────────────────────────────── */}
        {tab === "dashboard" && (
          <div>
            {/* Centros con ingreso de datos pendiente — últimas 3 semanas epidemiológicas */}
            <div style={{
              background: totalPendientesDashboard ? "#FFF8E1" : P.verdeLight,
              border: `1px solid ${totalPendientesDashboard ? "#ffc107" : P.verde}`,
              borderRadius: 12, padding: "16px 20px", marginBottom: 18,
            }}>
              <div style={{ display: "flex", gap: 16, alignItems: "flex-start", marginBottom: totalPendientesDashboard ? 14 : 0 }}>
                <div style={{ fontSize: 26, lineHeight: 1 }}>{totalPendientesDashboard ? "⏳" : "✅"}</div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 800, color: totalPendientesDashboard ? "#856404" : P.verde }}>
                    Ingreso de Datos Pendiente — Últimas 5 Semanas Epidemiológicas
                  </div>
                  <div style={{ fontSize: 12, color: totalPendientesDashboard ? "#856404" : P.verde, marginTop: 4 }}>
                    {totalPendientesDashboard
                      ? `Hay establecimientos con registros pendientes en las últimas 5 semanas epidemiológicas.`
                      : `Todos los establecimientos están al día en las últimas 5 semanas epidemiológicas.`}
                  </div>
                </div>
              </div>

              {totalPendientesDashboard > 0 && (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {pendientesPorSE.map(({ se, pendientes }) => (
                    <div key={se} style={{
                      background: pendientes.length ? "#fff" : "transparent",
                      border: pendientes.length ? `1px solid ${P.border}` : "none",
                      borderRadius: 8, padding: pendientes.length ? "10px 12px" : 0,
                    }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: pendientes.length ? "#856404" : P.verde, marginBottom: pendientes.length ? 8 : 0 }}>
                        {pendientes.length
                          ? <>La <b>{se}</b>{se === seActualDashboard ? " (actual)" : ""} tiene registros pendientes:</>
                          : <>La <b>{se}</b>{se === seActualDashboard ? " (actual)" : ""} está completa ✅</>}
                      </div>
                      {pendientes.length > 0 && (
                        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                          {pendientes.map(e => (
                            <span key={e} style={{
                              background: "#fff3cd", border: "1px solid #ffc107", borderRadius: 6,
                              padding: "4px 10px", fontSize: 11, fontWeight: 700, color: "#856404"
                            }}>
                              {e}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Filtros */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 10, padding: "14px 18px", marginBottom: 22, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: P.azulDark, alignSelf: "center", marginRight: 4 }}>🔍 Filtros</div>
              {[
                { label: "Polo", options: POLOS, val: filtroPolo, set: setFiltroPolo },
                { label: "Semana Epidemiológica", options: semanas, val: filtroSemana, set: setFiltroSemana },
                { label: "Establecimiento", options: ["Todos", ...ESTABLECIMIENTOS], val: filtroEstab, set: setFiltroEstab },
              ].map(f => (
                <div key={f.label}>
                  <div style={{ fontSize: 10, color: P.muted, marginBottom: 4, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.4px" }}>{f.label}</div>
                  <select value={f.val} onChange={e => f.set(e.target.value)} style={{
                    background: P.azulLight, color: P.azulDark, border: `1px solid ${P.azulMid}`,
                    borderRadius: 7, padding: "7px 12px", fontSize: 13, cursor: "pointer", fontWeight: 600
                  }}>
                    {f.options.map(o => <option key={o}>{o}</option>)}
                  </select>
                </div>
              ))}
              <div style={{ fontSize: 12, color: P.muted, paddingBottom: 2, marginLeft: "auto" }}>
                <b style={{ color: P.azul, fontSize: 16 }}>{filtrados.length}</b> registros
              </div>
            </div>

            {/* KPI Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(148px, 1fr))", gap: 14, marginBottom: 22 }}>
              {[
                { label: "Demanda Total",    val: kpis.demanda,            color: P.azul,    bg: P.azulLight },
                { label: "Atendidos",         val: kpis.atendidos,          color: P.verde,   bg: P.verdeLight },
                { label: "Respiratorias",     val: kpis.respiratorias,      color: P.amber,   bg: "#FEF3C7" },
                { label: "Abandonos",         val: kpis.abandonos,          color: P.rojo,    bg: P.rojoLight },
                { label: "Derivaciones",      val: kpis.derivaciones,       color: "#7B3FA0", bg: "#F3E8FF" },
                { label: "Tasa Abandono",     val: `${kpis.tasaAbandono}%`, color: P.rojo,    bg: P.rojoLight },
                { label: "% Respiratorio",    val: `${kpis.tasaResp}%`,     color: P.amber,   bg: "#FEF3C7" },
                { label: "T° Espera Prom.",   val: `${kpis.promEspera} min`,color: "#1A7A9A", bg: "#E0F4FA" },
              ].map(k => (
                <div key={k.label} style={{ background: k.bg, border: `1px solid ${P.border}`, borderLeft: `4px solid ${k.color}`, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 22, fontWeight: 800, color: k.color }}>{k.val}</div>
                  <div style={{ fontSize: 11, color: P.muted, marginTop: 3, fontWeight: 600 }}>{k.label}</div>
                </div>
              ))}
            </div>

            {/* Absorción de demanda por establecimiento */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, marginBottom: 14, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Absorción de la Demanda por Establecimiento</div>
                  <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Porcentaje de demanda que representa cada establecimiento sobre el total de la red filtrada</div>
                </div>
                <div style={{ background: P.azulLight, border: `1px solid ${P.azulMid}`, borderRadius: 8, padding: "8px 12px", textAlign: "right" }}>
                  <div style={{ fontSize: 10, color: P.muted, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.4px" }}>Total red</div>
                  <div style={{ fontSize: 18, color: P.azulDark, fontWeight: 800 }}>{kpis.demanda.toLocaleString("es-CL")}</div>
                </div>
              </div>
              {dataAbsorcionDemanda.length > 0 ? (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 18, alignItems: "stretch" }}>
                  <ResponsiveContainer width="100%" height={Math.max(260, dataAbsorcionDemanda.length * 34)}>
                    <BarChart
                      data={dataAbsorcionDemanda}
                      layout="vertical"
                      margin={{ top: 4, right: 28, left: 16, bottom: 4 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} horizontal={false} />
                      <XAxis type="number" domain={[0, 100]} tickFormatter={v => `${v}%`} tick={{ fontSize: 11, fill: P.muted }} />
                      <YAxis type="category" dataKey="establecimiento" width={142} tick={{ fontSize: 11, fill: P.text }} />
                      <Tooltip
                        contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }}
                        formatter={(value, name, props) => {
                          if (name === "Absorción") return [`${value}%`, name];
                          return [props.payload.demanda.toLocaleString("es-CL"), "Demanda"];
                        }}
                        labelFormatter={label => `Establecimiento: ${label}`}
                      />
                      <Bar dataKey="absorcion" name="Absorción" fill={P.azul} radius={[0, 6, 6, 0]} barSize={18} />
                    </BarChart>
                  </ResponsiveContainer>
                  <div style={{ border: `1px solid ${P.border}`, borderRadius: 10, overflow: "auto", alignSelf: "stretch", WebkitOverflowScrolling: "touch" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 320 }}>
                      <thead>
                        <tr style={{ background: P.azulDark }}>
                          {["Establecimiento", "Demanda", "% Red", "Atendidos"].map(h => (
                            <th key={h} style={{ padding: "10px 12px", textAlign: h === "Establecimiento" ? "left" : "right", color: "#fff", fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {dataAbsorcionDemanda.map((r, i) => (
                          <tr key={r.establecimiento} style={{ background: i % 2 === 0 ? "#fff" : P.bg }}>
                            <td style={{ ...tdS, whiteSpace: "normal", fontWeight: 700 }}>{r.establecimiento}</td>
                            <td style={{ ...tdS, textAlign: "right", color: P.azul, fontWeight: 800 }}>{r.demanda.toLocaleString("es-CL")}</td>
                            <td style={{ ...tdS, textAlign: "right" }}>
                              <span style={{ background: P.verdeLight, color: P.verde, borderRadius: 20, padding: "3px 9px", fontWeight: 800 }}>{r.absorcion.toFixed(1)}%</span>
                            </td>
                            <td style={{ ...tdS, textAlign: "right", color: P.verde, fontWeight: 700 }}>{r.atendidos.toLocaleString("es-CL")}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : (
                <div style={{ height: 120, display: "flex", alignItems: "center", justifyContent: "center", color: P.muted, fontSize: 13 }}>
                  Sin datos de demanda para calcular absorción con los filtros aplicados.
                </div>
              )}
            </div>

            {/* ── Atenciones totales por establecimiento ─────── */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Atenciones Totales por Establecimiento</div>
                <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Suma de pacientes atendidos según filtros aplicados</div>
              </div>
              {dataAbsorcionDemanda.length > 0 ? (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 18, alignItems: "stretch" }}>
                  <ResponsiveContainer width="100%" height={Math.max(260, dataAbsorcionDemanda.length * 34)}>
                    <BarChart
                      data={[...dataAbsorcionDemanda].sort((a,b) => b.atendidos - a.atendidos)}
                      layout="vertical"
                      margin={{ top: 4, right: 28, left: 16, bottom: 4 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} horizontal={false} />
                      <XAxis type="number" tick={{ fontSize: 11, fill: P.muted }} tickFormatter={v => v.toLocaleString("es-CL")} />
                      <YAxis type="category" dataKey="establecimiento" width={142} tick={{ fontSize: 11, fill: P.text }} />
                      <Tooltip
                        contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }}
                        formatter={(val) => [val.toLocaleString("es-CL"), "Atendidos"]}
                        labelFormatter={label => `Establecimiento: ${label}`}
                      />
                      <Bar dataKey="atendidos" name="Atendidos" fill={P.verde} radius={[0,6,6,0]} barSize={18} />
                    </BarChart>
                  </ResponsiveContainer>
                  <div style={{ border: `1px solid ${P.border}`, borderRadius: 10, overflow: "auto", alignSelf: "stretch", WebkitOverflowScrolling: "touch" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 320 }}>
                      <thead>
                        <tr style={{ background: P.azulDark }}>
                          {["Establecimiento", "Atendidos", "Respiratorias", "% Resp."].map(h => (
                            <th key={h} style={{ padding: "10px 12px", textAlign: h === "Establecimiento" ? "left" : "right", color: "#fff", fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {[...dataAbsorcionDemanda]
                          .sort((a,b) => b.atendidos - a.atendidos)
                          .map((r, i) => {
                            const regsEstab = filtrados.filter(reg => reg.establecimiento === r.establecimiento);
                            const totalResp = regsEstab.reduce((s, reg) => s + Number(reg.atenciones_respiratorias || 0), 0);
                            const pctResp   = r.atendidos > 0 ? ((totalResp / r.atendidos) * 100).toFixed(1) : "—";
                            return (
                              <tr key={r.establecimiento} style={{ background: i % 2 === 0 ? "#fff" : P.bg }}>
                                <td style={{ ...tdS, whiteSpace: "normal", fontWeight: 700 }}>{r.establecimiento}</td>
                                <td style={{ ...tdS, textAlign: "right", color: P.verde, fontWeight: 800 }}>{r.atendidos.toLocaleString("es-CL")}</td>
                                <td style={{ ...tdS, textAlign: "right", color: P.amber, fontWeight: 700 }}>{totalResp.toLocaleString("es-CL")}</td>
                                <td style={{ ...tdS, textAlign: "right" }}>
                                  <span style={{ background: "#fff8e1", color: P.amber, borderRadius: 20, padding: "3px 9px", fontWeight: 800 }}>{pctResp}%</span>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : (
                <div style={{ height: 120, display: "flex", alignItems: "center", justifyContent: "center", color: P.muted, fontSize: 13 }}>
                  Sin datos para mostrar con los filtros aplicados.
                </div>
              )}
            </div>

                        {/* ── ALERTA TIEMPO ESPERA ────────────────────── */}
            {Number(kpis.promEspera) > 180 && (
              <div style={{
                background: "#fdecea", border: "2px solid #C0392B", borderRadius: 8,
                padding: "14px 20px", marginBottom: 14,
                display: "flex", alignItems: "center", gap: 14,
              }}>
                <div style={{ fontSize: 28, flexShrink: 0 }}>🚨</div>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 800, color: "#C0392B" }}>
                    Alerta: Tiempo de Espera Promedio Superó el Umbral
                  </div>
                  <div style={{ fontSize: 12, color: "#922B21", marginTop: 3 }}>
                    El promedio de la red es <b>{kpis.promEspera} min</b>, superando el umbral de <b>180 min</b>. Se recomienda revisar la distribución de atenciones y refuerzos disponibles.
                  </div>
                </div>
              </div>
            )}

            {/* ── REFUERZOS POR ESTABLECIMIENTO ────────────────── */}
            {(() => {
              const ESTAMENTOS = [
                { key: "refuerzo_medico",         label: "Médico",         color: "#1A3A6B" },
                { key: "refuerzo_enfermera",       label: "Enfermera",      color: "#C0392B" },
                { key: "refuerzo_tens",            label: "TENS",           color: "#1A7A4A" },
                { key: "refuerzo_kinesiologo",     label: "Kinesiólogo",    color: "#B45309" },
                { key: "refuerzo_administrativo",  label: "Administrativo", color: "#6B3FA0" },
              ];
              const conRef = filtrados.filter(r => r.tiene_refuerzo);
              if (conRef.length === 0) return null;
              const porEstab = {};
              conRef.forEach(r => {
                const e = r.establecimiento;
                if (!porEstab[e]) porEstab[e] = { dias: 0, est: {} };
                porEstab[e].dias++;
                ESTAMENTOS.forEach(({ key, label }) => {
                  if (r[key]) porEstab[e].est[label] = (porEstab[e].est[label] || 0) + 1;
                });
              });
              return (
                <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
                  <div style={{ marginBottom: 12 }}>
                    <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>⚡ Refuerzos por Establecimiento</div>
                    <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Días con refuerzo y estamentos activos según filtros aplicados</div>
                  </div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 500 }}>
                      <thead>
                        <tr style={{ background: P.azulDark }}>
                          <th style={{ padding: "8px 12px", color: "#fff", fontWeight: 700, textAlign: "left" }}>Establecimiento</th>
                          <th style={{ padding: "8px 10px", color: "#fff", fontWeight: 700, textAlign: "center" }}>Días c/ Refuerzo</th>
                          {ESTAMENTOS.map(e => (
                            <th key={e.key} style={{ padding: "8px 10px", color: "#fff", fontWeight: 700, textAlign: "center", whiteSpace: "nowrap" }}>{e.label}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(porEstab).sort((a,b) => b[1].dias - a[1].dias).map(([estab, data], i) => (
                          <tr key={estab} style={{ background: i % 2 === 0 ? "#fff" : P.azulLight }}>
                            <td style={{ padding: "8px 12px", fontWeight: 700, color: P.azulDark, borderBottom: `1px solid ${P.border}` }}>{estab}</td>
                            <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 800, color: P.azul, borderBottom: `1px solid ${P.border}` }}>{data.dias}</td>
                            {ESTAMENTOS.map(({ label, color }) => {
                              const tiene = (data.est[label] || 0) > 0;
                              return (
                                <td key={label} style={{ padding: "8px 10px", textAlign: "center", borderBottom: `1px solid ${P.border}` }}>
                                  {tiene
                                    ? <span style={{ background: color, color: "#fff", borderRadius: 20, padding: "3px 12px", fontWeight: 700, fontSize: 13 }}>✓</span>
                                    : <span style={{ color: P.grisMid }}>—</span>
                                  }
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })()}

            {/* Gráfico comportamiento diario — ancho completo */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Comportamiento Diario de Variables</div>
                  <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Suma de atenciones por día según filtros aplicados</div>
                </div>
              </div>
              {dataXDia.length > 0 ? (
                <ResponsiveContainer width="100%" height={260}>
                  <LineChart data={dataXDia} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
                    <XAxis dataKey="dia" tick={{ fontSize: 10, fill: P.muted }} interval={dataXDia.length > 30 ? Math.floor(dataXDia.length / 8) : 0} angle={dataXDia.length > 20 ? -30 : 0} textAnchor={dataXDia.length > 20 ? "end" : "middle"} height={dataXDia.length > 20 ? 45 : 30} />
                    <YAxis tick={{ fontSize: 11, fill: P.muted }} width={45} />
                    <Tooltip contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 11 }} />
                    <Legend wrapperStyle={{ fontSize: 10 }} />
                    <Line type="monotone" dataKey="demanda"       name="Demanda Total"     stroke={P.azul}    strokeWidth={2} dot={dataXDia.length > 30 ? false : { r: 3 }} activeDot={{ r: 5 }} />
                    <Line type="monotone" dataKey="atendidos"     name="Atendidos"         stroke={P.verde}   strokeWidth={2} dot={dataXDia.length > 30 ? false : { r: 3 }} activeDot={{ r: 5 }} />
                    <Line type="monotone" dataKey="respiratorias" name="Respiratorias"     stroke={P.amber}   strokeWidth={1.5} dot={false} activeDot={{ r: 4 }} />
                    <Line type="monotone" dataKey="abandonos"     name="Abandonos"         stroke={P.rojo}    strokeWidth={1.5} dot={false} activeDot={{ r: 4 }} />
                    <Line type="monotone" dataKey="derivaciones"  name="Derivaciones"      stroke="#7B3FA0"   strokeWidth={1.5} dot={false} activeDot={{ r: 4 }} strokeDasharray="5 3" />
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div style={{ height: 120, display: "flex", alignItems: "center", justifyContent: "center", color: P.muted, fontSize: 13 }}>
                  Sin datos para mostrar. Ingresa registros para ver el comportamiento diario.
                </div>
              )}
            </div>

            {/* Gráfico: Atenciones diarias vs % Respiratorio */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Atenciones Diarias vs % Respiratorio</div>
                <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Barras: total atendidos por día · Línea: porcentaje de atenciones respiratorias</div>
              </div>
              {dataXDia.length > 0 ? (
                <ResponsiveContainer width="100%" height={260}>
                  <ComposedChart data={dataXDia} margin={{ top: 4, right: 32, left: 0, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
                    <XAxis dataKey="dia" tick={{ fontSize: 10, fill: P.muted }} interval={dataXDia.length > 30 ? Math.floor(dataXDia.length / 8) : 0} angle={dataXDia.length > 20 ? -30 : 0} textAnchor={dataXDia.length > 20 ? "end" : "middle"} height={dataXDia.length > 20 ? 45 : 30} />
                    <YAxis yAxisId="abs" orientation="left" tick={{ fontSize: 11, fill: P.muted }} label={{ value: "Atendidos", angle: -90, position: "insideLeft", fontSize: 10, fill: P.muted }} />
                    <YAxis yAxisId="pct" orientation="right" tickFormatter={v => `${v}%`} domain={[0, 100]} tick={{ fontSize: 11, fill: P.amber }} label={{ value: "%Resp.", angle: 90, position: "insideRight", fontSize: 10, fill: P.amber }} />
                    <Tooltip
                      contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }}
                      formatter={(value, name) => name === "% Respiratorio" ? `${value}%` : value}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar yAxisId="abs" dataKey="atendidos" name="Atendidos" fill={P.azul} radius={[4,4,0,0]} opacity={0.85} />
                    <Line
                      yAxisId="pct"
                      type="monotone"
                      dataKey={d => d.atendidos > 0 ? +((d.respiratorias / d.atendidos) * 100).toFixed(1) : 0}
                      name="% Respiratorio"
                      stroke={P.amber}
                      strokeWidth={2.5}
                      dot={{ r: 5, fill: P.amber, stroke: "#fff", strokeWidth: 2 }}
                      activeDot={{ r: 7 }}
                    />
                  </ComposedChart>
                </ResponsiveContainer>
              ) : (
                <div style={{ height: 100, display: "flex", alignItems: "center", justifyContent: "center", color: P.muted, fontSize: 13 }}>
                  Sin datos para mostrar con los filtros aplicados.
                </div>
              )}
            </div>

            {/* Gráfico: Atendidos vs Respiratorias por semana + acumulado */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Atenciones Totales v/s Respiratorias por Semana · Acumulado</div>
                <div style={{ fontSize: 11, color: P.muted, marginTop: 2 }}>Barras: valores semanales · Líneas: acumulado progresivo</div>
              </div>
              {dataRespAcum.length > 0 ? (
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={dataRespAcum} margin={{ top: 4, right: 24, left: 0, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
                    <XAxis dataKey="semana" tick={{ fontSize: 11, fill: P.muted }} />
                    <YAxis yAxisId="semanal" orientation="left" tick={{ fontSize: 11, fill: P.muted }} label={{ value: "Semanal", angle: -90, position: "insideLeft", fontSize: 10, fill: P.muted }} />
                    <YAxis yAxisId="acum" orientation="right" tick={{ fontSize: 11, fill: P.muted }} label={{ value: "Acumulado", angle: 90, position: "insideRight", fontSize: 10, fill: P.muted }} />
                    <Tooltip contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar yAxisId="semanal" dataKey="atendidos"     name="Atendidos (sem.)"     fill={P.azul}  radius={[4,4,0,0]} opacity={0.85} />
                    <Bar yAxisId="semanal" dataKey="respiratorias" name="Respiratorias (sem.)"  fill={P.amber} radius={[4,4,0,0]} opacity={0.85} />
                    <Line yAxisId="acum" type="monotone" dataKey="acumAtendidos"     name="Acum. Atendidos"     stroke={P.azulDark}  strokeWidth={2.5} dot={{ r: 4 }} strokeDasharray="6 2" />
                    <Line yAxisId="acum" type="monotone" dataKey="acumRespiratorias" name="Acum. Respiratorias"  stroke="#B45309"     strokeWidth={2.5} dot={{ r: 4 }} strokeDasharray="6 2" />
                  </BarChart>
                </ResponsiveContainer>
              ) : (
                <div style={{ height: 100, display: "flex", alignItems: "center", justifyContent: "center", color: P.muted, fontSize: 13 }}>
                  Sin datos para mostrar con los filtros aplicados.
                </div>
              )}
            </div>

            {/* Gráficos fila 2 */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 18 }}>
              <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18 }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14, color: P.azulDark }}>Demanda vs Atendidos por Semana Epidemiológica</div>
                {dataXSemana.length > 0 ? (
                  <ResponsiveContainer width="100%" height={210}>
                    <BarChart data={dataXSemana}>
                      <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
                      <XAxis dataKey="semana" tick={{ fontSize: 11, fill: P.muted }} />
                      <YAxis tick={{ fontSize: 11, fill: P.muted }} />
                      <Tooltip contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8 }} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="demanda" name="Demanda" fill={P.azul} radius={[4,4,0,0]} />
                      <Bar dataKey="atendidos" name="Atendidos" fill={P.verde} radius={[4,4,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                ) : <div style={{ color: P.muted, fontSize: 13, padding: 20 }}>Sin datos</div>}
              </div>
              <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18 }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4, color: P.azulDark }}>Destino de Derivaciones</div>
                <div style={{ fontSize: 11, color: P.muted, marginBottom: 14 }}>Total de derivaciones por hospital</div>
                {dataDerivaciones.length > 0 ? (() => {
                  const total = dataDerivaciones.reduce((a, d) => a + d.value, 0);
                  const COLORS = ["#1A3A6B", "#C0392B", "#1A7A4A", "#B45309"];
                  return (
                    <div>
                      {/* Donut chart */}
                      <ResponsiveContainer width="100%" height={180}>
                        <PieChart>
                          <Pie
                            data={dataDerivaciones}
                            cx="50%" cy="50%"
                            innerRadius={48} outerRadius={78}
                            dataKey="value"
                            paddingAngle={2}
                            label={({ cx, cy, midAngle, innerRadius, outerRadius, percent }) => {
                              const RADIAN = Math.PI / 180;
                              const radius = innerRadius + (outerRadius - innerRadius) * 1.35;
                              const x = cx + radius * Math.cos(-midAngle * RADIAN);
                              const y = cy + radius * Math.sin(-midAngle * RADIAN);
                              return percent > 0.05 ? (
                                <text x={x} y={y} fill={P.text} textAnchor={x > cx ? "start" : "end"} dominantBaseline="central" fontSize={11} fontWeight={700}>
                                  {`${(percent*100).toFixed(0)}%`}
                                </text>
                              ) : null;
                            }}
                            labelLine={false}
                          >
                            {dataDerivaciones.map((_, i) => (
                              <Cell key={i} fill={COLORS[i % COLORS.length]} stroke="#fff" strokeWidth={2} />
                            ))}
                          </Pie>
                          <Tooltip
                            contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }}
                            formatter={(val, name) => [`${val} (${((val/total)*100).toFixed(1)}%)`, name]}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                      {/* Leyenda manual con barras */}
                      <div style={{ display: "flex", flexDirection: "column", gap: 7, marginTop: 4 }}>
                        {dataDerivaciones.map((d, i) => {
                          const pct = total > 0 ? ((d.value / total) * 100) : 0;
                          return (
                            <div key={d.name}>
                              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                                  <div style={{ width: 10, height: 10, borderRadius: 2, background: COLORS[i % COLORS.length], flexShrink: 0 }} />
                                  <span style={{ fontSize: 11, fontWeight: 600, color: P.text }}>{d.name}</span>
                                </div>
                                <span style={{ fontSize: 11, fontWeight: 700, color: COLORS[i % COLORS.length] }}>{d.value} · {pct.toFixed(1)}%</span>
                              </div>
                              <div style={{ height: 5, background: P.grisMid, borderRadius: 3, overflow: "hidden" }}>
                                <div style={{ height: "100%", width: `${pct}%`, background: COLORS[i % COLORS.length], borderRadius: 3, transition: "width 0.4s" }} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })() : <div style={{ color: P.muted, fontSize: 13, padding: 20 }}>Sin datos con filtros seleccionados</div>}
              </div>
            </div>


          </div>
        )}

        {tab === "dashboard" && <>
          <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: "16px 20px", marginTop: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: mostrarComparador ? 20 : 0 }}>
              <div>
                <div style={{ fontSize: 15, fontWeight: 800, color: P.azulDark }}>📊 Comparar Períodos</div>
                <div style={{ fontSize: 12, color: P.muted }}>Compara métricas entre dos rangos de semanas epidemiológicas</div>
              </div>
              <button onClick={() => setMostrarComparador(v => !v)} style={{
                background: mostrarComparador ? P.azulLight : P.azul, color: mostrarComparador ? P.azul : "#fff",
                border: `1px solid ${P.azul}`, borderRadius: 8, padding: "7px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer"
              }}>{mostrarComparador ? "Ocultar" : "Abrir comparador"}</button>
            </div>
            {mostrarComparador && <ComparadorPeriodos
              semanasOpts={semanasOpts}
              registros={registros}
              filtroPolo={filtroPolo} filtroEstab={filtroEstab}
              p1desde={cmpP1desde} setP1desde={setCmpP1desde}
              p1hasta={cmpP1hasta} setP1hasta={setCmpP1hasta}
              p2desde={cmpP2desde} setP2desde={setCmpP2desde}
              p2hasta={cmpP2hasta} setP2hasta={setCmpP2hasta}
              calcMetricasRango={calcMetricasRango}
              inpS={inpS} P={P}
            />}
          </div>
          <ResumenAmbulancias
            registrosAmbulancias={registrosAmbulancias}
            filtroEstab={filtroEstab} filtroPolo={filtroPolo}
            POLO_MAP={POLO_MAP} P={P}
          />
        </>
        }

        {/* ── FORMULARIO ──────────────────────────────────────── */}
        {tab === "formulario" && (
          <div style={{ maxWidth: 740 }}>
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 800, color: P.azulDark }}>{editId ? "✏️ Editar Registro" : "➕ Nuevo Registro"}</div>
              <div style={{ fontSize: 12, color: P.muted, marginTop: 4 }}>Los campos marcados con * son obligatorios</div>
            </div>
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 26 }}>
              
              {/* Identificación */}
              <div style={{ fontSize: 12, fontWeight: 800, color: P.azul, borderBottom: `2px solid ${P.azulMid}`, paddingBottom: 6, marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Identificación del Registro
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
                <Fld label="Fecha *" name="fecha" type="date" value={form.fecha} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="Semana Epidemiológica" name="semana_epi" value={form.semana_epi} onChange={handleChange} disabled ls={labelS} is={inpS} />
              </div>
              <div style={{ marginBottom: 20 }}>
                <label style={labelS}>Establecimiento *</label>
                <select name="establecimiento" value={form.establecimiento} onChange={handleChange} style={inpS}>
                  <option value="">Seleccionar establecimiento...</option>
                  {ESTABLECIMIENTOS_GROUPED.map(g => (
                    <optgroup key={g.comuna} label={g.comuna}>
                      {g.items.map(e => <option key={e} value={e}>{e}</option>)}
                    </optgroup>
                  ))}
                </select>
              </div>

              {/* Producción */}
              <div style={{ fontSize: 12, fontWeight: 800, color: P.azul, borderBottom: `2px solid ${P.azulMid}`, paddingBottom: 6, marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Producción del Turno
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 20 }}>
                <Fld label="Demanda Total *" name="demanda_total" type="number" value={form.demanda_total} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="Pacientes Atendidos" name="pacientes_atendidos" type="number" value={form.pacientes_atendidos} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="At. Respiratorias" name="atenciones_respiratorias" type="number" value={form.atenciones_respiratorias} onChange={handleChange} ls={labelS} is={inpS} />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
                <Fld label="Abandonos" name="abandonos" type="number" value={form.abandonos} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="Tiempo Espera Prom. (min)" name="tiempo_espera" type="number" value={form.tiempo_espera} onChange={handleChange} ls={labelS} is={inpS} />
              </div>

              {/* Derivaciones por hospital */}
              <div style={{ fontSize: 11, fontWeight: 800, color: P.azul, borderBottom: `1px solid ${P.azulMid}`, paddingBottom: 5, marginBottom: 14, textTransform: "uppercase", letterSpacing: "0.4px" }}>
                Derivaciones a Hospital
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 20 }}>
                <Fld label="HEC" name="derivaciones_hec" type="number" value={form.derivaciones_hec} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="HCSBA" name="derivaciones_hcsba" type="number" value={form.derivaciones_hcsba} onChange={handleChange} ls={labelS} is={inpS} />
                <Fld label="HUAP" name="derivaciones_huap" type="number" value={form.derivaciones_huap} onChange={handleChange} ls={labelS} is={inpS} />
              </div>

              {/* Refuerzo */}
              <div style={{ fontSize: 12, fontWeight: 800, color: P.azul, borderBottom: `2px solid ${P.azulMid}`, paddingBottom: 6, marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Refuerzo de Recursos Humanos
              </div>
              <div style={{ background: P.azulLight, border: `1px solid ${P.azulMid}`, borderRadius: 9, padding: 16, marginBottom: 20 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                  <input type="checkbox" name="tiene_refuerzo" id="refuerzo" checked={form.tiene_refuerzo} onChange={handleChange}
                    style={{ width: 16, height: 16, accentColor: P.verde }} />
                  <label htmlFor="refuerzo" style={{ fontSize: 13, fontWeight: 700, cursor: "pointer", color: P.azulDark }}>
                    ⚡ Este turno contó con refuerzo de RRHH
                  </label>
                </div>
                {form.tiene_refuerzo && (
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: P.azulDark, marginBottom: 10 }}>
                      Selecciona los profesionales con refuerzo:
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
                      {[
                        { key: "refuerzo_medico",        label: "👨‍⚕️ Médico" },
                        { key: "refuerzo_enfermera",     label: "👩‍⚕️ Enfermera" },
                        { key: "refuerzo_tens",          label: "🩺 TENS" },
                        { key: "refuerzo_kinesiologo",   label: "🏃 Kinesiólogo" },
                        { key: "refuerzo_administrativo",label: "💼 Administrativo" },
                      ].map(({ key, label }) => (
                        <div key={key} style={{
                          display: "flex", alignItems: "center", justifyContent: "space-between",
                          background: form[key] ? "#d4edda" : "#fff",
                          border: `1.5px solid ${form[key] ? P.verde : P.azulMid}`,
                          borderRadius: 8, padding: "8px 14px", transition: "all 0.15s"
                        }}>
                          <label style={{ fontSize: 13, fontWeight: 600, color: P.azulDark, cursor: "pointer", flex: 1 }}
                            htmlFor={key}>{label}</label>
                          <div style={{ display: "flex", gap: 8 }}>
                            <label style={{ fontSize: 12, fontWeight: 700, color: form[key] ? P.verde : P.grisMid, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
                              <input type="radio" name={key} id={key} checked={form[key] === true}
                                onChange={() => setForm(f => ({ ...f, [key]: true }))}
                                style={{ accentColor: P.verde }} /> Sí
                            </label>
                            <label style={{ fontSize: 12, fontWeight: 700, color: !form[key] ? "#c0392b" : P.grisMid, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
                              <input type="radio" name={key} checked={form[key] === false}
                                onChange={() => setForm(f => ({ ...f, [key]: false }))}
                                style={{ accentColor: "#c0392b" }} /> No
                            </label>
                          </div>
                        </div>
                      ))}
                    </div>
                    <Fld label="Horas de Refuerzo" name="horas_refuerzo" type="number" value={form.horas_refuerzo} onChange={handleChange} ls={labelS} is={inpS} />
                  </div>
                )}
              </div>

              {/* Observaciones */}
              <div style={{ marginBottom: 22 }}>
                <label style={labelS}>Observaciones</label>
                <textarea name="observaciones" value={form.observaciones} onChange={handleChange} rows={3}
                  style={{ ...inpS, resize: "vertical", fontFamily: "inherit" }}
                  placeholder="Notas adicionales del turno..." />
              </div>

              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={handleSubmit} style={{
                  background: P.azul, color: "#fff", border: "none", borderRadius: 8,
                  padding: "10px 28px", cursor: "pointer", fontSize: 14, fontWeight: 700
                }}>
                  {editId ? "💾 Actualizar Registro" : "✅ Guardar Registro"}
                </button>
                {editId && (
                  <button onClick={() => { setForm(EMPTY_FORM); setEditId(null); }} style={{
                    background: P.grisMid, color: P.text, border: "none",
                    borderRadius: 8, padding: "10px 18px", cursor: "pointer", fontSize: 14
                  }}>Cancelar</button>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ── TABLA ───────────────────────────────────────────── */}
        {tab === "tabla" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 12 }}>
              <div style={{ fontSize: 16, fontWeight: 800, color: P.azulDark }}>Registros ({filtrados.length})</div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <select value={filtroPolo} onChange={e => setFiltroPolo(e.target.value)} style={{ ...inpS, width: "auto" }}>
                  {POLOS.map(p => <option key={p}>{p}</option>)}
                </select>
                <select value={filtroSemana} onChange={e => setFiltroSemana(e.target.value)} style={{ ...inpS, width: "auto" }}>
                  {semanas.map(s => <option key={s}>{s}</option>)}
                </select>
                <select value={filtroEstab} onChange={e => setFiltroEstab(e.target.value)} style={{ ...inpS, width: "auto" }}>
                  {estabs.map(c => <option key={c}>{c}</option>)}
                </select>
                <button onClick={() => { setTab("formulario"); setForm(EMPTY_FORM); setEditId(null); }} style={{
                  background: P.azul, color: "#fff", border: "none",
                  borderRadius: 7, padding: "8px 18px", cursor: "pointer", fontSize: 13, fontWeight: 700
                }}>+ Nuevo</button>
              </div>
            </div>
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: P.azul }}>
                      {["Fecha","SE","Establecimiento","Demanda","Atendidos","Resp.","Abandonos","HEC","HCSBA","HUAP","T° Espera","Refuerzo","Acciones"].map(h => (
                        <th key={h} style={{ padding: "11px 13px", textAlign: "left", color: "#fff", fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtrados.length === 0 && (
                      <tr><td colSpan={12} style={{ padding: 48, textAlign: "center", color: P.muted }}>No hay registros. ¡Ingresa el primero!</td></tr>
                    )}
                    {filtrados.map((r, i) => (
                      <tr key={r.id} style={{ background: i % 2 === 0 ? "#fff" : P.bg }}>
                        <td style={tdS}>{r.fecha}</td>
                        <td style={tdS}>
                          <span style={{ background: P.azulLight, color: P.azul, padding: "2px 9px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>{r.semana_epi}</span>
                        </td>
                        <td style={{ ...tdS, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>{r.establecimiento}</td>
                        <td style={{ ...tdS, color: P.azul, fontWeight: 700 }}>{r.demanda_total}</td>
                        <td style={{ ...tdS, color: P.verde, fontWeight: 700 }}>{r.pacientes_atendidos}</td>
                        <td style={{ ...tdS, color: P.amber }}>{r.atenciones_respiratorias}</td>
                        <td style={tdS}>{r.abandonos}</td>
                        <td style={{ ...tdS, color: "#7B3FA0", fontWeight: 600 }}>{r.derivaciones_hec || "—"}</td>
                        <td style={{ ...tdS, color: "#7B3FA0", fontWeight: 600 }}>{r.derivaciones_hcsba || "—"}</td>
                        <td style={{ ...tdS, color: "#7B3FA0", fontWeight: 600 }}>{r.derivaciones_huap || "—"}</td>
                        <td style={{ ...tdS, color: "#1A7A9A", fontWeight: 600 }}>{r.tiempo_espera ? `${r.tiempo_espera} min` : "—"}</td>
                        <td style={tdS}>
                          {r.tiene_refuerzo
                            ? <span style={{ background: P.verdeLight, color: P.verde, padding: "2px 9px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>⚡ {r.tipo_refuerzo}</span>
                            : <span style={{ color: P.grisMid }}>—</span>}
                        </td>
                        <td style={tdS}>
                          <div style={{ display: "flex", gap: 5 }}>
                            <button onClick={() => handleEdit(r)} style={{ background: P.azulLight, color: P.azul, border: "none", borderRadius: 5, padding: "4px 10px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>✏️</button>
                            {deleteConfirm === r.id
                              ? <span style={{ display: "flex", gap: 4 }}>
                                  <button onClick={() => handleDelete(r.id)} style={{ background: P.rojo, color: "#fff", border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>Confirmar</button>
                                  <button onClick={() => setDeleteConfirm(null)} style={{ background: P.grisMid, color: P.text, border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", fontSize: 11 }}>✕</button>
                                </span>
                              : <button onClick={() => setDeleteConfirm(r.id)} style={{ background: P.rojoLight, color: P.rojo, border: "none", borderRadius: 5, padding: "4px 10px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>🗑</button>
                            }
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ── RETENCIONES DE AMBULANCIAS ──────────────────────── */}
        {tab === "ambulancias" && (
          <div>
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 17, fontWeight: 800, color: P.azulDark }}>🚑 Retenciones de Ambulancias</div>
              <div style={{ fontSize: 12, color: P.muted, marginTop: 4 }}>
                Registra el tiempo que una ambulancia permanece retenida en el establecimiento al momento del traslado. Completa hasta {AMBULANCIA_FILAS} líneas y guarda en un solo paso.
              </div>
            </div>

            {/* Formulario rápido de 5 líneas */}
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: 22, marginBottom: 24 }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 680 }}>
                  <thead>
                    <tr>
                      {["#", "Fecha", "Establecimiento", "Horario Traslado", "T° Retención (min)"].map(h => (
                        <th key={h} style={{ textAlign: "left", padding: "0 8px 8px", fontSize: 11, color: P.muted, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.4px" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {formAmbulancias.map((row, idx) => (
                      <tr key={idx}>
                        <td style={{ padding: 6, color: P.muted, fontWeight: 700, fontSize: 12 }}>{idx + 1}</td>
                        <td style={{ padding: 6 }}>
                          <input type="date" value={row.fecha}
                            onChange={e => handleChangeAmbulancia(idx, "fecha", e.target.value)}
                            style={inpS} />
                        </td>
                        <td style={{ padding: 6, minWidth: 220 }}>
                          <select value={row.establecimiento}
                            onChange={e => handleChangeAmbulancia(idx, "establecimiento", e.target.value)}
                            style={inpS}>
                            <option value="">Seleccionar...</option>
                            {ESTABLECIMIENTOS_GROUPED.map(g => (
                              <optgroup key={g.comuna} label={g.comuna}>
                                {g.items.map(e => <option key={e} value={e}>{e}</option>)}
                              </optgroup>
                            ))}
                          </select>
                        </td>
                        <td style={{ padding: 6 }}>
                          <input type="time" value={row.hora_traslado}
                            onChange={e => handleChangeAmbulancia(idx, "hora_traslado", e.target.value)}
                            style={inpS} />
                        </td>
                        <td style={{ padding: 6 }}>
                          <input type="number" min="0" placeholder="min" value={row.tiempo_retencion}
                            onChange={e => handleChangeAmbulancia(idx, "tiempo_retencion", e.target.value)}
                            style={inpS} />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ display: "flex", gap: 10, marginTop: 18 }}>
                <button onClick={handleSubmitAmbulancias} style={{
                  background: P.azul, color: "#fff", border: "none", borderRadius: 8,
                  padding: "10px 28px", cursor: "pointer", fontSize: 14, fontWeight: 700
                }}>✅ Guardar Líneas</button>
                <button onClick={() => setFormAmbulancias(Array.from({ length: AMBULANCIA_FILAS }, () => ({ ...EMPTY_AMBULANCIA_ROW })))} style={{
                  background: P.grisMid, color: P.text, border: "none",
                  borderRadius: 8, padding: "10px 18px", cursor: "pointer", fontSize: 14
                }}>Limpiar</button>
              </div>
            </div>

            {/* KPI Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 22 }}>
              {[
                { label: "Total Retenciones", val: kpisAmbulancias.total, color: P.azul, bg: P.azulLight },
                { label: "T° Retención Prom.", val: `${kpisAmbulancias.promedio} min`, color: "#1A7A9A", bg: "#E0F4FA" },
                { label: "T° Retención Total", val: `${kpisAmbulancias.totalMin} min`, color: P.amber, bg: "#FEF3C7" },
                { label: "Establecimiento c/ más retenciones", val: kpisAmbulancias.establecimientoTop, color: P.rojo, bg: P.rojoLight },
              ].map(k => (
                <div key={k.label} style={{ background: k.bg, border: `1px solid ${P.border}`, borderLeft: `4px solid ${k.color}`, borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: k.label.includes("Establecimiento") ? 14 : 22, fontWeight: 800, color: k.color }}>{k.val}</div>
                  <div style={{ fontSize: 11, color: P.muted, marginTop: 3, fontWeight: 600 }}>{k.label}</div>
                </div>
              ))}
            </div>

            {/* Gráfico por establecimiento */}
            {dataAmbulanciasXEstablecimiento.length > 0 && (
              <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 22 }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark, marginBottom: 14 }}>Retenciones por Establecimiento</div>
                <ResponsiveContainer width="100%" height={Math.max(220, dataAmbulanciasXEstablecimiento.length * 36)}>
                  <BarChart data={dataAmbulanciasXEstablecimiento} layout="vertical" margin={{ top: 4, right: 28, left: 16, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} horizontal={false} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: P.muted }} />
                    <YAxis type="category" dataKey="establecimiento" width={150} tick={{ fontSize: 11, fill: P.text }} />
                    <Tooltip
                      contentStyle={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, fontSize: 12 }}
                      formatter={(value, name, props) => name === "retenciones"
                        ? [value, "N° Retenciones"]
                        : [`${props.payload.promedioMin} min`, "Promedio"]}
                    />
                    <Bar dataKey="retenciones" name="retenciones" fill={P.azul} radius={[0, 6, 6, 0]} barSize={18} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Tabla de registros */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
              <div style={{ fontSize: 16, fontWeight: 800, color: P.azulDark }}>Registros de Retención ({registrosAmbulancias.length})</div>
            </div>
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: P.azul }}>
                      {["Fecha", "Establecimiento", "Horario Traslado", "T° Retención", "Acciones"].map(h => (
                        <th key={h} style={{ padding: "11px 13px", textAlign: "left", color: "#fff", fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {registrosAmbulanciasOrdenados.length === 0 && (
                      <tr><td colSpan={5} style={{ padding: 48, textAlign: "center", color: P.muted }}>No hay registros de retención. ¡Ingresa el primero arriba!</td></tr>
                    )}
                    {registrosAmbulanciasOrdenados.map((r, i) => (
                      <tr key={r.id} style={{ background: i % 2 === 0 ? "#fff" : P.bg }}>
                        <td style={tdS}>{r.fecha}</td>
                        <td style={{ ...tdS, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis" }}>{r.establecimiento}</td>
                        <td style={tdS}>{r.hora_traslado}</td>
                        <td style={{ ...tdS, color: "#1A7A9A", fontWeight: 700 }}>{r.tiempo_retencion} min</td>
                        <td style={tdS}>
                          {deleteConfirmAmb === r.id
                            ? <span style={{ display: "flex", gap: 4 }}>
                                <button onClick={() => handleDeleteAmbulancia(r.id)} style={{ background: P.rojo, color: "#fff", border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>Confirmar</button>
                                <button onClick={() => setDeleteConfirmAmb(null)} style={{ background: P.grisMid, color: P.text, border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", fontSize: 11 }}>✕</button>
                              </span>
                            : <button onClick={() => setDeleteConfirmAmb(r.id)} style={{ background: P.rojoLight, color: P.rojo, border: "none", borderRadius: 5, padding: "4px 10px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>🗑</button>
                          }
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
        {/* ── IMPORTAR EXCEL ──────────────────────────────── */}
        {tab === "importar" && (
          <div style={{ maxWidth: 600, margin: "0 auto" }}>
            <div style={{ fontSize: 20, fontWeight: 800, color: P.azulDark, marginBottom: 6 }}>📥 Importar desde Excel</div>
            <p style={{ color: P.muted, fontSize: 13, marginBottom: 24 }}>
              Sube el archivo Excel con el formato estándar de atenciones. Los datos se cargan automáticamente a Supabase.
            </p>

            {/* Botón de carga explícito — funciona en móvil */}
            <div style={{ textAlign: "center", marginBottom: 20 }}>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleImportExcel}
                disabled={importando}
                style={{ display: "none" }}
                id="import-excel-file"
              />
              <label htmlFor="import-excel-file" style={{
                display: "inline-block",
                background: importando ? P.grisMid : P.azul,
                color: "#fff",
                padding: "14px 32px",
                borderRadius: 10,
                fontSize: 15,
                fontWeight: 700,
                cursor: importando ? "not-allowed" : "pointer",
                boxShadow: "0 2px 8px rgba(0,0,0,0.15)",
              }}>
                {importando ? "⏳ Procesando..." : "📂 Seleccionar archivo Excel"}
              </label>
              <div style={{ fontSize: 12, color: P.muted, marginTop: 10 }}>
                Formato .xlsx — mismo formato del reporte semanal SSMC
              </div>
            </div>

            {importando && (
              <div style={{ padding: 16, background: P.azulLight, borderRadius: 10, textAlign: "center", color: P.azul, fontWeight: 700, marginBottom: 16 }}>
                ⏳ Importando datos... por favor espera
              </div>
            )}

            {importResultado && (
              <div style={{
                padding: 16, borderRadius: 10, fontWeight: 700, fontSize: 14, marginBottom: 16,
                background: importResultado.ok ? "#d4edda" : "#fde8e8",
                color: importResultado.ok ? "#155724" : "#721c24",
                border: `1px solid ${importResultado.ok ? "#c3e6cb" : "#f5c6cb"}`
              }}>
                {importResultado.msg}
                {importResultado.ok && (
                  <div style={{ marginTop: 8, fontSize: 12, fontWeight: 400 }}>
                    Los datos ya aparecen en el Dashboard y la Tabla de Datos.
                  </div>
                )}
              </div>
            )}

            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark, marginBottom: 12 }}>📋 Formato esperado del Excel</div>
              {[
                "Fila 1: Semanas epidemiológicas (SE 01, SE 02…)",
                "Fila 2: Fechas de cada día (01/01/2026, 02/01/2026…)",
                "Fila 3: Encabezados (ESTABLECIMIENTO, TOTAL DEMANDA…)",
                "Filas siguientes: Un establecimiento por fila con sus datos diarios",
              ].map((item, i) => (
                <div key={i} style={{ display: "flex", gap: 10, marginBottom: 8, fontSize: 12, color: P.text }}>
                  <span style={{ color: P.azul, fontWeight: 700, flexShrink: 0 }}>{i+1}.</span>
                  <span>{item}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── TIEMPOS DE ESPERA ───────────────────────────────── */}
        {tab === "tiempos" && (
          <TabTiemposEspera
            registros={registros}
            filtroPolo={filtroPolo}
            filtroEstab={filtroEstab}
            filtroSemana={filtroSemana}
            P={P} inpS={inpS}
          />
        )}

        {/* ── PROYECCIONES ─────────────────────────────────────── */}
        {tab === "proyecciones" && (
          <Proyecciones
            registros={registros}
            filtroPolo={filtroPolo}
            filtroEstab={filtroEstab}
            metodo={proyMetodo} setMetodo={setProyMetodo}
            semanas={proySemanas} setSemanas={setProyeSemanas}
            semanasOpts={semanasOpts}
            P={P} inpS={inpS}
          />
        )}

      <style>{`
  @keyframes fadeIn { from { opacity:0; transform:translateY(-5px); } to { opacity:1; transform:translateY(0); } }
  * { box-sizing: border-box; }
  body { margin: 0; }
  input, select, textarea { font-family: inherit; }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: #F4F6FA; }
  ::-webkit-scrollbar-thumb { background: #C2D0E4; border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: #1A3A6B; }
  @media (max-width: 640px) {
    .nav-desktop { display: none !important; }
    .nav-mobile  { display: block !important; }
  }
  @media (min-width: 641px) {
    .nav-desktop { display: flex !important; }
    .nav-mobile  { display: none !important; }
  }
  @media print {
    body * { visibility: hidden; }
    #pdf-preview, #pdf-preview * { visibility: visible; }
    #pdf-preview { position: absolute; top: 0; left: 0; width: 100%; padding: 20px; }
  }
`}</style>
      {mostrarPDF && <ModalPDF
        registros={registros}
        registrosAmbulancias={registrosAmbulancias}
        onClose={() => setMostrarPDF(false)}
        P={P} getPolo={getPolo}
      />}
    </div>
  );
}

function Fld({ label, name, type="text", value, onChange, disabled, ls, is }) {
  return (
    <div>
      <label style={ls}>{label}</label>
      <input type={type} name={name} value={value} onChange={onChange} disabled={disabled}
        style={{ ...is, opacity: disabled ? 0.6 : 1 }} min={type === "number" ? "0" : undefined} />
    </div>
  );
}

// ── Comparador de períodos ────────────────────────────────────────────────────
function ComparadorPeriodos({ semanasOpts, registros, filtroPolo, filtroEstab, p1desde, setP1desde, p1hasta, setP1hasta, p2desde, setP2desde, p2hasta, setP2hasta, calcMetricasRango, inpS, P }) {
  const [metrica, setMetrica] = React.useState("demanda");
  const [tipoGrafico, setTipoGrafico] = React.useState("barras");

  const m1 = calcMetricasRango(p1desde, p1hasta);
  const m2 = calcMetricasRango(p2desde, p2hasta);
  const selStyle = { ...inpS, width: 90, padding: "6px 8px", fontSize: 13 };

  const METRICAS = [
    { k: "demanda",    label: "Demanda Total",            field: "demanda_total",            agg: "sum" },
    { k: "resp",       label: "Atenciones Respiratorias", field: "atenciones_respiratorias", agg: "sum" },
    { k: "abandonos",  label: "Abandonos",                field: "abandonos",                agg: "sum" },
    { k: "espera",     label: "T° Espera Prom. (min)",    field: "tiempo_espera",            agg: "avg" },
  ];

  // Calcular datos por SE para los gráficos
  const calcDataGrafico = (desde, hasta, label, color) => {
    if (!desde || !hasta) return [];
    const rango = semanasOpts.filter(se => se >= desde && se <= hasta);
    const met = METRICAS.find(m => m.k === metrica);
    return rango.map((se, i) => {
      const regs = registros.filter(r =>
        r.semana_epi === se &&
        (filtroPolo === "Todos" || getPolo(r.establecimiento) === filtroPolo) &&
        (filtroEstab === "Todos" || r.establecimiento === filtroEstab)
      );
      let val = 0;
      if (met.agg === "sum") {
        val = regs.reduce((a, r) => a + Number(r[met.field] || 0), 0);
      } else {
        const valid = regs.filter(r => r[met.field] != null && r[met.field] !== "");
        val = valid.length ? Math.round(valid.reduce((a, r) => a + Number(r[met.field]), 0) / valid.length) : 0;
      }
      return { se: `SE ${String(i+1).padStart(2,"0")}`, seReal: se, [label]: val };
    });
  };

  // Merge data for combined chart (align by position)
  const data1 = calcDataGrafico(p1desde, p1hasta, "P1", "#2980b9");
  const data2 = calcDataGrafico(p2desde, p2hasta, "P2", "#27ae60");
  const maxLen = Math.max(data1.length, data2.length);
  const dataGrafico = Array.from({ length: maxLen }, (_, i) => ({
    idx: `Sem ${i+1}`,
    seP1: data1[i]?.seReal || "",
    seP2: data2[i]?.seReal || "",
    P1: data1[i]?.P1 ?? null,
    P2: data2[i]?.P2 ?? null,
  }));

  const diff = (v1, v2, invert=false) => {
    if (v1==null || v2==null || Number(v2)===0) return null;
    const pct = (((Number(v1)-Number(v2))/Math.abs(Number(v2)))*100).toFixed(1);
    const up = Number(pct) > 0;
    const good = invert ? !up : up;
    return <span style={{ fontSize: 11, fontWeight: 700, color: good ? "#27ae60" : "#e74c3c", marginLeft: 6 }}>
      {up ? "▲" : "▼"}{Math.abs(pct)}%
    </span>;
  };

  const Row = ({ label, k, invert=false }) => (
    <tr>
      <td style={{ padding: "8px 12px", fontSize: 13, color: P.muted, borderBottom: `1px solid ${P.border}` }}>{label}</td>
      <td style={{ padding: "8px 12px", fontSize: 13, fontWeight: 700, color: "#2980b9", borderBottom: `1px solid ${P.border}`, textAlign: "right" }}>
        {m1 ? (m1[k] ?? "—") : "—"}
      </td>
      <td style={{ padding: "8px 12px", fontSize: 13, fontWeight: 700, color: "#27ae60", borderBottom: `1px solid ${P.border}`, textAlign: "right" }}>
        {m2 ? (m2[k] ?? "—") : "—"}
        {m1 && m2 ? diff(m2[k], m1[k], invert) : null}
      </td>
    </tr>
  );

  const hayDatos = data1.length > 0 || data2.length > 0;
  const metActual = METRICAS.find(m => m.k === metrica);

  return (
    <div>
      {/* Selectores de período */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
        {[
          { label: "📘 Período 1", desde: p1desde, setDesde: setP1desde, hasta: p1hasta, setHasta: setP1hasta, color: "#2980b9" },
          { label: "📗 Período 2", desde: p2desde, setDesde: setP2desde, hasta: p2hasta, setHasta: setP2hasta, color: "#27ae60" },
        ].map(({ label, desde, setDesde, hasta, setHasta, color }) => (
          <div key={label} style={{ background: P.azulLight, borderRadius: 10, padding: "14px 16px", border: `2px solid ${color}30` }}>
            <div style={{ fontSize: 13, fontWeight: 800, color, marginBottom: 10 }}>{label}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, color: P.muted }}>Desde</span>
              <select value={desde} onChange={e => setDesde(e.target.value)} style={selStyle}>
                <option value="">SE...</option>
                {semanasOpts.map(se => <option key={se}>{se}</option>)}
              </select>
              <span style={{ fontSize: 12, color: P.muted }}>Hasta</span>
              <select value={hasta} onChange={e => setHasta(e.target.value)} style={selStyle}>
                <option value="">SE...</option>
                {semanasOpts.map(se => <option key={se}>{se}</option>)}
              </select>
            </div>
            {desde && hasta && <div style={{ fontSize: 11, color: P.muted, marginTop: 6 }}>
              {semanasOpts.filter(se => se >= desde && se <= hasta).length} semanas seleccionadas
            </div>}
          </div>
        ))}
      </div>

      {hayDatos ? (<>
        {/* Controles de gráfico */}
        <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: P.muted }}>Métrica:</div>
          {METRICAS.map(m => (
            <button key={m.k} onClick={() => setMetrica(m.k)} style={{
              padding: "5px 12px", borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: "pointer",
              background: metrica === m.k ? P.azul : P.azulLight,
              color: metrica === m.k ? "#fff" : P.azulDark,
              border: `1px solid ${metrica === m.k ? P.azul : P.border}`,
            }}>{m.label}</button>
          ))}
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            {[{ v: "barras", label: "▤ Barras" }, { v: "lineas", label: "〜 Líneas" }].map(t => (
              <button key={t.v} onClick={() => setTipoGrafico(t.v)} style={{
                padding: "5px 12px", borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: "pointer",
                background: tipoGrafico === t.v ? P.azulDark : P.azulLight,
                color: tipoGrafico === t.v ? "#fff" : P.azulDark,
                border: `1px solid ${P.border}`,
              }}>{t.label}</button>
            ))}
          </div>
        </div>

        {/* Gráfico */}
        <div style={{ background: P.azulLight, borderRadius: 12, padding: "16px 8px", marginBottom: 20 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: P.azulDark, marginBottom: 12, paddingLeft: 8 }}>
            {metActual?.label} — comparación semana a semana
          </div>
          <ResponsiveContainer width="100%" height={280}>
            {tipoGrafico === "barras" ? (
              <BarChart data={dataGrafico} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e0e8f0" />
                <XAxis dataKey="idx" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip
                  formatter={(val, name) => [val ?? "—", name === "P1" ? `P1 (${p1desde}→${p1hasta})` : `P2 (${p2desde}→${p2hasta})`]}
                  labelFormatter={(label, payload) => {
                    if (!payload?.length) return label;
                    return `${label} | P1: ${payload[0]?.payload?.seP1 || "—"} · P2: ${payload[0]?.payload?.seP2 || "—"}`;
                  }}
                />
                <Legend formatter={n => n === "P1" ? `📘 P1 (${p1desde}→${p1hasta})` : `📗 P2 (${p2desde}→${p2hasta})`} />
                <Bar dataKey="P1" fill="#2980b9" radius={[4,4,0,0]} />
                <Bar dataKey="P2" fill="#27ae60" radius={[4,4,0,0]} />
              </BarChart>
            ) : (
              <LineChart data={dataGrafico} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e0e8f0" />
                <XAxis dataKey="idx" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip
                  formatter={(val, name) => [val ?? "—", name === "P1" ? `P1 (${p1desde}→${p1hasta})` : `P2 (${p2desde}→${p2hasta})`]}
                  labelFormatter={(label, payload) => {
                    if (!payload?.length) return label;
                    return `${label} | P1: ${payload[0]?.payload?.seP1 || "—"} · P2: ${payload[0]?.payload?.seP2 || "—"}`;
                  }}
                />
                <Legend formatter={n => n === "P1" ? `📘 P1 (${p1desde}→${p1hasta})` : `📗 P2 (${p2desde}→${p2hasta})`} />
                <Line dataKey="P1" stroke="#2980b9" strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
                <Line dataKey="P2" stroke="#27ae60" strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
              </LineChart>
            )}
          </ResponsiveContainer>
        </div>

        {/* Tabla resumen */}
        <div style={{ overflowX: "auto" }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: P.azulDark, marginBottom: 10 }}>Resumen total de cada período</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ padding: "10px 12px", background: P.azulLight, color: P.azulDark, fontSize: 12, fontWeight: 800, textAlign: "left" }}>Métrica</th>
                <th style={{ padding: "10px 12px", background: "#2980b915", color: "#2980b9", fontSize: 12, fontWeight: 800, textAlign: "right" }}>
                  📘 {p1desde}{p1hasta && p1hasta !== p1desde ? ` → ${p1hasta}` : ""}
                </th>
                <th style={{ padding: "10px 12px", background: "#27ae6015", color: "#27ae60", fontSize: 12, fontWeight: 800, textAlign: "right" }}>
                  📗 {p2desde}{p2hasta && p2hasta !== p2desde ? ` → ${p2hasta}` : ""}
                </th>
              </tr>
            </thead>
            <tbody>
              <Row label="Semanas comparadas"        k="semanas" />
              <Row label="Demanda Total"             k="demanda" />
              <Row label="Pacientes Atendidos"       k="atendidos" />
              <Row label="Atenciones Respiratorias"  k="resp" />
              <Row label="Abandonos"                 k="abandonos"    invert={true} />
              <Row label="Tasa de Abandono (%)"      k="tasaAbandono" invert={true} />
              <Row label="% Respiratorio"            k="pctResp" />
              <Row label="T° Espera Prom. (min)"     k="espera"       invert={true} />
            </tbody>
          </table>
          <div style={{ fontSize: 11, color: P.muted, marginTop: 8 }}>▲▼ variación del Período 2 respecto al Período 1 (referencia)</div>
        </div>
      </>) : (
        <div style={{ textAlign: "center", padding: 24, color: P.muted, fontSize: 13 }}>
          Selecciona un rango de SE en cada período para comparar
        </div>
      )}
    </div>
  );
}

// ── Resumen Ambulancias ───────────────────────────────────────────────────────
function ResumenAmbulancias({ registrosAmbulancias, filtroEstab, filtroPolo, POLO_MAP, P }) {
  const ambFiltrados = registrosAmbulancias.filter(r =>
    (filtroEstab === "Todos" || r.establecimiento === filtroEstab) &&
    (filtroPolo === "Todos" || POLO_MAP[r.establecimiento] === filtroPolo)
  );
  if (ambFiltrados.length === 0) return null;

  const tiempos = ambFiltrados.filter(r => r.tiempo_retencion != null).map(r => Number(r.tiempo_retencion));
  const promedio = tiempos.length ? Math.round(tiempos.reduce((a,b) => a+b, 0) / tiempos.length) : 0;
  const maximo   = tiempos.length ? Math.max(...tiempos) : 0;

  const porEstab = {};
  ambFiltrados.forEach(r => {
    if (!porEstab[r.establecimiento]) porEstab[r.establecimiento] = { total: 0, tiempos: [] };
    porEstab[r.establecimiento].total++;
    if (r.tiempo_retencion != null) porEstab[r.establecimiento].tiempos.push(Number(r.tiempo_retencion));
  });

  return (
    <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: "16px 20px", marginTop: 24 }}>
      <div style={{ fontSize: 15, fontWeight: 800, color: P.azulDark, marginBottom: 4 }}>🚑 Retenciones de Ambulancias</div>
      <div style={{ fontSize: 12, color: P.muted, marginBottom: 16 }}>Resumen según filtros aplicados</div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
        {[
          { label: "Total Retenciones", val: ambFiltrados.length, color: P.azul },
          { label: "Tiempo Promedio",   val: `${promedio} min`,   color: "#e67e22" },
          { label: "Tiempo Máximo",     val: `${maximo} min`,     color: "#e74c3c" },
        ].map(({ label, val, color }) => (
          <div key={label} style={{ background: P.azulLight, borderRadius: 10, padding: "12px 16px", borderLeft: `4px solid ${color}` }}>
            <div style={{ fontSize: 11, color: P.muted, fontWeight: 600 }}>{label}</div>
            <div style={{ fontSize: 20, fontWeight: 800, color }}>{val}</div>
          </div>
        ))}
      </div>

      <div style={{ fontSize: 12, fontWeight: 800, color: P.azulDark, marginBottom: 10 }}>Por establecimiento</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 10 }}>
        {Object.entries(porEstab).sort((a,b) => b[1].total - a[1].total).map(([estab, data]) => {
          const prom = data.tiempos.length ? Math.round(data.tiempos.reduce((a,b)=>a+b,0)/data.tiempos.length) : 0;
          return (
            <div key={estab} style={{ background: "#fff", border: `1px solid ${P.border}`, borderRadius: 8, padding: "10px 14px" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: P.azulDark, marginBottom: 4 }}>{estab}</div>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ fontSize: 12, color: P.muted }}>Retenciones: <b style={{ color: P.azul }}>{data.total}</b></span>
                <span style={{ fontSize: 12, color: P.muted }}>Prom: <b style={{ color: "#e67e22" }}>{prom} min</b></span>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Proyecciones ──────────────────────────────────────────────────────────────
function Proyecciones({ registros, filtroPolo, filtroEstab, metodo, setMetodo, semanas, setSemanas, semanasOpts, P, inpS }) {

  // Calcular SE actual y anterior para excluirlas (datos incompletos)
  const hoy = new Date();
  const hoyStr = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,"0")}-${String(hoy.getDate()).padStart(2,"0")}`;
  const seActual = getEpiWeek(hoyStr);
  const haceUnaSemana = new Date(hoy);
  haceUnaSemana.setDate(haceUnaSemana.getDate() - 7);
  const haceUnaSemanaStr = `${haceUnaSemana.getFullYear()}-${String(haceUnaSemana.getMonth()+1).padStart(2,"0")}-${String(haceUnaSemana.getDate()).padStart(2,"0")}`;
  const seAnterior = getEpiWeek(haceUnaSemanaStr);

  // Filtrar registros según polo y establecimiento, excluyendo SE actual y anterior
  const base = registros.filter(r =>
    (filtroPolo === "Todos" || getPolo(r.establecimiento) === filtroPolo) &&
    (filtroEstab === "Todos" || r.establecimiento === filtroEstab) &&
    r.semana_epi !== seActual &&
    r.semana_epi !== seAnterior
  );

  // Agrupar por SE → { demanda, resp }
  const porSE = {};
  base.forEach(r => {
    const se = r.semana_epi;
    if (!se) return;
    if (!porSE[se]) porSE[se] = { demanda: 0, resp: 0, n: 0 };
    porSE[se].demanda += Number(r.demanda_total || 0);
    porSE[se].resp    += Number(r.atenciones_respiratorias || 0);
    porSE[se].n++;
  });

  const seOrdenadas = semanasOpts.filter(se => porSE[se]);
  const historial   = seOrdenadas.map(se => ({ se, ...porSE[se] }));

  // Calcular proyección
  const proyectar = () => {
    if (historial.length < 2) return [];
    const resultado = [];
    const N = Number(semanas);

    for (let i = 1; i <= N; i++) {
      const seNum  = historial.length + i;
      const seLabel = `SE ${String(seNum).padStart(2, "0")} (proy)`;

      let demanda, resp;

      if (metodo === "historico") {
        // Promedio de todas las semanas históricas
        demanda = Math.round(historial.reduce((a, s) => a + s.demanda, 0) / historial.length);
        resp    = Math.round(historial.reduce((a, s) => a + s.resp, 0)    / historial.length);
      } else {
        // Tendencia: regresión lineal simple sobre últimas 8 semanas
        const ultimas = historial.slice(-8);
        const n = ultimas.length;
        const xs = ultimas.map((_, j) => j + 1);
        const meanX = xs.reduce((a, b) => a + b, 0) / n;

        const linReg = (vals) => {
          const meanY = vals.reduce((a, b) => a + b, 0) / n;
          const num   = xs.reduce((a, x, j) => a + (x - meanX) * (vals[j] - meanY), 0);
          const den   = xs.reduce((a, x) => a + (x - meanX) ** 2, 0);
          const slope = den ? num / den : 0;
          const inter = meanY - slope * meanX;
          return Math.max(0, Math.round(inter + slope * (n + i)));
        };

        demanda = linReg(ultimas.map(s => s.demanda));
        resp    = linReg(ultimas.map(s => s.resp));
      }

      resultado.push({ se: seLabel, demanda, resp, proyectado: true });
    }
    return resultado;
  };

  const proyecciones = proyectar();

  // Datos para el gráfico: últimas 12 semanas históricas + proyecciones
  const ultimasHist = historial.slice(-12).map(s => ({ ...s, proyectado: false }));
  const dataGrafico = [...ultimasHist, ...proyecciones];

  const btnStyle = (activo) => ({
    padding: "6px 14px", borderRadius: 6, fontSize: 12, fontWeight: 700,
    cursor: "pointer", border: `1px solid ${P.border}`,
    background: activo ? P.azul : P.azulLight,
    color: activo ? "#fff" : P.azulDark,
  });

  if (historial.length < 2) return (
    <div style={{ textAlign: "center", padding: 60, color: P.muted, fontSize: 14 }}>
      Necesitas al menos 2 semanas de datos históricos para generar proyecciones.
    </div>
  );

  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <div style={{ fontSize: 20, fontWeight: 800, color: P.azulDark, marginBottom: 4 }}>📈 Proyección de Demanda</div>
      <div style={{ fontSize: 13, color: P.muted, marginBottom: 24 }}>
        Estimación basada en datos históricos según filtros aplicados
      </div>

      {/* Controles */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: "16px 20px", marginBottom: 24, display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, color: P.muted, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>Método</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={btnStyle(metodo === "historico")} onClick={() => setMetodo("historico")}>📊 Promedio histórico</button>
            <button style={btnStyle(metodo === "tendencia")} onClick={() => setMetodo("tendencia")}>📉 Tendencia reciente</button>
          </div>
        </div>
        <div>
          <div style={{ fontSize: 11, fontWeight: 700, color: P.muted, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>Semanas a proyectar</div>
          <div style={{ display: "flex", gap: 8 }}>
            {[2, 4, 6, 8].map(n => (
              <button key={n} style={btnStyle(semanas === n)} onClick={() => setSemanas(n)}>{n}</button>
            ))}
          </div>
        </div>
        <div style={{ marginLeft: "auto", background: P.azulLight, borderRadius: 8, padding: "10px 16px", fontSize: 12, color: P.azulDark }}>
          <b>{historial.length}</b> semanas históricas · <b>{proyecciones.length}</b> proyectadas
        </div>
      </div>

      {/* Nota semana excluida */}
      <div style={{ background: "#fff3cd", border: "1px solid #ffc107", borderRadius: 8, padding: "10px 16px", marginBottom: 16, fontSize: 12, color: "#856404" }}>
        ⚠️ La semana en curso (<b>{seActual}</b>) y la anterior (<b>{seAnterior}</b>) fueron excluidas automáticamente por tener datos incompletos o aún en carga.
      </div>

      {/* Gráfico Demanda */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: "16px", marginBottom: 16 }}>
        <div style={{ fontSize: 14, fontWeight: 800, color: P.azulDark, marginBottom: 16 }}>
          Demanda Total
          <span style={{ fontSize: 11, fontWeight: 400, color: P.muted, marginLeft: 12 }}>— histórico &nbsp;&nbsp; - - - proyectado</span>
        </div>
        <ResponsiveContainer width="100%" height={260}>
          <ComposedChart data={dataGrafico} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e0e8f0" />
            <XAxis dataKey="se" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={50} />
            <YAxis tick={{ fontSize: 11 }} />
            <Tooltip formatter={(val, name) => [val, name === "demanda" ? "Demanda" : "Proyección"]} />
            <Bar dataKey="demanda" fill="#2980b9" radius={[4,4,0,0]}
              label={false}
              isAnimationActive={false}
            />
            <Line
              dataKey="demanda"
              stroke="#2980b9" strokeWidth={0} dot={false}
            />
          </ComposedChart>
        </ResponsiveContainer>
        <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
          {dataGrafico.map((d, i) => d.proyectado && (
            <div key={i} style={{ background: "#fff3cd", border: "1px solid #ffc107", borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, color: "#856404" }}>
              {d.se.replace(" (proy)", "")}: {d.demanda.toLocaleString()}
            </div>
          ))}
        </div>
      </div>

      {/* Gráfico Respiratorias */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: "16px", marginBottom: 24 }}>
        <div style={{ fontSize: 14, fontWeight: 800, color: P.azulDark, marginBottom: 16 }}>
          Atenciones Respiratorias
          <span style={{ fontSize: 11, fontWeight: 400, color: P.muted, marginLeft: 12 }}>— histórico &nbsp;&nbsp; - - - proyectado</span>
        </div>
        <ResponsiveContainer width="100%" height={260}>
          <ComposedChart data={dataGrafico} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#e0e8f0" />
            <XAxis dataKey="se" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={50} />
            <YAxis tick={{ fontSize: 11 }} />
            <Tooltip formatter={(val) => [val, "Respiratorias"]} />
            <Bar dataKey="resp" fill="#e67e22" radius={[4,4,0,0]} isAnimationActive={false} />
          </ComposedChart>
        </ResponsiveContainer>
        <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
          {dataGrafico.map((d, i) => d.proyectado && (
            <div key={i} style={{ background: "#fff3cd", border: "1px solid #ffc107", borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, color: "#856404" }}>
              {d.se.replace(" (proy)", "")}: {d.resp.toLocaleString()}
            </div>
          ))}
        </div>
      </div>

      {/* Tabla resumen proyecciones */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: "18px", marginBottom: 24 }}>
        <div style={{ fontSize: 14, fontWeight: 800, color: P.azulDark, marginBottom: 16 }}>Detalle de proyecciones</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ padding: "10px 12px", background: P.azulLight, color: P.azulDark, fontSize: 12, fontWeight: 800, textAlign: "left", borderRadius: "8px 0 0 0" }}>Semana</th>
              <th style={{ padding: "10px 12px", background: P.azulLight, color: P.azulDark, fontSize: 12, fontWeight: 800, textAlign: "right" }}>Demanda proyectada</th>
              <th style={{ padding: "10px 12px", background: P.azulLight, color: P.azulDark, fontSize: 12, fontWeight: 800, textAlign: "right", borderRadius: "0 8px 0 0" }}>Respiratorias proyectadas</th>
            </tr>
          </thead>
          <tbody>
            {proyecciones.map((p, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? "#fffbf0" : "#fff" }}>
                <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700, color: "#856404", borderBottom: `1px solid ${P.border}` }}>
                  {p.se}
                </td>
                <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700, color: P.azul, textAlign: "right", borderBottom: `1px solid ${P.border}` }}>
                  {p.demanda.toLocaleString()}
                </td>
                <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 700, color: "#e67e22", textAlign: "right", borderBottom: `1px solid ${P.border}` }}>
                  {p.resp.toLocaleString()}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ fontSize: 11, color: P.muted, marginTop: 10 }}>
          {metodo === "historico"
            ? `Método: promedio de las ${historial.length} semanas históricas disponibles`
            : `Método: regresión lineal sobre las últimas ${Math.min(8, historial.length)} semanas`}
        </div>
      </div>
    </div>
  );
}

// ── Tab Tiempos de Espera ─────────────────────────────────────────────────────
function TabTiemposEspera({ registros, P, inpS }) {
  const [vistaEstab, setVistaEstab] = React.useState("promedio");
  const [tPolo,    setTPolo]    = React.useState("Todos");
  const [tEstab,   setTEstab]   = React.useState("Todos");
  const [tSEDesde, setTSEDesde] = React.useState("");
  const [tSEHasta, setTSEHasta] = React.useState("");

  // Opciones disponibles
  const semanasDisp = React.useMemo(() =>
    [...new Set(registros.map(r => r.semana_epi).filter(Boolean))].sort()
  , [registros]);
  const establDisp = React.useMemo(() =>
    [...new Set(registros.map(r => r.establecimiento).filter(Boolean))].sort()
  , [registros]);

  const selStyle = { ...inpS, fontSize: 12, padding: "6px 10px" };

  // Aplicar filtros propios
  const base = registros.filter(r =>
    r.tiempo_espera != null && r.tiempo_espera !== "" &&
    (tPolo  === "Todos" || getPolo(r.establecimiento) === tPolo) &&
    (tEstab === "Todos" || r.establecimiento === tEstab) &&
    (!tSEDesde || (r.semana_epi >= tSEDesde)) &&
    (!tSEHasta || (r.semana_epi <= tSEHasta))
  );

  if (base.length === 0) return (
    <div style={{ textAlign: "center", padding: 60, color: P.muted, fontSize: 14 }}>
      Sin datos de tiempo de espera con los filtros aplicados.
    </div>
  );

  // ── Por establecimiento ───────────────────────────────────────────────────
  const porEstab = {};
  base.forEach(r => {
    const e = r.establecimiento;
    if (!porEstab[e]) porEstab[e] = { vals: [] };
    porEstab[e].vals.push(Number(r.tiempo_espera));
  });
  const dataEstab = Object.entries(porEstab).map(([e, d]) => {
    const sorted = [...d.vals].sort((a,b) => a-b);
    const prom = Math.round(d.vals.reduce((a,b) => a+b,0) / d.vals.length);
    return {
      establecimiento: e,
      promedio: prom,
      minimo: Math.min(...d.vals),
      maximo: Math.max(...d.vals),
      n: d.vals.length,
      alerta: prom > 180,
    };
  }).sort((a,b) => b.promedio - a.promedio);

  // ── Por SE (red completa) ─────────────────────────────────────────────────
  const porSE = {};
  base.forEach(r => {
    const se = r.semana_epi;
    if (!se) return;
    if (!porSE[se]) porSE[se] = { vals: [] };
    porSE[se].vals.push(Number(r.tiempo_espera));
  });
  const dataSE = Object.entries(porSE).sort(([a],[b]) => a.localeCompare(b)).map(([se, d]) => ({
    se,
    promedio: Math.round(d.vals.reduce((a,b) => a+b,0) / d.vals.length),
    maximo: Math.max(...d.vals),
    minimo: Math.min(...d.vals),
  }));

  // ── Por día (red completa) ────────────────────────────────────────────────
  const porDia = {};
  base.forEach(r => {
    const dia = r.fecha?.slice(0,10);
    if (!dia) return;
    if (!porDia[dia]) porDia[dia] = { vals: [] };
    porDia[dia].vals.push(Number(r.tiempo_espera));
  });
  const dataDia = Object.entries(porDia).sort(([a],[b]) => a.localeCompare(b)).map(([dia, d]) => ({
    dia,
    promedio: Math.round(d.vals.reduce((a,b) => a+b,0) / d.vals.length),
  }));

  // ── Días críticos ─────────────────────────────────────────────────────────
  const diasCriticos = base
    .filter(r => Number(r.tiempo_espera) > 180)
    .map(r => ({ fecha: r.fecha?.slice(0,10), estab: r.establecimiento, valor: Number(r.tiempo_espera), se: r.semana_epi }))
    .sort((a,b) => b.valor - a.valor)
    .slice(0, 20);

  // ── KPIs generales ────────────────────────────────────────────────────────
  const todos = base.map(r => Number(r.tiempo_espera));
  const promGeneral = Math.round(todos.reduce((a,b) => a+b,0) / todos.length);
  const maxGeneral  = Math.max(...todos);
  const nAlerta     = dataEstab.filter(e => e.alerta).length;

  // ── Tendencia últimas 8 SE ────────────────────────────────────────────────
  const ultimas8 = dataSE.slice(-8);
  const tendencia = ultimas8.length >= 2
    ? ultimas8[ultimas8.length-1].promedio - ultimas8[0].promedio
    : 0;

  const btnStyle = (activo) => ({
    padding: "5px 14px", borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: "pointer",
    background: activo ? P.azul : P.azulLight, color: activo ? "#fff" : P.azulDark,
    border: `1px solid ${activo ? P.azul : P.border}`,
  });

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto" }}>
      <div style={{ fontSize: 20, fontWeight: 800, color: P.azulDark, marginBottom: 4 }}>⏱️ Tiempos de Espera</div>
      <div style={{ fontSize: 13, color: P.muted, marginBottom: 16 }}>Análisis por establecimiento, semana epidemiológica y día</div>

      {/* Filtros propios */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 12, padding: "14px 18px", marginBottom: 20, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end" }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: P.azulDark, alignSelf: "center" }}>🔍 Filtros</div>

        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, marginBottom: 4, textTransform: "uppercase" }}>Polo</div>
          <select value={tPolo} onChange={e => { setTPolo(e.target.value); setTEstab("Todos"); }} style={selStyle}>
            <option value="Todos">Todos</option>
            {["Polo Cerrillos Maipú","Polo Santiago Estación Central"].map(p => <option key={p}>{p}</option>)}
          </select>
        </div>

        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, marginBottom: 4, textTransform: "uppercase" }}>Establecimiento</div>
          <select value={tEstab} onChange={e => setTEstab(e.target.value)} style={selStyle}>
            <option value="Todos">Todos</option>
            {establDisp
              .filter(e => tPolo === "Todos" || getPolo(e) === tPolo)
              .map(e => <option key={e}>{e}</option>)}
          </select>
        </div>

        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, marginBottom: 4, textTransform: "uppercase" }}>SE Desde</div>
          <select value={tSEDesde} onChange={e => setTSEDesde(e.target.value)} style={{ ...selStyle, width: 90 }}>
            <option value="">Todas</option>
            {semanasDisp.map(se => <option key={se}>{se}</option>)}
          </select>
        </div>

        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: P.muted, marginBottom: 4, textTransform: "uppercase" }}>SE Hasta</div>
          <select value={tSEHasta} onChange={e => setTSEHasta(e.target.value)} style={{ ...selStyle, width: 90 }}>
            <option value="">Todas</option>
            {semanasDisp.map(se => <option key={se}>{se}</option>)}
          </select>
        </div>

        <button onClick={() => { setTPolo("Todos"); setTEstab("Todos"); setTSEDesde(""); setTSEHasta(""); }}
          style={{ padding: "6px 14px", borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: "pointer",
            background: P.azulLight, color: P.azulDark, border: `1px solid ${P.border}` }}>
          ✕ Limpiar
        </button>

        <div style={{ marginLeft: "auto", fontSize: 12, color: P.muted, alignSelf: "center" }}>
          <b style={{ color: P.azulDark }}>{base.length}</b> registros con T° espera
        </div>
      </div>

      {/* KPIs */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 24 }}>
        {[
          { label: "Promedio Red",     val: `${promGeneral} min`, color: promGeneral > 180 ? "#e74c3c" : P.azul,  alerta: promGeneral > 180 },
          { label: "Máximo registrado",val: `${maxGeneral} min`,  color: maxGeneral > 180  ? "#e74c3c" : P.amber, alerta: maxGeneral > 180  },
          { label: "Centros en alerta",val: `${nAlerta} de ${dataEstab.length}`, color: nAlerta > 0 ? "#e74c3c" : P.verde, alerta: nAlerta > 0 },
          { label: "Tendencia (últ. 8 SE)", val: tendencia === 0 ? "Estable" : `${tendencia > 0 ? "▲" : "▼"} ${Math.abs(tendencia)} min`,
            color: tendencia > 15 ? "#e74c3c" : tendencia < -15 ? P.verde : P.amber, alerta: tendencia > 15 },
        ].map(({ label, val, color, alerta }) => (
          <div key={label} style={{
            background: alerta ? "#fdecea" : P.azulLight, borderRadius: 10, padding: "12px 16px",
            borderLeft: `4px solid ${color}`, border: alerta ? `1.5px solid #e74c3c` : undefined,
          }}>
            <div style={{ fontSize: 11, color: P.muted, fontWeight: 600 }}>{label}</div>
            <div style={{ fontSize: 18, fontWeight: 800, color }}>{val}</div>
            {alerta && <div style={{ fontSize: 10, color: "#e74c3c", fontWeight: 700, marginTop: 2 }}>⚠️ Sobre umbral</div>}
          </div>
        ))}
      </div>

      {/* Por establecimiento */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark }}>Por Establecimiento</div>
            <div style={{ fontSize: 11, color: P.muted }}>Comparación entre centros — umbral de alerta: 180 min</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button style={btnStyle(vistaEstab === "promedio")} onClick={() => setVistaEstab("promedio")}>Promedio</button>
            <button style={btnStyle(vistaEstab === "rango")}   onClick={() => setVistaEstab("rango")}>Rango min/max</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 18 }}>
          <ResponsiveContainer width="100%" height={Math.max(220, dataEstab.length * 38)}>
            <BarChart data={dataEstab} layout="vertical" margin={{ top: 4, right: 50, left: 16, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11 }} tickFormatter={v => `${v}m`} />
              <YAxis type="category" dataKey="establecimiento" width={148} tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v, n) => [`${v} min`, n === "promedio" ? "Promedio" : n === "maximo" ? "Máximo" : "Mínimo"]} />
              <ReferenceLine x={180} stroke="#e74c3c" strokeDasharray="4 2"
                label={{ value: "180 min", position: "top", fontSize: 10, fill: "#e74c3c" }} />
              {vistaEstab === "promedio" ? (
                <Bar dataKey="promedio" radius={[0,6,6,0]} barSize={20}>
                  {dataEstab.map((d,i) => <Cell key={i} fill={d.promedio > 180 ? "#e74c3c" : P.azul} />)}
                </Bar>
              ) : (<>
                <Bar dataKey="minimo"   fill={`${P.verde}90`} radius={[0,0,0,0]} barSize={10} stackId="a" />
                <Bar dataKey="promedio" fill={`${P.azul}90`}  radius={[0,0,0,0]} barSize={10} stackId="a" />
                <Bar dataKey="maximo"   fill={`${P.amber}90`} radius={[0,6,6,0]} barSize={10} stackId="a" />
              </>)}
            </BarChart>
          </ResponsiveContainer>

          {/* Tabla ranking */}
          <div style={{ overflow: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 280 }}>
              <thead>
                <tr style={{ background: P.azulDark }}>
                  {["#", "Establecimiento", "Prom.", "Mín.", "Máx.", "Días"].map(h => (
                    <th key={h} style={{ padding: "8px 10px", color: "#fff", fontWeight: 700, textAlign: h === "Establecimiento" ? "left" : "right", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {dataEstab.map((d, i) => (
                  <tr key={d.establecimiento} style={{ background: d.alerta ? "#fdecea" : i % 2 === 0 ? "#fff" : P.bg }}>
                    <td style={{ padding: "8px 10px", textAlign: "right", color: P.muted, fontWeight: 700 }}>{i+1}</td>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: d.alerta ? "#e74c3c" : P.azulDark }}>
                      {d.alerta && "⚠️ "}{d.establecimiento}
                    </td>
                    <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 800, color: d.alerta ? "#e74c3c" : P.azul }}>{d.promedio}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", color: P.verde }}>{d.minimo}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", color: P.amber }}>{d.maximo}</td>
                    <td style={{ padding: "8px 10px", textAlign: "right", color: P.muted }}>{d.n}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Por SE */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark, marginBottom: 4 }}>Promedio de la Red por Semana Epidemiológica</div>
        <div style={{ fontSize: 11, color: P.muted, marginBottom: 14 }}>Evolución semanal del tiempo de espera promedio</div>
        <ResponsiveContainer width="100%" height={240}>
          <ComposedChart data={dataSE} margin={{ top: 4, right: 24, left: 0, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
            <XAxis dataKey="se" tick={{ fontSize: 10 }} interval={dataSE.length > 20 ? Math.floor(dataSE.length/8) : 0} angle={dataSE.length > 15 ? -30 : 0} textAnchor={dataSE.length > 15 ? "end" : "middle"} height={dataSE.length > 15 ? 45 : 25} />
            <YAxis tick={{ fontSize: 11 }} tickFormatter={v => `${v}m`} />
            <Tooltip formatter={(v, n) => [`${v} min`, n === "promedio" ? "Promedio red" : n === "maximo" ? "Máximo" : "Mínimo"]} />
            <ReferenceLine y={180} stroke="#e74c3c" strokeDasharray="4 2" label={{ value: "⚠️ 180 min", position: "right", fontSize: 10, fill: "#e74c3c" }} />
            <Bar dataKey="promedio" radius={[4,4,0,0]} barSize={dataSE.length > 20 ? 6 : 14}>
              {dataSE.map((d,i) => <Cell key={i} fill={d.promedio > 180 ? "#e74c3c" : P.azul} />)}
            </Bar>
            <Line type="monotone" dataKey="promedio" stroke={P.azul} strokeWidth={2} dot={false} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      {/* Por día */}
      <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 8, padding: 18, marginBottom: 14 }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: P.azulDark, marginBottom: 4 }}>Promedio de la Red por Día</div>
        <div style={{ fontSize: 11, color: P.muted, marginBottom: 14 }}>Evolución diaria del tiempo de espera — barras rojas superan 180 min</div>
        <ResponsiveContainer width="100%" height={240}>
          <ComposedChart data={dataDia} margin={{ top: 4, right: 24, left: 0, bottom: 4 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={P.grisMid} />
            <XAxis dataKey="dia" tick={{ fontSize: 10 }} interval={dataDia.length > 20 ? Math.floor(dataDia.length/8) : 0} angle={dataDia.length > 15 ? -30 : 0} textAnchor={dataDia.length > 15 ? "end" : "middle"} height={dataDia.length > 15 ? 45 : 25} />
            <YAxis tick={{ fontSize: 11 }} tickFormatter={v => `${v}m`} />
            <Tooltip formatter={v => [`${v} min`, "Promedio red"]} />
            <ReferenceLine y={180} stroke="#e74c3c" strokeDasharray="4 2" label={{ value: "⚠️ 180 min", position: "right", fontSize: 10, fill: "#e74c3c" }} />
            <Bar dataKey="promedio" radius={[3,3,0,0]} barSize={dataDia.length > 60 ? 3 : 8}>
              {dataDia.map((d,i) => <Cell key={i} fill={d.promedio > 180 ? "#e74c3c" : `${P.azul}90`} />)}
            </Bar>
            <Line type="monotone" dataKey="promedio" stroke={P.azul} strokeWidth={2} dot={false} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      {/* Días críticos */}
      {diasCriticos.length > 0 && (
        <div style={{ background: "#fdecea", border: "1.5px solid #e74c3c", borderRadius: 8, padding: 18 }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#c0392b", marginBottom: 4 }}>⚠️ Días Críticos (sobre 180 min)</div>
          <div style={{ fontSize: 11, color: "#c0392b", marginBottom: 14 }}>Top 20 registros con mayor tiempo de espera según filtros aplicados</div>
          <div style={{ overflow: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 340 }}>
              <thead>
                <tr style={{ background: "#e74c3c" }}>
                  {["Fecha", "SE", "Establecimiento", "T° Espera"].map(h => (
                    <th key={h} style={{ padding: "8px 12px", color: "#fff", fontWeight: 700, textAlign: h === "T° Espera" ? "right" : "left", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {diasCriticos.map((d, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "#fff5f5" : "#fff" }}>
                    <td style={{ padding: "8px 12px", color: P.text }}>{d.fecha}</td>
                    <td style={{ padding: "8px 12px", color: P.muted }}>{d.se}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 700, color: "#c0392b" }}>{d.estab}</td>
                    <td style={{ padding: "8px 12px", textAlign: "right", fontWeight: 800, color: "#c0392b" }}>{d.valor} min</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Modal PDF ─────────────────────────────────────────────────────────────────
function ModalPDF({ registros, registrosAmbulancias, onClose, P, getPolo }) {
  const [pdfPolo,    setPdfPolo]    = React.useState("Todos");
  const [pdfEstab,   setPdfEstab]   = React.useState("Todos");
  const [pdfSEDesde, setPdfSEDesde] = React.useState("");
  const [pdfSEHasta, setPdfSEHasta] = React.useState("");
  const [secciones,  setSecciones]  = React.useState({
    kpis: true, absorcion: true, refuerzos: true,
    respiratorias: true, ambulancias: true, tabla: true,
  });

  const semanasDisp = [...new Set(registros.map(r => r.semana_epi).filter(Boolean))].sort();
  const establDisp  = [...new Set(registros.map(r => r.establecimiento).filter(Boolean))].sort();

  const base = registros.filter(r =>
    (pdfPolo  === "Todos" || getPolo(r.establecimiento) === pdfPolo) &&
    (pdfEstab === "Todos" || r.establecimiento === pdfEstab) &&
    (!pdfSEDesde || r.semana_epi >= pdfSEDesde) &&
    (!pdfSEHasta || r.semana_epi <= pdfSEHasta)
  );

  const tot = base.reduce((a,r) => ({
    demanda:   a.demanda   + Number(r.demanda_total||0),
    atendidos: a.atendidos + Number(r.pacientes_atendidos||0),
    resp:      a.resp      + Number(r.atenciones_respiratorias||0),
    abandonos: a.abandonos + Number(r.abandonos||0),
  }), { demanda:0, atendidos:0, resp:0, abandonos:0 });

  const conEspera    = base.filter(r => r.tiempo_espera != null && r.tiempo_espera !== "");
  const promEspera   = conEspera.length ? Math.round(conEspera.reduce((a,r)=>a+Number(r.tiempo_espera),0)/conEspera.length) : 0;
  const tasaAbandono = tot.demanda ? ((tot.abandonos/tot.demanda)*100).toFixed(1) : 0;
  const pctResp      = tot.atendidos ? ((tot.resp/tot.atendidos)*100).toFixed(1) : 0;

  const totalRed = base.reduce((a,r)=>a+Number(r.demanda_total||0),0);
  const absorMap = {};
  base.forEach(r => {
    const e = r.establecimiento;
    if (!absorMap[e]) absorMap[e] = { demanda:0, atendidos:0 };
    absorMap[e].demanda   += Number(r.demanda_total||0);
    absorMap[e].atendidos += Number(r.pacientes_atendidos||0);
  });
  const absorData = Object.entries(absorMap)
    .map(([e,d]) => ({ estab:e, demanda:d.demanda, atendidos:d.atendidos, pct: totalRed?((d.demanda/totalRed)*100).toFixed(1):0 }))
    .sort((a,b)=>b.demanda-a.demanda);

  const ESTAMENTOS = ["Médico","Enfermera","TENS","Kinesiólogo","Administrativo"];
  const EKEYS      = ["refuerzo_medico","refuerzo_enfermera","refuerzo_tens","refuerzo_kinesiologo","refuerzo_administrativo"];
  const refMap = {};
  base.filter(r=>r.tiene_refuerzo).forEach(r => {
    const e = r.establecimiento;
    if (!refMap[e]) refMap[e] = { dias:0, est:{} };
    refMap[e].dias++;
    EKEYS.forEach((k,i) => { if(r[k]) refMap[e].est[ESTAMENTOS[i]] = (refMap[e].est[ESTAMENTOS[i]]||0)+1; });
  });

  const titulo  = `Reporte Urgencias APS${pdfEstab!=="Todos"?` — ${pdfEstab}`:pdfPolo!=="Todos"?` — ${pdfPolo}`:""}`;
  const periodo = pdfSEDesde||pdfSEHasta ? `${pdfSEDesde||"Inicio"} → ${pdfSEHasta||"Fin"}` : "Período completo";
  const sel = { background:"#fff", border:`1px solid ${P.border}`, borderRadius:6, padding:"6px 10px", fontSize:12, color:P.text };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", zIndex:1000, display:"flex", alignItems:"flex-start", justifyContent:"center", overflowY:"auto", padding:"20px 0" }}>
      <div style={{ background:"#fff", borderRadius:10, width:"min(800px,95vw)", margin:"auto" }}>

        <div style={{ background:P.azul, padding:"14px 22px", borderRadius:"10px 10px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div style={{ color:"#fff", fontWeight:700, fontSize:14 }}>📄 Exportar Reporte PDF</div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,0.2)", border:"none", borderRadius:6, color:"#fff", padding:"3px 10px", cursor:"pointer", fontSize:16 }}>✕</button>
        </div>

        <div style={{ padding:"18px 22px" }}>
          {/* Filtros */}
          <div style={{ background:P.azulLight, borderRadius:8, padding:"12px 16px", marginBottom:16 }}>
            <div style={{ fontSize:12, fontWeight:800, color:P.azulDark, marginBottom:10 }}>Filtros del reporte</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10 }}>
              <div>
                <div style={{ fontSize:10, color:P.muted, fontWeight:600, marginBottom:3 }}>Polo</div>
                <select value={pdfPolo} onChange={e=>{ setPdfPolo(e.target.value); setPdfEstab("Todos"); }} style={sel}>
                  <option value="Todos">Todos</option>
                  {["Polo Cerrillos Maipú","Polo Santiago Estación Central"].map(p=><option key={p}>{p}</option>)}
                </select>
              </div>
              <div>
                <div style={{ fontSize:10, color:P.muted, fontWeight:600, marginBottom:3 }}>Establecimiento</div>
                <select value={pdfEstab} onChange={e=>setPdfEstab(e.target.value)} style={sel}>
                  <option value="Todos">Todos</option>
                  {establDisp.filter(e=>pdfPolo==="Todos"||getPolo(e)===pdfPolo).map(e=><option key={e}>{e}</option>)}
                </select>
              </div>
              <div>
                <div style={{ fontSize:10, color:P.muted, fontWeight:600, marginBottom:3 }}>SE Desde</div>
                <select value={pdfSEDesde} onChange={e=>setPdfSEDesde(e.target.value)} style={sel}>
                  <option value="">Inicio</option>
                  {semanasDisp.map(se=><option key={se}>{se}</option>)}
                </select>
              </div>
              <div>
                <div style={{ fontSize:10, color:P.muted, fontWeight:600, marginBottom:3 }}>SE Hasta</div>
                <select value={pdfSEHasta} onChange={e=>setPdfSEHasta(e.target.value)} style={sel}>
                  <option value="">Fin</option>
                  {semanasDisp.map(se=><option key={se}>{se}</option>)}
                </select>
              </div>
            </div>
          </div>

          {/* Secciones */}
          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:12, fontWeight:800, color:P.azulDark, marginBottom:8 }}>Secciones a incluir</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))", gap:6 }}>
              {[
                { key:"kpis",          label:"📊 KPIs generales" },
                { key:"absorcion",     label:"🏥 Absorción" },
                { key:"refuerzos",     label:"⚡ Refuerzos" },
                { key:"respiratorias", label:"🫁 Respiratorias" },
                { key:"ambulancias",   label:"🚑 Ambulancias" },
                { key:"tabla",         label:"📋 Tabla de registros" },
              ].map(({ key, label }) => (
                <label key={key} style={{
                  display:"flex", alignItems:"center", gap:7, cursor:"pointer",
                  background: secciones[key] ? P.azulLight : "#f9fafb",
                  border:`1px solid ${secciones[key] ? P.azul : P.border}`,
                  borderRadius:6, padding:"7px 11px", fontSize:12,
                  fontWeight: secciones[key] ? 700 : 400,
                  color: secciones[key] ? P.azulDark : P.muted,
                }}>
                  <input type="checkbox" checked={secciones[key]}
                    onChange={() => setSecciones(s=>({...s,[key]:!s[key]}))}
                    style={{ accentColor:P.azul, width:13, height:13 }}
                  />
                  {label}
                </label>
              ))}
            </div>
          </div>

          {/* Preview */}
          <div id="pdf-preview" style={{ border:`1px solid ${P.border}`, borderRadius:8, padding:18, background:"#fff", fontSize:11, maxHeight:420, overflowY:"auto" }}>
            <div style={{ borderBottom:`3px solid ${P.azul}`, paddingBottom:8, marginBottom:12, display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
              <div>
                <div style={{ fontSize:14, fontWeight:800, color:P.azul }}>{titulo}</div>
                <div style={{ fontSize:10, color:P.muted, marginTop:2 }}>SSMC · Ministerio de Salud · {periodo}</div>
              </div>
              <div style={{ fontSize:10, color:P.muted, textAlign:"right" }}>
                {new Date().toLocaleDateString("es-CL")}<br/>{base.length} registros
              </div>
            </div>

            {secciones.kpis && (
              <div style={{ marginBottom:12 }}>
                <div style={{ fontSize:11, fontWeight:800, color:P.azulDark, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Indicadores Generales</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:6 }}>
                  {[
                    { label:"Demanda Total",   val:tot.demanda.toLocaleString("es-CL"),   color:P.azul },
                    { label:"Atendidos",       val:tot.atendidos.toLocaleString("es-CL"), color:P.verde },
                    { label:"Tasa Abandono",   val:`${tasaAbandono}%`,                    color:"#C0392B" },
                    { label:"% Respiratorio",  val:`${pctResp}%`,                         color:"#B45309" },
                    { label:"Abandonos",       val:tot.abandonos.toLocaleString("es-CL"), color:"#C0392B" },
                    { label:"Respiratorias",   val:tot.resp.toLocaleString("es-CL"),       color:"#B45309" },
                    { label:"T° Espera Prom.", val:`${promEspera} min`,                   color:promEspera>180?"#C0392B":P.azul },
                    { label:"N° Semanas",      val:[...new Set(base.map(r=>r.semana_epi).filter(Boolean))].length, color:P.muted },
                  ].map(({ label, val, color }) => (
                    <div key={label} style={{ background:P.azulLight, borderRadius:5, padding:"6px 8px", borderLeft:`3px solid ${color}` }}>
                      <div style={{ fontSize:9, color:P.muted }}>{label}</div>
                      <div style={{ fontSize:13, fontWeight:800, color }}>{val}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {secciones.absorcion && absorData.length > 0 && (
              <div style={{ marginBottom:12 }}>
                <div style={{ fontSize:11, fontWeight:800, color:P.azulDark, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Absorción por Establecimiento</div>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:10 }}>
                  <thead><tr style={{ background:P.azul }}>
                    {["Establecimiento","Demanda","% Red","Atendidos"].map(h=>(
                      <th key={h} style={{ padding:"5px 7px", color:"#fff", fontWeight:700, textAlign:h==="Establecimiento"?"left":"right" }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>{absorData.map((r,i)=>(
                    <tr key={r.estab} style={{ background:i%2===0?"#fff":P.azulLight }}>
                      <td style={{ padding:"4px 7px", fontWeight:600 }}>{r.estab}</td>
                      <td style={{ padding:"4px 7px", textAlign:"right", color:P.azul, fontWeight:700 }}>{r.demanda.toLocaleString("es-CL")}</td>
                      <td style={{ padding:"4px 7px", textAlign:"right", fontWeight:700 }}>{r.pct}%</td>
                      <td style={{ padding:"4px 7px", textAlign:"right", color:P.verde, fontWeight:700 }}>{r.atendidos.toLocaleString("es-CL")}</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            )}

            {secciones.refuerzos && Object.keys(refMap).length > 0 && (
              <div style={{ marginBottom:12 }}>
                <div style={{ fontSize:11, fontWeight:800, color:P.azulDark, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Refuerzos por Establecimiento</div>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:10 }}>
                  <thead><tr style={{ background:P.azul }}>
                    <th style={{ padding:"5px 7px", color:"#fff", fontWeight:700, textAlign:"left" }}>Establecimiento</th>
                    <th style={{ padding:"5px 7px", color:"#fff", fontWeight:700, textAlign:"center" }}>Días</th>
                    {ESTAMENTOS.map(e=><th key={e} style={{ padding:"5px 7px", color:"#fff", fontWeight:700, textAlign:"center" }}>{e}</th>)}
                  </tr></thead>
                  <tbody>{Object.entries(refMap).sort((a,b)=>b[1].dias-a[1].dias).map(([estab,data],i)=>(
                    <tr key={estab} style={{ background:i%2===0?"#fff":P.azulLight }}>
                      <td style={{ padding:"4px 7px", fontWeight:600 }}>{estab}</td>
                      <td style={{ padding:"4px 7px", textAlign:"center", fontWeight:700, color:P.azul }}>{data.dias}</td>
                      {ESTAMENTOS.map(e=><td key={e} style={{ padding:"4px 7px", textAlign:"center" }}>{data.est[e]?`${data.est[e]}d`:"—"}</td>)}
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            )}

            {secciones.ambulancias && registrosAmbulancias.length > 0 && (() => {
              const ambBase = registrosAmbulancias.filter(r =>
                (pdfEstab==="Todos"||r.establecimiento===pdfEstab) &&
                (pdfPolo==="Todos"||getPolo(r.establecimiento)===pdfPolo)
              );
              const tiempos = ambBase.filter(r=>r.tiempo_retencion!=null).map(r=>Number(r.tiempo_retencion));
              return (
                <div style={{ marginBottom:12 }}>
                  <div style={{ fontSize:11, fontWeight:800, color:P.azulDark, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>Retenciones de Ambulancias</div>
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:6 }}>
                    {[
                      { label:"Total",    val:ambBase.length,                                                                     color:P.azul },
                      { label:"Prom.",    val:tiempos.length?`${Math.round(tiempos.reduce((a,b)=>a+b,0)/tiempos.length)} min`:"—",color:"#B45309" },
                      { label:"Máximo",   val:tiempos.length?`${Math.max(...tiempos)} min`:"—",                                    color:"#C0392B" },
                    ].map(({label,val,color})=>(
                      <div key={label} style={{ background:P.azulLight, borderRadius:5, padding:"6px 8px", borderLeft:`3px solid ${color}` }}>
                        <div style={{ fontSize:9, color:P.muted }}>{label}</div>
                        <div style={{ fontSize:13, fontWeight:800, color }}>{val}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

            {secciones.tabla && base.length > 0 && (
              <div>
                <div style={{ fontSize:11, fontWeight:800, color:P.azulDark, marginBottom:6, textTransform:"uppercase", letterSpacing:"0.5px" }}>
                  Registros ({Math.min(base.length,50)} de {base.length})
                </div>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:9 }}>
                  <thead><tr style={{ background:P.azul }}>
                    {["Fecha","SE","Establecimiento","Demanda","Atendidos","Resp.","Abandon.","T°Esp."].map(h=>(
                      <th key={h} style={{ padding:"4px 6px", color:"#fff", fontWeight:700, textAlign:h==="Establecimiento"?"left":"right" }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>{base.slice(0,50).map((r,i)=>(
                    <tr key={r.id||i} style={{ background:i%2===0?"#fff":P.azulLight }}>
                      <td style={{ padding:"3px 6px" }}>{r.fecha}</td>
                      <td style={{ padding:"3px 6px" }}>{r.semana_epi}</td>
                      <td style={{ padding:"3px 6px", fontWeight:600 }}>{r.establecimiento}</td>
                      <td style={{ padding:"3px 6px", textAlign:"right" }}>{r.demanda_total}</td>
                      <td style={{ padding:"3px 6px", textAlign:"right", color:P.verde }}>{r.pacientes_atendidos}</td>
                      <td style={{ padding:"3px 6px", textAlign:"right", color:"#B45309" }}>{r.atenciones_respiratorias}</td>
                      <td style={{ padding:"3px 6px", textAlign:"right", color:"#C0392B" }}>{r.abandonos}</td>
                      <td style={{ padding:"3px 6px", textAlign:"right" }}>{r.tiempo_espera?`${r.tiempo_espera}m`:"—"}</td>
                    </tr>
                  ))}</tbody>
                </table>
                {base.length > 50 && <div style={{ fontSize:9, color:P.muted, marginTop:3 }}>* Primeros 50 registros</div>}
              </div>
            )}

            <div style={{ borderTop:`1px solid ${P.border}`, paddingTop:6, marginTop:12, display:"flex", justifyContent:"space-between", fontSize:9, color:P.muted }}>
              <span>Servicio de Salud Metropolitano Central — Urgencias APS</span>
              <span>{new Date().toLocaleDateString("es-CL")} — Dashboard Urgencias SSMC</span>
            </div>
          </div>

          <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:14 }}>
            <button onClick={onClose} style={{ padding:"7px 18px", borderRadius:6, border:`1px solid ${P.border}`, background:"#fff", color:P.text, cursor:"pointer", fontSize:12, fontWeight:600 }}>
              Cancelar
            </button>
            <button onClick={() => window.print()} style={{ padding:"7px 22px", borderRadius:6, border:"none", background:P.rojo, color:"#fff", cursor:"pointer", fontSize:12, fontWeight:700 }}>
              🖨️ Imprimir / Guardar PDF
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
