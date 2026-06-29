import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

const root = path.resolve("C:/Users/julio/Desktop/TFM");
const outDir = path.join(root, "04_EVIDENCE", "Excel_visibilidad_26062026");
const normDir = path.join(outDir, "normalized");
fs.mkdirSync(normDir, { recursive: true });

const validasDir = path.join(root, "05_LOGS", "Validas");
const realDir = path.join(validasDir, "Prueba_24_06_2026");
const fpDir = path.join(validasDir, "Prueba_24_06_2026_FP");
const wazuhDir = path.join(root, "10_WAZUH", "TFM_WAZUH_VISIBILIDAD_EVIDENCE");
const baseXlsx = path.join(root, "04_EVIDENCE", "Analisis_Tecnicas_TFM_V.4.xlsx");

const realSummaryJson = path.join(realDir, "TFM_TEC_Run_CANDIDATE_v5_20260624_145022_summary.json");
const realSummaryCsv = path.join(realDir, "TFM_TEC_Run_CANDIDATE_v5_20260624_145022_summary.csv");
const fpSummaryJson = path.join(fpDir, "TFM_FP_20260626_101948_summary.json");
const fpSummaryCsv = path.join(fpDir, "TFM_FP_20260626_101948_summary.csv");
const fpHitsCsv = path.join(fpDir, "TFM_FP_20260626_101948_vr_hits.csv");
const wazuhBaseSummary = path.join(wazuhDir, "ENTREGA_WAZUH_VISIBILIDAD_20260622_114027", "base", "wazuh_base_summary.txt");
const wazuhCustomSummary = path.join(wazuhDir, "ENTREGA_WAZUH_VISIBILIDAD_20260622_114027", "custom", "wazuh_custom_summary.txt");
const wazuhRuleCount = path.join(wazuhDir, "ENTREGA_WAZUH_VISIBILIDAD_20260622_114027", "custom", "custom_rule_count_110xxx.txt");
const wazuh110201 = path.join(wazuhDir, "ENTREGA_WAZUH_VISIBILIDAD_20260622_114027", "custom", "rule_110201_count.txt");

const tecMeta = {
  "TEC-001": { mitre: "T1059.001", tecnica: "PowerShell", tactica: "Execution" },
  "TEC-002": { mitre: "T1059.003", tecnica: "Windows Command Shell", tactica: "Execution" },
  "TEC-003": { mitre: "T1053.005", tecnica: "Scheduled Task", tactica: "Execution / Persistence" },
  "TEC-004": { mitre: "T1547.001", tecnica: "Registry Run Key", tactica: "Persistence / Privilege Escalation" },
  "TEC-005": { mitre: "T1569.002", tecnica: "Service Execution", tactica: "Execution / Persistence" },
  "TEC-006": { mitre: "T1518.001", tecnica: "Security Software Discovery", tactica: "Discovery" },
  "TEC-007": { mitre: "T1486", tecnica: "Data Encrypted for Impact", tactica: "Impact" },
  "TEC-008": { mitre: "T1485", tecnica: "Data Destruction", tactica: "Impact" },
  "TEC-009": { mitre: "T1074.001 / T1560.001 / T1048.003", tecnica: "Staging / Archive / Controlled HTTP Upload", tactica: "Collection / Exfiltration" },
};

function rel(p) {
  return path.relative(root, p).replaceAll(path.sep, "\\");
}

function walk(dir) {
  if (!fs.existsSync(dir)) return [];
  let out = [];
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, ent.name);
    if (ent.isDirectory()) out = out.concat(walk(p));
    else out.push(p);
  }
  return out;
}

function readText(p) {
  return fs.existsSync(p) ? fs.readFileSync(p, "utf8").replace(/^\uFEFF/, "") : "";
}

function readJson(p) {
  return JSON.parse(readText(p));
}

function sha256(p) {
  const h = crypto.createHash("sha256");
  h.update(fs.readFileSync(p));
  return h.digest("hex");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === "\"") {
        if (text[i + 1] === "\"") {
          field += "\"";
          i++;
        } else {
          quoted = false;
        }
      } else {
        field += c;
      }
    } else if (c === "\"") {
      quoted = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c !== "\r") {
      field += c;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function csvObjects(p) {
  if (!fs.existsSync(p)) return [];
  const rows = parseCsv(readText(p));
  if (!rows.length) return [];
  const header = rows[0].map((h) => h.trim());
  return rows.slice(1).filter((r) => r.some((c) => c !== "")).map((r) => {
    const o = {};
    header.forEach((h, i) => {
      o[h] = r[i] ?? "";
    });
    return o;
  });
}

function csvEscape(v) {
  if (v === null || v === undefined) return "";
  const s = String(v);
  return /[",\r\n]/.test(s) ? `"${s.replaceAll("\"", "\"\"")}"` : s;
}

function writeCsv(name, rows, columns) {
  const p = path.join(normDir, name);
  const lines = [columns.map(csvEscape).join(",")];
  for (const row of rows) {
    lines.push(columns.map((c) => csvEscape(row[c])).join(","));
  }
  fs.writeFileSync(p, `${lines.join("\r\n")}\r\n`, "utf8");
  return p;
}

function toDate(value) {
  if (value === null || value === undefined || value === "") return null;
  let s = String(value).trim();
  const ps = /^\/Date\((\d+)\)\/$/.exec(s);
  if (ps) return new Date(Number(ps[1]));
  const n = Number(s.replace(",", "."));
  if (Number.isFinite(n)) {
    if (n > 1e12) return new Date(n);
    if (n > 1e9) return new Date(n * 1000);
  }
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function iso(value) {
  const d = value instanceof Date ? value : toDate(value);
  return d ? d.toISOString() : "";
}

function inWindow(d, start, end) {
  return d && d >= start && d <= end;
}

function trimEvidence(s, max = 900) {
  const one = String(s ?? "").replace(/\s+/g, " ").trim();
  return one.length > max ? `${one.slice(0, max)}...` : one;
}

function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function inferProfile(file, row) {
  const val = String(row.ArtifactProfile || row.Profile || "").toUpperCase();
  if (val.includes("P1")) return "P1";
  if (val.includes("P2")) return "P2";
  if (val.includes("P3")) return "P3";
  if (val.includes("P4")) return "P4";
  const name = path.basename(file).toUpperCase();
  if (name.startsWith("P1_EVENT")) return "P1";
  if (name.startsWith("FORENSIC_EVENT")) return "P2";
  if (name.startsWith("TEC")) return "P3";
  if (name.startsWith("CU")) return "P4";
  return "NO_CLASIFICADO";
}

function profileRank(profile) {
  return { P1: 4, P2: 3, P3: 2, P4: 1 }[profile] || 0;
}

function sourceBucket(row) {
  const eventId = String(row.EventID || "").trim();
  const signal = `${row.SignalType || ""} ${row.Source || ""} ${row.Channel || ""}`.toLowerCase();
  if (eventId === "4104" || signal.includes("4104")) return "PowerShell 4104";
  if (eventId === "7045" || signal.includes("7045")) return "System 7045";
  if (eventId === "1" || signal.includes("sysmon id 1")) return "Sysmon ID 1";
  if (eventId === "3" || signal.includes("network")) return "Sysmon ID 3";
  if (eventId === "11" || signal.includes("sysmon id 11")) return "Sysmon ID 11";
  if (["12", "13", "14"].includes(eventId) || signal.includes("registry")) return "Sysmon ID 12/13/14";
  if (eventId === "26") return "Sysmon ID 26";
  return "Otra fuente";
}

function parseWazuhSummary(text) {
  const getFirst = (re) => {
    const m = text.match(re);
    return m ? Number(m[1].replaceAll(".", "")) : 0;
  };
  return {
    archives: getFirst(/Total archives delta:\s*\r?\n\s*(\d+)/i),
    alerts: getFirst(/Total alerts delta:\s*\r?\n\s*(\d+)/i),
    ps4104: getFirst(/PowerShell 4104:\s*\r?\n\s*(\d+)/i),
    sysmon1: getFirst(/Sysmon Event ID 1:\s*\r?\n\s*(\d+)/i),
    sysmon11: getFirst(/Sysmon Event ID 11:\s*\r?\n\s*(\d+)/i),
    system7045: getFirst(/System 7045:\s*\r?\n\s*(\d+)/i),
  };
}

function parseRuleCounts(text) {
  const out = {};
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(/^\s*(\d+)\s+"id":"(\d+)"/);
    if (m) out[m[2]] = Number(m[1]);
  }
  return out;
}

const realSummary = readJson(realSummaryJson);
const fpSummary = readJson(fpSummaryJson);
const realRowsRaw = csvObjects(realSummaryCsv);
const fpRowsRaw = csvObjects(fpSummaryCsv);
const fpHitRowsRaw = csvObjects(fpHitsCsv);

const realStart = toDate(realSummary.Campaign.Start);
const realEnd = toDate(realSummary.Campaign.End);
const fpStart = toDate(fpSummary.StartTimeUtc);
const fpEnd = toDate(fpSummary.EndTimeUtc);

const realTechWindows = {};
for (const r of realRowsRaw) {
  realTechWindows[r.TEC_ID] = { start: toDate(r.Start), end: toDate(r.End) };
  if (tecMeta[r.TEC_ID]) {
    const nameTactic = String(r.NameTactic || "");
    const splitAtMitre = nameTactic.search(/\s\/\sT\d/);
    if (splitAtMitre > 0) {
      tecMeta[r.TEC_ID].tecnica = nameTactic.slice(0, splitAtMitre).trim();
      tecMeta[r.TEC_ID].mitre = nameTactic.slice(splitAtMitre + 3).trim();
    }
  }
}

const fpWindows = {};
for (const r of fpRowsRaw) fpWindows[r.FP_ID] = { start: toDate(r.StartTimeLocal), end: toDate(r.EndTimeLocal) };

const inputExts = new Set([".csv", ".json", ".jsonl", ".txt", ".log", ".xml", ".conf"]);
const sourceFiles = [
  baseXlsx,
  ...walk(validasDir).filter((p) => inputExts.has(path.extname(p).toLowerCase())),
  ...walk(wazuhDir).filter((p) => inputExts.has(path.extname(p).toLowerCase()) || path.extname(p).toLowerCase() === ".sha256"),
].filter((p, i, a) => fs.existsSync(p) && a.indexOf(p) === i);

const fuentes = sourceFiles.map((p) => {
  const st = fs.statSync(p);
  const rp = rel(p);
  let rol = "OTRO";
  let camp = "NO_CLASIFICADO";
  if (p === baseXlsx) {
    rol = "EXCEL_BASE_V4";
    camp = "BASE";
  } else if (rp.includes("10_WAZUH")) {
    rol = "WAZUH_EVIDENCE";
    camp = "WAZUH";
  } else if (rp.includes("Prueba_24_06_2026_FP")) {
    rol = rp.includes("summary") ? "FP_SUMMARY" : rp.includes("vr_hits") ? "FP_HITS" : "VR_EXPORT_FP_OR_MIXED";
    camp = "FP_26_06";
  } else if (rp.includes("Prueba_24_06_2026")) {
    rol = rp.includes("summary") ? "REAL_SUMMARY" : "VR_EXPORT_REAL";
    camp = "REAL_24_06";
  }
  return {
    ruta: rp,
    tipo: path.extname(p).toLowerCase().replace(".", "") || "file",
    tamano_bytes: st.size,
    fecha_modificacion: st.mtime.toISOString(),
    sha256: sha256(p),
    rol,
    campana_asignada: camp,
  };
});

const csvFiles = walk(validasDir).filter((p) => path.extname(p).toLowerCase() === ".csv" && !p.endsWith("_summary.csv") && !p.endsWith("_vr_hits.csv"));
const alertas = [];
const fileWindows = [];
const discrepancias = [];
const fpClientWindowRows = [];

for (const file of csvFiles) {
  const rows = csvObjects(file);
  let counts = { ruta: rel(file), filas: rows.length, real: 0, fp: 0, fuera: 0, sin_tiempo: 0, min: "", max: "" };
  let minDate = null;
  let maxDate = null;
  const isCanonicalReal = file.startsWith(realDir + path.sep);
  const isFpFolder = file.startsWith(fpDir + path.sep);
  for (const row of rows) {
    const d = toDate(row.DetectionTime || row._ts);
    if (!d) {
      counts.sin_tiempo++;
      continue;
    }
    if (!minDate || d < minDate) minDate = d;
    if (!maxDate || d > maxDate) maxDate = d;
    if (inWindow(d, realStart, realEnd)) {
      counts.real++;
      if (isCanonicalReal) {
        const profile = inferProfile(file, row);
        alertas.push({
          timestamp: d.toISOString(),
          profile,
          TEC: row.ID_Tecnica_Interna || row.TEC || "",
          MITRE: row.MITRE_ID || "",
          IOA_DetectionName: row.AlertTitle || row.AlertName || row.DetectionName || "",
          Artifact: path.basename(file).replace(/-\d{4}-\d{2}-\d{2}T.*$/, ""),
          host: row.Hostname || "",
          source: row.Source || row.Channel || "",
          raw_evidence: trimEvidence(row.Evidence || row.CU_Evidence || row.CommandLine || row.AlertDescription || ""),
          campana: "TEC_20260624",
          tipo_campana: "REAL",
          validacion: "INCLUIDO_REAL_CANONICO_POR_TIMESTAMP_Y_CARPETA",
          EventID: row.EventID || "",
          SignalType: row.SignalType || "",
          source_file: rel(file),
        });
      }
    } else if (inWindow(d, fpStart, fpEnd)) {
      counts.fp++;
      if (isFpFolder) {
        const profile = inferProfile(file, row);
        fpClientWindowRows.push({
          timestamp: d.toISOString(),
          profile,
          TEC: row.ID_Tecnica_Interna || row.TEC || "",
          DetectionName: row.AlertTitle || row.AlertName || "",
          source_file: rel(file),
        });
      }
    } else {
      counts.fuera++;
    }
  }
  counts.min = minDate ? minDate.toISOString() : "";
  counts.max = maxDate ? maxDate.toISOString() : "";
  fileWindows.push(counts);

  if (counts.fuera > 0 || counts.sin_tiempo > 0 || (isFpFolder && counts.real > 0)) {
    let tipo = "FECHAS_MEZCLADAS_O_FUERA_DE_VENTANA";
    let decision = "Filas fuera de ventana excluidas del resultado definitivo.";
    if (isFpFolder && counts.real > 0) decision = "Filas de ventana REAL en carpeta FP excluidas para evitar doble conteo.";
    if (counts.sin_tiempo > 0) decision += " Filas sin timestamp quedan NO_CLASIFICADO/REVISAR.";
    discrepancias.push({
      tipo,
      severidad: counts.sin_tiempo > 0 ? "WARN" : "INFO",
      ruta: rel(file),
      detalle: `filas=${counts.filas}; real=${counts.real}; fp=${counts.fp}; fuera=${counts.fuera}; sin_tiempo=${counts.sin_tiempo}; min=${counts.min}; max=${counts.max}`,
      decision,
    });
  }
}

const alertByTec = {};
const alertByProfile = { P1: 0, P2: 0, P3: 0, P4: 0, NO_CLASIFICADO: 0 };
const visibilityCounts = {};
for (const tec of Object.keys(tecMeta)) visibilityCounts[tec] = {};
for (const a of alertas) {
  if (!alertByTec[a.TEC]) alertByTec[a.TEC] = [];
  alertByTec[a.TEC].push(a);
  alertByProfile[a.profile] = (alertByProfile[a.profile] || 0) + 1;
  if (!visibilityCounts[a.TEC]) visibilityCounts[a.TEC] = {};
  const bucket = sourceBucket(a);
  visibilityCounts[a.TEC][bucket] = (visibilityCounts[a.TEC][bucket] || 0) + 1;
}

const tec009 = realSummary.TEC009 || {};
const tec009HasUpload = String(tec009.HttpStatus || tec009.SummaryObject?.HttpStatusCode || "") === "200"
  && String(tec009.UploadSucceeded || tec009.SummaryObject?.UploadSucceeded || "").toLowerCase() === "true"
  && Boolean(tec009.LocalSHA256 || tec009.SummaryObject?.LocalSHA256)
  && Boolean(tec009.SummaryObject?.ReceiverResponse);
const tec009ReceiverLogFound = sourceFiles.some((p) => path.basename(p).toLowerCase() === "receiver_log.jsonl");
if (tec009HasUpload && !tec009ReceiverLogFound) {
  discrepancias.push({
    tipo: "TEC009_EVIDENCIA_UPLOAD_PARCIAL",
    severidad: "WARN",
    ruta: rel(realSummaryJson),
    detalle: "HTTP 200, UploadSucceeded=True, ZIP/SHA256 y ReceiverResponse con bytes/sha256 presentes; receiver_log.jsonl y ZIP recibido no localizados en rutas revisadas.",
    decision: "No marcar como exfiltracion contextual completa; documentar como upload HTTP controlado con limitacion de receptor.",
  });
}

if (fpClientWindowRows.length > 0 && Number(fpSummary.VelociraptorHits?.Total || 0) === 0) {
  discrepancias.push({
    tipo: "FP_CLIENT_EVENT_EXPORTS_NO_RECONCILIADOS_CON_VR_HITS",
    severidad: "WARN",
    ruta: rel(fpDir),
    detalle: `${fpClientWindowRows.length} filas CLIENT_EVENT en ventana FP, pero vr_hits.csv y summary declaran TotalHits=0.`,
    decision: "No incluir como FP_HITS definitivos; conservar en discrepancias para revision metodologica.",
  });
}

const orderedTec = Object.keys(tecMeta);
const tecnicas = orderedTec.map((tec) => {
  const r = realRowsRaw.find((x) => x.TEC_ID === tec) || {};
  const alerts = alertByTec[tec] || [];
  const maxProfile = alerts.map((a) => a.profile).sort((a, b) => profileRank(b) - profileRank(a))[0] || "NO";
  const sys = num(r.SysmonRelevantEvents);
  const ps = num(r.PowerShell4104RelevantEvents);
  let obs = r.Comment || "";
  let limit = "";
  if (tec === "TEC-007") obs = "Simulacion controlada de cifrado para impacto; no ransomware real. " + obs;
  if (tec === "TEC-008") obs = "Borrado controlado de ficheros dummy; Sysmon ID 26 es evidencia forense, no alerta individual. " + obs;
  if (tec === "TEC-009") {
    obs = `${tec009HasUpload ? "HTTP 200/UploadSucceeded/ZIP/SHA256/bytes en ReceiverResponse." : "Upload HTTP no completamente documentado."} ${obs}`;
    limit = tec009ReceiverLogFound ? "" : "receiver_log.jsonl y ZIP recibido no localizados; no se afirma exfiltracion contextual completa sin matiz.";
  }
  return {
    TEC_ID: tec,
    MITRE: tecMeta[tec].mitre,
    tecnica: tecMeta[tec].tecnica,
    tactica: tecMeta[tec].tactica,
    runner: realSummary.Campaign.RunnerPath || "",
    inicio: iso(r.Start),
    fin: iso(r.End),
    duracion_seg: num(r.DurationSeconds),
    estado_ejecucion: r.Status || "NO_CLASIFICADO",
    Sysmon_events: sys,
    PowerShell_4104: ps,
    System_7045: tec === "TEC-005" ? (visibilityCounts[tec]["System 7045"] || 0) : 0,
    visibilidad_sistema: sys > 0 || ps > 0 || tec === "TEC-009" ? "SI" : "NO",
    visibilidad_velociraptor: alerts.length > 0 ? "SI" : "NO",
    deteccion_vr: alerts.length > 0 ? "SI_CLIENT_EVENT" : "NO",
    perfil_max_vr: maxProfile,
    alerta_externa_jsonl_discord: "NO_JSONL_DISCORD_VALIDO_LOCALIZADO",
    evidencia_principal: trimEvidence(r.FoundEvidence || ""),
    observaciones: trimEvidence(obs, 1200),
    limitaciones: limit,
  };
});

if (tec009HasUpload) visibilityCounts["TEC-009"]["receiver HTTP"] = 1;

const visSources = ["Sysmon ID 1", "Sysmon ID 3", "Sysmon ID 11", "Sysmon ID 12/13/14", "Sysmon ID 26", "PowerShell 4104", "System 7045", "receiver HTTP"];
const visibilidad = [];
for (const source of visSources) {
  const row = { fuente: source };
  for (const tec of orderedTec) row[tec] = visibilityCounts[tec]?.[source] || 0;
  row.observaciones = source === "Sysmon ID 26" ? "Evidencia forense; no alerta individual." : "";
  visibilidad.push(row);
}

const fpRunner = fpRowsRaw.map((r) => {
  const hits = fpHitRowsRaw.filter((h) => h.FP_ID === r.FP_ID);
  return {
    FP_ID: r.FP_ID,
    descripcion: r.Description,
    estado: r.Status,
    inicio: iso(r.StartTimeLocal),
    fin: iso(r.EndTimeLocal),
    duracion_seg: num(r.DurationSec),
    artefactos_generados: r.Artifacts || "",
    genero_alerta_vr: hits.length > 0 ? "SI" : "NO",
    perfil_generado: [...new Set(hits.map((h) => h.Profile).filter(Boolean))].join("; "),
    TEC_asociada: [...new Set(hits.map((h) => h.TEC).filter(Boolean))].join("; "),
    interpretacion: hits.length > 0 ? "REVISAR_FP" : "Sin hits en vr_hits.csv/summary.",
    observaciones: trimEvidence(r.Details || ""),
  };
});

let fpHits = fpHitRowsRaw;
if (fpHits.length === 0) {
  fpHits = [{
    RunId: fpSummary.RunId,
    EventTime: "",
    MatchedByTime: "",
    MatchedByRunId: "",
    FP_ID: "RESUMEN",
    Profile: "",
    TEC: "",
    DetectionName: "SIN_HITS",
    Severity: "",
    Confidence: "",
    Artifact: "",
    Source: "",
    PotentialFP: "NO",
    EvidenceExcerpt: "TotalHits=0 en summary.json y vr_hits.csv sin filas de datos.",
  }];
}

const wazuhBase = parseWazuhSummary(readText(wazuhBaseSummary));
const wazuhCustom = parseWazuhSummary(readText(wazuhCustomSummary));
const ruleCounts = { "110201": Number(readText(wazuh110201).trim() || 0), ...parseRuleCounts(readText(wazuhRuleCount)) };
for (const rid of ["110201", "110202", "110203", "110301", "110401", "110402", "110501"]) {
  if (!(rid in ruleCounts)) ruleCounts[rid] = 0;
}

const totalVrAlerts = alertas.length;
const comparacion = [
  {
    herramienta_campana: "Velociraptor custom REAL 24/06",
    eventos_crudos: "No aplica: se consolidan filas CLIENT_EVENT exportadas, no archivo bruto tipo archives.",
    alertas: totalVrAlerts,
    PowerShell_4104: tecnicas.reduce((s, r) => s + num(r.PowerShell_4104), 0),
    Sysmon_ID_1: visSources.includes("Sysmon ID 1") ? orderedTec.reduce((s, tec) => s + (visibilityCounts[tec]["Sysmon ID 1"] || 0), 0) : 0,
    Sysmon_ID_11: orderedTec.reduce((s, tec) => s + (visibilityCounts[tec]["Sysmon ID 11"] || 0), 0),
    System_7045: orderedTec.reduce((s, tec) => s + (visibilityCounts[tec]["System 7045"] || 0), 0),
    P1_P2_P3_P4: `P1=${alertByProfile.P1 || 0}; P2=${alertByProfile.P2 || 0}; P3=${alertByProfile.P3 || 0}; P4=${alertByProfile.P4 || 0}`,
    reglas_110xxx: "No aplica",
    TEC_009: tec009HasUpload ? "Upload controlado documentado por summary; receiver_log ausente." : "No cerrado",
    ruido: "Filas fuera de ventana excluidas; JSONL/Discord no usados como deteccion.",
    limitaciones: "CLIENT_EVENT detecta; SERVER_EVENT/JSONL no localizado para REAL.",
    lectura_metodologica: "Separar visibilidad sistema, deteccion CLIENT_EVENT y salida externa.",
  },
  {
    herramienta_campana: "Wazuh base",
    eventos_crudos: wazuhBase.archives,
    alertas: wazuhBase.alerts,
    PowerShell_4104: wazuhBase.ps4104,
    Sysmon_ID_1: wazuhBase.sysmon1,
    Sysmon_ID_11: wazuhBase.sysmon11,
    System_7045: wazuhBase.system7045,
    P1_P2_P3_P4: "No equivalente; sin reglas custom TFM.",
    reglas_110xxx: 0,
    TEC_009: "Sin regla custom 110201.",
    ruido: "58 alertas base sobre 902 archives.",
    limitaciones: "Visibilidad base sin IOA TFM custom.",
    lectura_metodologica: "Base aporta telemetria/alertas genericas, no deteccion TFM especifica.",
  },
  {
    herramienta_campana: "Wazuh custom",
    eventos_crudos: wazuhCustom.archives,
    alertas: wazuhCustom.alerts,
    PowerShell_4104: wazuhCustom.ps4104,
    Sysmon_ID_1: wazuhCustom.sysmon1,
    Sysmon_ID_11: wazuhCustom.sysmon11,
    System_7045: wazuhCustom.system7045,
    P1_P2_P3_P4: "Equivalencia por niveles/reglas 110xxx.",
    reglas_110xxx: Object.values(ruleCounts).reduce((a, b) => a + b, 0),
    TEC_009: "110201=0; gap TEC-009.",
    ruido: "110 alertas sobre 18.285 archives; 56 alertas TFM 110xxx.",
    limitaciones: "110201 no dispara; revisar forma real de upload.",
    lectura_metodologica: "Custom mejora cobertura, pero deja gap TEC-009.",
  },
];

const ruleDetails = {
  "110201": { severidad: 15, regla: "TFM P1 TEC-009 posible exfiltracion HTTP ZIP mediante PowerShell", tecnica: "TEC-009", mitre: "T1048.003 / T1560.001", limitacion: "0 alertas; gap TEC-009." },
  "110202": { severidad: 14, regla: "TFM P1 TEC-008 borrado destructivo mediante PowerShell", tecnica: "TEC-008", mitre: "T1485", limitacion: "" },
  "110203": { severidad: 15, regla: "TFM P1 TEC-007 comportamiento de cifrado mediante PowerShell", tecnica: "TEC-007", mitre: "T1486", limitacion: "" },
  "110301": { severidad: 10, regla: "TFM P2 PowerShell sospechoso detectado en ScriptBlock 4104", tecnica: "TEC-001", mitre: "T1059.001", limitacion: "" },
  "110401": { severidad: 8, regla: "TFM P3 posible persistencia Run Key mediante Sysmon", tecnica: "TEC-004", mitre: "T1547.001", limitacion: "0 alertas en resumen 110xxx." },
  "110402": { severidad: 8, regla: "TFM P3 instalacion de servicio Windows", tecnica: "TEC-005", mitre: "T1569.002", limitacion: "" },
  "110501": { severidad: 5, regla: "TFM P4 ejecucion de PowerShell observada por Sysmon", tecnica: "TEC-001/007", mitre: "T1059.001", limitacion: "Regla de baja prioridad; puede generar ruido." },
};
const wazuhDetalle = Object.keys(ruleDetails).map((id) => ({
  rule_id: id,
  severidad: ruleDetails[id].severidad,
  descripcion: ruleDetails[id].regla,
  tecnica: ruleDetails[id].tecnica,
  MITRE: ruleDetails[id].mitre,
  eventos_alertas: ruleCounts[id] || 0,
  limitaciones: ruleDetails[id].limitacion,
}));

const kpis = [
  { KPI: "TEC reales planificadas", valor: 9, detalle: "TEC-001 a TEC-009" },
  { KPI: "TEC reales ejecutadas", valor: realSummary.Campaign.TechniquesOK + realSummary.Campaign.TechniquesWARN + realSummary.Campaign.TechniquesFAIL, detalle: realSummary.Campaign.GlobalStatus },
  { KPI: "TEC OK", valor: realSummary.Campaign.TechniquesOK, detalle: "" },
  { KPI: "TEC WARN", valor: realSummary.Campaign.TechniquesWARN, detalle: "" },
  { KPI: "TEC FAIL", valor: realSummary.Campaign.TechniquesFAIL, detalle: "" },
  { KPI: "FP planificados", valor: 10, detalle: "FP-001 a FP-010" },
  { KPI: "FP OK", valor: fpSummary.TestCounts.OK, detalle: "" },
  { KPI: "FP WARN", valor: fpSummary.TestCounts.WARN, detalle: "" },
  { KPI: "FP FAIL", valor: fpSummary.TestCounts.FAIL, detalle: "" },
  { KPI: "FP SKIPPED", valor: fpSummary.TestCounts.SKIPPED, detalle: "" },
  { KPI: "Total alertas VR REAL incluidas", valor: totalVrAlerts, detalle: "Filas CLIENT_EVENT en carpeta real y ventana del runner." },
  { KPI: "Alertas VR P1", valor: alertByProfile.P1 || 0, detalle: "" },
  { KPI: "Alertas VR P2", valor: alertByProfile.P2 || 0, detalle: "" },
  { KPI: "Alertas VR P3", valor: alertByProfile.P3 || 0, detalle: "" },
  { KPI: "Alertas VR P4", valor: alertByProfile.P4 || 0, detalle: "" },
  { KPI: "Hits FP", valor: fpSummary.VelociraptorHits.Total, detalle: "summary.json / vr_hits.csv" },
  { KPI: "Wazuh base alerts", valor: wazuhBase.alerts, detalle: "" },
  { KPI: "Wazuh custom alerts", valor: wazuhCustom.alerts, detalle: "" },
  { KPI: "Reglas Wazuh 110xxx activadas", valor: Object.entries(ruleCounts).filter(([id, c]) => id !== "110201" && c > 0).length, detalle: "110201 no activada." },
  { KPI: "Gaps", valor: 1, detalle: "110201=0 / TEC-009" },
];

const chartEventsByTec = tecnicas.map((r) => ({ TEC_ID: r.TEC_ID, Sysmon_events: r.Sysmon_events, PowerShell_4104: r.PowerShell_4104, System_7045: r.System_7045 }));
const chartAlertsByProfile = ["P1", "P2", "P3", "P4"].map((p) => ({ profile: p, alertas: alertByProfile[p] || 0 }));
const chartFp = ["OK", "WARN", "FAIL", "SKIPPED", "HITS"].map((s) => ({ estado: s, valor: s === "HITS" ? fpSummary.VelociraptorHits.Total : fpSummary.TestCounts[s] || 0 }));
const chartWazuh = [
  { metrica: "archives", base: wazuhBase.archives, custom: wazuhCustom.archives },
  { metrica: "alerts", base: wazuhBase.alerts, custom: wazuhCustom.alerts },
  { metrica: "PowerShell 4104", base: wazuhBase.ps4104, custom: wazuhCustom.ps4104 },
  { metrica: "Sysmon ID 1", base: wazuhBase.sysmon1, custom: wazuhCustom.sysmon1 },
  { metrica: "Sysmon ID 11", base: wazuhBase.sysmon11, custom: wazuhCustom.sysmon11 },
  { metrica: "System 7045", base: wazuhBase.system7045, custom: wazuhCustom.system7045 },
];
const chartWazuhRules = Object.keys(ruleDetails).map((rid) => ({ rule_id: rid, alertas: ruleCounts[rid] || 0 }));
const heatmapVisibility = visibilidad.map((r) => ({ ...r }));
const heatmapCapacity = [
  { capacidad: "PowerShell 4104", VR_custom: 2, Wazuh_base: wazuhBase.ps4104 > 0 ? 1 : 0, Wazuh_custom: wazuhCustom.ps4104 > 0 ? 2 : 0 },
  { capacidad: "Sysmon ID 1", VR_custom: orderedTec.some((t) => visibilityCounts[t]["Sysmon ID 1"]) ? 2 : 0, Wazuh_base: wazuhBase.sysmon1 > 0 ? 1 : 0, Wazuh_custom: wazuhCustom.sysmon1 > 0 ? 2 : 0 },
  { capacidad: "Sysmon ID 11", VR_custom: orderedTec.some((t) => visibilityCounts[t]["Sysmon ID 11"]) ? 1 : 0, Wazuh_base: wazuhBase.sysmon11 > 0 ? 1 : 0, Wazuh_custom: wazuhCustom.sysmon11 > 0 ? 1 : 0 },
  { capacidad: "System 7045", VR_custom: orderedTec.some((t) => visibilityCounts[t]["System 7045"]) ? 2 : 0, Wazuh_base: wazuhBase.system7045 > 0 ? 1 : 0, Wazuh_custom: wazuhCustom.system7045 > 0 ? 2 : 0 },
  { capacidad: "TEC-009", VR_custom: tec009HasUpload ? 1 : 0, Wazuh_base: 0, Wazuh_custom: ruleCounts["110201"] > 0 ? 2 : 0 },
  { capacidad: "Salida externa", VR_custom: 0, Wazuh_base: 1, Wazuh_custom: 1 },
  { capacidad: "Ruido/FP", VR_custom: fpSummary.VelociraptorHits.Total === 0 ? 2 : 1, Wazuh_base: 1, Wazuh_custom: 1 },
];

const qualityChecks = [
  { check: "TEC-001 a TEC-009 presentes", status: tecnicas.length === 9 && tecnicas.every((r) => r.estado_ejecucion) ? "OK" : "FAIL", detail: tecnicas.map((r) => r.TEC_ID).join(", ") },
  { check: "FP-001 a FP-010 presentes o justificados", status: fpRunner.length === 10 ? "OK" : "FAIL", detail: fpRunner.map((r) => r.FP_ID).join(", ") },
  { check: "VR real separado de FP", status: "OK", detail: "Separacion aplicada por ventana/carpeta; CSV mixtos quedan en DISCREPANCIAS y no se agregan como definitivos." },
  { check: "TEC-009 no inflado", status: tec009ReceiverLogFound ? "OK" : "WARN", detail: "Upload documentado en summary; receiver_log/ZIP recibido no localizados." },
  { check: "Wazuh 110201=gap", status: ruleCounts["110201"] === 0 ? "OK" : "FAIL", detail: `110201=${ruleCounts["110201"]}` },
  { check: "FP hits", status: fpSummary.VelociraptorHits.Total === 0 ? "OK" : "WARN", detail: `TotalHits=${fpSummary.VelociraptorHits.Total}` },
  { check: "Fuentes con SHA-256", status: fuentes.every((f) => f.sha256) ? "OK" : "FAIL", detail: `${fuentes.length} fuentes hasheadas` },
  { check: "Artifact-tool unavailable", status: "WARN", detail: "Se usa generacion local con Node/Excel COM por ausencia de @oai/artifact-tool en el entorno." },
];

writeCsv("tecnicas_real_vr.csv", tecnicas, ["TEC_ID", "MITRE", "tecnica", "tactica", "runner", "inicio", "fin", "duracion_seg", "estado_ejecucion", "Sysmon_events", "PowerShell_4104", "System_7045", "visibilidad_sistema", "visibilidad_velociraptor", "deteccion_vr", "perfil_max_vr", "alerta_externa_jsonl_discord", "evidencia_principal", "observaciones", "limitaciones"]);
writeCsv("visibilidad_sistema.csv", visibilidad, ["fuente", ...orderedTec, "observaciones"]);
writeCsv("alertas_vr.csv", alertas, ["timestamp", "profile", "TEC", "MITRE", "IOA_DetectionName", "Artifact", "host", "source", "raw_evidence", "campana", "tipo_campana", "validacion", "EventID", "SignalType", "source_file"]);
writeCsv("fp_runner.csv", fpRunner, ["FP_ID", "descripcion", "estado", "inicio", "fin", "duracion_seg", "artefactos_generados", "genero_alerta_vr", "perfil_generado", "TEC_asociada", "interpretacion", "observaciones"]);
writeCsv("fp_hits.csv", fpHits, ["RunId", "EventTime", "MatchedByTime", "MatchedByRunId", "FP_ID", "Profile", "TEC", "DetectionName", "Severity", "Confidence", "Artifact", "Source", "PotentialFP", "EvidenceExcerpt"]);
writeCsv("comparacion_vr_wazuh.csv", comparacion, ["herramienta_campana", "eventos_crudos", "alertas", "PowerShell_4104", "Sysmon_ID_1", "Sysmon_ID_11", "System_7045", "P1_P2_P3_P4", "reglas_110xxx", "TEC_009", "ruido", "limitaciones", "lectura_metodologica"]);
writeCsv("wazuh_detalle.csv", wazuhDetalle, ["rule_id", "severidad", "descripcion", "tecnica", "MITRE", "eventos_alertas", "limitaciones"]);
writeCsv("discrepancias.csv", discrepancias, ["tipo", "severidad", "ruta", "detalle", "decision"]);
writeCsv("fuentes.csv", fuentes, ["ruta", "tipo", "tamano_bytes", "fecha_modificacion", "sha256", "rol", "campana_asignada"]);
writeCsv("kpis.csv", kpis, ["KPI", "valor", "detalle"]);
writeCsv("chart_events_by_tec.csv", chartEventsByTec, ["TEC_ID", "Sysmon_events", "PowerShell_4104", "System_7045"]);
writeCsv("chart_alerts_by_profile.csv", chartAlertsByProfile, ["profile", "alertas"]);
writeCsv("chart_fp.csv", chartFp, ["estado", "valor"]);
writeCsv("chart_wazuh.csv", chartWazuh, ["metrica", "base", "custom"]);
writeCsv("chart_wazuh_rules.csv", chartWazuhRules, ["rule_id", "alertas"]);
writeCsv("heatmap_visibility.csv", heatmapVisibility, ["fuente", ...orderedTec, "observaciones"]);
writeCsv("heatmap_capacity.csv", heatmapCapacity, ["capacidad", "VR_custom", "Wazuh_base", "Wazuh_custom"]);
writeCsv("file_window_audit.csv", fileWindows, ["ruta", "filas", "real", "fp", "fuera", "sin_tiempo", "min", "max"]);
writeCsv("fp_client_event_window_rows_revisar.csv", fpClientWindowRows, ["timestamp", "profile", "TEC", "DetectionName", "source_file"]);

const dataQuality = {
  generated_at: new Date().toISOString(),
  inputs: {
    real_summary: rel(realSummaryJson),
    fp_summary: rel(fpSummaryJson),
    fp_hits: rel(fpHitsCsv),
    wazuh_base_summary: rel(wazuhBaseSummary),
    wazuh_custom_summary: rel(wazuhCustomSummary),
    base_xlsx: rel(baseXlsx),
  },
  windows: {
    real_start_utc: realStart.toISOString(),
    real_end_utc: realEnd.toISOString(),
    fp_start_utc: fpStart.toISOString(),
    fp_end_utc: fpEnd.toISOString(),
  },
  metrics: {
    tecnicas: tecnicas.length,
    fp_tests: fpRunner.length,
    alertas_vr_real_incluidas: totalVrAlerts,
    alertas_vr_por_perfil: alertByProfile,
    fp_hits: fpSummary.VelociraptorHits.Total,
    fp_client_event_rows_revisar: fpClientWindowRows.length,
    wazuh_base: wazuhBase,
    wazuh_custom: wazuhCustom,
    wazuh_rule_counts: ruleCounts,
    discrepancias: discrepancias.length,
    fuentes: fuentes.length,
  },
  quality_checks: qualityChecks,
  tec009: {
    upload_summary_evidence: tec009HasUpload,
    receiver_log_found: tec009ReceiverLogFound,
    http_status: String(tec009.HttpStatus || tec009.SummaryObject?.HttpStatusCode || ""),
    upload_succeeded: String(tec009.UploadSucceeded || tec009.SummaryObject?.UploadSucceeded || ""),
    local_sha256: tec009.LocalSHA256 || tec009.SummaryObject?.LocalSHA256 || "",
    receiver_response: tec009.SummaryObject?.ReceiverResponse || "",
  },
};

fs.writeFileSync(path.join(outDir, "TFM_VISIBILIDAD_DETECCION_FP_WAZUH_FINAL_26062026_DATA_QUALITY.json"), JSON.stringify(dataQuality, null, 2), "utf8");
fs.writeFileSync(path.join(normDir, "dataset_summary.json"), JSON.stringify({
  kpis,
  qualityChecks,
  dataQuality,
}, null, 2), "utf8");

console.log(JSON.stringify({
  ok: true,
  totalVrAlerts,
  alertByProfile,
  fpHits: fpSummary.VelociraptorHits.Total,
  discrepancias: discrepancias.length,
  fuentes: fuentes.length,
  tec009ReceiverLogFound,
}, null, 2));
