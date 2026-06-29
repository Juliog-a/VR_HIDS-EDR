import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const root = process.cwd();
const generatedAt = new Date().toISOString();

const p = {
  campaignRoot: path.join(root, '06_CONTROLLED_RERUN', 'OUTPUT', 'BENCHMARK', 'BENCH_20260619_FINAL01'),
  traceDir: path.join(root, '04_EVIDENCE', 'Excel_benchmark_26062026'),
  normalizedDir: path.join(root, '04_EVIDENCE', 'Excel_benchmark_26062026', 'normalized'),
  finalDir: path.join(root, '08_MEMORIA', 'Excel_benchmark'),
  canonicalXlsx: path.join(root, '04_EVIDENCE', 'ENTREGA_MEMORIA_EXCELES_REGENERADOS', '02_BENCHMARK_RENDIMIENTO_VALIDACION_DATOS_REGENERADO.xlsx'),
  historicalXlsx: path.join(root, '04_EVIDENCE', 'Analisis_Benchmark.xlsx'),
  previousDefinitiveDir: path.join(root, '04_EVIDENCE', 'ENTREGA_MEMORIA_EXCELES_DEFINITIVOS'),
  previousReviewedDir: path.join(root, '04_EVIDENCE', 'ENTREGA_MEMORIA_EXCELES_DEFINITIVOS_REVIEWED'),
  contextDir: path.join(root, '00_CONTEXT'),
};

fs.mkdirSync(p.traceDir, { recursive: true });
fs.mkdirSync(p.normalizedDir, { recursive: true });
fs.mkdirSync(p.finalDir, { recursive: true });

const scenarioOrder = ['BASELINE_NO_VR', 'VR_IDLE', 'VR_TEC_RUNNER'];
const expectedScenarios = new Set(scenarioOrder);
const expectedReps = [1, 2, 3];

function stripBom(s) {
  return s.replace(/^\uFEFF/, '');
}

function readText(file) {
  return stripBom(fs.readFileSync(file, 'utf8'));
}

function readJson(file) {
  return JSON.parse(readText(file));
}

function sha256(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function statOrNull(file) {
  try {
    return fs.statSync(file);
  } catch {
    return null;
  }
}

function walk(dir, out = []) {
  if (!fs.existsSync(dir)) return out;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, out);
    else out.push(full);
  }
  return out;
}

function parseCsv(text) {
  text = stripBom(text);
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const c = text[i];
    const n = text[i + 1];
    if (inQuotes) {
      if (c === '"' && n === '"') {
        field += '"';
        i += 1;
      } else if (c === '"') {
        inQuotes = false;
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field);
      field = '';
    } else if (c === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else if (c !== '\r') {
      field += c;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  if (rows.length === 0) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1)
    .filter((r) => r.some((v) => v !== ''))
    .map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ''])));
}

function csvEscape(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeCsv(file, rows, headers) {
  const lines = [headers.map(csvEscape).join(',')];
  for (const row of rows) {
    lines.push(headers.map((h) => csvEscape(row[h])).join(','));
  }
  fs.writeFileSync(file, `${lines.join('\r\n')}\r\n`, 'utf8');
}

function num(v, fallback = null) {
  if (v === null || v === undefined || v === '') return fallback;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : fallback;
}

function bool(v) {
  if (typeof v === 'boolean') return v;
  if (v === null || v === undefined) return false;
  return String(v).trim().toLowerCase() === 'true';
}

function mean(values) {
  const xs = values.map((v) => num(v)).filter((v) => v !== null);
  if (xs.length === 0) return null;
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

function max(values) {
  const xs = values.map((v) => num(v)).filter((v) => v !== null);
  if (xs.length === 0) return null;
  return Math.max(...xs);
}

function round(v, places = 2) {
  if (v === null || v === undefined || !Number.isFinite(Number(v))) return '';
  const f = 10 ** places;
  return Math.round(Number(v) * f) / f;
}

function repFromPath(file) {
  const parts = file.split(/[\\/]/);
  const rep = parts.find((part) => /^REP_\d+$/i.test(part));
  return rep ? Number(rep.replace(/\D/g, '')) : null;
}

function roleForFile(file) {
  const base = path.basename(file).toLowerCase();
  if (base === 'summary.json') return 'run_summary';
  if (base === 'samples.csv') return 'raw_samples';
  if (base === 'process_samples.csv') return 'process_samples';
  if (base === 'environment.txt') return 'environment';
  if (base.includes('process_list')) return 'process_list';
  if (base.endsWith('.hashes.txt') || base.endsWith('.sha256')) return 'hash_inventory';
  if (base.endsWith('.xlsx')) return 'excel_reference';
  if (base.endsWith('.md')) return 'context';
  return 'support';
}

function sourceRow(file, role, campaign, obs = '') {
  const st = statOrNull(file);
  return {
    Ruta: path.resolve(file),
    Tipo: path.extname(file).replace('.', '').toLowerCase() || 'file',
    TamanoBytes: st ? st.size : '',
    FechaModificacion: st ? st.mtime.toISOString() : '',
    'SHA-256': st ? sha256(file) : '',
    Rol: role ?? roleForFile(file),
    Campana: campaign ?? '',
    Observaciones: obs,
  };
}

function sourceHash(file) {
  const st = statOrNull(file);
  return st ? sha256(file) : '';
}

const summaryFiles = walk(p.campaignRoot).filter((file) => path.basename(file).toLowerCase() === 'summary.json');
const runs = [];
const rawSamples = [];
const processSamples = [];
const sources = [];
const discrepancies = [];

sources.push(sourceRow(p.canonicalXlsx, 'canonical_regenerated_xlsx', 'BENCH_20260619_FINAL01', 'Fuente canonica solicitada para benchmark definitivo.'));
sources.push(sourceRow(p.historicalXlsx, 'historical_reference_xlsx', 'HISTORICO', 'Referencia historica/formato; no usada como resultado definitivo.'));

for (const dir of [p.previousDefinitiveDir, p.previousReviewedDir]) {
  for (const file of walk(dir).filter((f) => f.toLowerCase().endsWith('.xlsx'))) {
    sources.push(sourceRow(file, 'previous_excel_reference', 'REFERENCIA_FORMATO', 'Usado solo como referencia visual/estructural.'));
  }
}

const contextFiles = [
  'PROJECT_STATE.md',
  'CURRENT_TASK.md',
  'DECISIONS.md',
  'ARTIFACT_INDEX.md',
  'TEST_MATRIX.md',
  'EVIDENCE_INDEX.md',
  path.join('SESSION_LOGS', 'LAST_SESSION.md'),
  'CODEX_RULES.md',
  'VM_ACTIVE_PATHS.md',
];
for (const rel of contextFiles) {
  const file = path.join(p.contextDir, rel);
  if (fs.existsSync(file)) {
    sources.push(sourceRow(file, 'context_project_state', 'CONTEXTO', 'Contexto obligatorio AGENTS.md.'));
  }
}

for (const summaryFile of summaryFiles) {
  const runDir = path.dirname(summaryFile);
  const summary = readJson(summaryFile);
  const scenario = summary.Scenario ?? path.basename(path.dirname(runDir));
  const repetition = repFromPath(summaryFile);
  const runId = summary.RunId ?? path.basename(runDir);
  const samplesFile = path.join(runDir, 'samples.csv');
  const processFile = path.join(runDir, 'process_samples.csv');
  const raw = fs.existsSync(samplesFile) ? parseCsv(readText(samplesFile)) : [];
  const proc = fs.existsSync(processFile) ? parseCsv(readText(processFile)) : [];

  for (const file of fs.readdirSync(runDir).map((f) => path.join(runDir, f))) {
    if (statOrNull(file)?.isFile()) {
      sources.push(sourceRow(file, roleForFile(file), 'BENCH_20260619_FINAL01', `RunId=${runId}; Scenario=${scenario}; REP_${String(repetition).padStart(2, '0')}`));
    }
  }

  for (const row of raw) {
    rawSamples.push({
      Scenario: scenario,
      Repetition: repetition,
      SourceFile: path.resolve(samplesFile),
      ...row,
    });
  }
  for (const row of proc) {
    processSamples.push({
      Scenario: scenario,
      Repetition: repetition,
      SourceFile: path.resolve(processFile),
      ...row,
    });
  }

  const clientRows = proc.filter((r) => String(r.Role).toUpperCase() === 'CLIENT_SERVICE').length;
  const serverGuiRows = proc.filter((r) => String(r.Role).toUpperCase() === 'SERVER_GUI').length;
  const notepadRows = proc.filter((r) => {
    const joined = [r.ProcessName, r.ExecutablePath, r.CommandLine].join(' ').toLowerCase();
    return joined.includes('notepad.exe');
  }).length;
  const serverGuiExcluded = bool(summary.ServerGuiExcludedFromClientMetrics) && bool(summary.IncludeServerGuiInTotal) === false;
  const isBaseline = scenario === 'BASELINE_NO_VR';
  const scenarioValid = String(summary.ScenarioValidity).toUpperCase() === 'VALID';
  const serviceOk = isBaseline
    ? String(summary.ClientServiceStatusStart).toLowerCase() === 'stopped' && String(summary.ClientServiceStatusEnd).toLowerCase() === 'stopped'
    : String(summary.ClientServiceStatusStart).toLowerCase() === 'running' && String(summary.ClientServiceStatusEnd).toLowerCase() === 'running';
  const validForClientBenchmark = scenarioValid
    && serviceOk
    && serverGuiExcluded
    && notepadRows === 0
    && (isBaseline
      ? clientRows === 0 && num(summary.VR_Client_ProcessCountMax, 0) === 0
      : clientRows > 0 && num(summary.VR_Client_ProcessCountMax, 0) >= 1);

  const warningList = Array.isArray(summary.Warnings) ? summary.Warnings.slice() : [];
  if (bool(summary.RunnerStillRunningAtEnd)) {
    warningList.push('RunnerStillRunningAtEnd=True tratado como WARN no bloqueante.');
  }

  runs.push({
    RunId: runId,
    Scenario: scenario,
    Repetition: repetition,
    StartTime: summary.StartTime ?? '',
    EndTime: summary.EndTime ?? '',
    DurationSeconds: round(num(summary.ActualDurationSec, num(summary.DurationSec, 0))),
    ScenarioValidity: summary.ScenarioValidity ?? '',
    ValidForClientBenchmark: validForClientBenchmark,
    ClientServiceStatusStart: summary.ClientServiceStatusStart ?? '',
    ClientServiceStatusEnd: summary.ClientServiceStatusEnd ?? '',
    ClientRows: clientRows,
    RunnerObservedSamples: num(summary.RunnerObservedSamples, 0),
    VR_Client_ProcessCountMax: num(summary.VR_Client_ProcessCountMax, 0),
    Avg_VR_Client_CPU_Percent: round(num(summary.Avg_VR_Client_CPU_Percent, 0)),
    Max_VR_Client_CPU_Percent: round(num(summary.Max_VR_Client_CPU_Percent, 0)),
    P95_VR_Client_CPU_Percent: round(num(summary.P95_VR_Client_CPU_Percent, 0)),
    Avg_VR_Client_RAM_MB: round(num(summary.Avg_VR_Client_RAM_MB, 0)),
    Max_VR_Client_RAM_MB: round(num(summary.Max_VR_Client_RAM_MB, 0)),
    P95_VR_Client_RAM_MB: round(num(summary.P95_VR_Client_RAM_MB, 0)),
    Warnings: [...new Set(warningList)].join(' | '),
    SourceFile: path.resolve(summaryFile),
    'SHA-256': sourceHash(summaryFile),
    SchemaVersion: summary.SchemaVersion ?? '',
    Samples: num(summary.Samples, raw.length),
    RawSamplesRows: raw.length,
    ProcessSampleRows: proc.length,
    ServerGuiRows: serverGuiRows,
    ServerGuiExcludedFromClientMetrics: bool(summary.ServerGuiExcludedFromClientMetrics),
    IncludeServerGuiInTotal: bool(summary.IncludeServerGuiInTotal),
    ServerGuiDetected: bool(summary.ServerGuiDetected),
    NotepadRows: notepadRows,
    RunnerStillRunningAtEnd: bool(summary.RunnerStillRunningAtEnd),
    RunnerStarted: bool(summary.RunnerStarted),
  });
}

runs.sort((a, b) => {
  const sa = scenarioOrder.indexOf(a.Scenario);
  const sb = scenarioOrder.indexOf(b.Scenario);
  if (sa !== sb) return sa - sb;
  return a.Repetition - b.Repetition;
});

const byScenario = new Map();
for (const scenario of scenarioOrder) byScenario.set(scenario, []);
for (const run of runs) {
  if (!byScenario.has(run.Scenario)) byScenario.set(run.Scenario, []);
  byScenario.get(run.Scenario).push(run);
}

const scenarioSummary = [];
for (const scenario of scenarioOrder) {
  const rs = byScenario.get(scenario) ?? [];
  const valid = rs.filter((r) => r.ScenarioValidity === 'VALID' && r.ValidForClientBenchmark === true);
  const warnings = rs.flatMap((r) => r.Warnings ? [r.Warnings] : []);
  let interpretation = '';
  if (scenario === 'BASELINE_NO_VR') {
    interpretation = 'Referencia de laboratorio con servicio Velociraptor detenido; CPU/RAM del cliente igual a 0 por diseno.';
  } else if (scenario === 'VR_IDLE') {
    interpretation = 'Cliente Velociraptor activo en reposo; coste operativo observado bajo en CPU y estable en RAM.';
  } else if (scenario === 'VR_TEC_RUNNER') {
    interpretation = 'Cliente bajo carga controlada de runner TEC; RunnerStillRunningAtEnd se conserva como WARN no bloqueante.';
  }
  scenarioSummary.push({
    Scenario: scenario,
    RepetitionsExpected: 3,
    RepetitionsFound: rs.length,
    ValidRuns: valid.length,
    AvgDurationSeconds: round(mean(valid.map((r) => r.DurationSeconds))),
    AvgCPUPercent: round(mean(valid.map((r) => r.Avg_VR_Client_CPU_Percent))),
    MaxCPUPercent: round(max(valid.map((r) => r.Max_VR_Client_CPU_Percent))),
    AvgRAMMB: round(mean(valid.map((r) => r.Avg_VR_Client_RAM_MB))),
    MaxRAMMB: round(max(valid.map((r) => r.Max_VR_Client_RAM_MB))),
    RunnerObservedSamples: valid.reduce((acc, r) => acc + num(r.RunnerObservedSamples, 0), 0),
    Warnings: warnings.length ? [...new Set(warnings)].join(' | ') : '',
    Interpretation: interpretation,
  });
}

const countsByScenario = Object.fromEntries([...byScenario.entries()].map(([k, v]) => [k, v.length]));
const validRuns = runs.filter((r) => r.ScenarioValidity === 'VALID' && r.ValidForClientBenchmark === true);
const duplicateRunIds = Object.entries(runs.reduce((acc, r) => {
  acc[r.RunId] = (acc[r.RunId] ?? 0) + 1;
  return acc;
}, {})).filter(([, count]) => count > 1);
const missing = [];
for (const scenario of scenarioOrder) {
  const repsFound = new Set((byScenario.get(scenario) ?? []).map((r) => r.Repetition));
  for (const rep of expectedReps) {
    if (!repsFound.has(rep)) missing.push({ Scenario: scenario, Repetition: rep });
  }
}

const checks = [
  {
    Check: 'Fuente canonica XLSX disponible',
    Status: fs.existsSync(p.canonicalXlsx) ? 'OK' : 'FAIL',
    Evidence: p.canonicalXlsx,
    Severity: fs.existsSync(p.canonicalXlsx) ? 'INFO' : 'CRITICA',
    Decision: 'Debe existir; no se usa Analisis_Benchmark.xlsx como resultado definitivo.',
  },
  {
    Check: 'SchemaVersion 1.1 presente',
    Status: runs.length > 0 && runs.every((r) => r.SchemaVersion === '1.1') ? 'OK' : 'FAIL',
    Evidence: [...new Set(runs.map((r) => r.SchemaVersion))].join(', '),
    Severity: 'CRITICA',
    Decision: 'Todos los runs validos deben estar en schema 1.1.',
  },
  {
    Check: 'Exactamente 3 escenarios',
    Status: new Set(runs.map((r) => r.Scenario)).size === 3 && runs.every((r) => expectedScenarios.has(r.Scenario)) ? 'OK' : 'FAIL',
    Evidence: JSON.stringify(countsByScenario),
    Severity: 'CRITICA',
    Decision: 'Escenarios esperados: BASELINE_NO_VR, VR_IDLE, VR_TEC_RUNNER.',
  },
  {
    Check: '3 repeticiones por escenario',
    Status: scenarioOrder.every((s) => (byScenario.get(s) ?? []).length === 3) ? 'OK' : 'FAIL',
    Evidence: JSON.stringify(countsByScenario),
    Severity: 'CRITICA',
    Decision: 'Cada escenario debe tener REP_01, REP_02 y REP_03.',
  },
  {
    Check: '9/9 runs validos',
    Status: runs.length === 9 && validRuns.length === 9 ? 'OK' : 'FAIL',
    Evidence: `${validRuns.length}/${runs.length}`,
    Severity: 'CRITICA',
    Decision: 'No se genera Excel final si faltan runs validos.',
  },
  {
    Check: 'Muestras CPU suficientes',
    Status: runs.length > 0 && runs.every((r) => num(r.Samples, 0) > 0 && num(r.RawSamplesRows, 0) > 0) ? 'OK' : 'FAIL',
    Evidence: `samples.csv combinados=${rawSamples.length}`,
    Severity: 'CRITICA',
    Decision: 'CPU/RAM se interpretan como metricas de laboratorio, no universales.',
  },
  {
    Check: 'Muestras RAM suficientes',
    Status: runs.length > 0 && runs.every((r) => num(r.Samples, 0) > 0 && num(r.RawSamplesRows, 0) > 0) ? 'OK' : 'FAIL',
    Evidence: `samples.csv combinados=${rawSamples.length}`,
    Severity: 'CRITICA',
    Decision: 'Se exige muestra observada por run.',
  },
  {
    Check: 'BASELINE_NO_VR presente',
    Status: (byScenario.get('BASELINE_NO_VR') ?? []).length === 3 ? 'OK' : 'FAIL',
    Evidence: `${(byScenario.get('BASELINE_NO_VR') ?? []).length} runs`,
    Severity: 'CRITICA',
    Decision: 'Referencia sin cliente VR activo.',
  },
  {
    Check: 'VR_IDLE presente',
    Status: (byScenario.get('VR_IDLE') ?? []).length === 3 ? 'OK' : 'FAIL',
    Evidence: `${(byScenario.get('VR_IDLE') ?? []).length} runs`,
    Severity: 'CRITICA',
    Decision: 'Cliente VR activo en reposo.',
  },
  {
    Check: 'VR_TEC_RUNNER presente',
    Status: (byScenario.get('VR_TEC_RUNNER') ?? []).length === 3 ? 'OK' : 'FAIL',
    Evidence: `${(byScenario.get('VR_TEC_RUNNER') ?? []).length} runs`,
    Severity: 'CRITICA',
    Decision: 'Carga controlada de runner TEC.',
  },
  {
    Check: 'SERVER_GUI excluido del calculo principal',
    Status: runs.every((r) => r.IncludeServerGuiInTotal === false && r.ServerGuiExcludedFromClientMetrics === true) ? 'OK' : 'FAIL',
    Evidence: `SERVER_GUI rows en process_samples=${runs.reduce((acc, r) => acc + num(r.ServerGuiRows, 0), 0)}`,
    Severity: 'CRITICA',
    Decision: 'SERVER_GUI puede estar observado, pero no suma en metricas de cliente.',
  },
  {
    Check: 'notepad.exe runner = 0',
    Status: runs.every((r) => num(r.NotepadRows, 0) === 0) ? 'OK' : 'FAIL',
    Evidence: `notepad rows=${runs.reduce((acc, r) => acc + num(r.NotepadRows, 0), 0)}`,
    Severity: 'CRITICA',
    Decision: 'Evita contaminacion del runner por notepad.exe.',
  },
  {
    Check: 'RunnerStillRunningAtEnd evaluado',
    Status: runs.every((r) => Object.prototype.hasOwnProperty.call(r, 'RunnerStillRunningAtEnd')) ? 'WARN' : 'FAIL',
    Evidence: `true=${runs.filter((r) => r.RunnerStillRunningAtEnd === true).length}; false=${runs.filter((r) => r.RunnerStillRunningAtEnd === false).length}`,
    Severity: 'NO_BLOQUEANTE',
    Decision: 'Si aparece en VR_TEC_RUNNER se trata como WARN no bloqueante.',
  },
  {
    Check: 'ClientRows > 0 cuando aplica',
    Status: runs.every((r) => r.Scenario === 'BASELINE_NO_VR' ? num(r.ClientRows, 0) === 0 : num(r.ClientRows, 0) > 0) ? 'OK' : 'FAIL',
    Evidence: runs.map((r) => `${r.RunId}:${r.ClientRows}`).join('; '),
    Severity: 'CRITICA',
    Decision: 'BASELINE debe no observar cliente; VR_IDLE/VR_TEC_RUNNER deben observarlo.',
  },
  {
    Check: 'Duplicados e inconsistencias',
    Status: duplicateRunIds.length === 0 && missing.length === 0 ? 'OK' : 'FAIL',
    Evidence: `duplicados=${duplicateRunIds.length}; missing=${missing.length}`,
    Severity: 'CRITICA',
    Decision: 'Duplicados o ausencias bloquean el Excel final.',
  },
  {
    Check: 'Graficas generables con datos reales',
    Status: scenarioSummary.every((r) => r.ValidRuns === 3 && r.AvgDurationSeconds !== '') ? 'OK' : 'FAIL',
    Evidence: JSON.stringify(scenarioSummary.map((r) => ({ Scenario: r.Scenario, AvgCPU: r.AvgCPUPercent, AvgRAM: r.AvgRAMMB }))),
    Severity: 'CRITICA',
    Decision: 'No se generan graficas vacias.',
  },
];

if (runs.some((r) => r.RunnerStillRunningAtEnd)) {
  discrepancies.push({
    Tipo: 'WARN_NO_BLOQUEANTE',
    Ambito: 'VR_TEC_RUNNER',
    Detalle: 'RunnerStillRunningAtEnd=True en runs VR_TEC_RUNNER.',
    Decision: 'Se conserva como advertencia metodologica; no invalida el benchmark porque no afecta a CPU/RAM del cliente Velociraptor ya muestreadas.',
    Fuente: runs.filter((r) => r.RunnerStillRunningAtEnd).map((r) => r.SourceFile).join(' | '),
  });
}
if (runs.some((r) => r.ServerGuiDetected)) {
  discrepancies.push({
    Tipo: 'INFO_METODOLOGICA',
    Ambito: 'SERVER_GUI',
    Detalle: 'SERVER_GUI detectado en process_samples.',
    Decision: 'Excluido del calculo principal mediante IncludeServerGuiInTotal=False y ServerGuiExcludedFromClientMetrics=True.',
    Fuente: p.campaignRoot,
  });
}
discrepancies.push({
  Tipo: 'INFO_REFERENCIA',
  Ambito: 'Analisis_Benchmark.xlsx',
  Detalle: 'Excel historico localizado.',
  Decision: 'Usado solo como referencia historica/formato, no como resultado definitivo.',
  Fuente: p.historicalXlsx,
});
if (duplicateRunIds.length > 0) {
  discrepancies.push({
    Tipo: 'ERROR_DUPLICADO',
    Ambito: 'RunId',
    Detalle: JSON.stringify(duplicateRunIds),
    Decision: 'Bloqueante si persiste.',
    Fuente: p.campaignRoot,
  });
}
for (const miss of missing) {
  discrepancies.push({
    Tipo: 'ERROR_AUSENCIA',
    Ambito: `${miss.Scenario} REP_${String(miss.Repetition).padStart(2, '0')}`,
    Detalle: 'Run esperado no localizado.',
    Decision: 'Bloqueante; repetir prueba.',
    Fuente: p.campaignRoot,
  });
}

const blockingFailures = checks.filter((c) => c.Status === 'FAIL' && c.Severity === 'CRITICA');
const apto = blockingFailures.length === 0;

const kpi = {
  TotalRuns: runs.length,
  ValidRuns: validRuns.length,
  ScenarioCount: new Set(runs.map((r) => r.Scenario)).size,
  RepetitionsPerScenario: countsByScenario,
  AvgCPU_BASELINE_NO_VR: round(scenarioSummary.find((r) => r.Scenario === 'BASELINE_NO_VR')?.AvgCPUPercent ?? 0),
  AvgCPU_VR_IDLE: round(scenarioSummary.find((r) => r.Scenario === 'VR_IDLE')?.AvgCPUPercent ?? 0),
  AvgCPU_VR_TEC_RUNNER: round(scenarioSummary.find((r) => r.Scenario === 'VR_TEC_RUNNER')?.AvgCPUPercent ?? 0),
  AvgRAM_BASELINE_NO_VR: round(scenarioSummary.find((r) => r.Scenario === 'BASELINE_NO_VR')?.AvgRAMMB ?? 0),
  AvgRAM_VR_IDLE: round(scenarioSummary.find((r) => r.Scenario === 'VR_IDLE')?.AvgRAMMB ?? 0),
  AvgRAM_VR_TEC_RUNNER: round(scenarioSummary.find((r) => r.Scenario === 'VR_TEC_RUNNER')?.AvgRAMMB ?? 0),
  PeakCPU: round(max(runs.map((r) => r.Max_VR_Client_CPU_Percent))),
  PeakRAM: round(max(runs.map((r) => r.Max_VR_Client_RAM_MB))),
};

const dataset = {
  generatedAt,
  status: apto ? 'APTO' : 'NO_APTO',
  paths: Object.fromEntries(Object.entries(p).map(([k, v]) => [k, path.resolve(v)])),
  kpi,
  runs,
  scenarioSummary,
  rawSamples,
  processSamples,
  checks,
  sources,
  discrepancies,
  blockingFailures,
};

const headers = {
  runs: [
    'RunId', 'Scenario', 'Repetition', 'StartTime', 'EndTime', 'DurationSeconds',
    'ScenarioValidity', 'ValidForClientBenchmark', 'ClientServiceStatusStart',
    'ClientServiceStatusEnd', 'ClientRows', 'RunnerObservedSamples',
    'VR_Client_ProcessCountMax', 'Avg_VR_Client_CPU_Percent',
    'Max_VR_Client_CPU_Percent', 'P95_VR_Client_CPU_Percent',
    'Avg_VR_Client_RAM_MB', 'Max_VR_Client_RAM_MB', 'P95_VR_Client_RAM_MB',
    'Warnings', 'SourceFile', 'SHA-256', 'SchemaVersion', 'Samples',
    'RawSamplesRows', 'ProcessSampleRows', 'ServerGuiRows',
    'ServerGuiExcludedFromClientMetrics', 'IncludeServerGuiInTotal',
    'ServerGuiDetected', 'NotepadRows', 'RunnerStillRunningAtEnd', 'RunnerStarted',
  ],
  scenario: [
    'Scenario', 'RepetitionsExpected', 'RepetitionsFound', 'ValidRuns',
    'AvgDurationSeconds', 'AvgCPUPercent', 'MaxCPUPercent', 'AvgRAMMB',
    'MaxRAMMB', 'RunnerObservedSamples', 'Warnings', 'Interpretation',
  ],
  checks: ['Check', 'Status', 'Evidence', 'Severity', 'Decision'],
  sources: ['Ruta', 'Tipo', 'TamanoBytes', 'FechaModificacion', 'SHA-256', 'Rol', 'Campana', 'Observaciones'],
  discrepancies: ['Tipo', 'Ambito', 'Detalle', 'Decision', 'Fuente'],
};

const rawHeaders = Array.from(new Set(rawSamples.flatMap((r) => Object.keys(r))));
const procHeaders = Array.from(new Set(processSamples.flatMap((r) => Object.keys(r))));

writeCsv(path.join(p.normalizedDir, 'runs_validos.csv'), runs, headers.runs);
writeCsv(path.join(p.normalizedDir, 'resumen_escenario.csv'), scenarioSummary, headers.scenario);
writeCsv(path.join(p.normalizedDir, 'raw_samples.csv'), rawSamples, rawHeaders);
writeCsv(path.join(p.normalizedDir, 'process_samples.csv'), processSamples, procHeaders);
writeCsv(path.join(p.normalizedDir, 'calidad_datos.csv'), checks, headers.checks);
writeCsv(path.join(p.normalizedDir, 'trazabilidad.csv'), sources, headers.sources);
writeCsv(path.join(p.normalizedDir, 'discrepancias.csv'), discrepancies, headers.discrepancies);

fs.writeFileSync(path.join(p.traceDir, 'benchmark_dataset.json'), JSON.stringify(dataset, null, 2), 'utf8');
fs.writeFileSync(path.join(p.traceDir, 'TFM_BENCHMARK_RENDIMIENTO_VELOCIRAPTOR_FINAL_26062026_DATA_QUALITY.json'), JSON.stringify({
  generatedAt,
  status: dataset.status,
  decision: apto ? 'Datos suficientes para generar Excel final.' : 'Datos insuficientes; no generar Excel final.',
  kpi,
  checks,
  blockingFailures,
  methodologicalWarnings: discrepancies.filter((d) => d.Tipo.startsWith('WARN') || d.Tipo.startsWith('INFO')),
}, null, 2), 'utf8');

function mdTable(rows, cols) {
  const header = `| ${cols.join(' | ')} |`;
  const sep = `| ${cols.map(() => '---').join(' | ')} |`;
  const body = rows.map((r) => `| ${cols.map((c) => String(r[c] ?? '').replace(/\|/g, '\\|')).join(' | ')} |`);
  return [header, sep, ...body].join('\n');
}

const precheck = [
  '# BENCHMARK_PRECHECK_26062026',
  '',
  `Fecha generacion: ${generatedAt}`,
  '',
  `Decision: **${apto ? 'APTO' : 'NO APTO'}**`,
  '',
  '## Fuente canonica',
  '',
  `- ${path.resolve(p.canonicalXlsx)}`,
  '- Analisis_Benchmark.xlsx queda limitado a referencia historica/formato.',
  '',
  '## Validaciones obligatorias',
  '',
  mdTable(checks, ['Check', 'Status', 'Evidence', 'Severity', 'Decision']),
  '',
  '## KPIs del precheck',
  '',
  mdTable(Object.entries(kpi).map(([Metrica, Valor]) => ({ Metrica, Valor: typeof Valor === 'object' ? JSON.stringify(Valor) : Valor })), ['Metrica', 'Valor']),
  '',
  '## Resumen por escenario',
  '',
  mdTable(scenarioSummary, headers.scenario),
  '',
  '## Discrepancias y decisiones',
  '',
  mdTable(discrepancies, headers.discrepancies),
  '',
  '## Decision metodologica',
  '',
  apto
    ? 'Los datos actuales son suficientes para generar el Excel definitivo. RunnerStillRunningAtEnd=True se mantiene como WARN no bloqueante en VR_TEC_RUNNER. SERVER_GUI esta observado pero excluido del calculo principal. CPU/RAM se interpretan solo como coste operativo observado en laboratorio.'
    : 'No se debe generar Excel definitivo hasta corregir los fallos criticos indicados.',
  '',
].join('\n');
fs.writeFileSync(path.join(p.traceDir, 'BENCHMARK_PRECHECK_26062026.md'), precheck, 'utf8');

if (!apto) {
  const missingLines = missing.length
    ? missing.map((m) => `- Falta ${m.Scenario} REP_${String(m.Repetition).padStart(2, '0')}`).join('\n')
    : '- Revisar fallos criticos del precheck.';
  const needed = [
    '# BENCHMARK_PRUEBAS_NECESARIAS_26062026',
    '',
    `Fecha generacion: ${generatedAt}`,
    '',
    '## Motivo',
    '',
    'El precheck ha detectado datos insuficientes o inconsistentes. No se genera Excel final.',
    '',
    '## Pruebas faltantes o inconsistentes',
    '',
    missingLines,
    '',
    '## Comandos orientativos',
    '',
    'Ejecutar desde la raiz del TFM en la VM/laboratorio donde se genero BENCH_20260619_FINAL01. Ajustar solo la repeticion y escenario indicados por el precheck.',
    '',
    '```powershell',
    'Set-Location C:\\Users\\seguridad\\Desktop\\TFM',
    '.\\06_CONTROLLED_RERUN\\support\\TFM_Benchmark_VR_Resource_Usage_v1_role_fix.ps1 -Scenario <ESCENARIO> -OutputRoot .\\06_CONTROLLED_RERUN\\OUTPUT\\BENCHMARK\\BENCH_20260619_FINAL01\\REP_<NN>',
    '```',
    '',
    '## Salida esperada',
    '',
    '- summary.json con SchemaVersion 1.1 y ScenarioValidity=VALID.',
    '- samples.csv con muestras de CPU/RAM.',
    '- process_samples.csv con CLIENT_SERVICE cuando aplique.',
    '- SERVER_GUI excluido del calculo principal.',
    '- notepad.exe ausente como runner.',
    '',
    '## Archivo a devolver',
    '',
    'Comprimir la carpeta del run generado y entregar summary.json, samples.csv y process_samples.csv.',
    '',
  ].join('\n');
  fs.writeFileSync(path.join(p.traceDir, 'BENCHMARK_PRUEBAS_NECESARIAS_26062026.md'), needed, 'utf8');
}

console.log(JSON.stringify({
  status: dataset.status,
  totalRuns: runs.length,
  validRuns: validRuns.length,
  scenarios: countsByScenario,
  rawSamples: rawSamples.length,
  processSamples: processSamples.length,
  traceDir: p.traceDir,
}, null, 2));
