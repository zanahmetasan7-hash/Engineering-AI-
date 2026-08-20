// Engineering AI Core — прототип
// Загрузка/вставка BOM (CSV, Excel) -> ИИ-анализ (напряжение, температура, статус производства) -> отчёт с рисками и рекомендациями
// ИИ-движок вызывает Claude API напрямую из браузера (см. analyzeBatch / generateSummary)
// Дизайн-система: Инженерный Кибернетический Минимализм

import React, { useState, useMemo, useRef, useCallback } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  Cpu, Upload, ClipboardList, Sparkles, AlertTriangle, AlertCircle,
  CheckCircle2, Loader2, ChevronDown, ChevronRight, Download, RotateCcw, Zap, X,
  KeyRound, Eye, EyeOff,
} from 'lucide-react';

const BATCH_SIZE = 4;
const MAX_ROWS = 40;
const MODEL = 'claude-sonnet-5';

const RISK_LABELS = { high: 'ВЫСОКИЙ', medium: 'СРЕДНИЙ', low: 'НИЗКИЙ', error: 'ОШИБКА' };

const SAMPLE_ROWS = [
  { 'Номер детали': 'NE555P', 'Наименование': 'Таймер', 'Производитель': 'Texas Instruments', 'Кол-во': '10', 'Напряжение': '4.5-16V', 'Раб. температура': '0...70°C', 'Статус производства': 'Active' },
  { 'Номер детали': 'LM7805CT', 'Наименование': 'Стабилизатор напряжения 5V', 'Производитель': 'STMicroelectronics', 'Кол-во': '5', 'Напряжение': '7-25V (вход)', 'Раб. температура': '0...125°C', 'Статус производства': 'Active' },
  { 'Номер детали': 'GRM188R71C104KA01', 'Наименование': 'Конденсатор керамич. 0.1мкФ', 'Производитель': 'Murata', 'Кол-во': '50', 'Напряжение': '16V', 'Раб. температура': '-55...125°C', 'Статус производства': 'Active' },
  { 'Номер детали': 'ATmega328P-PU', 'Наименование': 'Микроконтроллер', 'Производитель': 'Microchip', 'Кол-во': '20', 'Напряжение': '1.8-5.5V', 'Раб. температура': '-40...85°C', 'Статус производства': 'Active' },
  { 'Номер детали': 'IRF540N', 'Наименование': 'MOSFET транзистор', 'Производитель': 'Infineon', 'Кол-во': '15', 'Напряжение': '100V', 'Раб. температура': '-55...175°C', 'Статус производства': '' },
  { 'Номер детали': 'PIC16F84A-04/P', 'Наименование': 'Микроконтроллер (устар. серия)', 'Производитель': 'Microchip', 'Кол-во': '8', 'Напряжение': '4-5.5V', 'Раб. температура': '0...70°C', 'Статус производства': '' },
  { 'Номер детали': 'CFR-25JB-10K', 'Наименование': 'Резистор 10 кОм 0.25Вт', 'Производитель': 'Yageo', 'Кол-во': '100', 'Напряжение': '-', 'Раб. температура': '-55...155°C', 'Статус производства': 'Active' },
];
const SAMPLE_CONDITIONS = { voltage: '12V', tempRange: '-10...50°C' };

// ---------- helpers ----------

function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function detectLabelKey(headers) {
  const kws = ['номер', 'парт', 'part', 'деталь', 'designator', 'компонент', 'наимен', 'назв', 'description', 'name', 'артикул'];
  return headers.find((h) => kws.some((k) => h.toLowerCase().includes(k))) || headers[0] || '';
}

function formatRowLabel(row, labelKey) {
  const val = labelKey && row[labelKey] !== undefined ? row[labelKey] : '';
  const s = String(val).trim();
  if (s) return s;
  return `Компонент ${parseInt(row._id, 10) + 1}`;
}

function parseCSVText(text) {
  const res = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
  const headers = res.meta && res.meta.fields ? res.meta.fields : [];
  const rows = (res.data || []).filter((r) => Object.values(r).some((v) => String(v).trim() !== ''));
  return { headers, rows };
}

async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv' || ext === 'txt') {
    const text = await file.text();
    return parseCSVText(text);
  }
  if (ext === 'xlsx' || ext === 'xls') {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const headers = json.length ? Object.keys(json[0]) : [];
    return { headers, rows: json };
  }
  throw new Error('unsupported');
}

function extractJSON(text) {
  const cleaned = text.replace(/```json/gi, '').replace(/```/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const match = cleaned.match(/[[{][\s\S]*[\]}]/);
    if (match) return JSON.parse(match[0]);
    throw e;
  }
}

function authError() {
  const err = new Error('Неверный API-ключ. Проверьте, что он скопирован полностью и активен.');
  err.code = 'auth';
  return err;
}

async function callClaude(system, userContent, apiKey) {
  if (!apiKey || !apiKey.trim()) throw authError();
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': apiKey.trim(),
      'anthropic-version': '2023-06-01',
      // Requests the API's CORS opt-in so this call can run directly from the browser
      // with a user-supplied key, instead of through a backend proxy.
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 1000,
      // Structured JSON output only — no need for extended thinking, and disabling it
      // keeps the full max_tokens budget available for the actual response.
      thinking: { type: 'disabled' },
      system,
      messages: [{ role: 'user', content: userContent }],
    }),
  });
  if (!response.ok) {
    if (response.status === 401) throw authError();
    if (response.status === 429) throw new Error('Превышен лимит запросов к API. Подождите немного и повторите анализ.');
    throw new Error('Ошибка API: ' + response.status);
  }
  const data = await response.json();
  if (data.type === 'error') {
    if (data.error && data.error.type === 'authentication_error') throw authError();
    throw new Error((data.error && data.error.message) || 'API error');
  }
  const text = (data.content || []).map((b) => (b.type === 'text' ? b.text : '')).join('');
  return extractJSON(text);
}

const ANALYSIS_SYSTEM = `Ты — экспертная ИИ-система инженерного анализа BOM (спецификаций электронных компонентов) в приложении Engineering AI Core.

Для каждого компонента определи:
1. risk_level: "high" | "medium" | "low" — на основе соответствия параметров (напряжение, температура) условиям эксплуатации проекта, если они указаны, и рисков по статусу производства/доступности компонента.
2. issues — не более 2 пунктов, каждый одно короткое предложение. Пустой массив, если проблем нет.
3. recommendations — не более 2 пунктов, каждый одно короткое предложение.
4. production_status_note — не более 12 слов; если данных недостаточно, укажи, что статус нужно проверить у поставщика.

Отвечай СТРОГО в формате JSON-массива, без пояснений, без markdown, без тройных кавычек. Каждый элемент — объект с полем "id", совпадающим с id из входных данных:
{"id":"...","risk_level":"high|medium|low","issues":["..."],"recommendations":["..."],"production_status_note":"..."}
Будь предельно кратким — это критично для корректной работы системы.`;

const SUMMARY_SYSTEM = `Ты — экспертная ИИ-система инженерного анализа BOM в приложении Engineering AI Core. На основе результатов анализа отдельных компонентов составь краткое резюме проекта на русском языке (2-4 предложения) и список из 3-5 приоритетных рекомендаций для инженерной команды (каждая — одно короткое предложение).

Отвечай СТРОГО в формате JSON, без пояснений, без markdown, без тройных кавычек:
{"summary":"...","recommendations":["...","..."]}`;

async function analyzeBatch(rows, conditions, apiKey) {
  const payload = rows.map((r) => {
    const { _id, ...rest } = r;
    return { id: _id, ...rest };
  });
  const userContent = `Условия эксплуатации проекта:
Напряжение питания: ${conditions.voltage ? conditions.voltage : 'не указано'}
Диапазон температур окружающей среды: ${conditions.tempRange ? conditions.tempRange : 'не указано'}

Компоненты для анализа (JSON):
${JSON.stringify(payload)}

Верни JSON-массив результатов анализа для каждого компонента из списка выше.`;
  const result = await callClaude(ANALYSIS_SYSTEM, userContent, apiKey);
  return Array.isArray(result) ? result : [];
}

async function generateSummary(results, conditions, totalCount, apiKey) {
  const counts = { high: 0, medium: 0, low: 0 };
  results.forEach((r) => { if (counts[r.risk_level] !== undefined) counts[r.risk_level]++; });
  const keyIssues = results.filter((r) => r.risk_level === 'high').flatMap((r) => r.issues || []).slice(0, 15);
  const userContent = `Всего компонентов: ${totalCount}. Высокий риск: ${counts.high}, средний: ${counts.medium}, низкий: ${counts.low}.
Условия эксплуатации: напряжение ${conditions.voltage || 'не указано'}, температура ${conditions.tempRange || 'не указано'}.
Ключевые проблемы (компоненты с высоким риском): ${JSON.stringify(keyIssues)}

Составь резюме и рекомендации.`;
  return callClaude(SUMMARY_SYSTEM, userContent, apiKey);
}

function rowsToCSV(rows, results, labelKey) {
  const header = ['Компонент', 'Уровень риска', 'Проблемы', 'Рекомендации', 'Статус производства'];
  const lines = [header.join(';')];
  rows.forEach((r) => {
    const res = results[r._id] || {};
    const label = formatRowLabel(r, labelKey).replace(/;/g, ',');
    const risk = RISK_LABELS[res.risk_level] || '—';
    const issues = (res.issues || []).join(' | ').replace(/;/g, ',');
    const recs = (res.recommendations || []).join(' | ').replace(/;/g, ',');
    const status = (res.production_status_note || '').replace(/;/g, ',');
    lines.push([label, risk, issues, recs, status].join(';'));
  });
  return lines.join('\n');
}

// ---------- design system: Инженерный Кибернетический Минимализм ----------

const STYLES = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');

:root {
  --core-dark: #17181B;
  --core-dark-deep: #101113;
  --panel-carbon: #1D1F23;
  --panel-carbon-hi: #23262B;
  --seam: rgba(255,255,255,0.06);
  --seam-hi: rgba(255,255,255,0.14);
  --text-primary: #ECEDEF;
  --text-secondary: #9BA0AA;
  --text-tertiary: #6B707B;
  --func-blue: #4C8DFF;
  --func-blue-dim: #2E5AA8;
  --neon-core: #00E5FF;
  --neon-core-soft: rgba(0,229,255,0.2);
  --brass: #C8935B;
  --danger: #FF5C5C;
  --ok: #35D48E;
  --warn: #FFB648;
}

.eac-root {
  background:
    radial-gradient(circle at 12% 0%, rgba(76,141,255,0.05), transparent 45%),
    radial-gradient(circle at 88% 100%, rgba(0,229,255,0.04), transparent 40%),
    var(--core-dark);
  color: var(--text-primary);
  font-family: 'Inter', ui-sans-serif, system-ui, sans-serif;
}
.eac-display { font-family: 'Space Grotesk', ui-sans-serif, sans-serif; letter-spacing: -0.2px; }
.eac-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
.eac-surface { background: var(--panel-carbon); border: 1px solid var(--seam); position: relative; }
.eac-surface::before {
  content: ''; position: absolute; top: 0; left: 14px; right: 14px; height: 1px;
  background: linear-gradient(90deg, transparent, var(--seam-hi), transparent);
}
.eac-muted { color: var(--text-secondary); }
.eac-faint { color: var(--text-tertiary); }
.eac-accent { color: var(--brass); }

/* сенсорные узлы — контрольные точки сборки, подпись стиля */
.eac-node { position: absolute; width: 5px; height: 5px; border-radius: 50%; background: var(--text-tertiary); opacity: 0.5; }
.eac-node.tl { top: 9px; left: 9px; } .eac-node.tr { top: 9px; right: 9px; }
.eac-node.bl { bottom: 9px; left: 9px; } .eac-node.br { bottom: 9px; right: 9px; }
.eac-node.live { background: var(--neon-core); opacity: 0.9; box-shadow: 0 0 5px var(--neon-core-soft); }

.eac-btn-primary {
  background: linear-gradient(180deg, #6FA4FF, var(--func-blue));
  color: var(--core-dark);
  font-weight: 600;
  border: 1px solid var(--func-blue);
  border-bottom: 2px solid #1E3F80;
  box-shadow: 0 6px 18px -8px rgba(76,141,255,0.55);
  transition: transform .12s cubic-bezier(.2,.8,.2,1), border-color .15s ease, filter .15s ease;
}
.eac-btn-primary:hover { filter: brightness(1.06); }
.eac-btn-primary:active { transform: translateY(2px); border-bottom-width: 1px; }
.eac-btn-primary:disabled { opacity: .45; cursor: not-allowed; }

.eac-btn-ghost {
  background: linear-gradient(180deg, var(--panel-carbon-hi), var(--panel-carbon));
  border: 1px solid var(--seam-hi);
  border-bottom: 2px solid rgba(0,0,0,0.35);
  color: var(--text-primary);
  transition: transform .12s cubic-bezier(.2,.8,.2,1), border-color .15s ease;
}
.eac-btn-ghost:hover { border-color: var(--func-blue-dim); }
.eac-btn-ghost:active { transform: translateY(2px); border-bottom-width: 1px; }

.eac-input { background: var(--core-dark-deep); border: 1px solid var(--seam-hi); color: var(--text-primary); border-radius: 10px; }
.eac-input:focus { outline: none; border-color: var(--func-blue); box-shadow: 0 0 0 3px rgba(76,141,255,0.18); }

.eac-focus:focus-visible { outline: 2px solid var(--neon-core); outline-offset: 2px; }

.led-dot { display: inline-block; width: 7px; height: 7px; border-radius: 999px; flex-shrink: 0; }
.led-high { background: var(--danger); box-shadow: 0 0 6px 1px rgba(255,92,92,0.6); }
.led-medium { background: var(--warn); box-shadow: 0 0 6px 1px rgba(255,182,72,0.6); }
.led-low { background: var(--ok); box-shadow: 0 0 6px 1px rgba(53,212,142,0.6); }
.led-pending { background: var(--text-tertiary); }
.led-error { background: #8B8FA3; }

.eac-text-high { color: var(--danger); }
.eac-text-medium { color: var(--warn); }
.eac-text-low { color: var(--ok); }

.eac-status-pulse { animation: eac-pulse 2s ease-in-out infinite; }
@keyframes eac-pulse {
  0%, 100% { opacity: 1; box-shadow: 0 0 7px 1px var(--neon-core-soft); }
  50% { opacity: .45; box-shadow: 0 0 2px 0px transparent; }
}

.eac-shimmer { background: linear-gradient(90deg,#1D1F23 25%,#25282E 37%,#1D1F23 63%); background-size: 400% 100%; animation: eac-shimmer 1.4s ease infinite; }
@keyframes eac-shimmer { 0% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }

.eac-row:hover { background: rgba(76,141,255,0.05); }

@media (prefers-reduced-motion: reduce) {
  .eac-status-pulse, .eac-shimmer { animation: none; }
}
`;

// ---------- component ----------

export default function EngineeringAICore() {
  const [apiKey, setApiKey] = useState('');
  const [showKey, setShowKey] = useState(false);
  const [view, setView] = useState('input'); // 'input' | 'results'
  const [inputMode, setInputMode] = useState('upload'); // 'upload' | 'paste'
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [pasteText, setPasteText] = useState('');
  const [conditions, setConditions] = useState({ voltage: '', tempRange: '' });
  const [loadError, setLoadError] = useState('');
  const [loadNote, setLoadNote] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [results, setResults] = useState({});
  const [summary, setSummary] = useState(null);
  const [riskFilter, setRiskFilter] = useState('all');
  const [expandedIds, setExpandedIds] = useState(() => new Set());
  const [analysisError, setAnalysisError] = useState('');

  const fileInputRef = useRef(null);

  const labelKey = useMemo(() => detectLabelKey(headers), [headers]);

  const loadParsed = useCallback((parsed) => {
    if (!parsed.headers.length || !parsed.rows.length) {
      setLoadError('Не удалось распознать данные. Проверьте формат файла.');
      setLoadNote('');
      return;
    }
    const truncated = parsed.rows.length > MAX_ROWS;
    const withIds = parsed.rows.slice(0, MAX_ROWS).map((r, i) => ({ _id: String(i), ...r }));
    setHeaders(parsed.headers);
    setRows(withIds);
    setLoadError('');
    setLoadNote(truncated ? `Загружены первые ${MAX_ROWS} компонентов из ${parsed.rows.length}.` : '');
  }, []);

  const handleFiles = useCallback(async (fileList) => {
    const file = fileList && fileList[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'pdf') {
      setLoadError('Загрузка PDF появится в следующей версии. Пока используйте CSV, Excel или вставку данных.');
      return;
    }
    try {
      const parsed = await parseFile(file);
      loadParsed(parsed);
    } catch (e) {
      setLoadError('Не удалось прочитать файл. Проверьте, что это CSV или Excel.');
    }
  }, [loadParsed]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    handleFiles(e.dataTransfer.files);
  }, [handleFiles]);

  const handlePasteLoad = useCallback(() => {
    if (!pasteText.trim()) return;
    loadParsed(parseCSVText(pasteText));
  }, [pasteText, loadParsed]);

  const handleLoadSample = useCallback(() => {
    const withIds = SAMPLE_ROWS.map((r, i) => ({ _id: String(i), ...r }));
    setHeaders(Object.keys(SAMPLE_ROWS[0]));
    setRows(withIds);
    setConditions(SAMPLE_CONDITIONS);
    setLoadError('');
    setLoadNote('');
  }, []);

  const handleReset = useCallback(() => {
    setView('input');
    setHeaders([]);
    setRows([]);
    setPasteText('');
    setResults({});
    setSummary(null);
    setRiskFilter('all');
    setExpandedIds(new Set());
    setAnalysisError('');
    setLoadError('');
    setLoadNote('');
  }, []);

  const handleAnalyze = useCallback(async () => {
    if (!rows.length) return;
    if (!apiKey.trim()) {
      setAnalysisError('Введите API-ключ Anthropic выше, чтобы запустить анализ.');
      return;
    }
    setView('results');
    setIsAnalyzing(true);
    setResults({});
    setSummary(null);
    setAnalysisError('');

    const batches = chunkArray(rows, BATCH_SIZE);
    setProgress({ current: 0, total: batches.length });
    const allResults = {};
    let hadError = false;
    let firstErrorMessage = '';
    let stopEarly = false;

    for (let i = 0; i < batches.length; i++) {
      const batchIds = new Set(batches[i].map((r) => r._id));
      if (!stopEarly) {
        try {
          const batchResults = await analyzeBatch(batches[i], conditions, apiKey);
          (Array.isArray(batchResults) ? batchResults : []).forEach((r) => {
            if (r && r.id !== undefined && batchIds.has(String(r.id))) {
              const id = String(r.id);
              allResults[id] = {
                id,
                risk_level: ['high', 'medium', 'low'].includes(r.risk_level) ? r.risk_level : 'medium',
                issues: Array.isArray(r.issues) ? r.issues.map(String) : [],
                recommendations: Array.isArray(r.recommendations) ? r.recommendations.map(String) : [],
                production_status_note: r.production_status_note ? String(r.production_status_note) : '',
              };
            }
          });
        } catch (e) {
          hadError = true;
          if (!firstErrorMessage) firstErrorMessage = (e && e.message) ? e.message : 'Не удалось подключиться к Claude API.';
          if (e && e.code === 'auth') stopEarly = true;
        }
      }
      batches[i].forEach((r) => {
        if (!allResults[r._id]) {
          hadError = true;
          allResults[r._id] = {
            id: r._id,
            risk_level: 'error',
            issues: [stopEarly ? 'Анализ остановлен: проверьте API-ключ.' : 'Не удалось получить анализ ИИ для этого компонента.'],
            recommendations: [],
            production_status_note: '',
          };
        }
      });
      setResults({ ...allResults });
      setProgress({ current: i + 1, total: batches.length });
    }

    if (hadError) {
      setAnalysisError(firstErrorMessage || 'Часть компонентов не удалось проанализировать. Можно запустить анализ заново.');
    }

    if (!stopEarly) {
      try {
        const summaryResult = await generateSummary(Object.values(allResults), conditions, rows.length, apiKey);
        setSummary({
          summary: (summaryResult && summaryResult.summary) ? String(summaryResult.summary) : 'Резюме недоступно.',
          recommendations: (summaryResult && Array.isArray(summaryResult.recommendations)) ? summaryResult.recommendations.map(String) : [],
        });
      } catch (e) {
        setSummary({ summary: 'Не удалось сформировать сводное резюме. Результаты по отдельным компонентам доступны ниже.', recommendations: [] });
      }
    } else {
      setSummary({ summary: 'Резюме недоступно — не удалось подключиться к Claude API с указанным ключом.', recommendations: [] });
    }

    setIsAnalyzing(false);
  }, [rows, conditions, apiKey]);

  const toggleExpand = useCallback((id) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }, []);

  const riskCounts = useMemo(() => {
    const c = { high: 0, medium: 0, low: 0, error: 0 };
    rows.forEach((r) => {
      const res = results[r._id];
      if (res && c[res.risk_level] !== undefined) c[res.risk_level]++;
    });
    return c;
  }, [rows, results]);

  const filteredRows = useMemo(() => {
    if (riskFilter === 'all') return rows;
    return rows.filter((r) => results[r._id] && results[r._id].risk_level === riskFilter);
  }, [rows, results, riskFilter]);

  const handleExportCSV = useCallback(() => {
    const csv = rowsToCSV(rows, results, labelKey);
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bom-analysis-report.csv';
    a.click();
    URL.revokeObjectURL(url);
  }, [rows, results, labelKey]);

  const previewRows = rows.slice(0, 8);
  const progressPct = progress.total ? Math.round((progress.current / progress.total) * 100) : 0;

  return (
    <div className="eac-root min-h-screen">
      <style>{STYLES}</style>
      <div className="max-w-4xl mx-auto px-4 py-8 sm:px-6">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center eac-surface">
              <Cpu size={19} className="eac-accent" />
            </div>
            <div>
              <h1 className="eac-display text-lg font-semibold tracking-tight">Engineering AI Core</h1>
              <p className="eac-mono eac-faint text-[11px] tracking-wide uppercase">BOM Analysis Engine · ИИ-анализ спецификаций</p>
            </div>
          </div>
          <span
            className={`led-dot ${isAnalyzing ? 'led-medium eac-status-pulse' : 'led-low'}`}
            title={isAnalyzing ? 'Анализ выполняется' : 'Готово'}
          ></span>
        </div>

        <div className="eac-surface rounded-2xl p-5 mb-4">
          <div className="flex items-center gap-2 mb-3">
            <KeyRound size={14} className="eac-accent" />
            <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide">Anthropic API Key</p>
            <span className={`led-dot ${apiKey.trim() ? 'led-low' : 'led-pending'}`}></span>
          </div>
          <div className="relative">
            <input
              type={showKey ? 'text' : 'password'}
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder="sk-ant-..."
              autoComplete="off"
              spellCheck="false"
              className="eac-input eac-focus w-full px-3 py-2 pr-10 text-sm eac-mono"
            />
            <button
              type="button"
              onClick={() => setShowKey((v) => !v)}
              className="eac-focus eac-muted absolute inset-y-0 right-0 flex items-center px-3"
              aria-label={showKey ? 'Скрыть ключ' : 'Показать ключ'}
            >
              {showKey ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </div>
          <p className="eac-faint text-xs mt-2 leading-relaxed">
            Ключ хранится только в памяти этой вкладки браузера и никуда не сохраняется — используется исключительно для прямых запросов к api.anthropic.com. Получить ключ можно в console.anthropic.com. Для публичного продакшн-приложения не встраивайте ключ в клиентский код — держите его на backend-сервере (см. FastAPI-слой в исходном плане).
          </p>
        </div>

        {view === 'input' && (
          <div className="eac-surface rounded-2xl p-5">
            <div className="flex gap-2 mb-4">
              <button
                type="button"
                onClick={() => setInputMode('upload')}
                className={`eac-focus px-4 py-2 rounded-xl text-sm font-medium flex items-center gap-2 ${inputMode === 'upload' ? 'eac-btn-primary' : 'eac-btn-ghost'}`}
              >
                <Upload size={14} /> Загрузить файл
              </button>
              <button
                type="button"
                onClick={() => setInputMode('paste')}
                className={`eac-focus px-4 py-2 rounded-xl text-sm font-medium flex items-center gap-2 ${inputMode === 'paste' ? 'eac-btn-primary' : 'eac-btn-ghost'}`}
              >
                <ClipboardList size={14} /> Вставить данные
              </button>
            </div>

            {inputMode === 'upload' && (
              <div
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInputRef.current && fileInputRef.current.click(); } }}
                tabIndex={0}
                role="button"
                aria-label="Загрузить файл BOM"
                className="eac-focus rounded-2xl p-8 text-center cursor-pointer transition-colors"
                style={{ border: `1px dashed ${isDragging ? '#4C8DFF' : 'rgba(255,255,255,0.14)'}`, background: isDragging ? 'rgba(76,141,255,0.06)' : 'transparent' }}
              >
                <Upload size={26} className="eac-faint mx-auto mb-3" />
                <p className="text-sm mb-1">Перетащите файл сюда или нажмите, чтобы выбрать</p>
                <p className="eac-mono eac-faint text-xs">CSV или Excel · до {MAX_ROWS} компонентов</p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".csv,.xlsx,.xls,.txt"
                  className="hidden"
                  onChange={(e) => handleFiles(e.target.files)}
                />
              </div>
            )}

            {inputMode === 'paste' && (
              <div>
                <textarea
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  placeholder="Номер детали, Наименование, Напряжение, Температура\nNE555P, Таймер, 4.5-16V, 0...70°C\nLM7805, Стабилизатор, 7-25V, 0...125°C"
                  rows={6}
                  className="eac-input eac-focus w-full p-3 text-sm eac-mono"
                />
                <button
                  type="button"
                  onClick={handlePasteLoad}
                  disabled={!pasteText.trim()}
                  className="eac-btn-primary eac-focus mt-3 px-4 py-2 rounded-xl text-sm font-medium"
                >
                  Загрузить данные
                </button>
              </div>
            )}

            <div className="flex items-center gap-3 my-4">
              <div className="flex-1" style={{ borderTop: '1px solid var(--seam)' }}></div>
              <span className="eac-mono eac-faint text-xs uppercase">или</span>
              <div className="flex-1" style={{ borderTop: '1px solid var(--seam)' }}></div>
            </div>
            <button
              type="button"
              onClick={handleLoadSample}
              className="eac-btn-ghost eac-focus w-full px-4 py-2 rounded-xl text-sm font-medium flex items-center justify-center gap-2"
            >
              <Sparkles size={15} className="eac-accent" /> Загрузить пример BOM
            </button>

            {loadError && (
              <div className="flex items-start gap-2 mt-3 text-sm eac-text-high">
                <AlertCircle size={16} className="flex-shrink-0 mt-0.5" />
                <span>{loadError}</span>
              </div>
            )}
            {loadNote && (
              <div className="flex items-start gap-2 mt-3 text-sm eac-faint">
                <span>{loadNote}</span>
              </div>
            )}

            <div className="mt-6 pt-6" style={{ borderTop: '1px solid var(--seam)' }}>
              <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-3">Условия эксплуатации проекта (опционально)</p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="text-xs eac-muted block mb-1">Напряжение питания</label>
                  <input
                    type="text"
                    value={conditions.voltage}
                    onChange={(e) => setConditions((c) => ({ ...c, voltage: e.target.value }))}
                    placeholder="например, 5V"
                    className="eac-input eac-focus w-full px-3 py-2 text-sm eac-mono"
                  />
                </div>
                <div>
                  <label className="text-xs eac-muted block mb-1">Диапазон температур окружающей среды</label>
                  <input
                    type="text"
                    value={conditions.tempRange}
                    onChange={(e) => setConditions((c) => ({ ...c, tempRange: e.target.value }))}
                    placeholder="например, -20...60°C"
                    className="eac-input eac-focus w-full px-3 py-2 text-sm eac-mono"
                  />
                </div>
              </div>
            </div>

            {rows.length > 0 && (
              <div className="mt-6 pt-6" style={{ borderTop: '1px solid var(--seam)' }}>
                <div className="flex items-center justify-between mb-3">
                  <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide">Предпросмотр · {rows.length} компонентов</p>
                  <button type="button" onClick={handleReset} className="eac-focus eac-faint text-xs flex items-center gap-1">
                    <X size={12} /> Очистить
                  </button>
                </div>
                <div className="overflow-x-auto rounded-xl eac-surface">
                  <table className="w-full text-xs">
                    <thead>
                      <tr style={{ borderBottom: '1px solid var(--seam)' }}>
                        {headers.map((h) => (
                          <th key={h} className="text-left px-3 py-2 eac-faint font-medium whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((r) => (
                        <tr key={r._id} style={{ borderTop: '1px solid var(--seam)' }}>
                          {headers.map((h) => (
                            <td key={h} className="px-3 py-2 eac-mono whitespace-nowrap">{String(r[h] !== undefined ? r[h] : '')}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {rows.length > previewRows.length && (
                  <p className="eac-faint text-xs mt-2">и ещё {rows.length - previewRows.length} компонентов</p>
                )}
                <button
                  type="button"
                  onClick={handleAnalyze}
                  disabled={!apiKey.trim()}
                  className="eac-btn-primary eac-focus w-full sm:w-auto mt-5 px-6 py-3 rounded-xl text-sm font-semibold flex items-center justify-center gap-2"
                >
                  <Zap size={16} /> Анализировать с ИИ
                </button>
                {!apiKey.trim() && (
                  <p className="text-xs mt-2 eac-text-medium">Введите API-ключ выше, чтобы начать анализ.</p>
                )}
              </div>
            )}
          </div>
        )}

        {view === 'results' && (
          <div>
            <div className="flex items-center justify-between mb-5">
              <button type="button" onClick={handleReset} className="eac-btn-ghost eac-focus px-3 py-2 rounded-xl text-sm flex items-center gap-2">
                <RotateCcw size={14} /> Новый анализ
              </button>
              {!isAnalyzing && (
                <button type="button" onClick={handleExportCSV} className="eac-btn-ghost eac-focus px-3 py-2 rounded-xl text-sm flex items-center gap-2">
                  <Download size={14} /> Экспорт CSV
                </button>
              )}
            </div>

            {isAnalyzing && (
              <div className="eac-surface rounded-2xl p-4 mb-5">
                <div className="flex items-center justify-between mb-2">
                  <span className="eac-mono text-xs eac-faint uppercase tracking-wide flex items-center gap-2">
                    <Loader2 size={14} className="animate-spin eac-accent" />
                    Анализ компонентов · батч {progress.current} из {progress.total}
                  </span>
                  <span className="eac-mono text-xs" style={{ color: 'var(--func-blue)' }}>{progressPct}%</span>
                </div>
                <div className="w-full rounded-full h-2" style={{ background: 'var(--core-dark-deep)' }}>
                  <div className="h-2 rounded-full" style={{ width: progressPct + '%', background: 'linear-gradient(90deg, var(--func-blue-dim), var(--func-blue), var(--neon-core))', transition: 'width 0.3s ease' }}></div>
                </div>
              </div>
            )}

            {analysisError && !isAnalyzing && (
              <div className="flex items-start justify-between gap-3 mb-5 text-sm eac-text-medium">
                <div className="flex items-start gap-2">
                  <AlertTriangle size={16} className="flex-shrink-0 mt-0.5" />
                  <span>{analysisError}</span>
                </div>
                <button type="button" onClick={handleAnalyze} className="eac-btn-ghost eac-focus px-3 py-1.5 rounded-lg text-xs flex items-center gap-1 flex-shrink-0">
                  <RotateCcw size={12} /> Повторить
                </button>
              </div>
            )}

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-5">
              <div className="eac-surface rounded-2xl p-4">
                <span className="eac-node tl"></span><span className="eac-node br"></span>
                <p className="eac-display text-2xl font-semibold eac-mono">{rows.length}</p>
                <p className="text-xs eac-faint mt-1">Всего компонентов</p>
              </div>
              <div className="eac-surface rounded-2xl p-4">
                <span className={`eac-node tl ${riskCounts.high > 0 ? 'live' : ''}`}></span><span className="eac-node br"></span>
                <div className="flex items-center gap-2">
                  <span className="led-dot led-high"></span>
                  <p className="eac-display text-2xl font-semibold eac-mono">{riskCounts.high}</p>
                </div>
                <p className="text-xs eac-faint mt-1">Высокий риск</p>
              </div>
              <div className="eac-surface rounded-2xl p-4">
                <span className="eac-node tl"></span><span className="eac-node br"></span>
                <div className="flex items-center gap-2">
                  <span className="led-dot led-medium"></span>
                  <p className="eac-display text-2xl font-semibold eac-mono">{riskCounts.medium}</p>
                </div>
                <p className="text-xs eac-faint mt-1">Средний риск</p>
              </div>
              <div className="eac-surface rounded-2xl p-4">
                <span className="eac-node tl"></span><span className="eac-node br"></span>
                <div className="flex items-center gap-2">
                  <span className="led-dot led-low"></span>
                  <p className="eac-display text-2xl font-semibold eac-mono">{riskCounts.low}</p>
                </div>
                <p className="text-xs eac-faint mt-1">Низкий риск</p>
              </div>
            </div>

            <div className="flex w-full h-2 rounded-full overflow-hidden mb-6" style={{ background: 'var(--core-dark-deep)' }}>
              {riskCounts.high > 0 && <div style={{ width: (riskCounts.high / rows.length * 100) + '%', background: 'var(--danger)' }}></div>}
              {riskCounts.medium > 0 && <div style={{ width: (riskCounts.medium / rows.length * 100) + '%', background: 'var(--warn)' }}></div>}
              {riskCounts.low > 0 && <div style={{ width: (riskCounts.low / rows.length * 100) + '%', background: 'var(--ok)' }}></div>}
              {riskCounts.error > 0 && <div style={{ width: (riskCounts.error / rows.length * 100) + '%', background: '#8B8FA3' }}></div>}
            </div>

            {(summary || isAnalyzing) && (
              <div className="eac-surface rounded-2xl p-4 mb-4">
                <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-2 flex items-center gap-2">
                  <Sparkles size={12} className="eac-accent" /> Резюме ИИ-анализа
                </p>
                {summary ? (
                  <p className="text-sm leading-relaxed">{summary.summary}</p>
                ) : (
                  <div className="space-y-2">
                    <div className="eac-shimmer h-3 rounded w-full"></div>
                    <div className="eac-shimmer h-3 rounded w-5/6"></div>
                    <div className="eac-shimmer h-3 rounded w-2/3"></div>
                  </div>
                )}
              </div>
            )}

            {summary && summary.recommendations && summary.recommendations.length > 0 && (
              <div className="eac-surface rounded-2xl p-4 mb-6">
                <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-3">Рекомендации</p>
                <ul className="space-y-2">
                  {summary.recommendations.map((rec, i) => (
                    <li key={i} className="flex items-start gap-2 text-sm">
                      <CheckCircle2 size={15} style={{ color: 'var(--ok)' }} className="flex-shrink-0 mt-0.5" />
                      <span>{rec}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className="flex flex-wrap gap-2 mb-4">
              {['all', 'high', 'medium', 'low'].map((f) => (
                <button
                  key={f}
                  type="button"
                  onClick={() => setRiskFilter(f)}
                  className={`eac-focus px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-2 ${riskFilter === f ? 'eac-btn-primary' : 'eac-btn-ghost'}`}
                >
                  {f !== 'all' && <span className={`led-dot led-${f}`}></span>}
                  {f === 'all' && `Все (${rows.length})`}
                  {f === 'high' && `Высокий (${riskCounts.high})`}
                  {f === 'medium' && `Средний (${riskCounts.medium})`}
                  {f === 'low' && `Низкий (${riskCounts.low})`}
                </button>
              ))}
            </div>

            <div className="eac-surface rounded-2xl overflow-hidden">
              <table className="w-full text-sm">
                <thead>
                  <tr style={{ borderBottom: '1px solid var(--seam)' }}>
                    <th className="text-left px-4 py-3 eac-faint font-medium text-[11px] uppercase tracking-wide">Компонент</th>
                    <th className="text-left px-4 py-3 eac-faint font-medium text-[11px] uppercase tracking-wide">Риск</th>
                    <th className="text-left px-4 py-3 eac-faint font-medium text-[11px] uppercase tracking-wide hidden sm:table-cell">Ключевая проблема</th>
                    <th className="px-4 py-3"></th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.length === 0 && (
                    <tr>
                      <td colSpan={4} className="px-4 py-6 text-center eac-faint text-sm">Нет компонентов с таким уровнем риска.</td>
                    </tr>
                  )}
                  {filteredRows.map((r) => {
                    const res = results[r._id];
                    const isExpanded = expandedIds.has(r._id);
                    const label = formatRowLabel(r, labelKey);
                    return (
                      <React.Fragment key={r._id}>
                        <tr
                          className="eac-row eac-focus cursor-pointer"
                          style={{ borderTop: '1px solid var(--seam)' }}
                          onClick={() => toggleExpand(r._id)}
                          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleExpand(r._id); } }}
                          tabIndex={0}
                          role="button"
                          aria-expanded={isExpanded}
                        >
                          <td className="px-4 py-3 eac-mono">{label}</td>
                          <td className="px-4 py-3">
                            {res ? (
                              res.risk_level === 'error' ? (
                                <span className="flex items-center gap-2 text-xs eac-mono" style={{ color: '#8B8FA3' }}>
                                  <span className="led-dot led-error"></span>{RISK_LABELS.error}
                                </span>
                              ) : (
                                <span className={`flex items-center gap-2 text-xs eac-mono eac-text-${res.risk_level}`}>
                                  <span className={`led-dot led-${res.risk_level}`}></span>{RISK_LABELS[res.risk_level]}
                                </span>
                              )
                            ) : (
                              <span className="flex items-center gap-2 text-xs eac-faint">
                                <span className="led-dot led-pending"></span>...
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-3 eac-faint hidden sm:table-cell">
                            <div className="truncate max-w-xs">
                              {res && res.issues && res.issues.length > 0 ? res.issues[0] : (res ? '—' : '')}
                            </div>
                          </td>
                          <td className="px-4 py-3 text-right">
                            {isExpanded ? <ChevronDown size={16} className="eac-faint inline" /> : <ChevronRight size={16} className="eac-faint inline" />}
                          </td>
                        </tr>
                        {isExpanded && (
                          <tr style={{ borderTop: '1px solid var(--seam)' }}>
                            <td colSpan={4} className="px-4 py-4" style={{ background: 'var(--core-dark-deep)' }}>
                              {res ? (
                                <div className="space-y-3">
                                  {res.issues && res.issues.length > 0 && (
                                    <div>
                                      <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-1">Проблемы</p>
                                      <ul className="space-y-1">
                                        {res.issues.map((iss, i) => (
                                          <li key={i} className="text-sm flex items-start gap-2">
                                            <AlertTriangle size={13} className="flex-shrink-0 mt-0.5" style={{ color: 'var(--warn)' }} />
                                            <span>{iss}</span>
                                          </li>
                                        ))}
                                      </ul>
                                    </div>
                                  )}
                                  {res.recommendations && res.recommendations.length > 0 && (
                                    <div>
                                      <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-1">Рекомендации</p>
                                      <ul className="space-y-1">
                                        {res.recommendations.map((rec, i) => (
                                          <li key={i} className="text-sm flex items-start gap-2">
                                            <CheckCircle2 size={13} style={{ color: 'var(--ok)' }} className="flex-shrink-0 mt-0.5" />
                                            <span>{rec}</span>
                                          </li>
                                        ))}
                                      </ul>
                                    </div>
                                  )}
                                  {res.production_status_note && (
                                    <div>
                                      <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-1">Статус производства</p>
                                      <p className="text-sm">{res.production_status_note}</p>
                                    </div>
                                  )}
                                  <div>
                                    <p className="eac-mono eac-faint text-[11px] uppercase tracking-wide mb-1">Исходные данные</p>
                                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-1">
                                      {headers.map((h) => (
                                        <p key={h} className="text-xs eac-mono eac-faint truncate">
                                          <span className="eac-faint">{h}:</span> {String(r[h] !== undefined ? r[h] : '')}
                                        </p>
                                      ))}
                                    </div>
                                  </div>
                                </div>
                              ) : (
                                <p className="text-sm eac-faint">Анализ ещё выполняется...</p>
                              )}
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
