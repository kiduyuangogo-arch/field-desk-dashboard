import { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import _ from "lodash";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Legend, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell, LabelList,
} from "recharts";
import {
  Upload, X, FileSpreadsheet, Search, ChevronLeft, ChevronRight,
  LayoutGrid, BarChart3, Table2, MapPinned, AlertCircle, Download,
  FileText, Sparkles, Send, SlidersHorizontal,
} from "lucide-react";

const PALETTE = ["#3f6b52", "#b5652b", "#5c7a8a", "#8a6d3b", "#6b5b8a", "#4a7c6f", "#a5763f", "#527a7a"];
const PAGE_SIZE = 25;
const IDENTIFIER_THRESHOLD = 0.85;
const SYSTEM_FIELD_NAMES = new Set(["objectid", "globalid", "creationdate", "creator", "editdate", "editor", "x", "y"]);

const CHART_KIND_OPTIONS = [
  { value: "bar2d", label: "Bar" },
  { value: "bar3d", label: "Bar (3D)" },
  { value: "hbar", label: "Bar (horizontal)" },
  { value: "line", label: "Line" },
  { value: "pie", label: "Pie" },
  { value: "donut", label: "Donut" },
];

function isSystemField(name) {
  return SYSTEM_FIELD_NAMES.has(String(name).trim().toLowerCase());
}

function detectType(values) {
  const clean = values
    .map((v) => (v === null || v === undefined ? "" : String(v).trim()))
    .filter((v) => v !== "");
  if (clean.length === 0) return "empty";
  const numCount = clean.filter((v) => v !== "" && !isNaN(Number(v))).length;
  if (numCount / clean.length >= 0.9) return "numeric";
  const dateCount = clean.filter((v) => !isNaN(Date.parse(v)) && /\d/.test(v)).length;
  if (dateCount / clean.length >= 0.9) return "date";
  return "categorical";
}

function profileColumns(rows) {
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  return keys.map((name) => {
    const values = rows.map((r) => r[name]);
    const type = detectType(values);
    const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const col = { name, type, filled: nonEmpty.length, total: rows.length, systemField: isSystemField(name) };
    if (type === "numeric") {
      const nums = nonEmpty.map(Number);
      col.min = Math.min(...nums);
      col.max = Math.max(...nums);
      col.mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    }
    if (type === "categorical") {
      col.unique = new Set(nonEmpty).size;
      col.identifierLike = nonEmpty.length > 0 && col.unique / nonEmpty.length >= IDENTIFIER_THRESHOLD && col.unique > 5;
      const commaCount = nonEmpty.filter((v) => String(v).includes(",")).length;
      col.multiSelect = nonEmpty.length > 0 && commaCount / nonEmpty.length >= 0.15;
    }
    return col;
  });
}

function bucketDate(d, span) {
  const date = new Date(d);
  if (span === "day") return date.toISOString().slice(0, 10);
  if (span === "week") {
    const day = date.getDay();
    const monday = new Date(date);
    monday.setDate(date.getDate() - ((day + 6) % 7));
    return monday.toISOString().slice(0, 10);
  }
  return date.toISOString().slice(0, 7);
}

function groupCounts(rows, field, opts = {}) {
  const { measure = "count", measureCol, limit = 12, splitMulti = false } = opts;
  if (splitMulti) {
    const tally = {};
    rows.forEach((r) => {
      const raw = r[field];
      const tags = raw && String(raw).trim() !== "" ? String(raw).split(",").map((t) => t.trim()).filter(Boolean) : ["No answer"];
      tags.forEach((t) => {
        tally[t] = (tally[t] || 0) + 1;
      });
    });
    let entries = Object.entries(tally).map(([name, value]) => ({ name, value, pct: value / rows.length }));
    entries = _.orderBy(entries, "value", "desc");
    if (entries.length > limit) {
      const top = entries.slice(0, limit - 1);
      const rest = entries.slice(limit - 1).reduce((a, e) => a + e.value, 0);
      entries = [...top, { name: "Other", value: rest, pct: 0 }];
    }
    return entries;
  }
  const grouped = _.groupBy(rows, (r) => {
    const v = r[field];
    return v === "" || v === null || v === undefined ? "No answer" : String(v);
  });
  let entries = Object.entries(grouped).map(([name, rs]) => {
    if (measure === "count") return { name, value: rs.length, pct: rs.length / rows.length };
    const nums = rs.map((r) => Number(r[measureCol])).filter((n) => !isNaN(n));
    const sum = nums.reduce((a, b) => a + b, 0);
    const value = measure === "avg" ? (nums.length ? sum / nums.length : 0) : sum;
    return { name, value, pct: rs.length / rows.length };
  });
  entries = _.orderBy(entries, "value", "desc");
  if (entries.length > limit) {
    const top = entries.slice(0, limit - 1);
    const restRows = entries.slice(limit - 1).reduce((a, e) => a + e.value, 0);
    entries = [...top, { name: "Other", value: restRows, pct: 0 }];
  }
  return entries;
}

function numericHistogram(rows, field, bins = 6) {
  const nums = rows.map((r) => Number(r[field])).filter((n) => !isNaN(n));
  if (!nums.length) return [];
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  if (min === max) return [{ name: `${round(min)}`, value: nums.length, pct: 1 }];
  const width = (max - min) / bins;
  const buckets = Array.from({ length: bins }, (_, i) => ({
    name: `${round(min + i * width)}–${round(min + (i + 1) * width)}`,
    value: 0,
  }));
  nums.forEach((n) => {
    let idx = Math.floor((n - min) / width);
    if (idx >= bins) idx = bins - 1;
    if (idx < 0) idx = 0;
    buckets[idx].value++;
  });
  return buckets.map((b) => ({ ...b, pct: b.value / nums.length }));
}

function defaultKindFor(col) {
  if (col.type === "numeric") return "bar2d";
  if (col.multiSelect) return "hbar";
  if (col.unique <= 4) return "donut";
  return "hbar";
}

function fieldInsight(col, data) {
  if (!data.length) return "";
  if (col.type === "numeric") {
    return `Average ${round(col.mean)}, ranging from ${round(col.min)} to ${round(col.max)} (${col.filled} of ${col.total} answered).`;
  }
  const top = data[0];
  return `Most common answer: "${top.name}" (${Math.round((top.pct || 0) * 100)}% of respondents).`;
}

function round(n) {
  if (n === undefined || n === null || isNaN(n)) return "—";
  return Math.round(n * 100) / 100;
}

function shadeColor(hex, percent) {
  const num = parseInt(hex.replace("#", ""), 16);
  let r = (num >> 16) + percent;
  let g = ((num >> 8) & 0x00ff) + percent;
  let b = (num & 0x0000ff) + percent;
  r = Math.min(255, Math.max(0, r));
  g = Math.min(255, Math.max(0, g));
  b = Math.min(255, Math.max(0, b));
  return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0")}`;
}

// Renders a pseudo-3D bar (front, top and side faces) since SVG has no true 3D.
function ThreeDBarShape(props) {
  const { x, y, width, height, fill } = props;
  const depth = 8;
  const top = shadeColor(fill, 30);
  const side = shadeColor(fill, -25);
  if (height <= 0) return null;
  return (
    <g>
      <polygon
        points={`${x + width},${y} ${x + width + depth},${y - depth} ${x + width + depth},${y + height - depth} ${x + width},${y + height}`}
        fill={side}
      />
      <polygon points={`${x},${y} ${x + depth},${y - depth} ${x + width + depth},${y - depth} ${x + width},${y}`} fill={top} />
      <rect x={x} y={y} width={width} height={height} fill={fill} />
    </g>
  );
}

async function downloadCardPng(cardEl, filename) {
  if (!cardEl) return;
  const canvas = await html2canvas(cardEl, { backgroundColor: "#ffffff", scale: 2 });
  const a = document.createElement("a");
  a.download = `${filename}.png`;
  a.href = canvas.toDataURL("image/png");
  a.click();
}

function ContourMark() {
  return (
    <svg viewBox="0 0 160 60" className="w-32 h-12 opacity-80" aria-hidden="true">
      <path d="M2 45 Q 30 20, 60 40 T 158 30" fill="none" stroke="#d9c9a3" strokeWidth="1.5" />
      <path d="M2 55 Q 35 32, 65 50 T 158 42" fill="none" stroke="#c7b689" strokeWidth="1.5" />
      <path d="M2 30 Q 25 10, 55 26 T 158 16" fill="none" stroke="#e6dcc0" strokeWidth="1.5" />
    </svg>
  );
}

function PieCard({ data, colorIdx, donut }) {
  return (
    <ResponsiveContainer width="100%" height={210}>
      <PieChart>
        <Pie data={data} dataKey="value" nameKey="name" innerRadius={donut ? 38 : 0} outerRadius={68} paddingAngle={2}>
          {data.map((_e, i) => (
            <Cell key={i} fill={PALETTE[(colorIdx + i) % PALETTE.length]} />
          ))}
        </Pie>
        <Tooltip formatter={(v, _n, item) => [`${v} (${Math.round((item.payload.pct || 0) * 100)}%)`, "Value"]} />
        <Legend wrapperStyle={{ fontSize: 11 }} />
      </PieChart>
    </ResponsiveContainer>
  );
}

function HBarCard({ data, colorIdx }) {
  const height = Math.max(160, data.length * 26);
  return (
    <ResponsiveContainer width="100%" height={height}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 32, left: 8, bottom: 4 }}>
        <XAxis type="number" tick={{ fontSize: 10 }} />
        <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 10 }} />
        <Tooltip formatter={(v, _n, item) => [`${v} (${Math.round((item.payload.pct || 0) * 100)}%)`, "Value"]} />
        <Bar dataKey="value" radius={[0, 3, 3, 0]}>
          <LabelList dataKey="value" position="right" style={{ fontSize: 10, fill: "#44403c" }} />
          {data.map((_e, i) => (
            <Cell key={i} fill={PALETTE[(colorIdx + i) % PALETTE.length]} />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

function Bar2DCard({ data, colorIdx, threeD }) {
  return (
    <ResponsiveContainer width="100%" height={210}>
      <BarChart data={data} margin={{ top: 16, right: threeD ? 14 : 4, left: 0, bottom: 34 }}>
        <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-25} textAnchor="end" interval={0} height={50} />
        <YAxis tick={{ fontSize: 10 }} width={28} />
        <Tooltip />
        <Bar dataKey="value" radius={threeD ? 0 : [3, 3, 0, 0]} fill={PALETTE[colorIdx % PALETTE.length]} shape={threeD ? ThreeDBarShape : undefined}>
          <LabelList dataKey="value" position="top" style={{ fontSize: 10, fill: "#44403c" }} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

function LineFieldCard({ data, colorIdx }) {
  return (
    <ResponsiveContainer width="100%" height={210}>
      <LineChart data={data} margin={{ top: 16, right: 8, left: 0, bottom: 34 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d3" />
        <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-25} textAnchor="end" interval={0} height={50} />
        <YAxis tick={{ fontSize: 10 }} width={28} />
        <Tooltip />
        <Line type="monotone" dataKey="value" stroke={PALETTE[colorIdx % PALETTE.length]} strokeWidth={2} dot={{ r: 3 }} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function GenericChart({ kind, data, colorIdx }) {
  if (kind === "pie") return <PieCard data={data} colorIdx={colorIdx} donut={false} />;
  if (kind === "donut") return <PieCard data={data} colorIdx={colorIdx} donut />;
  if (kind === "hbar") return <HBarCard data={data} colorIdx={colorIdx} />;
  if (kind === "line") return <LineFieldCard data={data} colorIdx={colorIdx} />;
  if (kind === "bar3d") return <Bar2DCard data={data} colorIdx={colorIdx} threeD />;
  return <Bar2DCard data={data} colorIdx={colorIdx} />;
}

function FieldChart({ col, rows, colorIdx, kind, onChangeKind }) {
  const ref = useRef(null);
  const data = useMemo(() => {
    if (col.type === "numeric") return numericHistogram(rows, col.name);
    if (col.multiSelect) return groupCounts(rows, col.name, { limit: 10, splitMulti: true });
    if (kind === "pie" || kind === "donut") return groupCounts(rows, col.name, { limit: 6 });
    return groupCounts(rows, col.name, { limit: 10 });
  }, [rows, col, kind]);

  const insight = fieldInsight(col, data);

  return (
    <div
      ref={ref}
      data-report-chart={col.name}
      className="bg-white border border-stone-300 rounded-md p-4 hover:shadow-md hover:border-stone-400 transition-all"
    >
      <div className="flex items-start justify-between gap-2 mb-2">
        <h3 className="text-sm text-stone-800 leading-snug" title={col.name}>{col.name}</h3>
        <button
          onClick={() => downloadCardPng(ref.current, col.name)}
          className="text-stone-400 hover:text-stone-700 shrink-0"
          title="Download this chart as a PNG"
        >
          <Download className="w-3.5 h-3.5" />
        </button>
      </div>
      <select
        value={kind}
        onChange={(e) => onChangeKind(col.name, e.target.value)}
        className="text-xs border border-stone-300 rounded-sm px-1.5 py-0.5 mb-2 text-stone-600 bg-stone-50"
      >
        {CHART_KIND_OPTIONS.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
      <GenericChart kind={kind} data={data} colorIdx={colorIdx} />
      {insight && <p className="text-xs text-stone-500 mt-2 border-t border-stone-100 pt-2">{insight}</p>}
    </div>
  );
}

// Mounted off-screen at all times so a PDF report can be generated from any tab,
// with a chart (matching whatever type the person chose) for every field.
function ReportRenderRoot({ source, kindFor }) {
  if (!source) return null;
  const fields = source.columns.filter((c) => !c.systemField && ((c.type === "categorical" && !c.identifierLike) || c.type === "numeric"));
  return (
    <div id="report-render-root" style={{ position: "absolute", left: -10000, top: 0, width: 560 }} aria-hidden="true">
      {fields.map((c, i) => (
        <div key={c.name} style={{ width: 560, marginBottom: 24 }}>
          <FieldChart col={c} rows={source.rows} colorIdx={i} kind={kindFor(c)} onChangeKind={() => {}} />
        </div>
      ))}
    </div>
  );
}

function StatCard({ label, value }) {
  return (
    <div className="bg-white border border-stone-300 rounded-md p-4 hover:shadow-sm transition-shadow">
      <div className="text-sm text-stone-500">{label}</div>
      <div className="font-serif text-2xl text-stone-900 mt-1">{value}</div>
    </div>
  );
}

export default function FieldDesk() {
  const [sources, setSources] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [tab, setTab] = useState("overview");
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [search, setSearch] = useState("");
  const [fieldSearch, setFieldSearch] = useState("");
  const [page, setPage] = useState(1);
  const [chartConfig, setChartConfig] = useState({});
  const [typeOverrides, setTypeOverrides] = useState({});
  const [chatMessages, setChatMessages] = useState({});
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const [reportBusy, setReportBusy] = useState(false);

  const active = sources.find((s) => s.id === activeId) || null;

  const addSource = useCallback((name, rows) => {
    const cleanRows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").trim() !== ""));
    if (!cleanRows.length) {
      setError(`${name}: no readable rows found.`);
      return;
    }
    const columns = profileColumns(cleanRows);
    const id = `${name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    setSources((prev) => [...prev, { id, name, rows: cleanRows, columns }]);
    setActiveId(id);
    setTab("overview");
    setError("");
  }, []);

  const handleFiles = useCallback((fileList) => {
    Array.from(fileList).forEach((file) => {
      const ext = file.name.split(".").pop().toLowerCase();
      if (ext === "csv") {
        Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          complete: (res) => addSource(file.name, res.data),
          error: (err) => setError(`${file.name}: ${err.message}`),
        });
      } else if (ext === "xlsx" || ext === "xls") {
        const reader = new FileReader();
        reader.onload = (evt) => {
          try {
            const wb = XLSX.read(evt.target.result, { type: "array" });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            addSource(file.name, data);
          } catch (err) {
            setError(`${file.name}: couldn't read that file. ${err.message}`);
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        setError(`${file.name}: unsupported format. Use CSV or Excel exports.`);
      }
    });
  }, [addSource]);

  const removeSource = (id) => {
    setSources((prev) => prev.filter((s) => s.id !== id));
    if (activeId === id) {
      const remaining = sources.filter((s) => s.id !== id);
      setActiveId(remaining.length ? remaining[0].id : null);
    }
  };

  const kindFor = useCallback((col) => {
    const key = `${activeId}:${col.name}`;
    return typeOverrides[key] || defaultKindFor(col);
  }, [activeId, typeOverrides]);

  const setKindFor = (fieldName, value) => {
    setTypeOverrides((prev) => ({ ...prev, [`${activeId}:${fieldName}`]: value }));
  };

  const allCatCols = active ? active.columns.filter((c) => !c.systemField && c.type === "categorical" && c.unique <= 40) : [];
  const catCols = allCatCols.filter((c) => !c.identifierLike);
  const identifierCols = allCatCols.filter((c) => c.identifierLike);
  const numCols = active ? active.columns.filter((c) => !c.systemField && c.type === "numeric") : [];
  const dateCols = active ? active.columns.filter((c) => c.type === "date") : [];
  const systemCols = active ? active.columns.filter((c) => c.systemField) : [];

  const chartableFields = active
    ? active.columns.filter((c) => !c.systemField && ((c.type === "categorical" && !c.identifierLike && c.unique <= 40) || c.type === "numeric"))
    : [];
  const filteredFields = chartableFields.filter((c) => c.name.toLowerCase().includes(fieldSearch.toLowerCase()));

  const cfg = chartConfig[activeId] || {
    groupCol: catCols[0]?.name || "",
    measure: "count",
    measureCol: numCols[0]?.name || "",
    timeBucket: "month",
    dateCol: dateCols[0]?.name || "",
  };
  const setCfg = (patch) =>
    setChartConfig((prev) => ({ ...prev, [activeId]: { ...cfg, ...patch } }));

  const selectedGroupCol = allCatCols.find((c) => c.name === cfg.groupCol);
  const groupIsMulti = !!selectedGroupCol?.multiSelect;

  const groupChartData = useMemo(() => {
    if (!active || !cfg.groupCol) return [];
    return groupCounts(active.rows, cfg.groupCol, {
      measure: groupIsMulti ? "count" : cfg.measure,
      measureCol: cfg.measureCol,
      splitMulti: groupIsMulti,
    });
  }, [active, cfg.groupCol, cfg.measure, cfg.measureCol, groupIsMulti]);

  const timeChartData = useMemo(() => {
    if (!active || !cfg.dateCol) return [];
    const grouped = _.groupBy(
      active.rows.filter((r) => r[cfg.dateCol] && !isNaN(Date.parse(r[cfg.dateCol]))),
      (r) => bucketDate(r[cfg.dateCol], cfg.timeBucket)
    );
    return Object.entries(grouped)
      .map(([name, rows]) => ({ name, value: rows.length }))
      .sort((a, b) => (a.name > b.name ? 1 : -1));
  }, [active, cfg.dateCol, cfg.timeBucket]);

  const filteredRows = useMemo(() => {
    if (!active) return [];
    if (!search.trim()) return active.rows;
    const q = search.toLowerCase();
    return active.rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
  }, [active, search]);

  const pageCount = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  const pageRows = filteredRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files?.length) handleFiles(e.dataTransfer.files);
  };

  const buildProfile = useCallback((source) => {
    const lines = [];
    lines.push(`Source: ${source.name}`);
    lines.push(`Rows: ${source.rows.length}, Columns: ${source.columns.length}`);
    source.columns.filter((c) => !c.systemField).forEach((c) => {
      if (c.type === "numeric") {
        lines.push(`- ${c.name} (numeric): min ${round(c.min)}, mean ${round(c.mean)}, max ${round(c.max)}, filled ${c.filled}/${c.total}`);
      } else if (c.type === "categorical") {
        const top = groupCounts(source.rows, c.name, { limit: 6, splitMulti: c.multiSelect })
          .map((e) => `${e.name} (${e.value}, ${Math.round((e.pct || 0) * 100)}%)`)
          .join(", ");
        lines.push(`- ${c.name} (categorical, ${c.unique} unique${c.multiSelect ? ", multi-select" : ""}): top values ${top}`);
      } else if (c.type === "date") {
        lines.push(`- ${c.name} (date field), filled ${c.filled}/${c.total}`);
      }
    });
    const sysNames = source.columns.filter((c) => c.systemField).map((c) => c.name);
    if (sysNames.length) {
      lines.push(`\nSystem/location fields (not treated as survey questions): ${sysNames.join(", ")}`);
    }
    return lines.join("\n");
  }, []);

  const generateReport = async () => {
    if (!active) return;
    setReportBusy(true);
    try {
      const doc = new jsPDF({ unit: "pt", format: "a4" });
      const margin = 48;
      let y = margin;
      doc.setFont("times", "bold");
      doc.setFontSize(20);
      doc.text("Field Desk report", margin, y);
      y += 22;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.setTextColor(90);
      doc.text(`${active.name} — generated ${new Date().toLocaleString()}`, margin, y);
      y += 24;
      doc.setTextColor(20);
      doc.setFontSize(12);
      doc.text(`${active.rows.length.toLocaleString()} rows across ${active.columns.length} fields.`, margin, y);
      y += 20;

      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.text("Question-by-question summary", margin, y);
      y += 8;
      doc.setDrawColor(200);
      doc.line(margin, y, 547, y);
      y += 16;

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      active.columns.filter((c) => !c.systemField).forEach((c) => {
        let line;
        if (c.type === "numeric") {
          line = `${c.name} — numeric: min ${round(c.min)}, mean ${round(c.mean)}, max ${round(c.max)} (${c.filled}/${c.total} answered)`;
        } else if (c.type === "categorical") {
          const top = groupCounts(active.rows, c.name, { limit: 4, splitMulti: c.multiSelect })
            .map((e) => `${e.name} ${Math.round((e.pct || 0) * 100)}%`)
            .join(", ");
          line = `${c.name}${c.multiSelect ? " (multiple answers allowed)" : ""} — ${top}`;
        } else {
          return;
        }
        const wrapped = doc.splitTextToSize(`•  ${line}`, 500);
        wrapped.forEach((wline, idx) => {
          if (y > 770) {
            doc.addPage();
            y = margin;
          }
          doc.text(wline, margin + (idx > 0 ? 12 : 0), y);
          y += 12;
        });
        y += 4;
      });

      const sysNames = active.columns.filter((c) => c.systemField).map((c) => c.name);
      if (sysNames.length) {
        if (y > 740) {
          doc.addPage();
          y = margin;
        }
        y += 8;
        doc.setFont("helvetica", "italic");
        doc.setFontSize(8);
        doc.setTextColor(120);
        doc.text(`System/location fields excluded from analysis: ${sysNames.join(", ")}`, margin, y);
        doc.setTextColor(20);
      }

      const chartNodes = document.querySelectorAll("#report-render-root [data-report-chart]");
      let imgY = margin;
      let onPage = 0;
      for (const node of chartNodes) {
        const canvas = await html2canvas(node, { backgroundColor: "#ffffff", scale: 2 });
        const dataUrl = canvas.toDataURL("image/png");
        const imgWidth = 480;
        const imgHeight = (canvas.height / canvas.width) * imgWidth;

        if (onPage === 0) {
          doc.addPage();
          imgY = margin;
        }
        doc.addImage(dataUrl, "PNG", margin, imgY, imgWidth, imgHeight);
        imgY += imgHeight + 26;
        onPage++;
        if (imgY > 660) onPage = 0;
      }

      doc.save(`${active.name.replace(/\.[^.]+$/, "")}-report.pdf`);
    } catch (err) {
      setError(`Couldn't generate the report: ${err.message}`);
    } finally {
      setReportBusy(false);
    }
  };

  const messages = chatMessages[activeId] || [];
  const setMessages = (updater) =>
    setChatMessages((prev) => ({ ...prev, [activeId]: typeof updater === "function" ? updater(prev[activeId] || []) : updater }));

  const sendChat = async () => {
    if (!chatInput.trim() || !active) return;
    const question = chatInput.trim();
    setChatInput("");
    setMessages((prev) => [...prev, { role: "user", text: question }]);
    setChatLoading(true);
    try {
      const res = await fetch("/.netlify/functions/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question, profile: buildProfile(active) }),
      });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.error || `Request failed (${res.status})`);
      }
      const data = await res.json();
      setMessages((prev) => [...prev, { role: "assistant", text: data.answer }]);
    } catch (err) {
      setMessages((prev) => [...prev, { role: "assistant", text: `Couldn't reach the assistant: ${err.message}`, isError: true }]);
    } finally {
      setChatLoading(false);
    }
  };

  return (
    <div className="w-full min-h-screen bg-stone-100 text-stone-900 font-sans flex flex-col">
      <header className="bg-emerald-900 text-stone-50 px-6 py-5 flex items-center justify-between">
        <div>
          <div className="flex items-center gap-2">
            <MapPinned className="w-5 h-5 text-amber-300" />
            <h1 className="font-serif text-2xl">Field Desk</h1>
          </div>
          <p className="text-emerald-100 text-sm mt-1">
            Bring in your survey exports and see what they say.
          </p>
        </div>
        <ContourMark />
      </header>

      {error && (
        <div className="bg-amber-50 border-b border-amber-300 text-amber-900 text-sm px-6 py-2 flex items-center gap-2">
          <AlertCircle className="w-4 h-4 shrink-0" />
          <span>{error}</span>
          <button onClick={() => setError("")} className="ml-auto text-amber-700 hover:text-amber-900">
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      <div className="flex flex-1 flex-col md:flex-row">
        <aside className="w-full md:w-64 shrink-0 bg-stone-50 border-b md:border-b-0 md:border-r border-stone-300 p-4">
          <label className="flex items-center justify-center gap-2 border-2 border-dashed border-stone-300 rounded p-4 text-sm text-stone-600 cursor-pointer hover:border-emerald-700 hover:text-emerald-800 transition-colors">
            <Upload className="w-4 h-4" />
            Add a file
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              multiple
              className="hidden"
              onChange={(e) => e.target.files?.length && handleFiles(e.target.files)}
            />
          </label>

          <p className="text-xs text-stone-500 mt-3">
            Works with exports from Survey123, KoboToolbox, Google Forms, or any CSV/Excel file.
          </p>

          {sources.length > 0 && (
            <ul className="mt-4 space-y-1">
              {sources.map((s) => (
                <li key={s.id}>
                  <div
                    className={`group flex items-center gap-2 px-3 py-2 rounded-sm cursor-pointer text-sm ${
                      s.id === activeId ? "bg-emerald-800 text-white" : "text-stone-700 hover:bg-stone-200"
                    }`}
                    onClick={() => {
                      setActiveId(s.id);
                      setTab("overview");
                      setPage(1);
                      setSearch("");
                      setFieldSearch("");
                    }}
                  >
                    <FileSpreadsheet className="w-4 h-4 shrink-0" />
                    <div className="min-w-0 flex-1">
                      <div className="truncate">{s.name}</div>
                      <div className={`text-xs ${s.id === activeId ? "text-emerald-200" : "text-stone-500"}`}>
                        {s.rows.length.toLocaleString()} rows
                      </div>
                    </div>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        removeSource(s.id);
                      }}
                      className={`opacity-0 group-hover:opacity-100 shrink-0 ${
                        s.id === activeId ? "text-emerald-200 hover:text-white" : "text-stone-400 hover:text-stone-700"
                      }`}
                      aria-label={`Remove ${s.name}`}
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </aside>

        <main className="flex-1 p-6">
          {!active ? (
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              className={`border-2 border-dashed rounded p-16 text-center transition-colors ${
                dragOver ? "border-emerald-600 bg-emerald-50" : "border-stone-300 bg-white"
              }`}
            >
              <Upload className="w-8 h-8 mx-auto text-stone-400 mb-3" />
              <p className="font-serif text-lg text-stone-700">Drop a survey export here</p>
              <p className="text-sm text-stone-500 mt-1 mb-4">CSV or Excel, from Survey123, Kobo, Google Forms, or elsewhere.</p>
              <label className="inline-block bg-emerald-800 text-white text-sm px-4 py-2 rounded-sm cursor-pointer hover:bg-emerald-900">
                Choose a file
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  multiple
                  className="hidden"
                  onChange={(e) => e.target.files?.length && handleFiles(e.target.files)}
                />
              </label>
            </div>
          ) : (
            <>
              <nav className="flex gap-6 border-b border-stone-300 mb-6 overflow-x-auto">
                {[
                  { id: "overview", label: "Overview", icon: LayoutGrid },
                  { id: "charts", label: "Charts", icon: BarChart3 },
                  { id: "table", label: "Table", icon: Table2 },
                  { id: "report", label: "Report", icon: FileText },
                  { id: "assistant", label: "Ask Claude", icon: Sparkles },
                ].map(({ id, label, icon: Icon }) => (
                  <button
                    key={id}
                    onClick={() => setTab(id)}
                    className={`flex items-center gap-1.5 pb-3 px-1 text-sm whitespace-nowrap rounded-t-sm transition-colors ${
                      tab === id
                        ? "border-b-2 border-amber-700 text-stone-900 font-medium"
                        : "text-stone-500 hover:text-stone-800 hover:bg-stone-50"
                    }`}
                  >
                    <Icon className="w-4 h-4" />
                    {label}
                  </button>
                ))}
              </nav>

              {tab === "overview" && (
                <div className="space-y-6">
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                    <StatCard label="Rows" value={active.rows.length.toLocaleString()} />
                    <StatCard label="Columns" value={active.columns.length} />
                    <StatCard label="Categorical fields" value={catCols.length} />
                    <StatCard label="Numeric fields" value={numCols.length} />
                  </div>

                  {(identifierCols.length > 0 || systemCols.length > 0) && (
                    <p className="text-xs text-stone-500 bg-stone-50 border border-stone-200 rounded-sm px-3 py-2">
                      {systemCols.length > 0 && (
                        <>System fields ({systemCols.map((c) => c.name).join(", ")}) are Survey123/GIS metadata, not questions, so they're kept out of the charts below. </>
                      )}
                      {identifierCols.length > 0 && (
                        <>{identifierCols.map((c) => c.name).join(", ")} {identifierCols.length === 1 ? "looks" : "look"} like identifier fields (nearly every value unique), so {identifierCols.length === 1 ? "it's" : "they're"} left out too. </>
                      )}
                      All of these are still visible in the Table tab.
                    </p>
                  )}

                  <div className="flex items-center justify-between">
                    <h2 className="font-serif text-lg">All questions ({filteredFields.length})</h2>
                    <div className="flex items-center gap-2 border border-stone-300 rounded-sm bg-white px-2 py-1">
                      <Search className="w-3.5 h-3.5 text-stone-400" />
                      <input
                        value={fieldSearch}
                        onChange={(e) => setFieldSearch(e.target.value)}
                        placeholder="Filter questions"
                        className="text-sm outline-none w-48"
                      />
                    </div>
                  </div>

                  {filteredFields.length === 0 ? (
                    <p className="text-sm text-stone-500">No fields match that filter.</p>
                  ) : (
                    <div className="grid sm:grid-cols-2 xl:grid-cols-3 gap-4">
                      {filteredFields.map((c, i) => (
                        <FieldChart
                          key={c.name}
                          col={c}
                          rows={active.rows}
                          colorIdx={i}
                          kind={kindFor(c)}
                          onChangeKind={setKindFor}
                        />
                      ))}
                    </div>
                  )}
                </div>
              )}

              {tab === "charts" && (
                <div className="space-y-6">
                  <div className="bg-white border border-stone-300 rounded-md p-4">
                    <h2 className="font-serif text-base mb-4 flex items-center gap-2">
                      <SlidersHorizontal className="w-4 h-4 text-stone-500" />
                      Custom pivot
                    </h2>
                    <div className="flex flex-col md:flex-row gap-6">
                      <div className="w-full md:w-56 shrink-0 space-y-4 text-sm">
                        <label className="block">
                          <span className="block text-xs text-stone-500 mb-1">Group by</span>
                          <select
                            value={cfg.groupCol}
                            onChange={(e) => setCfg({ groupCol: e.target.value })}
                            className="w-full border border-stone-300 rounded-sm px-2 py-1.5"
                          >
                            {catCols.map((c) => (
                              <option key={c.name} value={c.name}>{c.name}{c.multiSelect ? " (multi-select)" : ""}</option>
                            ))}
                            {identifierCols.length > 0 && (
                              <optgroup label="Identifier-like fields">
                                {identifierCols.map((c) => (
                                  <option key={c.name} value={c.name}>{c.name}</option>
                                ))}
                              </optgroup>
                            )}
                          </select>
                        </label>
                        {!groupIsMulti && (
                          <label className="block">
                            <span className="block text-xs text-stone-500 mb-1">Measure</span>
                            <select
                              value={cfg.measure}
                              onChange={(e) => setCfg({ measure: e.target.value })}
                              className="w-full border border-stone-300 rounded-sm px-2 py-1.5"
                            >
                              <option value="count">Count of rows</option>
                              {numCols.length > 0 && <option value="sum">Sum of</option>}
                              {numCols.length > 0 && <option value="avg">Average of</option>}
                            </select>
                          </label>
                        )}
                        {!groupIsMulti && cfg.measure !== "count" && (
                          <label className="block">
                            <span className="block text-xs text-stone-500 mb-1">Numeric field</span>
                            <select
                              value={cfg.measureCol}
                              onChange={(e) => setCfg({ measureCol: e.target.value })}
                              className="w-full border border-stone-300 rounded-sm px-2 py-1.5"
                            >
                              {numCols.map((c) => (
                                <option key={c.name} value={c.name}>{c.name}</option>
                              ))}
                            </select>
                          </label>
                        )}
                        <button
                          onClick={() => downloadCardPng(document.getElementById("group-chart"), cfg.groupCol || "chart")}
                          className="flex items-center gap-1.5 text-xs text-stone-600 border border-stone-300 rounded-sm px-2 py-1.5 hover:bg-stone-100 w-full justify-center"
                        >
                          <Download className="w-3.5 h-3.5" />
                          Download chart
                        </button>
                      </div>
                      <div className="flex-1 min-w-0">
                        {catCols.length === 0 ? (
                          <p className="text-sm text-stone-500">No categorical fields to group by.</p>
                        ) : (
                          <div id="group-chart" className="bg-white">
                            <ResponsiveContainer width="100%" height={340}>
                              <BarChart data={groupChartData} margin={{ top: 20, right: 8, left: 0, bottom: 50 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d3" />
                                <XAxis dataKey="name" tick={{ fontSize: 11 }} angle={-30} textAnchor="end" interval={0} />
                                <YAxis tick={{ fontSize: 11 }} />
                                <Tooltip formatter={(value, _n, item) => [`${round(value)} (${Math.round((item.payload.pct || 0) * 100)}%)`, "Value"]} />
                                <Bar dataKey="value" radius={[3, 3, 0, 0]}>
                                  <LabelList dataKey="value" position="top" style={{ fontSize: 11, fill: "#44403c" }} formatter={(v) => round(v)} />
                                  {groupChartData.map((_entry, i) => (
                                    <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
                                  ))}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {dateCols.length > 0 && (
                    <div className="bg-white border border-stone-300 rounded-md p-4">
                      <h2 className="font-serif text-base mb-4">Responses over time</h2>
                      <div className="flex flex-col md:flex-row gap-6">
                        <div className="w-full md:w-56 shrink-0 space-y-4 text-sm">
                          <label className="block">
                            <span className="block text-xs text-stone-500 mb-1">Date field</span>
                            <select
                              value={cfg.dateCol}
                              onChange={(e) => setCfg({ dateCol: e.target.value })}
                              className="w-full border border-stone-300 rounded-sm px-2 py-1.5"
                            >
                              {dateCols.map((c) => (
                                <option key={c.name} value={c.name}>{c.name}</option>
                              ))}
                            </select>
                          </label>
                          <label className="block">
                            <span className="block text-xs text-stone-500 mb-1">Group by</span>
                            <select
                              value={cfg.timeBucket}
                              onChange={(e) => setCfg({ timeBucket: e.target.value })}
                              className="w-full border border-stone-300 rounded-sm px-2 py-1.5"
                            >
                              <option value="day">Day</option>
                              <option value="week">Week</option>
                              <option value="month">Month</option>
                            </select>
                          </label>
                          <button
                            onClick={() => downloadCardPng(document.getElementById("time-chart"), `${cfg.dateCol}-over-time`)}
                            className="flex items-center gap-1.5 text-xs text-stone-600 border border-stone-300 rounded-sm px-2 py-1.5 hover:bg-stone-100 w-full justify-center"
                          >
                            <Download className="w-3.5 h-3.5" />
                            Download chart
                          </button>
                        </div>
                        <div className="flex-1 min-w-0">
                          <div id="time-chart" className="bg-white">
                            <ResponsiveContainer width="100%" height={280}>
                              <LineChart data={timeChartData} margin={{ top: 8, right: 8, left: 0, bottom: 8 }}>
                                <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d3" />
                                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                                <YAxis tick={{ fontSize: 11 }} />
                                <Tooltip />
                                <Line type="monotone" dataKey="value" stroke="#3f6b52" strokeWidth={2} dot={false} />
                              </LineChart>
                            </ResponsiveContainer>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {tab === "table" && (
                <div className="bg-white border border-stone-300 rounded-md">
                  <div className="p-3 border-b border-stone-200 flex items-center gap-2">
                    <Search className="w-4 h-4 text-stone-400" />
                    <input
                      value={search}
                      onChange={(e) => {
                        setSearch(e.target.value);
                        setPage(1);
                      }}
                      placeholder="Search this data"
                      className="flex-1 text-sm outline-none"
                    />
                    <span className="text-xs text-stone-500">
                      {filteredRows.length.toLocaleString()} of {active.rows.length.toLocaleString()} rows
                    </span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="bg-stone-50 border-b border-stone-200">
                          {active.columns.map((c) => (
                            <th key={c.name} className="text-left font-medium text-stone-600 px-3 py-2 whitespace-nowrap">
                              {c.name}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {pageRows.map((r, i) => (
                          <tr key={i} className={i % 2 ? "bg-stone-50" : "bg-white"}>
                            {active.columns.map((c) => (
                              <td
                                key={c.name}
                                className={`px-3 py-1.5 whitespace-nowrap ${
                                  c.type === "numeric" ? "font-mono text-right" : ""
                                }`}
                              >
                                {String(r[c.name] ?? "")}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="flex items-center justify-between p-3 border-t border-stone-200 text-sm">
                    <span className="text-stone-500">
                      Page {page} of {pageCount}
                    </span>
                    <div className="flex gap-2">
                      <button
                        disabled={page <= 1}
                        onClick={() => setPage((p) => p - 1)}
                        className="p-1.5 border border-stone-300 rounded-sm disabled:opacity-40 hover:bg-stone-100"
                      >
                        <ChevronLeft className="w-4 h-4" />
                      </button>
                      <button
                        disabled={page >= pageCount}
                        onClick={() => setPage((p) => p + 1)}
                        className="p-1.5 border border-stone-300 rounded-sm disabled:opacity-40 hover:bg-stone-100"
                      >
                        <ChevronRight className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {tab === "report" && (
                <div className="bg-white border border-stone-300 rounded-md p-6 max-w-2xl">
                  <h2 className="font-serif text-lg mb-2">Dataset report</h2>
                  <p className="text-sm text-stone-600 mb-4">
                    Generates a PDF with a question-by-question summary and a chart (matching whatever
                    chart type you picked on the Overview tab) for every question, complete with its
                    legend, border, and a one-line takeaway underneath.
                  </p>
                  <button
                    onClick={generateReport}
                    disabled={reportBusy}
                    className="flex items-center gap-2 bg-emerald-800 text-white text-sm px-4 py-2 rounded-sm hover:bg-emerald-900 disabled:opacity-50"
                  >
                    <FileText className="w-4 h-4" />
                    {reportBusy ? "Building report…" : "Download PDF report"}
                  </button>
                  <div className="mt-6 text-xs text-stone-500 font-mono whitespace-pre-wrap border-t border-stone-200 pt-4">
                    {buildProfile(active)}
                  </div>
                </div>
              )}

              {tab === "assistant" && (
                <div className="bg-white border border-stone-300 rounded-md p-4 max-w-2xl flex flex-col h-[520px]">
                  <div className="flex items-center gap-2 mb-3">
                    <Sparkles className="w-4 h-4 text-amber-700" />
                    <h2 className="font-serif text-base">Ask Claude about this data</h2>
                  </div>
                  <div className="flex-1 overflow-y-auto space-y-3 mb-3 pr-1">
                    {messages.length === 0 && (
                      <p className="text-sm text-stone-500">
                        Ask things like "what stands out in this data?" or "summarize the gender split."
                        Claude sees a summary of your fields, not the raw rows.
                      </p>
                    )}
                    {messages.map((m, i) => (
                      <div
                        key={i}
                        className={`text-sm rounded-sm px-3 py-2 max-w-[85%] ${
                          m.role === "user"
                            ? "bg-emerald-800 text-white ml-auto"
                            : m.isError
                              ? "bg-amber-50 text-amber-900 border border-amber-300"
                              : "bg-stone-100 text-stone-800"
                        }`}
                      >
                        {m.text}
                      </div>
                    ))}
                    {chatLoading && <div className="text-sm text-stone-400">Thinking…</div>}
                  </div>
                  <div className="flex gap-2 border-t border-stone-200 pt-3">
                    <input
                      value={chatInput}
                      onChange={(e) => setChatInput(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && sendChat()}
                      placeholder="Ask a question about this dataset"
                      className="flex-1 border border-stone-300 rounded-sm px-3 py-2 text-sm outline-none focus:border-emerald-700"
                    />
                    <button
                      onClick={sendChat}
                      disabled={chatLoading}
                      className="bg-emerald-800 text-white px-3 py-2 rounded-sm hover:bg-emerald-900 disabled:opacity-50"
                    >
                      <Send className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              )}
            </>
          )}
        </main>
      </div>
      <ReportRenderRoot source={active} kindFor={kindFor} />
    </div>
  );
}
