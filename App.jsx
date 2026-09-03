import { useState, useMemo, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import _ from "lodash";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell,
} from "recharts";
import {
  Upload, X, FileSpreadsheet, Search, ChevronLeft, ChevronRight,
  LayoutGrid, BarChart3, Table2, MapPinned, AlertCircle,
} from "lucide-react";

const PALETTE = ["#3f6b52", "#b5652b", "#5c7a8a", "#8a6d3b", "#6b5b8a", "#4a7c6f"];
const PAGE_SIZE = 25;

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
    const col = { name, type, filled: nonEmpty.length, total: rows.length };
    if (type === "numeric") {
      const nums = nonEmpty.map(Number);
      col.min = Math.min(...nums);
      col.max = Math.max(...nums);
      col.mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    }
    if (type === "categorical") {
      col.unique = new Set(nonEmpty).size;
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

function ContourMark() {
  return (
    <svg viewBox="0 0 160 60" className="w-32 h-12 opacity-80" aria-hidden="true">
      <path d="M2 45 Q 30 20, 60 40 T 158 30" fill="none" stroke="#d9c9a3" strokeWidth="1.5" />
      <path d="M2 55 Q 35 32, 65 50 T 158 42" fill="none" stroke="#c7b689" strokeWidth="1.5" />
      <path d="M2 30 Q 25 10, 55 26 T 158 16" fill="none" stroke="#e6dcc0" strokeWidth="1.5" />
    </svg>
  );
}

export default function FieldDesk() {
  const [sources, setSources] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [tab, setTab] = useState("overview");
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [chartConfig, setChartConfig] = useState({});

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

  const catCols = active ? active.columns.filter((c) => c.type === "categorical" && c.unique <= 40) : [];
  const numCols = active ? active.columns.filter((c) => c.type === "numeric") : [];
  const dateCols = active ? active.columns.filter((c) => c.type === "date") : [];

  const cfg = chartConfig[activeId] || {
    groupCol: catCols[0]?.name || "",
    measure: "count",
    measureCol: numCols[0]?.name || "",
    timeBucket: "month",
    dateCol: dateCols[0]?.name || "",
  };
  const setCfg = (patch) =>
    setChartConfig((prev) => ({ ...prev, [activeId]: { ...cfg, ...patch } }));

  const groupChartData = useMemo(() => {
    if (!active || !cfg.groupCol) return [];
    const grouped = _.groupBy(active.rows, (r) => {
      const v = r[cfg.groupCol];
      return v === "" || v === null || v === undefined ? "(blank)" : String(v);
    });
    let entries = Object.entries(grouped).map(([name, rows]) => {
      if (cfg.measure === "count") return { name, value: rows.length };
      const nums = rows.map((r) => Number(r[cfg.measureCol])).filter((n) => !isNaN(n));
      const sum = nums.reduce((a, b) => a + b, 0);
      return { name, value: cfg.measure === "avg" ? (nums.length ? sum / nums.length : 0) : sum };
    });
    entries = _.orderBy(entries, "value", "desc");
    if (entries.length > 12) {
      const top = entries.slice(0, 11);
      const rest = entries.slice(11).reduce((a, e) => a + e.value, 0);
      entries = [...top, { name: "Other", value: rest }];
    }
    return entries;
  }, [active, cfg.groupCol, cfg.measure, cfg.measureCol]);

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
              <nav className="flex gap-6 border-b border-stone-300 mb-6">
                {[
                  { id: "overview", label: "Overview", icon: LayoutGrid },
                  { id: "charts", label: "Charts", icon: BarChart3 },
                  { id: "table", label: "Table", icon: Table2 },
                ].map(({ id, label, icon: Icon }) => (
                  <button
                    key={id}
                    onClick={() => setTab(id)}
                    className={`flex items-center gap-1.5 pb-3 text-sm ${
                      tab === id
                        ? "border-b-2 border-amber-700 text-stone-900 font-medium"
                        : "text-stone-500 hover:text-stone-800"
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

                  {numCols.length > 0 && (
                    <div className="bg-white border border-stone-300 rounded-sm p-4">
                      <h2 className="font-serif text-base mb-3">Numeric fields</h2>
                      <div className="grid sm:grid-cols-2 gap-3 text-sm font-mono">
                        {numCols.map((c) => (
                          <div key={c.name} className="flex justify-between border-b border-stone-100 pb-1">
                            <span className="text-stone-600 truncate pr-2">{c.name}</span>
                            <span className="text-stone-900 shrink-0">
                              min {round(c.min)} · mean {round(c.mean)} · max {round(c.max)}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="grid md:grid-cols-2 gap-4">
                    {catCols.slice(0, 4).map((c, i) => (
                      <MiniBarChart key={c.name} title={c.name} rows={active.rows} colorIdx={i} />
                    ))}
                  </div>

                  {catCols.length === 0 && numCols.length === 0 && (
                    <p className="text-sm text-stone-500">
                      No categorical or numeric fields were detected — check the Table tab to see the raw data.
                    </p>
                  )}
                </div>
              )}

              {tab === "charts" && (
                <div className="space-y-6">
                  <div className="bg-white border border-stone-300 rounded-sm p-4">
                    <h2 className="font-serif text-base mb-3">Break down by field</h2>
                    <div className="flex flex-wrap gap-3 mb-4 text-sm">
                      <label className="flex items-center gap-2">
                        Group by
                        <select
                          value={cfg.groupCol}
                          onChange={(e) => setCfg({ groupCol: e.target.value })}
                          className="border border-stone-300 rounded-sm px-2 py-1"
                        >
                          {catCols.map((c) => (
                            <option key={c.name} value={c.name}>{c.name}</option>
                          ))}
                        </select>
                      </label>
                      <label className="flex items-center gap-2">
                        Measure
                        <select
                          value={cfg.measure}
                          onChange={(e) => setCfg({ measure: e.target.value })}
                          className="border border-stone-300 rounded-sm px-2 py-1"
                        >
                          <option value="count">Count of rows</option>
                          {numCols.length > 0 && <option value="sum">Sum of</option>}
                          {numCols.length > 0 && <option value="avg">Average of</option>}
                        </select>
                      </label>
                      {cfg.measure !== "count" && (
                        <select
                          value={cfg.measureCol}
                          onChange={(e) => setCfg({ measureCol: e.target.value })}
                          className="border border-stone-300 rounded-sm px-2 py-1 text-sm"
                        >
                          {numCols.map((c) => (
                            <option key={c.name} value={c.name}>{c.name}</option>
                          ))}
                        </select>
                      )}
                    </div>
                    {catCols.length === 0 ? (
                      <p className="text-sm text-stone-500">No categorical fields to group by.</p>
                    ) : (
                      <ResponsiveContainer width="100%" height={320}>
                        <BarChart data={groupChartData} margin={{ top: 8, right: 8, left: 0, bottom: 40 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#e7e2d3" />
                          <XAxis dataKey="name" tick={{ fontSize: 11 }} angle={-30} textAnchor="end" interval={0} />
                          <YAxis tick={{ fontSize: 11 }} />
                          <Tooltip />
                          <Bar dataKey="value" radius={[3, 3, 0, 0]}>
                            {groupChartData.map((_entry, i) => (
                              <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    )}
                  </div>

                  {dateCols.length > 0 && (
                    <div className="bg-white border border-stone-300 rounded-sm p-4">
                      <h2 className="font-serif text-base mb-3">Responses over time</h2>
                      <div className="flex flex-wrap gap-3 mb-4 text-sm">
                        <label className="flex items-center gap-2">
                          Date field
                          <select
                            value={cfg.dateCol}
                            onChange={(e) => setCfg({ dateCol: e.target.value })}
                            className="border border-stone-300 rounded-sm px-2 py-1"
                          >
                            {dateCols.map((c) => (
                              <option key={c.name} value={c.name}>{c.name}</option>
                            ))}
                          </select>
                        </label>
                        <label className="flex items-center gap-2">
                          Group by
                          <select
                            value={cfg.timeBucket}
                            onChange={(e) => setCfg({ timeBucket: e.target.value })}
                            className="border border-stone-300 rounded-sm px-2 py-1"
                          >
                            <option value="day">Day</option>
                            <option value="week">Week</option>
                            <option value="month">Month</option>
                          </select>
                        </label>
                      </div>
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
                  )}
                </div>
              )}

              {tab === "table" && (
                <div className="bg-white border border-stone-300 rounded-sm">
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
            </>
          )}
        </main>
      </div>
    </div>
  );
}

function round(n) {
  if (n === undefined || n === null || isNaN(n)) return "—";
  return Math.round(n * 100) / 100;
}

function StatCard({ label, value }) {
  return (
    <div className="bg-white border border-stone-300 rounded-sm p-4">
      <div className="text-sm text-stone-500">{label}</div>
      <div className="font-serif text-2xl text-stone-900 mt-1">{value}</div>
    </div>
  );
}

function MiniBarChart({ title, rows, colorIdx }) {
  const data = useMemo(() => {
    const grouped = _.groupBy(rows, (r) => {
      const v = r[title];
      return v === "" || v === null || v === undefined ? "(blank)" : String(v);
    });
    let entries = Object.entries(grouped).map(([name, rs]) => ({ name, value: rs.length }));
    entries = _.orderBy(entries, "value", "desc");
    if (entries.length > 8) {
      const top = entries.slice(0, 7);
      const rest = entries.slice(7).reduce((a, e) => a + e.value, 0);
      entries = [...top, { name: "Other", value: rest }];
    }
    return entries;
  }, [rows, title]);

  return (
    <div className="bg-white border border-stone-300 rounded-sm p-4">
      <h3 className="text-sm text-stone-700 mb-2 truncate">{title}</h3>
      <ResponsiveContainer width="100%" height={180}>
        <BarChart data={data} margin={{ top: 4, right: 4, left: 0, bottom: 4 }}>
          <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-20} textAnchor="end" height={40} />
          <YAxis tick={{ fontSize: 10 }} width={28} />
          <Tooltip />
          <Bar dataKey="value" radius={[3, 3, 0, 0]} fill={PALETTE[colorIdx % PALETTE.length]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
