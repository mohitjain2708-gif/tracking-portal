import React, { useMemo, useState } from "react";
import {
  Upload,
  FileSpreadsheet,
  Search,
  RefreshCw,
  CheckCircle2,
  AlertTriangle,
  Train,
  Table2,
  Info,
  Package,
  MapPin,
  Clock3,
  BarChart3,
  Ship,
  Download,
  Filter,
  Loader2,
  ArrowUpDown,
} from "lucide-react";
import * as XLSX from "xlsx";
const API_BASE = "https://tracking-portal-2t4o.onrender.com";

function normalize(v) {
  return v == null ? "" : String(v).trim();
}

function KPI({ title, value, icon: Icon, hint }) {
  return (
    <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 16, padding: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.05)", minWidth: 0 }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 12, color: "#64748b", overflowWrap: "anywhere" }}>{title}</div>
          <div style={{ marginTop: 4, fontSize: 24, fontWeight: 700, color: "#0f172a", lineHeight: 1.1 }}>{value}</div>
          {hint ? <div style={{ marginTop: 4, fontSize: 11, color: "#64748b", lineHeight: 1.3 }}>{hint}</div> : null}
        </div>
        <div style={{ background: "#f1f5f9", borderRadius: 12, padding: 8, flex: "0 0 auto" }}>
          <Icon size={16} color="#334155" />
        </div>
      </div>
    </div>
  );
}

function MiniStat({ label, value, tone = "slate" }) {
  const tones = {
    slate: { bg: "#f1f5f9", text: "#334155", border: "#e2e8f0" },
    green: { bg: "#dcfce7", text: "#15803d", border: "#bbf7d0" },
    amber: { bg: "#fef3c7", text: "#b45309", border: "#fde68a" },
    red: { bg: "#fee2e2", text: "#b91c1c", border: "#fecaca" },
    blue: { bg: "#dbeafe", text: "#1d4ed8", border: "#bfdbfe" },
  };
  const t = tones[tone] || tones.slate;
  return (
    <div style={{ background: t.bg, color: t.text, border: `1px solid ${t.border}`, borderRadius: 999, padding: "7px 10px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>
      <span style={{ opacity: 0.85 }}>{label}:</span> <span>{value}</span>
    </div>
  );
}

function getStatusTone(value) {
  const v = String(value || "").trim().toLowerCase();
  if (["arrived", "paid", "yes", "sent", "complete"].includes(v)) return "green";
  if (["not railed", "pending", "unpaid", "hold", "error"].includes(v)) return "red";
  if (["in transit", "moving"].includes(v)) return "blue";
  return "amber";
}

function StatusChip({ value }) {
  const tone = getStatusTone(value);
  const tones = {
    green: { bg: "#dcfce7", text: "#15803d", border: "#bbf7d0" },
    red: { bg: "#fee2e2", text: "#b91c1c", border: "#fecaca" },
    blue: { bg: "#dbeafe", text: "#1d4ed8", border: "#bfdbfe" },
    amber: { bg: "#fef3c7", text: "#b45309", border: "#fde68a" },
  };
  const t = tones[tone] || tones.amber;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", border: `1px solid ${t.border}`, background: t.bg, color: t.text, borderRadius: 999, padding: "4px 10px", fontSize: 11, fontWeight: 700, lineHeight: 1.2, whiteSpace: "nowrap" }}>
      {String(value || "") || "-"}
    </span>
  );
}

function compareValues(a, b) {
  const aNum = Number(a);
  const bNum = Number(b);
  const aIsNum = String(a).trim() !== "" && !Number.isNaN(aNum);
  const bIsNum = String(b).trim() !== "" && !Number.isNaN(bNum);
  if (aIsNum && bIsNum) return aNum - bNum;
  return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [rawPreview, setRawPreview] = useState([]);
  const [query, setQuery] = useState("");
  const [activeTab, setActiveTab] = useState("oonc");
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [paymentFilter, setPaymentFilter] = useState("all");
  const [loading, setLoading] = useState(false);
  const [sortField, setSortField] = useState("default");
  const [sortDirection, setSortDirection] = useState("asc");

  async function handleUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setError("");
    setNotice("");
    setLoading(true);
    setRows([]);
    setHeaders([]);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const API_BASE = "https://tracking-portal-2t4o.onrender.com";

const res = await fetch(`${API_BASE}/api/process-tracking`, {
        method: "POST",
        body: formData,
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || "Processing failed");
      }

      setHeaders(data.headers || []);
      setRows((data.rows || []).map((row, idx) => ({ ...row, __index: idx })));
      setNotice(`Workbook processed successfully. Tracked containers: ${data.tracked_containers || 0}`);

      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      if (wb.Sheets["OONC"]) {
        const raw = XLSX.utils.sheet_to_json(wb.Sheets["OONC"], { header: 1, defval: "", raw: false, blankrows: false });
        setRawPreview(raw);
      } else {
        setRawPreview([]);
      }
    } catch (err) {
      setError(String(err.message || err));
    } finally {
      setLoading(false);
    }
  }

  const filteredRows = useMemo(() => {
    let filtered = rows;

    const q = query.trim().toLowerCase();
    if (q) {
      filtered = filtered.filter((row) =>
        Object.values(row).some((v) => String(v || "").toLowerCase().includes(q))
      );
    }

    const railTransitKey = headers.find((h) => String(h).trim().toLowerCase() === "rail transit time");
    const paymentKey = headers.find((h) => String(h).trim().toLowerCase() === "payment status");

    if (statusFilter !== "all" && railTransitKey) {
      filtered = filtered.filter((row) => {
        const status = String(row[railTransitKey] || "").toLowerCase();
        if (statusFilter === "arrived") return status === "arrived";
        if (statusFilter === "moving") return status !== "arrived" && status !== "not railed" && status !== "";
        if (statusFilter === "not_railed") return status === "not railed";
        return true;
      });
    }

    if (paymentFilter !== "all" && paymentKey) {
      filtered = filtered.filter((row) => {
        const payment = String(row[paymentKey] || "").toLowerCase();
        if (paymentFilter === "paid") return ["paid", "yes", "sent", "complete"].includes(payment);
        if (paymentFilter === "pending") return ["pending", "unpaid"].includes(payment);
        return true;
      });
    }

    if (sortField !== "default") {
      filtered = [...filtered].sort((a, b) => {
        const result = compareValues(a[sortField], b[sortField]);
        return sortDirection === "asc" ? result : -result;
      });
    } else {
      filtered = [...filtered].sort((a, b) => (a.__index || 0) - (b.__index || 0));
    }

    return filtered;
  }, [rows, headers, query, statusFilter, paymentFilter, sortField, sortDirection]);

  const dashboard = useMemo(() => {
    const total = filteredRows.length;

    const getVal = (row, names) => {
      const key = headers.find((h) => names.includes(String(h).trim().toLowerCase()));
      return key ? normalize(row[key]) : "";
    };

    const withContainer = filteredRows.filter((r) => getVal(r, ["container no", "containerno", "container", "cntr no"]).length > 0).length;
    const arrived = filteredRows.filter((r) => getVal(r, ["rail transit time"]).toLowerCase() === "arrived").length;
    const notRailed = filteredRows.filter((r) => getVal(r, ["rail transit time"]).toLowerCase() === "not railed").length;
    const inTransit = total - arrived - notRailed;
    const paid = filteredRows.filter((r) => ["paid", "yes", "sent", "complete"].includes(getVal(r, ["payment status"]).toLowerCase())).length;
    const pending = filteredRows.filter((r) => ["pending", "unpaid"].includes(getVal(r, ["payment status"]).toLowerCase())).length;

    const birgunjCount = filteredRows.filter((r) => {
      const loc = getVal(r, ["last location"]);
      return loc.toUpperCase().includes("BIRGUNJ") || loc.toUpperCase().includes("BIRGANJ");
    }).length;

    const portCount = filteredRows.filter((r) => {
      const loc = getVal(r, ["last location"]).toUpperCase();
      return loc.includes("VISHAKAPATNAM") || loc.includes("VIZAG") || loc.includes("KOLKATA") || loc.includes("HALDIA") || loc.includes("PORT");
    }).length;

    const railCount = filteredRows.filter((r) => {
      const loc = getVal(r, ["last location"]).toUpperCase();
      return loc && !loc.includes("BIRGUNJ") && !loc.includes("BIRGANJ") && !loc.includes("VISHAKAPATNAM") && !loc.includes("VIZAG") && !loc.includes("KOLKATA") && !loc.includes("HALDIA") && !loc.includes("PORT");
    }).length;

    const locationCounts = {};
    filteredRows.forEach((r) => {
      const loc = getVal(r, ["last location"]) || "Unknown";
      locationCounts[loc] = (locationCounts[loc] || 0) + 1;
    });
    const topLocations = Object.entries(locationCounts).sort((a, b) => b[1] - a[1]).slice(0, 5);

    const partyCounts = {};
    filteredRows.forEach((r) => {
      const party = getVal(r, ["pary name", "party name"]) || "Unknown Party";
      partyCounts[party] = (partyCounts[party] || 0) + 1;
    });
    const topParties = Object.entries(partyCounts).sort((a, b) => b[1] - a[1]).slice(0, 5);

    return { total, withContainer, arrived, notRailed, inTransit, paid, pending, birgunjCount, portCount, railCount, topLocations, topParties };
  }, [filteredRows, headers]);

  function exportUpdatedExcel() {
    if (!headers.length) return;
    const data = [headers, ...filteredRows.map((row) => headers.map((h) => row[h] ?? ""))];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OONC_Result_View");
    XLSX.writeFile(wb, fileName ? `${fileName.replace(/\.(xlsx|xlsm|xls)$/i, "")}_OONC_Result.xlsx` : "OONC_Result.xlsx");
  }

  const pageStyle = {
    minHeight: "100vh",
    background: "#f8fafc",
    padding: "12px",
    color: "#0f172a",
    fontFamily: "Inter, Arial, sans-serif",
    overflowX: "hidden",
  };
  const sectionStyle = {
    width: "100%",
    maxWidth: "100%",
    margin: "0 auto",
    display: "grid",
    gap: 14,
  };
  const cardStyle = {
    background: "white",
    border: "1px solid #e2e8f0",
    borderRadius: 16,
    boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
    minWidth: 0,
  };
  const compactBtn = {
    border: 0,
    padding: "8px 12px",
    borderRadius: 10,
    fontWeight: 600,
    cursor: "pointer",
    fontSize: 12,
    whiteSpace: "nowrap",
  };

  return (
    <div style={pageStyle}>
      <div style={sectionStyle}>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#64748b" }}>Tracking Control Portal</div>
          <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", gap: 12, alignItems: "end" }}>
            <div style={{ minWidth: 0, flex: "1 1 600px" }}>
              <h1 style={{ margin: 0, fontSize: "clamp(22px, 3vw, 30px)", fontWeight: 700, letterSpacing: "-0.02em" }}>Professional Tracking Dashboard</h1>
              <p style={{ marginTop: 6, maxWidth: 980, color: "#475569", lineHeight: 1.45, fontSize: 13 }}>
                Upload your tracking workbook. The backend processes OONC containers automatically and shows the updated OONC result with dashboard summaries, filters, sorting, status chips, and Excel export.
              </p>
            </div>
            <label style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 14px", border: "1px solid #e2e8f0", borderRadius: 14, background: "white", cursor: "pointer", boxShadow: "0 1px 3px rgba(0,0,0,0.05)", whiteSpace: "nowrap" }}>
              {loading ? <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Upload size={16} />}
              <span style={{ fontWeight: 600, fontSize: 13 }}>{loading ? "Processing..." : "Upload Excel File"}</span>
              <input type="file" accept=".xlsx,.xlsm,.xls" style={{ display: "none" }} onChange={handleUpload} disabled={loading} />
            </label>
          </div>
        </div>

        {fileName ? (
          <div style={{ ...cardStyle, overflow: "hidden" }}>
            <div style={{ background: "linear-gradient(90deg, #0f172a 0%, #1e293b 55%, #334155 100%)", color: "white", padding: 16 }}>
              <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
                  <div style={{ background: "rgba(255,255,255,0.12)", borderRadius: 14, padding: 9, flex: "0 0 auto" }}><FileSpreadsheet size={20} /></div>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 17, fontWeight: 700, overflowWrap: "anywhere" }}>{fileName}</div>
                    <div style={{ fontSize: 12, color: "#cbd5e1" }}>Processed result ready for view</div>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button onClick={() => setActiveTab("oonc")} style={{ ...compactBtn, background: "white", color: "#0f172a" }}>
                    <RefreshCw size={13} style={{ marginRight: 6, verticalAlign: "middle" }} />Refresh View
                  </button>
                  <button onClick={exportUpdatedExcel} style={{ ...compactBtn, background: "#dbeafe", color: "#1d4ed8" }}>
                    <Download size={13} style={{ marginRight: 6, verticalAlign: "middle" }} />Export Updated Excel
                  </button>
                </div>
              </div>
            </div>
            <div style={{ padding: 12, display: "flex", gap: 8, flexWrap: "wrap", overflowX: "auto" }}>
              <MiniStat label="Containers Found" value={dashboard.withContainer} tone="blue" />
              <MiniStat label="Arrived" value={dashboard.arrived} tone="green" />
              <MiniStat label="In Transit" value={dashboard.inTransit} tone="amber" />
              <MiniStat label="Not Railed" value={dashboard.notRailed} tone="red" />
            </div>
          </div>
        ) : null}

        {error ? <div style={{ border: "1px solid #fecaca", background: "#fef2f2", color: "#b91c1c", padding: 12, borderRadius: 14, fontSize: 13 }}>{error}</div> : null}
        {notice ? <div style={{ border: "1px solid #bfdbfe", background: "#eff6ff", color: "#1d4ed8", padding: 12, borderRadius: 14, fontSize: 13 }}>{notice}</div> : null}

        <div style={{ display: "grid", gap: 12, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
          <KPI title="OONC Rows" value={dashboard.total} icon={Table2} hint="Rows shown in current view" />
          <KPI title="Containers Found" value={dashboard.withContainer} icon={Package} hint="Rows having a container number" />
          <KPI title="Arrived" value={dashboard.arrived} icon={CheckCircle2} hint="Gate in Birganj already reached" />
          <KPI title="Not Railed" value={dashboard.notRailed} icon={AlertTriangle} hint="No rail movement assigned yet" />
        </div>

        <div style={{ display: "grid", gap: 12, gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))" }}>
          <div style={cardStyle}><div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><MapPin size={15} /> Birgunj</div><div style={{ padding: 14 }}><div style={{ fontSize: 26, fontWeight: 700 }}>{dashboard.birgunjCount}</div><div style={{ marginTop: 4, color: "#64748b", fontSize: 12 }}>Containers currently showing Birgunj as last location</div></div></div>
          <div style={cardStyle}><div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><Ship size={15} /> Port</div><div style={{ padding: 14 }}><div style={{ fontSize: 26, fontWeight: 700 }}>{dashboard.portCount}</div><div style={{ marginTop: 4, color: "#64748b", fontSize: 12 }}>Containers still at India port side locations</div></div></div>
          <div style={cardStyle}><div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><Train size={15} /> Rail Movement</div><div style={{ padding: 14 }}><div style={{ fontSize: 26, fontWeight: 700 }}>{dashboard.railCount}</div><div style={{ marginTop: 4, color: "#64748b", fontSize: 12 }}>Containers currently moving through inland rail points</div></div></div>
        </div>

        <div style={{ display: "grid", gap: 12, gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))" }}>
          <div style={cardStyle}>
            <div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><Train size={15} /> Movement Summary</div>
            <div style={{ padding: 14, display: "grid", gap: 8 }}>
              <MiniStat label="In Transit" value={dashboard.inTransit} tone="blue" />
              <MiniStat label="Arrived" value={dashboard.arrived} tone="green" />
              <MiniStat label="Not Railed" value={dashboard.notRailed} tone="red" />
            </div>
          </div>

          <div style={cardStyle}>
            <div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><Clock3 size={15} /> Payment Snapshot</div>
            <div style={{ padding: 14, display: "grid", gap: 8 }}>
              <MiniStat label="Paid / Yes / Sent" value={dashboard.paid} tone="green" />
              <MiniStat label="Pending / Unpaid" value={dashboard.pending} tone="red" />
            </div>
          </div>

          <div style={cardStyle}>
            <div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><BarChart3 size={15} /> Top Parties</div>
            <div style={{ padding: 14, display: "grid", gap: 8 }}>
              {dashboard.topParties.length ? dashboard.topParties.map(([party, count]) => (
                <div key={party} style={{ display: "flex", justifyContent: "space-between", gap: 10, background: "#f8fafc", borderRadius: 12, padding: "8px 10px", fontSize: 12, minWidth: 0 }}>
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", paddingRight: 10 }}>{party}</span>
                  <StatusChip value={count} />
                </div>
              )) : <div style={{ color: "#64748b", fontSize: 12 }}>No party data available.</div>}
            </div>
          </div>
        </div>

        <div style={cardStyle}>
          <div style={{ padding: 14, borderBottom: "1px solid #e2e8f0", fontWeight: 700, display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}><MapPin size={15} /> Top Last Locations</div>
          <div style={{ padding: 14, display: "grid", gap: 10, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
            {dashboard.topLocations.length ? dashboard.topLocations.map(([loc, count]) => (
              <div key={loc} style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 14, padding: 12, minWidth: 0 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: "#0f172a", minHeight: 34, overflowWrap: "anywhere" }}>{loc}</div>
                <div style={{ marginTop: 8, fontSize: 22, fontWeight: 700, color: "#334155" }}>{count}</div>
              </div>
            )) : <div style={{ color: "#64748b", fontSize: 12 }}>No location data available.</div>}
          </div>
        </div>

        <div style={cardStyle}>
          <div style={{ padding: 14, borderBottom: "1px solid #e2e8f0" }}>
            <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
              <div style={{ fontSize: 18, fontWeight: 700 }}>Workbook Result View</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid #e2e8f0", background: "white", borderRadius: 12, padding: "7px 9px" }}>
                  <Filter size={13} color="#64748b" />
                  <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ border: 0, outline: "none", background: "transparent", fontSize: 12 }}>
                    <option value="all">All Movement</option>
                    <option value="arrived">Arrived</option>
                    <option value="moving">In Transit</option>
                    <option value="not_railed">Not Railed</option>
                  </select>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid #e2e8f0", background: "white", borderRadius: 12, padding: "7px 9px" }}>
                  <Filter size={13} color="#64748b" />
                  <select value={paymentFilter} onChange={(e) => setPaymentFilter(e.target.value)} style={{ border: 0, outline: "none", background: "transparent", fontSize: 12 }}>
                    <option value="all">All Payment</option>
                    <option value="paid">Paid / Yes / Sent</option>
                    <option value="pending">Pending / Unpaid</option>
                  </select>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid #e2e8f0", background: "white", borderRadius: 12, padding: "7px 9px" }}>
                  <ArrowUpDown size={13} color="#64748b" />
                  <select value={sortField} onChange={(e) => setSortField(e.target.value)} style={{ border: 0, outline: "none", background: "transparent", fontSize: 12, maxWidth: 150 }}>
                    <option value="default">Default Order</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                  <select value={sortDirection} onChange={(e) => setSortDirection(e.target.value)} style={{ border: 0, outline: "none", background: "transparent", fontSize: 12 }}>
                    <option value="asc">Asc</option>
                    <option value="desc">Desc</option>
                  </select>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, border: "1px solid #e2e8f0", background: "white", borderRadius: 12, padding: "7px 9px", minWidth: 220 }}>
                  <Search size={13} color="#64748b" />
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search any OONC field"
                    style={{ border: 0, outline: "none", width: "100%", fontSize: 12 }}
                  />
                </div>
              </div>
            </div>
          </div>

          <div style={{ padding: 12, display: "flex", flexWrap: "wrap", gap: 8 }}>
            <button onClick={() => setActiveTab("oonc")} style={{ ...compactBtn, background: activeTab === "oonc" ? "#0f172a" : "#f8fafc", color: activeTab === "oonc" ? "white" : "#0f172a" }}>OONC Result View</button>
            <button onClick={() => setActiveTab("raw")} style={{ ...compactBtn, background: activeTab === "raw" ? "#0f172a" : "#f8fafc", color: activeTab === "raw" ? "white" : "#0f172a" }}>Raw OONC Sheet</button>
          </div>

          {activeTab === "oonc" && (
            <div style={{ overflowX: "auto", overflowY: "auto", borderTop: "1px solid #e2e8f0", maxHeight: "56vh", width: "100%" }}>
              <table style={{ minWidth: "max-content", width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                <thead style={{ background: "#f1f5f9", color: "#334155", position: "sticky", top: 0, zIndex: 2 }}>
                  <tr>
                    {headers.map((h) => (
                      <th key={h} style={{ whiteSpace: "nowrap", padding: "10px 12px", textAlign: "left", fontWeight: 700, borderBottom: "1px solid #e2e8f0" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td colSpan={headers.length || 1} style={{ padding: 24, textAlign: "center", color: "#64748b" }}>
                        Upload the tracking workbook to preview the updated OONC sheet.
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((row, idx) => (
                      <tr key={idx} style={{ borderTop: "1px solid #f1f5f9" }}>
                        {headers.map((h) => {
                          const lower = String(h).trim().toLowerCase();
                          const value = row[h];
                          return (
                            <td key={h} style={{ whiteSpace: "nowrap", padding: "10px 12px", verticalAlign: "middle" }}>
                              {lower === "payment status" ? (
                                <StatusChip value={value} />
                              ) : lower === "rail transit time" ? (
                                <StatusChip value={value} />
                              ) : (lower === "gate in birganj" || lower === "gate in birgunj") && value ? (
  (() => {
    const dateKey = headers.find(
      (h) => String(h).trim().toLowerCase() === "last location (date)"
    );
    return dateKey ? String(row[dateKey] || "") : "";
  })()
) : (
                                String(value ?? "")
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

          {activeTab === "raw" && (
            <div style={{ overflowX: "auto", overflowY: "auto", borderTop: "1px solid #e2e8f0", maxHeight: "56vh", width: "100%" }}>
              <table style={{ minWidth: "max-content", width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                <tbody>
                  {rawPreview.map((row, idx) => (
                    <tr key={idx} style={{ borderTop: "1px solid #f1f5f9" }}>
                      {(row || []).map((cell, cidx) => (
                        <td key={`${idx}-${cidx}`} style={{ whiteSpace: "nowrap", padding: "10px 12px" }}>{String(cell ?? "")}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div style={cardStyle}>
          <div style={{ padding: 14, display: "flex", gap: 10, alignItems: "start", color: "#475569", fontSize: 12 }}>
            <Info size={15} style={{ marginTop: 2, flex: "0 0 auto" }} />
            <div>
              View is adjusted for large datasets: only the result table scrolls horizontally, the page itself stays within the Windows screen, chips are color-fixed, and sorting works with filters.
            </div>
          </div>
        </div>
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } } body { margin: 0; } * { box-sizing: border-box; }`}</style>
    </div>
  );
}
