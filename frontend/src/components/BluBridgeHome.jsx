import React, { useState, useRef, useCallback, useEffect } from "react";
import { Upload, Download, FileSpreadsheet, CheckCircle2, AlertCircle, X, ArrowRight, Database, TrendingUp, Award, LogOut } from "lucide-react";

const API = process.env.REACT_APP_BACKEND_URL + "/api";

const StatCard = ({ icon: Icon, value, label, color }) => (
  <div className="bg-white rounded-2xl p-6 flex items-center gap-4"
       style={{ boxShadow: "0 4px 16px -4px rgba(0,0,0,0.08)" }}>
    <div className={`w-12 h-12 rounded-xl flex items-center justify-center`}
         style={{ backgroundColor: color + "18" }}>
      <Icon size={22} style={{ color }} />
    </div>
    <div>
      <div className="text-2xl font-bold text-gray-900" style={{ fontFamily: "'Manrope', sans-serif" }}>{value}</div>
      <div className="text-xs text-gray-500 font-medium uppercase tracking-wider">{label}</div>
    </div>
  </div>
);

const Step = ({ num, title, desc, active }) => (
  <div className={`flex gap-4 p-4 rounded-xl transition-all duration-200 ${active ? "bg-blue-50 border border-blue-100" : ""}`}>
    <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold flex-shrink-0 mt-0.5
      ${active ? "bg-[#1A73E8] text-white" : "bg-gray-100 text-gray-500"}`}>
      {num}
    </div>
    <div>
      <div className={`text-sm font-semibold ${active ? "text-[#1A73E8]" : "text-gray-600"}`}>{title}</div>
      <div className="text-xs text-gray-400 mt-0.5 leading-relaxed">{desc}</div>
    </div>
  </div>
);

export default function BluBridgeHome() {
  const [file, setFile] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [status, setStatus] = useState("idle"); // idle | uploading | done | error
  const [errorMsg, setErrorMsg] = useState("");
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [downloadName, setDownloadName] = useState("");
  const [seeding, setSeeding] = useState(false);
  const [seedDone, setSeedDone] = useState(false);
  const [stats, setStats] = useState({ total: 300, ranked: 100, unranked: 200 });

  useEffect(() => {
    fetch(`${API}/colleges/stats`)
      .then(r => r.json())
      .then(d => setStats({ total: d.total, ranked: d.ranked, unranked: d.unranked }))
      .catch(() => {});
  }, [seedDone]);
  const fileRef = useRef();

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files[0];
    if (f) acceptFile(f);
  }, []);

  const acceptFile = (f) => {
    if (!f.name.match(/\.(xlsx|xls)$/i)) {
      setErrorMsg("Only .xlsx or .xls files are supported.");
      setStatus("error");
      return;
    }
    setFile(f);
    setStatus("idle");
    setErrorMsg("");
    setDownloadUrl(null);
  };

  const handleProcess = async () => {
    if (!file) return;
    setStatus("uploading");
    setErrorMsg("");
    setDownloadUrl(null);
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch(`${API}/process-excel`, { method: "POST", body: formData });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Processing failed");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const cd = res.headers.get("content-disposition") || "";
      const match = cd.match(/filename="(.+)"/);
      setDownloadUrl(url);
      setDownloadName(match ? match[1] : "BluBridge_Processed.xlsx");
      setStatus("done");
    } catch (e) {
      setErrorMsg(e.message || "Something went wrong. Please try again.");
      setStatus("error");
    }
  };

  const handleSeed = async () => {
    setSeeding(true);
    try {
      const res = await fetch(`${API}/colleges/seed`, { method: "POST" });
      const data = await res.json();
      if (data.seeded) setSeedDone(true);
    } catch { }
    setSeeding(false);
  };

  const reset = () => { setFile(null); setStatus("idle"); setErrorMsg(""); setDownloadUrl(null); };

  return (
    <div className="min-h-screen flex flex-col" style={{ fontFamily: "'Inter', sans-serif", backgroundColor: "#F5F2E9" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&family=Inter:wght@400;500;600&display=swap');`}</style>

      {/* Header */}
      <header data-testid="header"
        className="sticky top-0 z-50 flex items-center justify-center px-6 md:px-12 h-16 border-b border-gray-200/60"
        style={{ backgroundColor: "rgba(238, 234, 221, 0.9)", backdropFilter: "blur(12px)" }}>
        <img
          data-testid="logo"
          src="https://customer-assets.emergentagent.com/job_premium-design-ui/artifacts/dwo5zlsk_logo-black121.png"
          alt="BluBridge"
          className="object-contain"
        />
      </header>

      <main className="flex-1 px-4 py-10 md:py-14 max-w-5xl mx-auto w-full">

        {/* Hero */}
        <div className="mb-10 text-center">
          <div className="inline-flex items-center gap-2 bg-blue-50 text-[#1A73E8] text-xs font-semibold px-3 py-1.5 rounded-full mb-4 border border-blue-100">
            <TrendingUp size={12} /> NIRF 2025 Rankings
          </div>
          <h1 data-testid="hero-title"
            className="text-4xl md:text-5xl font-bold text-gray-900 tracking-tight mb-3"
            style={{ fontFamily: "'Manrope', sans-serif" }}>
            College Ranking Processor
          </h1>
          <p className="text-gray-500 text-base max-w-xl mx-auto leading-relaxed">
            Upload your Job Post applicant Excel. We auto-match college names and fill NIRF ranks — works with full names and short names.
          </p>
        </div>

        {/* Stats */}
        <div className="grid grid-cols-2 md:grid-cols-3 gap-4 mb-10">
          <StatCard icon={Database}   value={stats.total}    label="Total Colleges"  color="#1A73E8" />
          <StatCard icon={Award}      value={stats.ranked}   label="NIRF Ranked"     color="#0F9D58" />
          <StatCard icon={FileSpreadsheet} value={stats.unranked} label="Rank Band" color="#F4B400" />
        </div>

        <div className="grid md:grid-cols-2 gap-6 mb-6">

          {/* Upload Card */}
          <div data-testid="upload-card"
            className="bg-white rounded-3xl overflow-hidden"
            style={{ boxShadow: "0 20px 50px -12px rgba(0,0,0,0.10)" }}>
            <div className="h-1.5 w-full" style={{ background: "linear-gradient(90deg,#1A73E8,#4FA3F7)" }} />
            <div className="p-7">
              <h2 className="text-lg font-bold text-gray-900 mb-1" style={{ fontFamily: "'Manrope', sans-serif" }}>
                Upload Job Post Excel
              </h2>
              <p className="text-xs text-gray-400 mb-5">Supports .xlsx and .xls files with UG/PG college name columns</p>

              {/* How it works steps */}
              <div className="space-y-1 mb-6">
                <Step num="1" title="Upload Excel" desc="Job Post applicant file with college name columns" active={!file} />
                <Step num="2" title="Auto Match"   desc="We match full names & short names to NIRF rank"  active={!!file && status === "idle"} />
                <Step num="3" title="Download"     desc="Get processed Excel with ranks + ranking reference tab" active={status === "done"} />
              </div>

              {/* Drop Zone */}
              {!file ? (
                <div data-testid="drop-zone"
                  onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={handleDrop}
                  onClick={() => fileRef.current?.click()}
                  className={`border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all duration-200
                    ${dragOver ? "border-[#1A73E8] bg-blue-50" : "border-gray-200 bg-gray-50 hover:border-[#1A73E8] hover:bg-blue-50/50"}`}>
                  <Upload size={28} className={`mx-auto mb-3 ${dragOver ? "text-[#1A73E8]" : "text-gray-300"}`} />
                  <p className="text-sm font-medium text-gray-600">Drag & drop or <span className="text-[#1A73E8]">browse</span></p>
                  <p className="text-xs text-gray-400 mt-1">.xlsx or .xls</p>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden"
                    data-testid="file-input"
                    onChange={(e) => e.target.files[0] && acceptFile(e.target.files[0])} />
                </div>
              ) : (
                <div className="space-y-3">
                  <div className="flex items-center gap-3 p-4 bg-blue-50 rounded-xl border border-blue-100">
                    <FileSpreadsheet size={20} className="text-[#1A73E8] flex-shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-semibold text-gray-800 truncate">{file.name}</p>
                      <p className="text-xs text-gray-400">{(file.size / 1024).toFixed(1)} KB</p>
                    </div>
                    <button onClick={reset} className="text-gray-400 hover:text-gray-600 transition-colors">
                      <X size={16} />
                    </button>
                  </div>

                  {status === "error" && (
                    <div data-testid="error-msg" className="flex items-center gap-2 p-3 bg-red-50 rounded-xl text-sm text-red-600 border border-red-100">
                      <AlertCircle size={16} /> {errorMsg}
                    </div>
                  )}

                  {status !== "done" ? (
                    <button data-testid="process-btn"
                      onClick={handleProcess}
                      disabled={status === "uploading"}
                      className={`w-full h-12 rounded-xl font-bold text-sm text-white flex items-center justify-center gap-2 transition-all duration-300
                        ${status === "uploading"
                          ? "bg-gray-300 cursor-not-allowed"
                          : "bg-[#1A73E8] hover:bg-[#1557B0] shadow-lg shadow-blue-500/25 hover:-translate-y-0.5"}`}>
                      {status === "uploading" ? (
                        <><span className="animate-spin inline-block w-4 h-4 border-2 border-white border-t-transparent rounded-full" />
                          Processing...</>
                      ) : (
                        <><ArrowRight size={16} /> Process & Match Ranks</>
                      )}
                    </button>
                  ) : (
                    <div className="space-y-2">
                      <div className="flex items-center gap-2 p-3 bg-green-50 rounded-xl text-sm text-green-700 border border-green-100">
                        <CheckCircle2 size={16} /> Ranks matched successfully!
                      </div>
                      <button onClick={reset} className="w-full h-10 rounded-xl text-sm font-medium text-gray-500 border border-gray-200 hover:border-gray-300 hover:bg-gray-50 transition-all">
                        Upload another file
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>

          {/* Download Card */}
          <div data-testid="download-card"
            className="bg-white rounded-3xl overflow-hidden"
            style={{ boxShadow: "0 20px 50px -12px rgba(0,0,0,0.10)" }}>
            <div className="h-1.5 w-full" style={{ background: "linear-gradient(90deg,#0F9D58,#34C77B)" }} />
            <div className="p-7">
              <h2 className="text-lg font-bold text-gray-900 mb-1" style={{ fontFamily: "'Manrope', sans-serif" }}>
                Download Files
              </h2>
              <p className="text-xs text-gray-400 mb-6">Download processed Excel or NIRF reference lists</p>

              <div className="space-y-3">
                {/* Processed Excel (appears after processing) */}
                {downloadUrl && (
                  <a href={downloadUrl} download={downloadName} data-testid="download-processed-btn"
                    className="flex items-center gap-4 p-4 bg-green-50 rounded-xl border border-green-200
                      hover:bg-green-100 hover:border-green-300 transition-all group cursor-pointer">
                    <div className="w-10 h-10 rounded-xl bg-green-500 flex items-center justify-center flex-shrink-0">
                      <Download size={18} className="text-white" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-bold text-gray-800 truncate">{downloadName}</p>
                      <p className="text-xs text-green-600 font-medium">Processed — Ready to download</p>
                    </div>
                    <ArrowRight size={16} className="text-green-500 group-hover:translate-x-0.5 transition-transform" />
                  </a>
                )}

                {/* NIRF Rankings with Short Names */}
                <a href={`${API}/download-college-list-shortnames`} data-testid="download-shortnames-btn"
                  className="flex items-center gap-4 p-4 bg-blue-50 rounded-xl border border-blue-100
                    hover:bg-blue-100 hover:border-blue-200 transition-all group cursor-pointer">
                  <div className="w-10 h-10 rounded-xl bg-[#1A73E8] flex items-center justify-center flex-shrink-0">
                    <FileSpreadsheet size={18} className="text-white" />
                  </div>
                  <div className="flex-1">
                    <p className="text-sm font-bold text-gray-800">NIRF Rankings with Short Names</p>
                    <p className="text-xs text-blue-500 font-medium">300 colleges — Rank | Name | Short Name | City | State</p>
                  </div>
                  <ArrowRight size={16} className="text-[#1A73E8] group-hover:translate-x-0.5 transition-transform" />
                </a>

                {/* Original Rankings */}
                <a href={`${API}/download-college-list`} data-testid="download-rankings-btn"
                  className="flex items-center gap-4 p-4 bg-gray-50 rounded-xl border border-gray-100
                    hover:bg-gray-100 hover:border-gray-200 transition-all group cursor-pointer">
                  <div className="w-10 h-10 rounded-xl bg-gray-500 flex items-center justify-center flex-shrink-0">
                    <FileSpreadsheet size={18} className="text-white" />
                  </div>
                  <div className="flex-1">
                    <p className="text-sm font-bold text-gray-800">NIRF Rankings (Original)</p>
                    <p className="text-xs text-gray-400 font-medium">300 colleges — Rank | Name | City | State</p>
                  </div>
                  <ArrowRight size={16} className="text-gray-400 group-hover:translate-x-0.5 transition-transform" />
                </a>

                {/* All India Colleges */}
                <a href={`${API}/download-all-india-colleges`} data-testid="download-all-india-btn"
                  className="flex items-center gap-4 p-4 bg-emerald-50 rounded-xl border border-emerald-100
                    hover:bg-emerald-100 hover:border-emerald-200 transition-all group cursor-pointer">
                  <div className="w-10 h-10 rounded-xl bg-emerald-600 flex items-center justify-center flex-shrink-0">
                    <Database size={18} className="text-white" />
                  </div>
                  <div className="flex-1">
                    <p className="text-sm font-bold text-gray-800">All India Colleges</p>
                    <p className="text-xs text-emerald-600 font-medium">42,000+ colleges — Rank | Name | Short Name | City | State</p>
                  </div>
                  <ArrowRight size={16} className="text-emerald-500 group-hover:translate-x-0.5 transition-transform" />
                </a>
              </div>
            </div>
          </div>
        </div>

        {/* Excel Matcher Link */}
        <div className="mb-6">
          <a href="/excel-matcher" data-testid="excel-matcher-link"
            className="bg-white rounded-2xl p-5 flex items-center justify-between gap-4 hover:shadow-lg transition-all cursor-pointer block"
            style={{ boxShadow: "0 4px 16px -4px rgba(0,0,0,0.07)" }}>
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-orange-50 flex items-center justify-center">
                <ArrowRight size={18} className="text-orange-500" />
              </div>
              <div>
                <p className="text-sm font-bold text-gray-800">Excel Matcher — Shortlist / Reject</p>
                <p className="text-xs text-gray-400">Upload 2 sheets, match by phone number, get status tagged</p>
              </div>
            </div>
            <ArrowRight size={16} className="text-gray-400" />
          </a>
        </div>

        {/* DB Admin Card */}
        {/* <div className="bg-white rounded-2xl p-5 flex items-center justify-between gap-4"
             style={{ boxShadow: "0 4px 16px -4px rgba(0,0,0,0.07)" }}>
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-purple-50 flex items-center justify-center">
              <Database size={18} className="text-purple-500" />
            </div>
            <div>
              <p className="text-sm font-bold text-gray-800">College Database</p>
              <p className="text-xs text-gray-400">Seed / refresh all 300 colleges in MongoDB</p>
            </div>
          </div>
          <button data-testid="seed-btn"
            onClick={handleSeed}
            disabled={seeding || seedDone}
            className={`px-5 h-10 rounded-xl text-sm font-bold transition-all duration-200
              ${seedDone
                ? "bg-green-50 text-green-600 border border-green-200"
                : seeding
                  ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                  : "bg-purple-500 text-white hover:bg-purple-600 shadow-md shadow-purple-200"}`}>
            {seedDone ? <span className="flex items-center gap-1.5"><CheckCircle2 size={14} /> Seeded!</span>
              : seeding ? "Seeding..." : "Seed DB"}
          </button>
        </div> */}

      </main>

      {/* Footer */}
      <footer data-testid="footer" className="py-6 text-center" style={{ backgroundColor: "#222222" }}>
        <p className="text-sm text-gray-400">Copyright 2026 &copy; <span className="text-white font-semibold">Blubridge.com</span></p>
      </footer>
    </div>
  );
}
