import React, { useState, useEffect, useCallback } from "react";
import { Search, ChevronLeft, ChevronRight, Building2, MapPin, Filter, X } from "lucide-react";
import { useNavigate, useSearchParams } from "react-router-dom";

const API = process.env.REACT_APP_BACKEND_URL + "/api";

const RANK_FILTERS = [
  { key: "", label: "All" },
  { key: "top100", label: "Top 100" },
  { key: "101-150", label: "101–150" },
  { key: "151-200", label: "151–200" },
  { key: "201-300", label: "201–300" },
];

const TYPE_COLORS = {
  IIT: "bg-blue-100 text-blue-700",
  NIT: "bg-emerald-100 text-emerald-700",
  IIIT: "bg-purple-100 text-purple-700",
  "IISc/IISER": "bg-amber-100 text-amber-700",
  BITS: "bg-rose-100 text-rose-700",
  Deemed: "bg-cyan-100 text-cyan-700",
  Government: "bg-teal-100 text-teal-700",
  University: "bg-indigo-100 text-indigo-700",
  Private: "bg-orange-100 text-orange-700",
  Institute: "bg-gray-100 text-gray-600",
};

export default function CollegeDirectory() {
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();

  const [colleges, setColleges] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [pages, setPages] = useState(1);
  const [search, setSearch] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [rankFilter, setRankFilter] = useState(searchParams.get("rank") || "");
  const [typeFilter, setTypeFilter] = useState("");
  const [collegeTypes, setCollegeTypes] = useState({});
  const [loading, setLoading] = useState(true);

  // Debounce search
  useEffect(() => {
    const t = setTimeout(() => { setDebouncedSearch(search); setPage(1); }, 300);
    return () => clearTimeout(t);
  }, [search]);

  // Fetch colleges
  const fetchColleges = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: 50 });
      if (debouncedSearch) params.set("search", debouncedSearch);
      if (rankFilter) params.set("rank_filter", rankFilter);
      if (typeFilter) params.set("type_filter", typeFilter);
      const res = await fetch(`${API}/colleges/directory?${params}`);
      const data = await res.json();
      setColleges(data.colleges);
      setTotal(data.total);
      setPages(data.pages);
      setCollegeTypes(data.college_types || {});
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  }, [page, debouncedSearch, rankFilter, typeFilter]);

  useEffect(() => { fetchColleges(); }, [fetchColleges]);

  // Init rank filter from URL
  useEffect(() => {
    const r = searchParams.get("rank");
    if (r) setRankFilter(r);
  }, [searchParams]);

  const handleRankFilter = (key) => {
    setRankFilter(key);
    setPage(1);
    setSearchParams(key ? { rank: key } : {});
  };

  const RankBadge = ({ rank }) => {
    if (typeof rank === "number" || (typeof rank === "string" && /^\d+$/.test(rank))) {
      const n = parseInt(rank);
      const color = n <= 10 ? "bg-yellow-100 text-yellow-800 ring-yellow-300" :
        n <= 50 ? "bg-blue-100 text-blue-700" : "bg-gray-100 text-gray-700";
      return <span className={`inline-flex items-center justify-center w-12 h-7 rounded-lg text-xs font-bold ${color}`}>#{rank}</span>;
    }
    if (typeof rank === "string" && rank.includes("-")) {
      return <span className="inline-flex items-center justify-center px-2 h-7 rounded-lg text-xs font-bold bg-orange-100 text-orange-700">{rank}</span>;
    }
    return <span className="inline-flex items-center justify-center w-12 h-7 rounded-lg text-xs font-medium bg-gray-50 text-gray-400">NL</span>;
  };

  const TypeBadge = ({ type }) => (
    <span className={`inline-flex px-2 py-0.5 rounded-md text-[10px] font-bold uppercase tracking-wider ${TYPE_COLORS[type] || TYPE_COLORS.Institute}`}>
      {type}
    </span>
  );

  return (
    <div className="min-h-screen bg-[#F7F8FA]" style={{ fontFamily: "'Inter', sans-serif" }}>
      {/* Header */}
      <header data-testid="header" className="w-full py-4 px-6 bg-white border-b border-gray-100 flex items-center justify-center relative">
        <button data-testid="back-home-btn" onClick={() => navigate("/")}
          className="absolute left-6 flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-800 transition-colors">
          <ChevronLeft size={16} /> Home
        </button>
        <img data-testid="logo"
          src="https://customer-assets.emergentagent.com/job_premium-design-ui/artifacts/dwo5zlsk_logo-black121.png"
          alt="BluBridge" className="h-8" />
      </header>

      <div className="max-w-5xl mx-auto px-4 py-6">
        {/* Title */}
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">College Directory</h1>
          <p className="text-sm text-gray-500 mt-1">NIRF 2025 Engineering — {total} colleges</p>
        </div>

        {/* Search */}
        <div className="relative mb-4">
          <Search size={18} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
          <input data-testid="search-input" type="text" value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Search by college name, short name, city, or state..."
            className="w-full pl-10 pr-4 py-3 bg-white border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-200 focus:border-blue-400 outline-none transition-all" />
          {search && (
            <button onClick={() => setSearch("")} className="absolute right-3.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600">
              <X size={16} />
            </button>
          )}
        </div>

        {/* Filters Row */}
        <div className="flex flex-wrap gap-3 mb-4">
          {/* Rank Filter */}
          <div className="flex items-center gap-1.5 flex-wrap">
            <Filter size={14} className="text-gray-400" />
            {RANK_FILTERS.map(f => (
              <button key={f.key} data-testid={`rank-filter-${f.key || "all"}`}
                onClick={() => handleRankFilter(f.key)}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all
                  ${rankFilter === f.key ? "bg-[#1A73E8] text-white" : "bg-white border border-gray-200 text-gray-600 hover:border-blue-300"}`}>
                {f.label}
              </button>
            ))}
          </div>

          {/* Type Filter */}
          <div className="flex items-center gap-1.5 flex-wrap ml-auto">
            <Building2 size={14} className="text-gray-400" />
            <button onClick={() => { setTypeFilter(""); setPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all
                ${!typeFilter ? "bg-[#1A73E8] text-white" : "bg-white border border-gray-200 text-gray-600 hover:border-blue-300"}`}>
              All Types
            </button>
            {Object.entries(collegeTypes).sort((a, b) => b[1] - a[1]).map(([type, count]) => (
              <button key={type} data-testid={`type-filter-${type}`}
                onClick={() => { setTypeFilter(typeFilter === type ? "" : type); setPage(1); }}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-bold transition-all
                  ${typeFilter === type ? "bg-[#1A73E8] text-white" : "bg-white border border-gray-200 text-gray-600 hover:border-blue-300"}`}>
                {type} <span className="opacity-60 ml-0.5">{count}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Table */}
        <div className="bg-white rounded-2xl border border-gray-100 overflow-hidden" style={{ boxShadow: "0 2px 12px -4px rgba(0,0,0,0.06)" }}>
          {/* Table Header */}
          <div className="grid grid-cols-[60px_1fr_140px_180px_100px] gap-2 px-4 py-3 bg-gray-50 border-b border-gray-100 text-xs font-bold text-gray-500 uppercase tracking-wider">
            <div className="text-center">Rank</div>
            <div>College Name</div>
            <div>Type</div>
            <div className="flex items-center gap-1"><MapPin size={11} /> Location</div>
            <div>Short Name</div>
          </div>

          {/* Loading */}
          {loading && (
            <div className="flex items-center justify-center py-16">
              <div className="animate-spin w-6 h-6 border-2 border-[#1A73E8] border-t-transparent rounded-full" />
            </div>
          )}

          {/* No Results */}
          {!loading && colleges.length === 0 && (
            <div className="text-center py-16 text-gray-400 text-sm">No colleges found matching your criteria</div>
          )}

          {/* Rows */}
          {!loading && colleges.map((c, i) => (
            <div key={i} data-testid={`college-row-${i}`}
              className={`grid grid-cols-[60px_1fr_140px_180px_100px] gap-2 px-4 py-3 items-center text-sm border-b border-gray-50 hover:bg-blue-50/30 transition-colors ${i % 2 === 0 ? "" : "bg-gray-50/30"}`}>
              <div className="flex justify-center">
                <RankBadge rank={c.rank} />
              </div>
              <div className="font-medium text-gray-800 truncate" title={c.college_name}>{c.college_name}</div>
              <div><TypeBadge type={c.college_type} /></div>
              <div className="text-xs text-gray-500 truncate" title={`${c.city}, ${c.state}`}>
                {c.city}{c.city && c.state ? ", " : ""}{c.state}
              </div>
              <div className="text-xs text-blue-600 font-semibold truncate" title={(c.short_names || []).join(", ")}>
                {Array.isArray(c.short_names) ? c.short_names.join(", ") : c.short_names || ""}
              </div>
            </div>
          ))}
        </div>

        {/* Pagination */}
        {pages > 1 && (
          <div className="flex items-center justify-between mt-4 px-1">
            <p className="text-xs text-gray-400">
              Page {page} of {pages} — {total} colleges
            </p>
            <div className="flex gap-2">
              <button data-testid="prev-page" onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page <= 1}
                className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-bold bg-white border border-gray-200 text-gray-600 hover:border-blue-300 disabled:opacity-40 disabled:cursor-not-allowed transition-all">
                <ChevronLeft size={14} /> Prev
              </button>
              <button data-testid="next-page" onClick={() => setPage(p => Math.min(pages, p + 1))} disabled={page >= pages}
                className="flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-bold bg-white border border-gray-200 text-gray-600 hover:border-blue-300 disabled:opacity-40 disabled:cursor-not-allowed transition-all">
                Next <ChevronRight size={14} />
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
