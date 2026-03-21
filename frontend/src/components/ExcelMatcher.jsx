import React, { useState, useCallback } from "react";
import { Upload, FileSpreadsheet, ArrowRight, Download, CheckCircle2, AlertCircle, ChevronLeft, ArrowLeftRight, Layers } from "lucide-react";
import { useNavigate } from "react-router-dom";

const API = process.env.REACT_APP_BACKEND_URL + "/api";

export default function ExcelMatcher() {
  const navigate = useNavigate();
  const [step, setStep] = useState(1);
  const [leftFile, setLeftFile] = useState(null);
  const [rightFile, setRightFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [config, setConfig] = useState({ leftPhone: '', rightPhone: '', rightStatus: '', leftSheet: '' });
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');

  const handleUpload = useCallback(async () => {
    if (!leftFile || !rightFile) return;
    setLoading(true);
    setError('');
    try {
      const fd = new FormData();
      fd.append('left_file', leftFile);
      fd.append('right_file', rightFile);
      const res = await fetch(`${API}/excel-matcher/preview`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error((await res.json()).detail || 'Upload failed');
      const data = await res.json();
      setPreview(data);

      const phoneKw = ['phone', 'mobile', 'contact', 'number', 'tel'];
      const statusKw = ['status', 'shortlist', 'reject', 'result', 'decision'];
      const find = (cols, kw) => cols.find(c => kw.some(k => c.toLowerCase().includes(k))) || '';

      setConfig({
        leftSheet: data.left_sheets[0],
        leftPhone: find(data.left_columns, phoneKw),
        rightPhone: find(data.right_columns, phoneKw),
        rightStatus: find(data.right_columns, statusKw),
      });
      setStep(2);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [leftFile, rightFile]);

  const handleSheetChange = useCallback(async (sheetName) => {
    setConfig(p => ({ ...p, leftSheet: sheetName, leftPhone: '' }));
    setLoading(true);
    try {
      const res = await fetch(`${API}/excel-matcher/preview-sheet?left_sheet=${encodeURIComponent(sheetName)}`, { method: 'POST' });
      if (!res.ok) throw new Error((await res.json()).detail || 'Failed to load sheet');
      const data = await res.json();
      setPreview(p => ({ ...p, left_columns: data.left_columns, left_all_columns: data.left_all_columns, left_rows: data.left_rows, left_sample: data.left_sample }));
      const phoneKw = ['phone', 'mobile', 'contact', 'number', 'tel'];
      const find = (cols, kw) => cols.find(c => kw.some(k => c.toLowerCase().includes(k))) || '';
      setConfig(p => ({ ...p, leftPhone: find(data.left_columns, phoneKw) }));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const handleProcess = useCallback(async () => {
    if (!config.leftPhone || !config.rightPhone || !config.rightStatus) {
      setError('Please select all required columns');
      return;
    }
    setLoading(true);
    setError('');
    try {
      const params = new URLSearchParams({
        left_phone_col: config.leftPhone,
        right_phone_col: config.rightPhone,
        right_status_col: config.rightStatus,
        left_sheet: config.leftSheet,
      });
      const res = await fetch(`${API}/excel-matcher/process?${params}`, { method: 'POST' });
      if (!res.ok) throw new Error((await res.json()).detail || 'Processing failed');
      const matched = parseInt(res.headers.get('X-Matched') || '0');
      const total = parseInt(res.headers.get('X-Total') || '0');
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setResult({ url, matched, total, unmatched: total - matched });
      setStep(3);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [config]);

  const FileDropZone = ({ label, file, onFile, side, testId }) => (
    <div data-testid={testId}
      className={`relative border-2 border-dashed rounded-2xl p-6 text-center transition-all cursor-pointer
        ${file ? 'border-green-300 bg-green-50/50' : 'border-gray-200 hover:border-blue-300 hover:bg-blue-50/30'}`}
      onClick={() => document.getElementById(`file-${side}`).click()}
      onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add('border-blue-400', 'bg-blue-50'); }}
      onDragLeave={e => { e.currentTarget.classList.remove('border-blue-400', 'bg-blue-50'); }}
      onDrop={e => { e.preventDefault(); e.currentTarget.classList.remove('border-blue-400', 'bg-blue-50'); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}>
      <input id={`file-${side}`} type="file" accept=".xlsx,.xls" className="hidden"
        onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      {file ? (
        <div className="flex flex-col items-center gap-2">
          <CheckCircle2 size={28} className="text-green-500" />
          <p className="text-sm font-semibold text-gray-800 truncate max-w-[200px]">{file.name}</p>
          <p className="text-xs text-gray-400">{(file.size / 1024).toFixed(0)} KB</p>
        </div>
      ) : (
        <div className="flex flex-col items-center gap-2">
          <Upload size={28} className="text-gray-400" />
          <p className="text-sm font-semibold text-gray-600">{label}</p>
          <p className="text-xs text-gray-400">Drag & drop or click</p>
        </div>
      )}
    </div>
  );

  const ColSelect = ({ label, value, options, onChange, testId }) => (
    <div>
      <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1.5">{label}</label>
      <select data-testid={testId} value={value} onChange={e => onChange(e.target.value)}
        className="w-full px-3 py-2.5 border border-gray-200 rounded-xl text-sm bg-white focus:ring-2 focus:ring-blue-200 focus:border-blue-400 outline-none transition-all">
        <option value="">-- Select Column --</option>
        {options.map(c => <option key={c} value={c}>{c}</option>)}
      </select>
    </div>
  );

  return (
    <div className="min-h-screen bg-[#F7F8FA]" style={{ fontFamily: "'Inter', sans-serif" }}>
      <header data-testid="header" className="w-full py-4 px-6 bg-white border-b border-gray-100 flex items-center justify-center relative">
        <button data-testid="back-home-btn" onClick={() => navigate('/')}
          className="absolute left-6 flex items-center gap-1.5 text-sm text-gray-500 hover:text-gray-800 transition-colors">
          <ChevronLeft size={16} /> Home
        </button>
        <img data-testid="logo"
          src="https://customer-assets.emergentagent.com/job_premium-design-ui/artifacts/dwo5zlsk_logo-black121.png"
          alt="BluBridge" className="h-8" />
      </header>

      <div className="max-w-3xl mx-auto px-4 py-8">
        <div className="text-center mb-8">
          <div className="inline-flex items-center gap-2 bg-blue-50 text-[#1A73E8] px-4 py-1.5 rounded-full text-xs font-bold uppercase tracking-wider mb-3">
            <ArrowLeftRight size={14} /> Excel Matcher
          </div>
          <h1 className="text-2xl font-bold text-gray-900">Match & Tag Candidates</h1>
          <p className="text-sm text-gray-500 mt-1">Upload two Excel sheets, match by phone number, get Shortlist/Reject status</p>
        </div>

        {/* Progress */}
        <div className="flex items-center justify-center gap-3 mb-8">
          {[{n:1, t:'Upload'}, {n:2, t:'Configure'}, {n:3, t:'Download'}].map(({n, t}, i) => (
            <React.Fragment key={n}>
              {i > 0 && <div className={`w-8 h-0.5 ${step >= n ? 'bg-[#1A73E8]' : 'bg-gray-200'}`} />}
              <div className={`flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-bold
                ${step >= n ? 'bg-[#1A73E8] text-white' : 'bg-gray-100 text-gray-400'}`}>
                {n}. {t}
              </div>
            </React.Fragment>
          ))}
        </div>

        {error && (
          <div data-testid="error-msg" className="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl flex items-center gap-2 text-sm text-red-700">
            <AlertCircle size={16} /> {error}
          </div>
        )}

        {/* Step 1: Upload */}
        {step === 1 && (
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-gray-100">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
              <div>
                <h3 className="text-sm font-bold text-gray-700 mb-2 flex items-center gap-2">
                  <FileSpreadsheet size={16} className="text-blue-500" /> Left Sheet (Candidates)
                </h3>
                <FileDropZone label="Upload candidate list" file={leftFile} onFile={setLeftFile} side="left" testId="left-file-drop" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-gray-700 mb-2 flex items-center gap-2">
                  <FileSpreadsheet size={16} className="text-emerald-500" /> Right Sheet (Status)
                </h3>
                <FileDropZone label="Upload status sheet" file={rightFile} onFile={setRightFile} side="right" testId="right-file-drop" />
              </div>
            </div>
            <button data-testid="upload-btn" onClick={handleUpload}
              disabled={!leftFile || !rightFile || loading}
              className={`w-full py-3 rounded-xl text-sm font-bold flex items-center justify-center gap-2 transition-all
                ${leftFile && rightFile ? 'bg-[#1A73E8] text-white hover:bg-[#1557B0]' : 'bg-gray-100 text-gray-400 cursor-not-allowed'}`}>
              {loading ? <span className="animate-spin w-4 h-4 border-2 border-white border-t-transparent rounded-full" /> : <ArrowRight size={16} />}
              {loading ? 'Reading files...' : 'Upload & Preview'}
            </button>
          </div>
        )}

        {/* Step 2: Configure */}
        {step === 2 && preview && (
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-gray-100">
            <div className="grid grid-cols-2 gap-4 mb-4 text-center">
              <div className="bg-blue-50 rounded-xl p-3">
                <div className="text-lg font-bold text-[#1A73E8]">{preview.left_rows}</div>
                <div className="text-xs text-gray-500">Left Sheet Rows</div>
              </div>
              <div className="bg-emerald-50 rounded-xl p-3">
                <div className="text-lg font-bold text-emerald-600">{preview.right_rows}</div>
                <div className="text-xs text-gray-500">Right Sheet Rows (all tabs)</div>
              </div>
            </div>

            <div className="space-y-4 mb-6">
              {/* Left Sheet Config */}
              <div className="p-4 bg-blue-50/50 rounded-xl border border-blue-100">
                <h4 className="text-xs font-bold text-blue-700 uppercase tracking-wider mb-3">Left Sheet (Candidates)</h4>

                {/* Sheet Tab Selector */}
                {preview.left_sheets && preview.left_sheets.length > 1 && (
                  <div className="mb-3">
                    <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-1.5">
                      <Layers size={12} className="inline mr-1" /> Select Sheet Tab
                    </label>
                    <div className="flex flex-wrap gap-2">
                      {preview.left_sheets.map(s => (
                        <button key={s} data-testid={`sheet-tab-${s}`}
                          onClick={() => handleSheetChange(s)}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all
                            ${config.leftSheet === s
                              ? 'bg-[#1A73E8] text-white shadow-sm'
                              : 'bg-white border border-gray-200 text-gray-600 hover:border-blue-300'}`}>
                          {s}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                <ColSelect label="Phone Number Column" value={config.leftPhone} testId="left-phone-select"
                  options={preview.left_columns} onChange={v => setConfig(p => ({...p, leftPhone: v}))} />
              </div>

              {/* Right Sheet Config */}
              <div className="p-4 bg-emerald-50/50 rounded-xl border border-emerald-100">
                <h4 className="text-xs font-bold text-emerald-700 uppercase tracking-wider mb-1">Right Sheet (Status)</h4>
                <p className="text-xs text-emerald-600/70 mb-3 flex items-center gap-1">
                  <Layers size={11} /> Searches across all {preview.right_sheets?.length || 1} sheet tabs automatically
                </p>
                <div className="grid grid-cols-2 gap-3">
                  <ColSelect label="Phone Number Column" value={config.rightPhone} testId="right-phone-select"
                    options={preview.right_columns} onChange={v => setConfig(p => ({...p, rightPhone: v}))} />
                  <ColSelect label="Status Column" value={config.rightStatus} testId="right-status-select"
                    options={preview.right_columns} onChange={v => setConfig(p => ({...p, rightStatus: v}))} />
                </div>
              </div>
            </div>

            <div className="flex gap-3">
              <button data-testid="back-btn" onClick={() => setStep(1)}
                className="px-5 py-3 rounded-xl text-sm font-bold bg-gray-100 text-gray-600 hover:bg-gray-200 transition-all">
                Back
              </button>
              <button data-testid="process-btn" onClick={handleProcess} disabled={loading}
                className="flex-1 py-3 rounded-xl text-sm font-bold bg-[#1A73E8] text-white hover:bg-[#1557B0] flex items-center justify-center gap-2 transition-all">
                {loading ? <span className="animate-spin w-4 h-4 border-2 border-white border-t-transparent rounded-full" /> : <ArrowLeftRight size={16} />}
                {loading ? 'Matching...' : 'Match & Generate'}
              </button>
            </div>
          </div>
        )}

        {/* Step 3: Result */}
        {step === 3 && result && (
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-gray-100 text-center">
            <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <CheckCircle2 size={32} className="text-green-600" />
            </div>
            <h2 className="text-xl font-bold text-gray-900 mb-2">Matching Complete!</h2>
            <div className="grid grid-cols-3 gap-3 mb-6 max-w-md mx-auto">
              <div className="bg-gray-50 rounded-xl p-3">
                <div data-testid="total-count" className="text-lg font-bold text-gray-800">{result.total}</div>
                <div className="text-xs text-gray-500">Total</div>
              </div>
              <div className="bg-green-50 rounded-xl p-3">
                <div data-testid="matched-count" className="text-lg font-bold text-green-600">{result.matched}</div>
                <div className="text-xs text-gray-500">Matched</div>
              </div>
              <div className="bg-red-50 rounded-xl p-3">
                <div data-testid="unmatched-count" className="text-lg font-bold text-red-500">{result.unmatched}</div>
                <div className="text-xs text-gray-500">Not Found</div>
              </div>
            </div>
            <a data-testid="download-result-btn" href={result.url} download="Matched_Results.xlsx"
              className="inline-flex items-center gap-2 px-6 py-3 bg-[#1A73E8] text-white rounded-xl text-sm font-bold hover:bg-[#1557B0] transition-all">
              <Download size={16} /> Download Result
            </a>
            <button data-testid="start-over-btn" onClick={() => { setStep(1); setLeftFile(null); setRightFile(null); setPreview(null); setResult(null); setError(''); }}
              className="block mx-auto mt-4 text-sm text-gray-500 hover:text-gray-800 transition-colors">
              Start Over
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
