/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef } from "react";
import { 
  Upload, 
  FileText, 
  Music, 
  Play, 
  Download, 
  CheckCircle2, 
  AlertCircle,
  Loader2,
  Trash2,
  Plus,
  Sparkles,
  Settings,
  X,
  FileArchive,
  BookOpen,
  Sun,
  Moon
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import JSZip from "jszip";
import Fuse from "fuse.js";
import * as idb from "idb-keyval";
import { parseHymnList, parseLiturgySheet, HymnSlot, LiturgyData } from "./services/geminiService";
import { mergePPTX, generateSampleMaster, extractPptxText } from "./services/pptxService";

interface HymnArchiveItem {
  fileName: string;
  buffer: ArrayBuffer;
  content: string;
}

export default function App() {
  // Persistent Setup State (IndexedDB)
  const [masterFile, setMasterFile] = useState<{ name: string; buffer: ArrayBuffer } | null>(null);
  const [hymnArchive, setHymnArchive] = useState<HymnArchiveItem[]>([]);
  
  // Main Input State
  const [hymnListText, setHymnListText] = useState("");
  const [liturgyFile, setLiturgyFile] = useState<File | null>(null);
  
  // UI State
  const [showSetup, setShowSetup] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isAssembling, setIsAssembling] = useState(false);
  const [isIndexing, setIsIndexing] = useState(false);
  const [indexingProgress, setIndexingProgress] = useState("");
  const [status, setStatus] = useState<string>("");
  const [finalBlob, setFinalBlob] = useState<Blob | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [setupError, setSetupError] = useState<string | null>(null);
  const [parsedSlots, setParsedSlots] = useState<HymnSlot[]>([]);
  const [liturgyData, setLiturgyData] = useState<LiturgyData | null>(null);
  const [progress, setProgress] = useState(0);
  const [step, setStep] = useState<"input" | "review" | "result">("input");
  const [logs, setLogs] = useState<{ message: string; type: "info" | "success" | "error"; timestamp: string }[]>([]);
  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== "undefined") {
      return localStorage.getItem("theme") === "dark";
    }
    return false;
  });

  const toggleTheme = () => {
    setDarkMode(prev => {
      const next = !prev;
      localStorage.setItem("theme", next ? "dark" : "light");
      return next;
    });
  };

  const masterInputRef = useRef<HTMLInputElement>(null);
  const zipInputRef = useRef<HTMLInputElement>(null);
  const liturgyInputRef = useRef<HTMLInputElement>(null);
  const logEndRef = useRef<HTMLDivElement>(null);

  const addLog = useCallback((message: string, type: "info" | "success" | "error" = "info") => {
    setLogs(prev => [...prev, { message, type, timestamp: new Date().toLocaleTimeString() }]);
  }, []);

  // Auto-scroll logs
  React.useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [logs]);

  // Load persisted data on mount
  React.useEffect(() => {
    const loadPersisted = async () => {
      try {
        const persistedMaster = await idb.get("masterFile");
        const persistedArchive = await idb.get("hymnArchive");
        if (persistedMaster) setMasterFile(persistedMaster);
        if (persistedArchive) setHymnArchive(persistedArchive);
      } catch (err) {
        console.error("Failed to load persisted data", err);
      }
    };
    loadPersisted();
  }, []);

  // Load Zip Archive
  const handleZipUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsIndexing(true);
    setSetupError(null);
    setIndexingProgress("Reading ZIP file...");
    
    try {
      const zip = await JSZip.loadAsync(file);
      const archiveItems: HymnArchiveItem[] = [];
      
      const files = Object.entries(zip.files).filter(([path, f]) => !f.dir && path.toLowerCase().endsWith(".pptx"));
      
      if (files.length === 0) {
        throw new Error("No .pptx files found in the ZIP archive.");
      }

      let count = 0;
      for (const [path, f] of files) {
        try {
          const buffer = await f.async("arraybuffer");
          const fileName = path.split("/").pop() || path;
          setIndexingProgress(`Indexing hymn ${count + 1}/${files.length}: ${fileName}`);
          
          let slides: string[] = [];
          try {
            slides = await extractPptxText(buffer);
          } catch (extractionErr) {
            console.warn(`Could not extract text from ${fileName}, using filename only.`, extractionErr);
          }

          archiveItems.push({
            fileName,
            buffer,
            content: slides.join(" ")
          });
        } catch (fileErr) {
          console.error(`Error processing file ${path}:`, fileErr);
        }
        count++;
      }
      
      setHymnArchive(archiveItems);
      await idb.set("hymnArchive", archiveItems);
      setIndexingProgress("");
    } catch (err: any) {
      console.error("ZIP Upload Error:", err);
      setSetupError(err.message || "Failed to process ZIP file.");
    } finally {
      setIsIndexing(false);
      if (e.target) e.target.value = "";
    }
  };

  const handleMasterUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const masterData = { name: file.name, buffer };
      setMasterFile(masterData);
      await idb.set("masterFile", masterData);
    } catch (err) {
      setSetupError("Failed to read master file.");
    }
  };

  const handleResetSetup = async () => {
    if (confirm("Are you sure you want to reset the setup? This will remove the persisted master template and hymn archive.")) {
      await idb.del("masterFile");
      await idb.del("hymnArchive");
      setMasterFile(null);
      setHymnArchive([]);
    }
  };

  const handleAnalyze = async () => {
    if (!hymnListText.trim() || !liturgyFile) {
      setError("Please provide both the hymn list and the liturgy PDF.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);
    setFinalBlob(null);
    setProgress(0);
    setLogs([]);
    addLog("Starting analysis...", "info");

    try {
      // 1. Reading PDF
      setStatus("Reading liturgy PDF...");
      addLog("Reading liturgy PDF file...", "info");
      setProgress(20);
      const base64Pdf = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve((reader.result as string).split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(liturgyFile);
      });

      // 2. Extracting Liturgy
      setStatus("Extracting liturgy data with AI...");
      addLog("Sending PDF to Gemini for extraction...", "info");
      setProgress(50);
      const liturgy = await parseLiturgySheet(base64Pdf);
      setLiturgyData(liturgy);
      addLog("Liturgy data extracted successfully.", "success");

      // 3. Parsing Hymn List
      setStatus("Parsing hymn list with AI...");
      addLog("Parsing hymn list text...", "info");
      setProgress(80);
      const slots = await parseHymnList(hymnListText);
      if (!slots || slots.length === 0) {
        throw new Error("AI could not identify any hymns in the provided list.");
      }
      setParsedSlots(slots);
      addLog(`Identified ${slots.length} hymn slots.`, "success");
      
      setProgress(100);
      setStatus("Analysis complete. Please review the data below.");
      addLog("Analysis complete.", "success");
      setStep("review");
    } catch (err: any) {
      console.error("Analysis Error:", err);
      setError(err.message || "An error occurred during analysis.");
      setStatus("Analysis failed.");
      addLog(`Analysis failed: ${err.message}`, "error");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleAssemble = async () => {
    if (!masterFile || hymnArchive.length === 0 || !liturgyData || parsedSlots.length === 0) {
      setError("Missing required data for assembly.");
      return;
    }

    setIsAssembling(true);
    setError(null);
    setProgress(0);
    addLog("Starting assembly process...", "info");

    try {
      setStatus("Matching hymns and assembling presentation...");
      addLog("Initializing fuzzy search for hymns...", "info");
      setProgress(30);
      
      const fuse = new Fuse<HymnArchiveItem>(hymnArchive, {
        keys: [
          { name: "fileName", weight: 0.7 },
          { name: "content", weight: 0.3 }
        ],
        threshold: 0.4
      });

      const matchedHymns = new Map<string, ArrayBuffer>();
      const matchedMappings: HymnSlot[] = [];

      for (const slot of parsedSlots) {
        addLog(`Matching hymn for slot: ${slot.slot} (${slot.hymnName})...`, "info");
        let results = fuse.search(slot.hymnName);
        if (results.length === 0 || results[0].score! > 0.3) {
          const lyricsResults = fuse.search(slot.lyricsSnippet || slot.hymnName);
          if (lyricsResults.length > 0 && (!results.length || lyricsResults[0].score! < results[0].score!)) {
            results = lyricsResults;
          }
        }

        if (results.length > 0) {
          const bestMatch = results[0].item;
          matchedHymns.set(slot.slot, bestMatch.buffer);
          matchedMappings.push({ 
            ...slot, 
            hymnName: bestMatch.fileName // Use the actual file name for matching in mergePPTX
          });
          addLog(`Matched ${slot.slot} to ${bestMatch.fileName}`, "success");
        } else {
          matchedMappings.push(slot);
          addLog(`No match found for ${slot.slot} (${slot.hymnName}). Will fallback to lyrics.`, "info");
        }
      }

      setProgress(60);
      const blob = await mergePPTX(masterFile.buffer, matchedHymns, matchedMappings, liturgyData, addLog);
      setFinalBlob(blob);
      setProgress(100);
      setStatus("Presentation assembled successfully!");
      addLog("Presentation assembly finished.", "success");
      setStep("result");
    } catch (err: any) {
      console.error("Assembly Error:", err);
      setError(err.message || "An error occurred during assembly.");
      setStatus("Assembly failed.");
      addLog(`Assembly failed: ${err.message}`, "error");
    } finally {
      setIsAssembling(false);
    }
  };

  const downloadResult = () => {
    if (!finalBlob) return;
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Sunday_Mass_${new Date().toISOString().split('T')[0]}.pptx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className={`min-h-screen transition-colors duration-500 ${darkMode ? "atmosphere-dark" : "atmosphere-light"}`}>
      {/* Setup Modal */}
      <AnimatePresence>
        {showSetup && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-md"
          >
            <motion.div 
              initial={{ scale: 0.95, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              className={`${darkMode ? "glass-card-dark" : "bg-white"} rounded-3xl shadow-2xl w-full max-w-2xl overflow-hidden`}
            >
              <div className={`p-6 border-b ${darkMode ? "border-white/10 bg-white/5" : "border-stone-100 bg-stone-50"} flex items-center justify-between`}>
                <div className="flex items-center gap-3">
                  <Settings className="text-stone-400" />
                  <h2 className="text-xl font-bold">Initial Setup</h2>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={async () => {
                      const blob = await generateSampleMaster();
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement("a");
                      a.href = url;
                      a.download = "Sample_Master_Template.pptx";
                      a.click();
                    }}
                    className={`text-xs font-bold ${darkMode ? "text-emerald-400 bg-emerald-900/30" : "text-emerald-600 bg-emerald-50"} px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1.5`}
                  >
                    <Download size={14} />
                    Download Sample Master
                  </button>
                  <button onClick={() => setShowSetup(false)} className={`p-2 ${darkMode ? "hover:bg-white/10" : "hover:bg-stone-200"} rounded-full transition-colors`}>
                    <X size={20} />
                  </button>
                </div>
              </div>
              
              <div className="p-8 space-y-8">
                <p className={`${darkMode ? "text-stone-400" : "text-stone-500"} text-sm`}>Upload your master template and hymn archive once. These will be used for all future assemblies in this session.</p>
                
                {setupError && (
                  <div className={`p-4 ${darkMode ? "bg-red-900/20 border-red-900/30 text-red-400" : "bg-red-50 border-red-100 text-red-600"} border rounded-2xl flex items-center gap-3 text-sm`}>
                    <AlertCircle size={18} />
                    {setupError}
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Master Template */}
                  <div className="space-y-3">
                    <label className={`text-xs font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"}`}>Master Template</label>
                    <div 
                      onClick={() => masterInputRef.current?.click()}
                      className={`border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all ${
                        masterFile ? (darkMode ? "border-emerald-500/50 bg-emerald-500/10" : "border-emerald-200 bg-emerald-50") : 
                        (darkMode ? "border-white/10 hover:border-white/20" : "border-stone-200 hover:border-stone-300")
                      }`}
                    >
                      <input type="file" ref={masterInputRef} onChange={handleMasterUpload} className="hidden" accept=".pptx" />
                      {masterFile ? (
                        <div className={`${darkMode ? "text-emerald-400" : "text-emerald-700"} font-bold text-sm truncate`}>{masterFile.name}</div>
                      ) : (
                        <div className={darkMode ? "text-stone-500" : "text-stone-400"}>Upload Master .pptx</div>
                      )}
                    </div>
                  </div>

                  {/* Hymn Archive */}
                  <div className="space-y-3">
                    <label className={`text-xs font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"}`}>Hymn Archive (ZIP)</label>
                    <div 
                      onClick={() => !isIndexing && zipInputRef.current?.click()}
                      className={`border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all ${
                        isIndexing ? (darkMode ? "border-indigo-500/50 bg-indigo-500/10 cursor-wait" : "border-indigo-200 bg-indigo-50/50 cursor-wait") : 
                        hymnArchive.length > 0 ? (darkMode ? "border-indigo-500/50 bg-indigo-500/10" : "border-indigo-200 bg-indigo-50") : 
                        (darkMode ? "border-white/10 hover:border-white/20" : "border-stone-200 hover:border-stone-300")
                      }`}
                    >
                      <input type="file" ref={zipInputRef} onChange={handleZipUpload} className="hidden" accept=".zip" />
                      {isIndexing ? (
                        <div className="flex flex-col items-center gap-2">
                          <Loader2 className={`animate-spin ${darkMode ? "text-indigo-400" : "text-indigo-600"}`} size={20} />
                          <div className={`${darkMode ? "text-indigo-300" : "text-indigo-700"} font-bold text-[10px] uppercase tracking-tighter truncate w-full px-2`}>
                            {indexingProgress}
                          </div>
                        </div>
                      ) : hymnArchive.length > 0 ? (
                        <div className={`${darkMode ? "text-indigo-400" : "text-indigo-700"} font-bold text-sm`}>{hymnArchive.length} Hymns Loaded</div>
                      ) : (
                        <div className={darkMode ? "text-stone-500" : "text-stone-400"}>Upload Hymns .zip</div>
                      )}
                    </div>
                  </div>
                </div>

                <div className="flex items-center justify-between pt-4">
                  <button 
                    onClick={handleResetSetup}
                    className={`text-xs font-bold ${darkMode ? "text-red-400 hover:text-red-300" : "text-red-600 hover:text-red-700"} flex items-center gap-1.5`}
                  >
                    <Trash2 size={14} />
                    Reset Setup
                  </button>
                  <button 
                    onClick={() => setShowSetup(false)}
                    disabled={isIndexing}
                    className={`${darkMode ? "bg-white text-stone-900 hover:bg-stone-100" : "bg-stone-900 text-white hover:bg-stone-800"} font-bold py-3 px-8 rounded-xl disabled:bg-stone-100 disabled:text-stone-400 transition-all shadow-lg`}
                  >
                    Save & Continue
                  </button>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <div className="max-w-6xl mx-auto p-4 md:p-12">
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-6 mb-16">
          <div className="space-y-2">
            <h1 className={`text-5xl font-serif italic tracking-tight ${darkMode ? "text-white" : "text-stone-900"}`}>Sunday Mass Assembly</h1>
            <p className={`${darkMode ? "text-stone-400" : "text-stone-500"} text-lg`}>Generate your liturgical presentation in seconds.</p>
          </div>
          <div className="flex items-center gap-4">
            <button
              onClick={toggleTheme}
              className={`p-3 rounded-2xl transition-all ${darkMode ? "bg-white/10 text-yellow-400 hover:bg-white/20" : "bg-white/60 backdrop-blur-xl text-stone-600 border border-white/40 hover:bg-white/80 shadow-sm"}`}
            >
              {darkMode ? <Sun size={20} /> : <Moon size={20} />}
            </button>
            <button 
              onClick={() => setShowSetup(true)}
              className={`flex items-center gap-2 px-6 py-3 ${darkMode ? "bg-white/10 text-white border border-white/10 hover:bg-white/20" : "bg-white/60 backdrop-blur-xl border border-white/40 text-stone-600 hover:bg-white/80 shadow-sm"} rounded-2xl font-bold text-sm transition-all`}
            >
              <Settings size={18} />
              Setup Archive
              {(masterFile && hymnArchive.length > 0) && <CheckCircle2 size={16} className="text-emerald-500" />}
            </button>
          </div>
        </header>

        <div className="mb-12">
          <h2 className={`text-3xl font-bold ${darkMode ? "text-white" : "text-stone-900"}`}>Today</h2>
          <p className={`${darkMode ? "text-stone-400" : "text-stone-500"} text-xl`}>
            {new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' })}
          </p>
        </div>

        <main className="grid grid-cols-1 lg:grid-cols-12 gap-12">
          {/* Inputs */}
          <div className="lg:col-span-7 space-y-12">
            {step === "input" ? (
              <>
                {/* Hymn List */}
                <section className="space-y-4">
                  <div className="flex items-center gap-3">
                    <div className={`p-2 ${darkMode ? "bg-amber-900/30 text-amber-400" : "bg-amber-50 text-amber-600"} rounded-xl`}>
                      <Music size={24} />
                    </div>
                    <h2 className={`text-xl font-bold ${darkMode ? "text-white" : "text-stone-900"}`}>Hymn List</h2>
                  </div>
                  <textarea 
                    value={hymnListText}
                    onChange={(e) => setHymnListText(e.target.value)}
                    placeholder="Paste the hymn list for this Sunday..."
                    className={`w-full h-64 p-6 ${darkMode ? "bg-white/5 border-white/10 text-white focus:ring-amber-900/30 focus:border-amber-500" : "bg-white border-stone-200 text-stone-900 focus:ring-amber-50 focus:border-amber-400"} border rounded-3xl outline-none transition-all text-sm leading-relaxed shadow-sm resize-none font-mono`}
                  />
                </section>

                {/* Liturgy Sheet */}
                <section className="space-y-4">
                  <div className="flex items-center gap-3">
                    <div className={`p-2 ${darkMode ? "bg-emerald-900/30 text-emerald-400" : "bg-emerald-50 text-emerald-600"} rounded-xl`}>
                      <BookOpen size={24} />
                    </div>
                    <h2 className={`text-xl font-bold ${darkMode ? "text-white" : "text-stone-900"}`}>Liturgy Sheet</h2>
                  </div>
                  <div 
                    onClick={() => liturgyInputRef.current?.click()}
                    className={`border-2 border-dashed rounded-3xl p-12 text-center cursor-pointer transition-all ${
                      darkMode ? "bg-white/5 shadow-sm" : "bg-white shadow-sm"
                    } ${
                      liturgyFile ? (darkMode ? "border-emerald-500/50 bg-emerald-500/10" : "border-emerald-200 bg-emerald-50/30") : 
                      (darkMode ? "border-white/10 hover:border-white/20" : "border-stone-200 hover:border-stone-300")
                    }`}
                  >
                    <input type="file" ref={liturgyInputRef} onChange={(e) => e.target.files && setLiturgyFile(e.target.files[0])} className="hidden" accept=".pdf" />
                    {liturgyFile ? (
                      <div className="flex flex-col items-center gap-3">
                        <div className={`w-16 h-16 ${darkMode ? "bg-emerald-900/30 text-emerald-400" : "bg-emerald-100 text-emerald-600"} rounded-2xl flex items-center justify-center`}>
                          <FileText size={32} />
                        </div>
                        <p className={`font-bold ${darkMode ? "text-emerald-300" : "text-emerald-900"}`}>{liturgyFile.name}</p>
                        <p className={`text-xs ${darkMode ? "text-emerald-500" : "text-emerald-600"}`}>Click to change PDF</p>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        <Upload className={`mx-auto ${darkMode ? "text-stone-600" : "text-stone-300"}`} size={48} />
                        <div>
                          <p className={`font-bold ${darkMode ? "text-stone-400" : "text-stone-600"}`}>Upload Liturgy PDF</p>
                          <p className={`text-sm ${darkMode ? "text-stone-500" : "text-stone-400"} mt-1`}>AI will extract readings and psalm verses</p>
                        </div>
                      </div>
                    )}
                  </div>
                </section>
              </>
            ) : (
              <motion.div 
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                className="space-y-8"
              >
                <div className="flex items-center justify-between">
                  <h2 className={`text-2xl font-serif italic ${darkMode ? "text-white" : "text-stone-900"}`}>Review Extracted Data</h2>
                  <button 
                    onClick={() => setStep("input")}
                    className={`text-xs font-bold ${darkMode ? "text-stone-500 hover:text-stone-400" : "text-stone-400 hover:text-stone-600"} flex items-center gap-1`}
                  >
                    <X size={14} />
                    Start Over
                  </button>
                </div>

                {liturgyData && (
                  <div className={`${darkMode ? "glass-card-dark" : "glass-card"} p-8 card-rounded space-y-6`}>
                    <div>
                      <h4 className={`text-[10px] font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"} mb-2`}>Liturgy Title</h4>
                      <p className={`text-xl font-serif italic ${darkMode ? "text-white" : "text-stone-900"}`}>{liturgyData.title}</p>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div className="space-y-4">
                        <h4 className={`text-[10px] font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"}`}>Readings</h4>
                        <div className="space-y-4">
                          {[liturgyData.firstReading, liturgyData.secondReading, liturgyData.gospel].map((r, i) => (
                            <div key={i} className={`p-4 ${darkMode ? "bg-white/5 border-white/5" : "bg-white/40 border-white/20"} rounded-3xl border`}>
                              <p className={`text-xs font-bold ${darkMode ? "text-stone-200" : "text-stone-900"}`}>{r.title}</p>
                              <p className={`text-[10px] ${darkMode ? "text-stone-500" : "text-stone-500"} italic mb-2`}>{r.reference}</p>
                              <p className={`text-[11px] ${darkMode ? "text-stone-400" : "text-stone-600"} line-clamp-3`}>{r.text}</p>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="space-y-4">
                        <h4 className={`text-[10px] font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"}`}>Psalm & Response</h4>
                        <div className={`p-4 ${darkMode ? "bg-amber-900/20 border-amber-900/30" : "bg-amber-50/30 border-amber-100/50"} rounded-3xl border`}>
                          <p className={`text-xs font-bold ${darkMode ? "text-amber-400" : "text-amber-900"} mb-1`}>Response</p>
                          <p className={`text-sm font-serif italic ${darkMode ? "text-amber-200" : "text-amber-800"} mb-4`}>{liturgyData.psalm.response}</p>
                          <p className={`text-xs font-bold ${darkMode ? "text-amber-400" : "text-amber-900"} mb-2`}>Verses</p>
                          <div className="space-y-2">
                            {liturgyData.psalm.verses.map((v, i) => (
                              <p key={i} className={`text-[11px] ${darkMode ? "text-amber-300 bg-white/5 border-white/5" : "text-amber-700 bg-white/40 border-amber-100/30"} p-2 rounded-xl border line-clamp-2`}>
                                <span className="font-bold mr-2">{i+1}.</span>
                                {v}
                              </p>
                            ))}
                          </div>
                        </div>
                        <div className={`p-4 ${darkMode ? "bg-emerald-900/20 border-emerald-900/30" : "bg-emerald-50/30 border-emerald-100/50"} rounded-3xl border`}>
                          <p className={`text-xs font-bold ${darkMode ? "text-emerald-400" : "text-emerald-900"} mb-1`}>Faithful Response</p>
                          <p className={`text-sm italic ${darkMode ? "text-emerald-200" : "text-emerald-800"} mb-4`}>{liturgyData.prayerOfTheFaithfulResponse}</p>
                          {(liturgyData.prayerTitle || liturgyData.prayerText) && (
                            <div className={`pt-4 border-t ${darkMode ? "border-white/10" : "border-emerald-100/30"}`}>
                              <p className={`text-xs font-bold ${darkMode ? "text-emerald-400" : "text-emerald-900"} mb-1`}>{liturgyData.prayerTitle || "Prayer"}</p>
                              <p className={`text-[11px] ${darkMode ? "text-emerald-300" : "text-emerald-700"} line-clamp-3`}>{liturgyData.prayerText}</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                <div className={`${darkMode ? "glass-card-dark" : "glass-card"} p-8 card-rounded`}>
                  <h4 className={`text-[10px] font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"} mb-4`}>Hymn Slots</h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {parsedSlots.map((s, i) => (
                      <div key={i} className={`flex items-center justify-between p-3 ${darkMode ? "bg-white/5 border-white/5" : "bg-white/40 border-white/20"} rounded-2xl border`}>
                        <span className={`text-[10px] font-bold ${darkMode ? "text-stone-500" : "text-stone-400"} uppercase`}>{s.slot}</span>
                        <span className={`text-xs font-bold ${darkMode ? "text-stone-200" : "text-stone-700"}`}>{s.hymnName}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </motion.div>
            )}
          </div>

          {/* Actions & Status */}
          <div className="lg:col-span-5">
            <div className="sticky top-8 space-y-6">
              <div className={`${darkMode ? "glass-card-dark" : "glass-card"} p-8 card-rounded`}>
                <div className="flex items-center justify-between mb-8">
                  <h3 className={`text-lg font-bold ${darkMode ? "text-white" : "text-stone-900"}`}>
                    {step === "input" ? "Step 1: Analyze Data" : step === "review" ? "Step 2: Assemble PPTX" : "Step 3: Download"}
                  </h3>
                  <img 
                    src="https://picsum.photos/seed/liturgy/100/100" 
                    alt="User" 
                    className="w-10 h-10 rounded-full border-2 border-white/40 shadow-sm"
                    referrerPolicy="no-referrer"
                  />
                </div>
                
                <div className="space-y-4 mb-8">
                  <div className="flex items-center justify-between text-sm">
                    <span className={`font-medium ${darkMode ? "text-stone-400" : "text-stone-400"}`}>Master Template</span>
                    {masterFile ? <CheckCircle2 size={18} className="text-emerald-500" /> : <X size={18} className={`${darkMode ? "text-stone-700" : "text-stone-300"}`} />}
                  </div>
                  <div className="flex items-center justify-between text-sm">
                    <span className={`font-medium ${darkMode ? "text-stone-400" : "text-stone-400"}`}>Hymn Archive</span>
                    {hymnArchive.length > 0 ? <CheckCircle2 size={18} className="text-emerald-500" /> : <X size={18} className={`${darkMode ? "text-stone-700" : "text-stone-300"}`} />}
                  </div>
                  <div className="flex items-center justify-between text-sm">
                    <span className={`font-medium ${darkMode ? "text-stone-400" : "text-stone-400"}`}>Hymn List</span>
                    {hymnListText.trim() ? <CheckCircle2 size={18} className="text-emerald-500" /> : <X size={18} className={`${darkMode ? "text-stone-700" : "text-stone-300"}`} />}
                  </div>
                  <div className="flex items-center justify-between text-sm">
                    <span className={`font-medium ${darkMode ? "text-stone-400" : "text-stone-400"}`}>Liturgy PDF</span>
                    {liturgyFile ? <CheckCircle2 size={18} className="text-emerald-500" /> : <X size={18} className={`${darkMode ? "text-stone-700" : "text-stone-300"}`} />}
                  </div>
                </div>

                <AnimatePresence>
                  {(status || error) && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className={`p-5 rounded-2xl mb-6 text-sm font-medium ${
                        error 
                          ? (darkMode ? "bg-red-900/20 text-red-400 border border-red-900/30" : "bg-red-50 text-red-600 border border-red-100") 
                          : (darkMode ? "bg-blue-900/20 text-blue-400 border border-blue-900/30" : "bg-blue-50 text-blue-600 border border-blue-100")
                      }`}
                    >
                      <div className="flex items-center gap-3 mb-3">
                        {error ? <AlertCircle size={18} /> : <Loader2 size={18} className={(isAnalyzing || isAssembling) ? "animate-spin" : ""} />}
                        <span className="flex-1">{error || status}</span>
                      </div>
                      
                      {(isAnalyzing || isAssembling) && (
                        <div className={`w-full ${darkMode ? "bg-white/10" : "bg-blue-200/50"} rounded-full h-1.5 overflow-hidden`}>
                          <motion.div 
                            className={`h-full ${darkMode ? "bg-blue-500" : "bg-blue-600"}`}
                            initial={{ width: 0 }}
                            animate={{ width: `${progress}%` }}
                            transition={{ duration: 0.5 }}
                          />
                        </div>
                      )}
                    </motion.div>
                  )}
                </AnimatePresence>

                {step === "input" && (
                  <button 
                    onClick={handleAnalyze}
                    disabled={isAnalyzing || !hymnListText.trim() || !liturgyFile}
                    className={`w-full font-bold py-5 rounded-full transition-all shadow-lg flex items-center justify-center gap-3 group ${
                      isAnalyzing || !hymnListText.trim() || !liturgyFile
                        ? (darkMode ? "bg-white/5 text-stone-600 cursor-not-allowed" : "bg-stone-100 text-stone-400 cursor-not-allowed")
                        : (darkMode ? "bg-white text-black hover:bg-stone-200" : "bg-black text-white hover:bg-stone-800")
                    }`}
                  >
                    {isAnalyzing ? <Loader2 className="animate-spin" size={24} /> : (
                      <>
                        Analyze & Extract
                        <Sparkles size={20} className="group-hover:rotate-12 transition-transform" />
                      </>
                    )}
                  </button>
                )}

                {step === "review" && (
                  <button 
                    onClick={handleAssemble}
                    disabled={isAssembling || !masterFile || hymnArchive.length === 0}
                    className={`w-full font-bold py-5 rounded-full transition-all shadow-lg flex items-center justify-center gap-3 group ${
                      isAssembling || !masterFile || hymnArchive.length === 0
                        ? (darkMode ? "bg-white/5 text-stone-600 cursor-not-allowed" : "bg-stone-100 text-stone-400 cursor-not-allowed")
                        : "bg-emerald-600 text-white hover:bg-emerald-700"
                    }`}
                  >
                    {isAssembling ? <Loader2 className="animate-spin" size={24} /> : (
                      <>
                        Confirm & Assemble
                        <Play size={20} />
                      </>
                    )}
                  </button>
                )}

                {step === "result" && finalBlob && (
                  <motion.button 
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    onClick={downloadResult}
                    className="w-full bg-emerald-600 text-white font-bold py-5 rounded-full hover:bg-emerald-700 transition-all shadow-lg flex items-center justify-center gap-3"
                  >
                    <Download size={24} />
                    Download PPTX
                  </motion.button>
                )}

                {/* Activity Log */}
                <div className={`mt-8 border-t ${darkMode ? "border-white/10" : "border-stone-100"} pt-8`}>
                  <h4 className={`text-[10px] font-black uppercase tracking-widest ${darkMode ? "text-stone-500" : "text-stone-400"} mb-4`}>Activity Log</h4>
                  <div className={`${darkMode ? "bg-white/5 border border-white/5" : "bg-stone-900"} rounded-2xl p-4 h-48 overflow-y-auto font-mono text-[10px] space-y-1 scrollbar-hide`}>
                    {logs.length === 0 && <p className={`${darkMode ? "text-stone-600" : "text-stone-600"} italic`}>No activity yet...</p>}
                    {logs.map((log, i) => (
                      <div key={i} className="flex gap-2">
                        <span className={`${darkMode ? "text-stone-500" : "text-stone-500"} shrink-0`}>[{log.timestamp}]</span>
                        <span className={
                          log.type === "success" ? "text-emerald-400" :
                          log.type === "error" ? "text-red-400" :
                          (darkMode ? "text-stone-300" : "text-stone-300")
                        }>
                          {log.message}
                        </span>
                      </div>
                    ))}
                    <div ref={logEndRef} />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}
