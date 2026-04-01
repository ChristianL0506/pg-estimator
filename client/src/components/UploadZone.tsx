import { useState, useRef, useCallback, useEffect } from "react";
import { Upload, FileText, AlertCircle, CheckCircle2, FileStack, CloudLightning, AlertTriangle, WifiOff, Layers } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { useToast } from "@/hooks/use-toast";
import { queryClient, getAuthToken } from "@/lib/queryClient";
import { useQuery } from "@tanstack/react-query";
import { apiRequest } from "@/lib/queryClient";

interface UploadZoneProps {
  discipline: "mechanical" | "structural" | "civil";
  onProjectCreated: (projectId: string) => void;
}

type UploadState = "idle" | "uploading" | "processing" | "done" | "error";

export default function UploadZone({ discipline, onProjectCreated }: UploadZoneProps) {
  const [state, setState] = useState<UploadState>("idle");
  const [progress, setProgress] = useState(0);
  const [statusMsg, setStatusMsg] = useState("");
  const [warnings, setWarnings] = useState<string[]>([]);
  const [pdfQuality, setPdfQuality] = useState<string | null>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [hasRevisions, setHasRevisions] = useState(false);
  const [dualModel, setDualModel] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [pollWarning, setPollWarning] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const consecutiveFailuresRef = useRef(0);
  const consecutive404sRef = useRef(0);
  const { toast } = useToast();

  const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

  // Check if Gemini key is configured for dual-model option
  const { data: geminiStatus } = useQuery<{ configured: boolean }>({
    queryKey: ["/api/settings/gemini-key"],
    queryFn: async () => {
      const res = await apiRequest("GET", "/api/settings/gemini-key");
      return res.json();
    },
  });
  const geminiAvailable = geminiStatus?.configured ?? false;

  // Cleanup polling interval on unmount
  useEffect(() => {
    return () => { if (pollRef.current) clearInterval(pollRef.current); };
  }, []);

  const stopPolling = useCallback(() => {
    if (pollRef.current) { clearInterval(pollRef.current); pollRef.current = null; }
  }, []);

  const pollProgress = useCallback((jobId: string) => {
    let attempts = 0;
    consecutiveFailuresRef.current = 0;
    consecutive404sRef.current = 0;
    setPollWarning(null);

    pollRef.current = setInterval(async () => {
      attempts++;
      if (attempts > 1800) { // 60min max
        stopPolling();
        setState("error");
        setIsProcessing(false);
        setStatusMsg("Timeout — processing took too long.");
        return;
      }
      try {
        const authHeaders: Record<string, string> = {};
        const t = getAuthToken();
        if (t) authHeaders["Authorization"] = `Bearer ${t}`;
        const res = await fetch(`${API_BASE}/api/progress/${jobId}`, { headers: authHeaders });

        // Handle 404 — job not found (server restart / job loss)
        if (res.status === 404) {
          consecutive404sRef.current++;
          consecutiveFailuresRef.current = 0; // Reset general failure counter
          if (consecutive404sRef.current >= 3) {
            stopPolling();
            setState("error");
            setIsProcessing(false);
            setStatusMsg("Processing was interrupted. Your project may have partial results — check the project list.");
            setPollWarning(null);
            queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
          }
          return;
        }

        if (!res.ok) {
          // Non-404 error — count as general poll failure
          consecutiveFailuresRef.current++;
          consecutive404sRef.current = 0;
          if (consecutiveFailuresRef.current >= 15) {
            setPollWarning("Lost connection to server. Your upload may still be processing — refresh the page to check.");
          } else if (consecutiveFailuresRef.current >= 5) {
            setPollWarning("Connection interrupted. Checking status...");
          }
          return;
        }

        // Success — reset all failure counters
        consecutiveFailuresRef.current = 0;
        consecutive404sRef.current = 0;
        setPollWarning(null);

        const prog = await res.json();

        // Track warnings from backend
        if (prog.warnings && Array.isArray(prog.warnings)) {
          setWarnings(prog.warnings);
        }
        // Track PDF quality
        if (prog.pdfQuality) {
          setPdfQuality(prog.pdfQuality);
        }

        if (prog.status === "error") {
          stopPolling();
          setState("error");
          setIsProcessing(false);
          setStatusMsg(prog.error || "Processing failed.");
          return;
        }

        if (prog.status === "done") {
          stopPolling();
          setProgress(100);
          setStatusMsg(`Complete — ${prog.itemsFound} items found`);
          setState("done");
          setIsProcessing(false);
          queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
          if (prog.projectId) onProjectCreated(prog.projectId);
          return;
        }

        // Update progress display — two-phase aware
        if (prog.phase === "rendering") {
          // Phase 1: rendering uses 10-50% of progress bar
          const renderPct = prog.totalPages > 0 ? Math.round((prog.pagesProcessed / prog.totalPages) * 40) + 10 : 15;
          setProgress(Math.min(renderPct, 50));
          setStatusMsg(`Preparing pages for extraction... ${prog.pagesProcessed} of ${prog.totalPages}`);
        } else if (prog.phase === "extracting") {
          // Phase 2: extracting uses 50-95% of progress bar
          const extractPct = prog.totalPages > 0 ? Math.round((prog.pagesProcessed / prog.totalPages) * 45) + 50 : 55;
          setProgress(Math.min(extractPct, 95));
          setStatusMsg(`AI extraction in progress... ${prog.pagesProcessed} of ${prog.totalPages} pages — ${prog.itemsFound} items found`);
        } else if (prog.status === "uploading" || prog.phase === "uploading") {
          setProgress(8);
          setStatusMsg("Analyzing PDF...");
        } else if (prog.status === "processing" && prog.totalChunks > 0) {
          // Fallback for legacy progress without phase field
          const pct = prog.totalPages > 0 ? Math.round((prog.pagesProcessed / prog.totalPages) * 85) + 10 : 20;
          setProgress(Math.min(pct, 90));
          setStatusMsg(`Processing chunk ${prog.chunk}/${prog.totalChunks} — ${prog.pagesProcessed}/${prog.totalPages} pages — ${prog.itemsFound} items found...`);
        } else {
          setStatusMsg(`Status: ${prog.status}...`);
        }
      } catch {
        // Network error — count as poll failure
        consecutiveFailuresRef.current++;
        consecutive404sRef.current = 0;
        if (consecutiveFailuresRef.current >= 15) {
          setPollWarning("Lost connection to server. Your upload may still be processing — refresh the page to check.");
        } else if (consecutiveFailuresRef.current >= 5) {
          setPollWarning("Connection interrupted. Checking status...");
        }
      }
    }, 2000);
  }, [API_BASE, stopPolling, onProjectCreated]);

  const uploadFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".pdf")) {
      toast({ title: "Invalid file", description: "Please upload a PDF file.", variant: "destructive" });
      return;
    }
    if (isProcessing) {
      toast({ title: "Upload in progress", description: "Please wait for the current upload to finish.", variant: "destructive" });
      return;
    }
    setIsProcessing(true);
    setState("uploading");
    setProgress(5);
    setStatusMsg("Uploading PDF...");
    setWarnings([]);
    setPdfQuality(null);
    setPollWarning(null);

    const formData = new FormData();
    formData.append("file", file);
    formData.append("discipline", discipline);
    formData.append("hasRevisions", String(hasRevisions));
    if (dualModel && geminiAvailable) formData.append("dualModel", "true");

    try {
      const uploadHeaders: Record<string, string> = {};
      const tk = getAuthToken();
      if (tk) uploadHeaders["Authorization"] = `Bearer ${tk}`;
      const res = await fetch(`${API_BASE}/api/takeoff/upload`, { method: "POST", body: formData, headers: uploadHeaders });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Upload failed");
      }
      const { jobId, pageCount, totalChunks } = await res.json();
      setProgress(10);
      setState("processing");
      setStatusMsg(`PDF received (${pageCount} pages, ${totalChunks} chunk${totalChunks > 1 ? "s" : ""}) — running AI vision...`);
      pollProgress(jobId);
    } catch (err: any) {
      setState("error");
      setIsProcessing(false);
      setStatusMsg(err.message || "Upload failed");
      toast({ title: "Upload failed", description: err.message, variant: "destructive" });
    }
  }, [API_BASE, discipline, hasRevisions, dualModel, geminiAvailable, isProcessing, pollProgress, toast]);

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    if (isProcessing) return;
    const file = e.dataTransfer.files[0];
    if (file) uploadFile(file);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) uploadFile(file);
    e.target.value = "";
  };

  const reset = () => {
    stopPolling();
    setState("idle");
    setIsProcessing(false);
    setProgress(0);
    setStatusMsg("");
    setWarnings([]);
    setPdfQuality(null);
    setPollWarning(null);
    consecutiveFailuresRef.current = 0;
    consecutive404sRef.current = 0;
  };

  const disciplineLabels: Record<string, string> = {
    mechanical: "piping isometric",
    structural: "structural drawing",
    civil: "civil/site plan",
  };

  const showRevisionSelector = discipline === "mechanical" && state === "idle" && !isProcessing;

  return (
    <div className="space-y-3">
      {/* Revision toggle — card selector (mechanical only, when idle) */}
      {showRevisionSelector && (
        <div className="grid grid-cols-2 gap-3" data-testid="revision-selector">
          <button
            type="button"
            onClick={() => setHasRevisions(false)}
            className={`relative rounded-lg border-2 p-4 text-left transition-all cursor-pointer
              ${!hasRevisions
                ? "border-primary bg-primary/5 ring-1 ring-primary/20"
                : "border-border hover:border-muted-foreground/40"
              }`}
            data-testid="btn-original-drawings"
          >
            <div className="flex items-start gap-3">
              <div className={`mt-0.5 w-8 h-8 rounded-md flex items-center justify-center shrink-0
                ${!hasRevisions ? "bg-primary/15 text-primary" : "bg-muted text-muted-foreground"}`}>
                <FileStack size={16} />
              </div>
              <div className="min-w-0">
                <p className={`text-sm font-semibold ${!hasRevisions ? "text-primary" : "text-foreground"}`}>
                  Original Drawings
                </p>
                <p className="text-xs text-muted-foreground mt-1 leading-relaxed">
                  Fast processing — recommended for new/original drawing sets
                </p>
              </div>
            </div>
            {!hasRevisions && (
              <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-primary" />
            )}
          </button>

          <button
            type="button"
            onClick={() => setHasRevisions(true)}
            className={`relative rounded-lg border-2 p-4 text-left transition-all cursor-pointer
              ${hasRevisions
                ? "border-primary bg-primary/5 ring-1 ring-primary/20"
                : "border-border hover:border-muted-foreground/40"
              }`}
            data-testid="btn-revised-drawings"
          >
            <div className="flex items-start gap-3">
              <div className={`mt-0.5 w-8 h-8 rounded-md flex items-center justify-center shrink-0
                ${hasRevisions ? "bg-primary/15 text-primary" : "bg-muted text-muted-foreground"}`}>
                <CloudLightning size={16} />
              </div>
              <div className="min-w-0">
                <p className={`text-sm font-semibold ${hasRevisions ? "text-primary" : "text-foreground"}`}>
                  Revised Drawings (with clouds)
                </p>
                <p className="text-xs text-muted-foreground mt-1 leading-relaxed">
                  Detects revision clouds and tags changed items — use for revised/addenda drawings
                </p>
              </div>
            </div>
            {hasRevisions && (
              <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-primary" />
            )}
          </button>
        </div>
      )}

      {/* Dual-Model toggle (when Gemini is configured and idle) */}
      {geminiAvailable && state === "idle" && !isProcessing && (
        <div className="flex items-center gap-3 px-3 py-2 rounded-lg border border-border bg-muted/20">
          <Layers size={16} className={dualModel ? "text-primary" : "text-muted-foreground"} />
          <div className="flex-1 min-w-0">
            <p className="text-xs font-medium">Dual-Model Verification</p>
            <p className="text-[10px] text-muted-foreground">Run Claude + Gemini and cross-check results for higher confidence</p>
          </div>
          <button
            type="button"
            role="switch"
            aria-checked={dualModel}
            aria-label="Dual-Model Verification"
            onClick={() => setDualModel(!dualModel)}
            className={`relative w-9 h-5 rounded-full transition-colors ${dualModel ? "bg-primary" : "bg-muted-foreground/30"}`}
          >
            <span className={`absolute top-0.5 left-0.5 w-4 h-4 rounded-full bg-white transition-transform ${dualModel ? "translate-x-4" : "translate-x-0"}`} />
          </button>
        </div>
      )}

      {/* Main upload zone */}
      <div
        className={`relative rounded-xl border-2 border-dashed p-8 transition-all duration-200
          ${isProcessing && state !== "done" && state !== "error"
            ? "cursor-not-allowed opacity-80 border-muted bg-muted/10"
            : isDragOver ? "border-primary bg-primary/5 shadow-lg shadow-primary/10 scale-[1.01] cursor-pointer" : "border-border hover:border-primary/60 hover:bg-accent/30 hover:shadow-md cursor-pointer"
          }
          ${state !== "idle" ? "cursor-default" : ""}`}
        onDrop={handleDrop}
        onDragOver={e => { e.preventDefault(); if (!isProcessing) setIsDragOver(true); }}
        onDragLeave={() => setIsDragOver(false)}
        onClick={() => state === "idle" && !isProcessing && fileInputRef.current?.click()}
        data-testid="upload-zone"
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".pdf"
          className="hidden"
          onChange={handleFileChange}
          disabled={isProcessing}
          data-testid="input-file"
        />

        {state === "idle" && !isProcessing && (
          <div className="flex flex-col items-center gap-3 text-center">
            <div className="w-14 h-14 rounded-full bg-primary/10 flex items-center justify-center">
              <Upload size={24} className="text-primary" />
            </div>
            <div>
              <p className="font-medium text-sm">Drop {disciplineLabels[discipline]} PDF here</p>
              <p className="text-xs text-muted-foreground mt-1">or click to browse — up to 100MB</p>
            </div>
            <Button variant="outline" size="sm" className="mt-1" data-testid="btn-browse">
              Browse PDF
            </Button>
          </div>
        )}

        {(state === "uploading" || state === "processing") && (
          <div className="flex flex-col items-center gap-4 text-center">
            <div className="w-14 h-14 rounded-full bg-primary/10 flex items-center justify-center animate-pulse">
              <FileText size={24} className="text-primary" />
            </div>
            <div className="w-full max-w-sm">
              <Progress value={progress} className="h-2.5 progress-gradient" />
              <p className="text-xs text-muted-foreground mt-2">{statusMsg}</p>
              {/* Poll connection warning */}
              {pollWarning && (
                <div className="mt-3 flex items-start gap-2 text-left rounded-md bg-orange-500/10 border border-orange-500/30 p-2">
                  <WifiOff size={14} className="text-orange-600 dark:text-orange-400 mt-0.5 shrink-0" />
                  <p className="text-xs text-orange-700 dark:text-orange-300">{pollWarning}</p>
                </div>
              )}
              {/* Warnings from backend */}
              {warnings.length > 0 && (
                <div className="mt-3 flex items-start gap-2 text-left rounded-md bg-yellow-500/10 border border-yellow-500/30 p-2">
                  <AlertTriangle size={14} className="text-yellow-600 dark:text-yellow-400 mt-0.5 shrink-0" />
                  <div className="text-xs text-yellow-700 dark:text-yellow-300">
                    {warnings.map((w, i) => (
                      <p key={i}>{w}</p>
                    ))}
                  </div>
                </div>
              )}
              {/* Poor scan quality warning */}
              {pdfQuality === "poor_scan" && (
                <div className="mt-3 flex items-start gap-2 text-left rounded-md bg-amber-500/10 border border-amber-500/30 p-2">
                  <AlertTriangle size={14} className="text-amber-600 dark:text-amber-400 mt-0.5 shrink-0" />
                  <p className="text-xs text-amber-700 dark:text-amber-300">
                    Poor scan quality detected — results may have lower accuracy. Consider re-scanning with higher DPI for better extraction.
                  </p>
                </div>
              )}
            </div>
          </div>
        )}

        {state === "done" && (
          <div className="flex flex-col items-center gap-3 text-center animate-[fade-in_0.4s_ease-out]">
            <div className="w-16 h-16 rounded-full bg-gradient-to-br from-green-400 to-emerald-500 flex items-center justify-center shadow-lg shadow-green-500/25">
              <CheckCircle2 size={28} className="text-white" />
            </div>
            <p className="text-sm font-semibold text-green-600 dark:text-green-400">{statusMsg}</p>
            {warnings.length > 0 && (
              <div className="flex items-start gap-2 text-left rounded-md bg-yellow-500/10 border border-yellow-500/30 p-2 max-w-sm">
                <AlertTriangle size={14} className="text-yellow-600 dark:text-yellow-400 mt-0.5 shrink-0" />
                <div className="text-xs text-yellow-700 dark:text-yellow-300">
                  {warnings.map((w, i) => (
                    <p key={i}>{w}</p>
                  ))}
                </div>
              </div>
            )}
            {pdfQuality === "poor_scan" && (
              <div className="flex items-start gap-2 text-left rounded-md bg-amber-500/10 border border-amber-500/30 p-2 max-w-sm">
                <AlertTriangle size={14} className="text-amber-600 dark:text-amber-400 mt-0.5 shrink-0" />
                <p className="text-xs text-amber-700 dark:text-amber-300">
                  Poor scan quality detected — results may have lower accuracy. Consider re-scanning with higher DPI.
                </p>
              </div>
            )}
            <Button variant="outline" size="sm" onClick={reset} data-testid="btn-upload-another">
              Upload Another
            </Button>
          </div>
        )}

        {state === "error" && (
          <div className="flex flex-col items-center gap-3 text-center">
            <div className="w-14 h-14 rounded-full bg-destructive/10 flex items-center justify-center">
              <AlertCircle size={24} className="text-destructive" />
            </div>
            <p className="text-sm font-medium text-destructive">{statusMsg}</p>
            {warnings.length > 0 && (
              <div className="flex items-start gap-2 text-left rounded-md bg-yellow-500/10 border border-yellow-500/30 p-2 max-w-sm">
                <AlertTriangle size={14} className="text-yellow-600 dark:text-yellow-400 mt-0.5 shrink-0" />
                <div className="text-xs text-yellow-700 dark:text-yellow-300">
                  {warnings.map((w, i) => (
                    <p key={i}>{w}</p>
                  ))}
                </div>
              </div>
            )}
            <Button variant="outline" size="sm" onClick={reset} data-testid="btn-retry-upload">
              Try Again
            </Button>
          </div>
        )}
      </div>
    </div>
  );
}
