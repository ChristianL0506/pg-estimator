import { useState, useEffect } from "react";
import AppLayout from "@/components/AppLayout";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useToast } from "@/hooks/use-toast";

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

export default function SettingsPage() {
  const { toast } = useToast();
  const token = (window as any).__PG_AUTH_TOKEN__;

  // Anthropic state
  const [apiKey, setApiKey] = useState("");
  const [status, setStatus] = useState<{configured: boolean; masked: string | null; source: string | null} | null>(null);
  const [testing, setTesting] = useState(false);
  const [saving, setSaving] = useState(false);
  const [testResult, setTestResult] = useState<{success: boolean; message: string} | null>(null);

  // Gemini state
  const [geminiKey, setGeminiKey] = useState("");
  const [geminiStatus, setGeminiStatus] = useState<{configured: boolean; masked: string | null} | null>(null);
  const [geminiTesting, setGeminiTesting] = useState(false);
  const [geminiSaving, setGeminiSaving] = useState(false);
  const [geminiTestResult, setGeminiTestResult] = useState<{success: boolean; message: string} | null>(null);

  useEffect(() => {
    const headers = { Authorization: `Bearer ${token}` };
    fetch(`${API_BASE}/api/settings/api-key`, { headers }).then(r => r.json()).then(setStatus).catch(() => {});
    fetch(`${API_BASE}/api/settings/gemini-key`, { headers }).then(r => r.json()).then(setGeminiStatus).catch(() => {});
  }, []);

  // --- Anthropic handlers ---
  const handleTest = async () => {
    setTesting(true);
    setTestResult(null);
    try {
      const res = await fetch(`${API_BASE}/api/settings/test-api-key`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ apiKey: apiKey || undefined }),
      });
      const data = await res.json();
      setTestResult(data.success
        ? { success: true, message: "API key is working — connection to Anthropic confirmed." }
        : { success: false, message: data.error || "Test failed" });
    } catch (err: any) {
      setTestResult({ success: false, message: err.message || "Connection error" });
    }
    setTesting(false);
  };

  const handleSave = async () => {
    if (!apiKey.startsWith("sk-ant-")) {
      toast({ title: "Invalid key", description: "Anthropic API keys start with 'sk-ant-'", variant: "destructive" });
      return;
    }
    setSaving(true);
    try {
      const res = await fetch(`${API_BASE}/api/settings/api-key`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ apiKey }),
      });
      const data = await res.json();
      if (data.success) {
        setStatus({ configured: true, masked: data.masked, source: "user" });
        setApiKey("");
        toast({ title: "API key saved", description: "Your Anthropic API key is now active." });
      } else {
        toast({ title: "Error", description: data.error, variant: "destructive" });
      }
    } catch {
      toast({ title: "Error", description: "Failed to save", variant: "destructive" });
    }
    setSaving(false);
  };

  const handleRemove = async () => {
    try {
      await fetch(`${API_BASE}/api/settings/api-key`, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
      setStatus({ configured: false, masked: null, source: null });
      toast({ title: "API key removed", description: "Falling back to platform-provided key." });
    } catch {}
  };

  // --- Gemini handlers ---
  const handleGeminiTest = async () => {
    setGeminiTesting(true);
    setGeminiTestResult(null);
    try {
      const res = await fetch(`${API_BASE}/api/settings/test-gemini-key`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ apiKey: geminiKey || undefined }),
      });
      const data = await res.json();
      setGeminiTestResult(data.success
        ? { success: true, message: "Gemini key is working — connection to Google AI confirmed." }
        : { success: false, message: data.error || "Test failed" });
    } catch (err: any) {
      setGeminiTestResult({ success: false, message: err.message || "Connection error" });
    }
    setGeminiTesting(false);
  };

  const handleGeminiSave = async () => {
    if (!geminiKey || geminiKey.length < 10) {
      toast({ title: "Invalid key", description: "Please enter a valid Gemini API key", variant: "destructive" });
      return;
    }
    setGeminiSaving(true);
    try {
      const res = await fetch(`${API_BASE}/api/settings/gemini-key`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ apiKey: geminiKey }),
      });
      const data = await res.json();
      if (data.success) {
        setGeminiStatus({ configured: true, masked: data.masked });
        setGeminiKey("");
        toast({ title: "Gemini key saved", description: "Dual-model extraction is now available." });
      } else {
        toast({ title: "Error", description: data.error, variant: "destructive" });
      }
    } catch {
      toast({ title: "Error", description: "Failed to save", variant: "destructive" });
    }
    setGeminiSaving(false);
  };

  const handleGeminiRemove = async () => {
    try {
      await fetch(`${API_BASE}/api/settings/gemini-key`, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
      setGeminiStatus({ configured: false, masked: null });
      toast({ title: "Gemini key removed" });
    } catch {}
  };

  return (
    <AppLayout subtitle="Settings">
      <div className="max-w-2xl mx-auto space-y-6 p-5">
        <div>
          <h2 className="text-lg font-semibold text-foreground">AI API Configuration</h2>
          <p className="text-sm text-muted-foreground mt-1">
            Configure your API keys for PDF processing and dual-model verification.
          </p>
        </div>

        {/* ===== ANTHROPIC SECTION ===== */}
        <div className="rounded-lg border border-border bg-card p-4 space-y-3">
          <h3 className="font-medium text-sm">Anthropic API Key (Claude)</h3>
          <p className="text-xs text-muted-foreground">Primary extraction engine. Powers all AI-powered PDF extraction and the chat assistant.</p>

          {/* Status indicator */}
          <div className={`rounded-md border p-3 ${
            status?.source === "user" ? "border-green-200 bg-green-50 dark:border-green-800 dark:bg-green-950/20" :
            status?.configured ? "border-yellow-200 bg-yellow-50 dark:border-yellow-800 dark:bg-yellow-950/20" :
            "border-red-200 bg-red-50 dark:border-red-800 dark:bg-red-950/20"
          }`}>
            <div className="flex items-center gap-2">
              <div className={`w-2.5 h-2.5 rounded-full ${
                status?.source === "user" ? "bg-green-500" :
                status?.configured ? "bg-yellow-500" : "bg-red-500"
              }`} />
              <span className="font-medium text-xs">
                {status?.source === "user" ? "Your API Key Active" :
                 status?.configured ? "Platform Key (may timeout)" :
                 "Not Configured"}
              </span>
              {status?.masked && <span className="text-[10px] text-muted-foreground ml-auto">{status.masked}</span>}
            </div>
          </div>

          <div className="flex gap-2">
            <Input
              id="anthropic-key"
              type="password"
              placeholder="sk-ant-api03-..."
              value={apiKey}
              onChange={e => setApiKey(e.target.value)}
              className="font-mono text-xs"
              aria-label="Anthropic API Key"
            />
            <Button variant="outline" size="sm" onClick={handleTest} disabled={testing || (!apiKey && !status?.configured)}>
              {testing ? "Testing..." : "Test"}
            </Button>
            <Button size="sm" onClick={handleSave} disabled={saving || !apiKey}>
              {saving ? "Saving..." : "Save"}
            </Button>
          </div>

          {testResult && (
            <div className={`text-xs p-2 rounded ${
              testResult.success ? "bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-300" :
              "bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-300"
            }`}>
              {testResult.success ? "✓ " : "✗ "}{testResult.message}
            </div>
          )}

          {status?.source === "user" && (
            <Button variant="ghost" size="sm" onClick={handleRemove} className="text-xs text-muted-foreground">
              Remove saved key
            </Button>
          )}
        </div>

        {/* ===== GEMINI SECTION ===== */}
        <div className="rounded-lg border border-border bg-card p-4 space-y-3">
          <h3 className="font-medium text-sm">Google Gemini API Key</h3>
          <p className="text-xs text-muted-foreground">
            Optional — used as a fallback if Claude is not configured, and for dual-model verification
            when enabled on a takeoff.
          </p>

          {/* Status indicator */}
          <div className={`rounded-md border p-3 ${
            geminiStatus?.configured ? "border-green-200 bg-green-50 dark:border-green-800 dark:bg-green-950/20" :
            "border-gray-200 bg-gray-50 dark:border-gray-700 dark:bg-gray-900/20"
          }`}>
            <div className="flex items-center gap-2">
              <div className={`w-2.5 h-2.5 rounded-full ${geminiStatus?.configured ? "bg-green-500" : "bg-gray-400"}`} />
              <span className="font-medium text-xs">
                {geminiStatus?.configured ? "Gemini Key Active" : "Not Configured (optional)"}
              </span>
              {geminiStatus?.masked && <span className="text-[10px] text-muted-foreground ml-auto">{geminiStatus.masked}</span>}
            </div>
          </div>

          <div className="flex gap-2">
            <Input
              id="gemini-key"
              type="password"
              placeholder="AIzaSy..."
              value={geminiKey}
              onChange={e => setGeminiKey(e.target.value)}
              className="font-mono text-xs"
              aria-label="Google Gemini API Key"
            />
            <Button variant="outline" size="sm" onClick={handleGeminiTest} disabled={geminiTesting || (!geminiKey && !geminiStatus?.configured)}>
              {geminiTesting ? "Testing..." : "Test"}
            </Button>
            <Button size="sm" onClick={handleGeminiSave} disabled={geminiSaving || !geminiKey}>
              {geminiSaving ? "Saving..." : "Save"}
            </Button>
          </div>

          {geminiTestResult && (
            <div className={`text-xs p-2 rounded ${
              geminiTestResult.success ? "bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-300" :
              "bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-300"
            }`}>
              {geminiTestResult.success ? "✓ " : "✗ "}{geminiTestResult.message}
            </div>
          )}

          {geminiStatus?.configured && (
            <Button variant="ghost" size="sm" onClick={handleGeminiRemove} className="text-xs text-muted-foreground">
              Remove Gemini key
            </Button>
          )}

          <div className="text-[10px] text-muted-foreground space-y-1 pt-1 border-t border-border">
            <p>1. Go to <span className="font-mono bg-muted px-1 rounded">aistudio.google.com</span></p>
            <p>2. Click "Get API Key" and create a key</p>
            <p>3. Paste the key above</p>
          </div>
        </div>
      </div>
    </AppLayout>
  );
}
