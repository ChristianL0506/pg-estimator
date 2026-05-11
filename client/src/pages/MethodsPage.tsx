// Methods page — full factor tree viewer + editor.
//
// Lists every estimating method on the left (base methods Bill / Justin /
// Industry, plus any saved custom profiles). The right pane shows every
// numeric factor that drives the estimator, grouped by section (Pipe, Welds,
// Valves, Bolts, Threads, Other, Cost Params, etc.) with the value rendered
// in a sortable table.
//
// Base methods (bill / justin / industry) are read-only — the lock badge
// makes that clear. To edit, the user clicks "Save As Custom" which clones
// the current method (base + any in-flight UI edits the user already made
// on a base — though we only allow edits on customs to avoid confusion)
// into a new CustomEstimatorMethod via POST /api/custom-methods.
//
// Custom methods are fully editable: every leaf cell is an inline input,
// modified cells are flagged with a colored dot + "Revert" link. There's
// also a per-section bulk multiplier and a copy-from selector to pull a
// section's values from another method wholesale. Save persists the
// override map via PATCH /api/custom-methods/:id.
//
// Each method has an "Export to Excel" button that hits the server export
// endpoint (/api/methods/:key/export) and downloads the workbook.

import { useState, useMemo, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { Plus, Lock, Save, Trash2, Download, RotateCcw, Copy, Calculator, AlertCircle, CheckCircle2, Pencil, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { useToast } from "@/hooks/use-toast";
import { apiRequest } from "@/lib/queryClient";
import AppLayout from "@/components/AppLayout";
import type { CustomEstimatorMethod } from "@shared/schema";

// ----------------------- Path encoding -----------------------
// Override paths use '.' as the segment delimiter; we backslash-escape any
// literal '.' inside a segment so e.g. Bill's "0.25" size becomes "0\.25".
// Keeps the server's splitOverridePath in lockstep.
function encodeSegment(seg: string): string {
  return seg.replace(/\\/g, "\\\\").replace(/\./g, "\\.");
}
function joinPath(segs: string[]): string {
  return segs.map(encodeSegment).join(".");
}

// ----------------------- Section layout config -----------------------
// Each base-method block has a different shape. We declare per-method which
// top-level subtrees to render as a section, and the human label/icon for it.
// Anything not listed falls through to a generic "Other" renderer.
type SectionDef = { key: string; label: string; pathFromRoot: string[] };

const SECTIONS_BY_METHOD: Record<string, SectionDef[]> = {
  bill: [
    { key: "butt_welds_ei",        label: "Butt Welds (EI)",          pathFromRoot: ["labor_rates", "butt_welds_ei"] },
    { key: "flanged_joints",       label: "Flanged Joints (MH/joint)", pathFromRoot: ["labor_rates", "flanged_joints_mh_per_joint"] },
    { key: "manhours_per_eq_inch", label: "Manhours per Eq. Inch",     pathFromRoot: ["labor_rates", "manhours_per_eq_inch"] },
    { key: "pipe_handling",        label: "Pipe Handling (MH/LF)",     pathFromRoot: ["labor_rates", "pipe_handling_mh_per_lf"] },
    { key: "material_factors",     label: "Material Factors",          pathFromRoot: ["material_factors"] },
    { key: "material_factor_groups", label: "Material Factor Groups",  pathFromRoot: ["material_factor_groups"] },
  ],
  justin: [
    { key: "pipe",        label: "Pipe (MH/LF)",        pathFromRoot: ["labor_factors", "pipe"] },
    { key: "welds",       label: "Welds (MH/weld)",     pathFromRoot: ["labor_factors", "welds"] },
    { key: "valves",      label: "Valves (MH/valve)",   pathFromRoot: ["labor_factors", "valves"] },
    { key: "bolts",       label: "Bolts (MH/set)",      pathFromRoot: ["labor_factors", "bolts"] },
    { key: "threads",     label: "Threads (MH/joint)",  pathFromRoot: ["labor_factors", "threads"] },
    { key: "other",       label: "Other Factors",       pathFromRoot: ["labor_factors", "other"] },
    { key: "cost_params", label: "Cost Parameters",     pathFromRoot: ["cost_params"] },
  ],
  industry: [
    { key: "pipe",        label: "Pipe (MH/LF)",        pathFromRoot: ["labor_factors", "pipe"] },
    { key: "welds",       label: "Welds (MH/weld)",     pathFromRoot: ["labor_factors", "welds"] },
    { key: "valves",      label: "Valves (MH/valve)",   pathFromRoot: ["labor_factors", "valves"] },
    { key: "bolts",       label: "Bolts (MH/set)",      pathFromRoot: ["labor_factors", "bolts"] },
    { key: "threads",     label: "Threads (MH/joint)",  pathFromRoot: ["labor_factors", "threads"] },
    { key: "other",       label: "Other Factors",       pathFromRoot: ["labor_factors", "other"] },
    { key: "cost_params", label: "Cost Parameters",     pathFromRoot: ["cost_params"] },
  ],
};

// ----------------------- Data fetching -----------------------
type EstimatorMethodsResponse = {
  methods: { key: "bill" | "justin" | "industry"; name: string; description: string; source: string }[];
  data: { bill: any; justin: any; industry: any };
};

// Walks a subtree and lists every leaf as a {pathSegs, value} pair.
// Only numeric/string leaves become editable cells; nested non-leaf children
// are descended into.
type Leaf = { pathSegs: string[]; value: any };
function collectLeaves(node: any, prefix: string[] = []): Leaf[] {
  if (node === null || node === undefined) return [];
  if (typeof node !== "object") return [{ pathSegs: prefix, value: node }];
  const out: Leaf[] = [];
  for (const k of Object.keys(node)) {
    out.push(...collectLeaves(node[k], [...prefix, k]));
  }
  return out;
}

// Reads a value from a deep-cloned object by path segments (no escaping needed).
function readBySegs(obj: any, segs: string[]): any {
  let cur = obj;
  for (const s of segs) {
    if (cur === null || cur === undefined) return undefined;
    cur = cur[s];
  }
  return cur;
}

// Groups leaves under a section by their first sub-segment so we can render
// a row per "row key" (e.g. each pipe size) and a column per "col key"
// (e.g. standard_mh_per_lf vs rack_mh_per_lf).
function groupRowsCols(leaves: Leaf[]): { rowKey: string; cells: { colKey: string; segs: string[]; value: any }[] }[] {
  // The leaves all share the same prefix length depth=2 (row, col) in most
  // sections; if depth=1 (e.g. cost_params, material_factors) treat each
  // leaf as its own row with a single col.
  const rows = new Map<string, { rowKey: string; cells: { colKey: string; segs: string[]; value: any }[] }>();
  for (const lf of leaves) {
    const rowKey = lf.pathSegs[0] ?? "—";
    const colKey = lf.pathSegs[1] ?? "value";
    if (!rows.has(rowKey)) rows.set(rowKey, { rowKey, cells: [] });
    rows.get(rowKey)!.cells.push({ colKey, segs: lf.pathSegs, value: lf.value });
  }
  return Array.from(rows.values());
}

// Stable sort: numeric-prefix size keys ascending, then alphabetical.
function sortRowKey(a: string, b: string): number {
  const an = parseFloat(a);
  const bn = parseFloat(b);
  const aNum = !isNaN(an);
  const bNum = !isNaN(bn);
  if (aNum && bNum) return an - bn;
  if (aNum) return -1;
  if (bNum) return 1;
  return a.localeCompare(b);
}

// All distinct column keys across all rows, in first-seen order.
function uniqueColKeys(rows: ReturnType<typeof groupRowsCols>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const r of rows) {
    for (const c of r.cells) {
      if (!seen.has(c.colKey)) {
        seen.add(c.colKey);
        out.push(c.colKey);
      }
    }
  }
  return out;
}

// ----------------------- Component -----------------------

type MethodRef =
  | { kind: "base"; key: "bill" | "justin" | "industry"; name: string; description: string }
  | { kind: "custom"; id: string; name: string; baseMethod: "bill" | "justin" | "industry"; description: string };

export default function MethodsPage() {
  const { toast } = useToast();
  const qc = useQueryClient();
  const [selected, setSelected] = useState<MethodRef | null>(null);
  const [pendingOverrides, setPendingOverrides] = useState<Record<string, any>>({});
  const [bulkMultiplier, setBulkMultiplier] = useState<Record<string, string>>({});
  const [showSaveAs, setShowSaveAs] = useState(false);
  const [saveAsName, setSaveAsName] = useState("");
  const [showCopyFrom, setShowCopyFrom] = useState<string | null>(null);
  const [copyFromKey, setCopyFromKey] = useState<string>("");
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [renameValue, setRenameValue] = useState<string>("");
  const [renaming, setRenaming] = useState(false);

  const baseMethodsQ = useQuery<EstimatorMethodsResponse>({ queryKey: ["/api/estimator-methods"] });
  const customMethodsQ = useQuery<CustomEstimatorMethod[]>({ queryKey: ["/api/custom-methods"] });

  // Auto-select the first method when data first arrives.
  useEffect(() => {
    if (!selected && baseMethodsQ.data) {
      const m = baseMethodsQ.data.methods[0];
      if (m) setSelected({ kind: "base", key: m.key, name: m.name, description: m.description });
    }
  }, [baseMethodsQ.data, selected]);

  // Effective data for the selected method = base data + saved overrides + pending UI overrides.
  // We compute it client-side identically to the server's applyCustomOverrides so the user
  // sees exactly what the calculator will use.
  const effectiveData = useMemo(() => {
    if (!selected || !baseMethodsQ.data) return null;
    const baseKey = selected.kind === "base" ? selected.key : selected.baseMethod;
    const base = baseMethodsQ.data.data[baseKey];
    if (!base) return null;
    const cloned = JSON.parse(JSON.stringify(base));
    let saved: Record<string, any> = {};
    if (selected.kind === "custom") {
      const cm = customMethodsQ.data?.find(c => c.id === selected.id);
      if (cm) saved = cm.overrides || {};
    }
    const combined: Record<string, any> = { ...saved, ...pendingOverrides };
    for (const [keyPath, value] of Object.entries(combined)) {
      const parts: string[] = [];
      let buf = "";
      for (let i = 0; i < keyPath.length; i++) {
        const c = keyPath[i];
        if (c === "\\" && i + 1 < keyPath.length && keyPath[i + 1] === ".") {
          buf += ".";
          i++;
        } else if (c === ".") {
          parts.push(buf);
          buf = "";
        } else {
          buf += c;
        }
      }
      parts.push(buf);
      let cursor: any = cloned;
      for (let i = 0; i < parts.length - 1; i++) {
        if (cursor[parts[i]] === undefined || cursor[parts[i]] === null || typeof cursor[parts[i]] !== "object") {
          cursor[parts[i]] = {};
        }
        cursor = cursor[parts[i]];
      }
      cursor[parts[parts.length - 1]] = value;
    }
    return cloned;
  }, [selected, baseMethodsQ.data, customMethodsQ.data, pendingOverrides]);

  // Reset pending edits whenever the selection changes.
  useEffect(() => {
    setPendingOverrides({});
    setBulkMultiplier({});
  }, [selected?.kind, selected?.kind === "base" ? selected.key : selected?.kind === "custom" ? selected.id : null]);

  const baseKey = selected
    ? selected.kind === "base"
      ? selected.key
      : selected.baseMethod
    : null;
  const sectionDefs = baseKey ? SECTIONS_BY_METHOD[baseKey] || [] : [];
  const isEditable = selected?.kind === "custom";

  // ----------------------- Mutations -----------------------
  const saveAsMutation = useMutation({
    mutationFn: async (payload: { name: string; baseMethod: string; overrides: Record<string, any> }) => {
      const res = await apiRequest("POST", "/api/custom-methods", payload);
      return res.json();
    },
    onSuccess: (created: CustomEstimatorMethod) => {
      qc.invalidateQueries({ queryKey: ["/api/custom-methods"] });
      setShowSaveAs(false);
      setSaveAsName("");
      setPendingOverrides({});
      setSelected({ kind: "custom", id: created.id, name: created.name, baseMethod: created.baseMethod, description: created.description || "" });
      toast({ title: `Saved '${created.name}' as a new custom method` });
    },
    onError: (err: any) => {
      toast({ title: "Save failed", description: err?.message || String(err), variant: "destructive" });
    },
  });

  const patchMutation = useMutation({
    mutationFn: async (payload: { id: string; data: Partial<{ name: string; description: string; overrides: Record<string, any> }> }) => {
      const res = await apiRequest("PATCH", `/api/custom-methods/${payload.id}`, payload.data);
      return res.json();
    },
    onSuccess: (updated: CustomEstimatorMethod) => {
      qc.invalidateQueries({ queryKey: ["/api/custom-methods"] });
      setPendingOverrides({});
      setRenaming(false);
      // Keep our selection ref in sync with the renamed method.
      if (selected?.kind === "custom" && selected.id === updated.id) {
        setSelected({ kind: "custom", id: updated.id, name: updated.name, baseMethod: updated.baseMethod, description: updated.description || "" });
      }
      toast({ title: "Saved changes" });
    },
    onError: (err: any) => {
      toast({ title: "Save failed", description: err?.message || String(err), variant: "destructive" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: async (id: string) => apiRequest("DELETE", `/api/custom-methods/${id}`),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["/api/custom-methods"] });
      setShowDeleteConfirm(false);
      // Fall back to the first base method when the selection disappears.
      const first = baseMethodsQ.data?.methods[0];
      if (first) setSelected({ kind: "base", key: first.key, name: first.name, description: first.description });
      toast({ title: "Custom method deleted" });
    },
    onError: (err: any) => {
      toast({ title: "Delete failed", description: err?.message || String(err), variant: "destructive" });
    },
  });

  // ----------------------- Cell editing -----------------------
  function commitCell(segs: string[], rawValue: string, originalType: "number" | "string") {
    const keyPath = joinPath(segs);
    if (originalType === "number") {
      const n = parseFloat(rawValue);
      if (isNaN(n)) {
        toast({ title: "Not a number", description: `Value '${rawValue}' is not numeric`, variant: "destructive" });
        return;
      }
      setPendingOverrides(prev => ({ ...prev, [keyPath]: n }));
    } else {
      setPendingOverrides(prev => ({ ...prev, [keyPath]: rawValue }));
    }
  }

  function revertCell(segs: string[]) {
    const keyPath = joinPath(segs);
    // Pending takes priority — drop it. If there's also a saved override we
    // explicitly null it out so the PATCH below will clear it on save.
    const saved = selected?.kind === "custom" ? customMethodsQ.data?.find(c => c.id === selected.id)?.overrides || {} : {};
    setPendingOverrides(prev => {
      const next = { ...prev };
      if (keyPath in saved) {
        // We need a sentinel — undefined isn't preserved through JSON.
        // The cleanest approach: rebuild the saved map without this key.
        next[`__delete__${keyPath}`] = true;
      }
      delete next[keyPath];
      return next;
    });
  }

  // When applying bulkMultiplier to a section, walk every numeric leaf in
  // that section and stage an override of value × multiplier.
  function applyBulkMultiplier(section: SectionDef) {
    const factor = parseFloat(bulkMultiplier[section.key] || "");
    if (isNaN(factor) || factor === 1) return;
    if (!effectiveData) return;
    const subtree = readBySegs(effectiveData, section.pathFromRoot);
    if (!subtree) return;
    const leaves = collectLeaves(subtree).filter(l => typeof l.value === "number");
    const next: Record<string, any> = { ...pendingOverrides };
    for (const lf of leaves) {
      const fullSegs = [...section.pathFromRoot, ...lf.pathSegs];
      const kp = joinPath(fullSegs);
      next[kp] = Math.round(lf.value * factor * 10000) / 10000;
    }
    setPendingOverrides(next);
    setBulkMultiplier(prev => ({ ...prev, [section.key]: "" }));
    toast({ title: `Multiplied ${leaves.length} cells by ${factor}` });
  }

  // Copy an entire section from another method by enumerating that method's
  // leaves at the same path and staging each as an override.
  function applyCopyFrom(section: SectionDef, sourceKey: string) {
    if (!baseMethodsQ.data) return;
    // Source data: base data + (if custom) overrides applied.
    let sourceData: any;
    if (sourceKey.startsWith("custom:")) {
      const cmId = sourceKey.slice(7);
      const cm = customMethodsQ.data?.find(c => c.id === cmId);
      if (!cm) return;
      const base = baseMethodsQ.data.data[cm.baseMethod];
      if (!base) return;
      sourceData = JSON.parse(JSON.stringify(base));
      for (const [keyPath, value] of Object.entries(cm.overrides || {})) {
        const parts: string[] = [];
        let buf = "";
        for (let i = 0; i < keyPath.length; i++) {
          const c = keyPath[i];
          if (c === "\\" && i + 1 < keyPath.length && keyPath[i + 1] === ".") { buf += "."; i++; }
          else if (c === ".") { parts.push(buf); buf = ""; }
          else { buf += c; }
        }
        parts.push(buf);
        let cur = sourceData;
        for (let i = 0; i < parts.length - 1; i++) {
          if (cur[parts[i]] === undefined || cur[parts[i]] === null || typeof cur[parts[i]] !== "object") cur[parts[i]] = {};
          cur = cur[parts[i]];
        }
        cur[parts[parts.length - 1]] = value;
      }
    } else {
      const base = baseMethodsQ.data.data[sourceKey as "bill" | "justin" | "industry"];
      if (!base) return;
      sourceData = base;
    }
    const sourceSubtree = readBySegs(sourceData, section.pathFromRoot);
    if (!sourceSubtree) {
      toast({ title: "Source has no values for this section", variant: "destructive" });
      return;
    }
    const leaves = collectLeaves(sourceSubtree);
    const next: Record<string, any> = { ...pendingOverrides };
    let count = 0;
    for (const lf of leaves) {
      if (typeof lf.value !== "number" && typeof lf.value !== "string") continue;
      const fullSegs = [...section.pathFromRoot, ...lf.pathSegs];
      const kp = joinPath(fullSegs);
      next[kp] = lf.value;
      count++;
    }
    setPendingOverrides(next);
    setShowCopyFrom(null);
    setCopyFromKey("");
    toast({ title: `Copied ${count} cells` });
  }

  // ----------------------- Save handler -----------------------
  function handleSave() {
    if (!selected || selected.kind !== "custom") return;
    const cm = customMethodsQ.data?.find(c => c.id === selected.id);
    if (!cm) return;
    // Compose final overrides: start from saved, apply pending, honor __delete__ sentinels.
    const finalOverrides: Record<string, any> = { ...(cm.overrides || {}) };
    for (const [k, v] of Object.entries(pendingOverrides)) {
      if (k.startsWith("__delete__")) {
        delete finalOverrides[k.slice("__delete__".length)];
      } else {
        finalOverrides[k] = v;
      }
    }
    patchMutation.mutate({ id: cm.id, data: { overrides: finalOverrides } });
  }

  function handleRename() {
    if (!selected || selected.kind !== "custom" || !renameValue.trim()) return;
    patchMutation.mutate({ id: selected.id, data: { name: renameValue.trim() } });
  }

  // ----------------------- Export handler -----------------------
  async function handleExport() {
    if (!selected) return;
    const methodKey = selected.kind === "base" ? selected.key : `custom:${selected.id}`;
    try {
      const res = await apiRequest("GET", `/api/methods/${encodeURIComponent(methodKey)}/export`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${selected.name} - Factors.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      toast({ title: `Downloaded ${selected.name} factors` });
    } catch (err: any) {
      toast({ title: "Export failed", description: err?.message || String(err), variant: "destructive" });
    }
  }

  // ----------------------- Method list -----------------------
  const allMethods: MethodRef[] = useMemo(() => {
    const out: MethodRef[] = [];
    for (const m of baseMethodsQ.data?.methods || []) {
      out.push({ kind: "base", key: m.key, name: m.name, description: m.description });
    }
    for (const c of customMethodsQ.data || []) {
      out.push({ kind: "custom", id: c.id, name: c.name, baseMethod: c.baseMethod, description: c.description || "" });
    }
    return out;
  }, [baseMethodsQ.data, customMethodsQ.data]);

  // Counts so the user knows how many changes are staged.
  const pendingChangeCount = useMemo(() => {
    return Object.keys(pendingOverrides).filter(k => !k.startsWith("__delete__")).length
      + Object.keys(pendingOverrides).filter(k => k.startsWith("__delete__")).length;
  }, [pendingOverrides]);

  // ----------------------- Render -----------------------
  if (baseMethodsQ.isLoading || customMethodsQ.isLoading) {
    return (
      <AppLayout subtitle="Methods">
        <div className="p-6 text-sm text-muted-foreground">Loading methods…</div>
      </AppLayout>
    );
  }

  if (baseMethodsQ.error || !baseMethodsQ.data) {
    return (
      <AppLayout subtitle="Methods">
        <div className="p-6">
          <Card>
            <CardContent className="pt-6 text-sm text-destructive flex items-center gap-2">
              <AlertCircle size={16} /> Failed to load estimator methods. Try reloading the page.
            </CardContent>
          </Card>
        </div>
      </AppLayout>
    );
  }

  return (
    <AppLayout subtitle="Methods">
      <div className="flex h-full">
        {/* Left rail: method list */}
        <aside className="w-72 border-r border-border bg-card/50 overflow-y-auto shrink-0">
          <div className="px-4 py-4 border-b border-border">
            <h2 className="text-sm font-semibold">Estimating Methods</h2>
            <p className="text-[11px] text-muted-foreground mt-0.5">Browse, edit, and export factor tables</p>
          </div>

          {/* Base methods */}
          <div className="px-2 py-2">
            <p className="text-[10px] uppercase tracking-wider text-muted-foreground font-semibold px-2 mb-1">Base methods</p>
            {(baseMethodsQ.data?.methods || []).map(m => {
              const active = selected?.kind === "base" && selected.key === m.key;
              return (
                <button
                  key={m.key}
                  onClick={() => setSelected({ kind: "base", key: m.key, name: m.name, description: m.description })}
                  className={`w-full text-left px-3 py-2 rounded-md text-sm transition-colors flex items-start gap-2 ${
                    active ? "bg-primary/10 text-primary border border-primary/30" : "hover:bg-accent/50 border border-transparent"
                  }`}
                  data-testid={`method-${m.key}`}
                >
                  <Lock size={12} className="mt-1 shrink-0 text-muted-foreground" />
                  <div className="min-w-0">
                    <div className="font-medium truncate">{m.name}</div>
                    <div className="text-[10px] text-muted-foreground line-clamp-2">{m.description || m.source}</div>
                  </div>
                </button>
              );
            })}
          </div>

          {/* Custom methods */}
          <div className="px-2 py-2 border-t border-border">
            <div className="flex items-center justify-between px-2 mb-1">
              <p className="text-[10px] uppercase tracking-wider text-muted-foreground font-semibold">Custom profiles</p>
              <Button
                size="sm"
                variant="ghost"
                className="h-6 px-2 text-xs"
                onClick={() => {
                  // Create a blank custom method off Justin (most flexible default).
                  saveAsMutation.mutate({ name: `Custom ${(customMethodsQ.data?.length || 0) + 1}`, baseMethod: "justin", overrides: {} });
                }}
                data-testid="btn-new-custom"
              >
                <Plus size={12} className="mr-1" /> New
              </Button>
            </div>
            {(customMethodsQ.data || []).length === 0 ? (
              <p className="text-[11px] text-muted-foreground px-3 py-2">No custom profiles yet. Click "New" or use "Save As Custom" on a base method.</p>
            ) : (
              (customMethodsQ.data || []).map(c => {
                const active = selected?.kind === "custom" && selected.id === c.id;
                return (
                  <button
                    key={c.id}
                    onClick={() => setSelected({ kind: "custom", id: c.id, name: c.name, baseMethod: c.baseMethod, description: c.description || "" })}
                    className={`w-full text-left px-3 py-2 rounded-md text-sm transition-colors flex items-start gap-2 ${
                      active ? "bg-primary/10 text-primary border border-primary/30" : "hover:bg-accent/50 border border-transparent"
                    }`}
                    data-testid={`custom-${c.id}`}
                  >
                    <Pencil size={12} className="mt-1 shrink-0 text-muted-foreground" />
                    <div className="min-w-0 flex-1">
                      <div className="font-medium truncate">{c.name}</div>
                      <div className="text-[10px] text-muted-foreground">
                        based on {c.baseMethod === "bill" ? "Bill" : c.baseMethod === "justin" ? "Justin" : "Industry"}
                        {" · "}
                        {Object.keys(c.overrides || {}).length} override{Object.keys(c.overrides || {}).length === 1 ? "" : "s"}
                      </div>
                    </div>
                  </button>
                );
              })
            )}
          </div>
        </aside>

        {/* Main pane */}
        <main className="flex-1 overflow-auto">
          {!selected ? (
            <div className="p-6 text-sm text-muted-foreground">Select a method to view its factors.</div>
          ) : (
            <div className="p-6 space-y-4 max-w-6xl">
              {/* Header */}
              <div className="flex items-start justify-between gap-4 flex-wrap">
                <div className="min-w-0 flex-1">
                  {selected.kind === "custom" && renaming ? (
                    <div className="flex items-center gap-2">
                      <Input
                        value={renameValue}
                        onChange={e => setRenameValue(e.target.value)}
                        className="max-w-sm h-9 text-base"
                        autoFocus
                        onKeyDown={e => { if (e.key === "Enter") handleRename(); if (e.key === "Escape") setRenaming(false); }}
                      />
                      <Button size="sm" onClick={handleRename} disabled={patchMutation.isPending}><CheckCircle2 size={14} /></Button>
                      <Button size="sm" variant="ghost" onClick={() => setRenaming(false)}><X size={14} /></Button>
                    </div>
                  ) : (
                    <div className="flex items-center gap-2 flex-wrap">
                      <h1 className="text-xl font-bold">{selected.name}</h1>
                      {selected.kind === "base" ? (
                        <Badge variant="secondary" className="text-[10px]"><Lock size={10} className="mr-1" /> Read-only</Badge>
                      ) : (
                        <>
                          <Badge variant="outline" className="text-[10px]">Custom · based on {baseKey === "bill" ? "Bill" : baseKey === "justin" ? "Justin" : "Industry"}</Badge>
                          <Button size="sm" variant="ghost" className="h-7 px-2" onClick={() => { setRenameValue(selected.name); setRenaming(true); }}>
                            <Pencil size={12} />
                          </Button>
                        </>
                      )}
                    </div>
                  )}
                  <p className="text-xs text-muted-foreground mt-1 max-w-2xl">{selected.description}</p>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  {selected.kind === "base" && (
                    <Button size="sm" variant="outline" onClick={() => { setSaveAsName(`My ${selected.name}`); setShowSaveAs(true); }} data-testid="btn-save-as">
                      <Copy size={13} className="mr-1.5" /> Save As Custom
                    </Button>
                  )}
                  {selected.kind === "custom" && (
                    <>
                      <Button
                        size="sm"
                        onClick={handleSave}
                        disabled={pendingChangeCount === 0 || patchMutation.isPending}
                        data-testid="btn-save-overrides"
                      >
                        <Save size={13} className="mr-1.5" />
                        {patchMutation.isPending ? "Saving…" : pendingChangeCount > 0 ? `Save ${pendingChangeCount} change${pendingChangeCount === 1 ? "" : "s"}` : "Saved"}
                      </Button>
                      <Button size="sm" variant="outline" onClick={() => setPendingOverrides({})} disabled={pendingChangeCount === 0}>
                        Discard
                      </Button>
                      <Button size="sm" variant="ghost" className="text-destructive" onClick={() => setShowDeleteConfirm(true)}>
                        <Trash2 size={13} />
                      </Button>
                    </>
                  )}
                  <Button size="sm" variant="outline" onClick={handleExport} data-testid="btn-export-method">
                    <Download size={13} className="mr-1.5" /> Export to Excel
                  </Button>
                </div>
              </div>

              <Separator />

              {/* Sections */}
              <Accordion type="multiple" defaultValue={sectionDefs.map(s => s.key)} className="space-y-2">
                {sectionDefs.map(section => {
                  const subtree = effectiveData ? readBySegs(effectiveData, section.pathFromRoot) : null;
                  return (
                    <AccordionItem key={section.key} value={section.key} className="border rounded-md px-3 bg-card">
                      <AccordionTrigger className="text-sm font-semibold py-3 hover:no-underline">
                        <div className="flex items-center gap-2">
                          <Calculator size={14} className="text-muted-foreground" />
                          {section.label}
                          {subtree && typeof subtree === "object" && (
                            <span className="text-[10px] text-muted-foreground font-normal ml-1">
                              ({Object.keys(subtree).length} entries)
                            </span>
                          )}
                        </div>
                      </AccordionTrigger>
                      <AccordionContent className="pb-3">
                        {!subtree || typeof subtree !== "object" ? (
                          <p className="text-xs text-muted-foreground italic">No data for this section.</p>
                        ) : (
                          <FactorSection
                            section={section}
                            subtree={subtree}
                            isEditable={isEditable}
                            pendingOverrides={pendingOverrides}
                            savedOverrides={selected?.kind === "custom" ? (customMethodsQ.data?.find(c => c.id === selected.id)?.overrides || {}) : {}}
                            allMethodsForCopy={allMethods.filter(m => !(selected.kind === "base" && m.kind === "base" && m.key === selected.key) && !(selected.kind === "custom" && m.kind === "custom" && m.id === selected.id))}
                            onCommitCell={commitCell}
                            onRevertCell={revertCell}
                            bulkMultiplier={bulkMultiplier[section.key] || ""}
                            onBulkChange={v => setBulkMultiplier(prev => ({ ...prev, [section.key]: v }))}
                            onBulkApply={() => applyBulkMultiplier(section)}
                            showCopyFrom={showCopyFrom === section.key}
                            onShowCopyFrom={() => { setShowCopyFrom(section.key); setCopyFromKey(""); }}
                            onCloseCopyFrom={() => { setShowCopyFrom(null); setCopyFromKey(""); }}
                            copyFromKey={copyFromKey}
                            onCopyFromKeyChange={setCopyFromKey}
                            onApplyCopyFrom={() => copyFromKey && applyCopyFrom(section, copyFromKey)}
                          />
                        )}
                      </AccordionContent>
                    </AccordionItem>
                  );
                })}
              </Accordion>
            </div>
          )}
        </main>
      </div>

      {/* Save As Custom dialog */}
      <Dialog open={showSaveAs} onOpenChange={setShowSaveAs}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Save as Custom Method</DialogTitle>
          </DialogHeader>
          <div className="space-y-2 py-2">
            <label className="text-xs font-medium">Name</label>
            <Input value={saveAsName} onChange={e => setSaveAsName(e.target.value)} placeholder="My custom method" autoFocus />
            <p className="text-[11px] text-muted-foreground">
              Creates an editable copy of {selected?.name}. Existing factors are inherited from the base method; only changes you make are stored as overrides.
            </p>
          </div>
          <DialogFooter>
            <Button variant="ghost" onClick={() => setShowSaveAs(false)}>Cancel</Button>
            <Button
              onClick={() => {
                if (!selected || selected.kind !== "base" || !saveAsName.trim()) return;
                saveAsMutation.mutate({ name: saveAsName.trim(), baseMethod: selected.key, overrides: {} });
              }}
              disabled={!saveAsName.trim() || saveAsMutation.isPending}
            >
              {saveAsMutation.isPending ? "Saving…" : "Save"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete confirm */}
      <Dialog open={showDeleteConfirm} onOpenChange={setShowDeleteConfirm}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Delete this custom method?</DialogTitle>
          </DialogHeader>
          <p className="text-sm text-muted-foreground py-2">
            This removes <strong>{selected?.kind === "custom" ? selected.name : ""}</strong> and all its overrides. Any saved estimates that reference it will fall back to its base method on next recalculation.
          </p>
          <DialogFooter>
            <Button variant="ghost" onClick={() => setShowDeleteConfirm(false)}>Cancel</Button>
            <Button variant="destructive" onClick={() => selected?.kind === "custom" && deleteMutation.mutate(selected.id)} disabled={deleteMutation.isPending}>
              {deleteMutation.isPending ? "Deleting…" : "Delete"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </AppLayout>
  );
}

// ----------------------- FactorSection -----------------------
// Renders a single accordion section's factor table. Editable when the parent
// is a custom method. Cell-level changes call onCommitCell, which the parent
// stages in pendingOverrides. The colored dot indicates a cell differs from
// its base value.

function FactorSection(props: {
  section: SectionDef;
  subtree: any;
  isEditable: boolean;
  pendingOverrides: Record<string, any>;
  savedOverrides: Record<string, any>;
  allMethodsForCopy: MethodRef[];
  onCommitCell: (segs: string[], rawValue: string, originalType: "number" | "string") => void;
  onRevertCell: (segs: string[]) => void;
  bulkMultiplier: string;
  onBulkChange: (v: string) => void;
  onBulkApply: () => void;
  showCopyFrom: boolean;
  onShowCopyFrom: () => void;
  onCloseCopyFrom: () => void;
  copyFromKey: string;
  onCopyFromKeyChange: (v: string) => void;
  onApplyCopyFrom: () => void;
}) {
  const { section, subtree, isEditable, pendingOverrides, savedOverrides, allMethodsForCopy } = props;

  // Build the row-by-col grid. Each cell's full path = section.pathFromRoot + leaf.pathSegs.
  const leaves = useMemo(() => collectLeaves(subtree), [subtree]);
  const rows = useMemo(() => groupRowsCols(leaves), [leaves]);
  rows.sort((a, b) => sortRowKey(a.rowKey, b.rowKey));
  const colKeys = useMemo(() => uniqueColKeys(rows), [rows]);

  // A cell is "modified" if it differs from the base value. We can tell because
  // pendingOverrides or savedOverrides contain the joined keyPath.
  function cellIsModified(fullSegs: string[]): boolean {
    const kp = joinPath(fullSegs);
    if (kp in pendingOverrides) return true;
    if (kp in savedOverrides) return true;
    return false;
  }

  return (
    <div className="space-y-3">
      {/* Toolbar: bulk + copy-from (only for editable customs) */}
      {isEditable && (
        <div className="flex items-center gap-2 flex-wrap text-xs">
          <div className="flex items-center gap-1">
            <span className="text-muted-foreground">Multiply all by</span>
            <Input
              type="number"
              step="0.01"
              value={props.bulkMultiplier}
              onChange={e => props.onBulkChange(e.target.value)}
              placeholder="1.10"
              className="h-7 w-24 text-xs"
            />
            <Button size="sm" variant="outline" className="h-7 px-2" onClick={props.onBulkApply} disabled={!props.bulkMultiplier}>
              Apply
            </Button>
          </div>
          <Separator orientation="vertical" className="h-5" />
          {!props.showCopyFrom ? (
            <Button size="sm" variant="outline" className="h-7 px-2" onClick={props.onShowCopyFrom}>
              <Copy size={11} className="mr-1" /> Copy from…
            </Button>
          ) : (
            <div className="flex items-center gap-1">
              <Select value={props.copyFromKey} onValueChange={props.onCopyFromKeyChange}>
                <SelectTrigger className="h-7 text-xs w-44"><SelectValue placeholder="Source method" /></SelectTrigger>
                <SelectContent>
                  {allMethodsForCopy.map(m => (
                    <SelectItem key={m.kind === "base" ? `b-${m.key}` : `c-${m.id}`} value={m.kind === "base" ? m.key : `custom:${m.id}`}>
                      {m.name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <Button size="sm" variant="outline" className="h-7 px-2" onClick={props.onApplyCopyFrom} disabled={!props.copyFromKey}>
                Copy
              </Button>
              <Button size="sm" variant="ghost" className="h-7 w-7 p-0" onClick={props.onCloseCopyFrom}>
                <X size={12} />
              </Button>
            </div>
          )}
        </div>
      )}

      {/* Data table */}
      <div className="overflow-x-auto border rounded">
        <table className="text-xs w-full">
          <thead className="bg-muted/40">
            <tr>
              <th className="text-left px-3 py-2 font-semibold sticky left-0 bg-muted/40">Key</th>
              {colKeys.map(c => (
                <th key={c} className="text-right px-3 py-2 font-semibold whitespace-nowrap">{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(row => (
              <tr key={row.rowKey} className="border-t hover:bg-muted/20">
                <td className="px-3 py-1.5 sticky left-0 bg-card font-medium">{row.rowKey}</td>
                {colKeys.map(colKey => {
                  const cell = row.cells.find(c => c.colKey === colKey);
                  if (!cell) return <td key={colKey} className="px-3 py-1.5 text-right text-muted-foreground/40">—</td>;
                  const fullSegs = [...section.pathFromRoot, ...cell.segs];
                  const modified = cellIsModified(fullSegs);
                  const cellType: "number" | "string" = typeof cell.value === "number" ? "number" : "string";
                  return (
                    <td key={colKey} className="px-2 py-1 text-right whitespace-nowrap">
                      <div className="flex items-center justify-end gap-1">
                        {isEditable ? (
                          <Input
                            type={cellType === "number" ? "number" : "text"}
                            step="0.01"
                            defaultValue={cell.value}
                            onBlur={e => {
                              const raw = e.target.value;
                              const orig = String(cell.value);
                              if (raw !== orig) props.onCommitCell(fullSegs, raw, cellType);
                            }}
                            className={`h-7 text-xs text-right w-20 ${modified ? "border-amber-500 ring-1 ring-amber-500/30" : ""}`}
                          />
                        ) : (
                          <span className="tabular-nums">{typeof cell.value === "number" ? cell.value : String(cell.value)}</span>
                        )}
                        {modified && isEditable && (
                          <Button
                            size="sm"
                            variant="ghost"
                            className="h-6 w-6 p-0"
                            onClick={() => props.onRevertCell(fullSegs)}
                            title="Revert to base value"
                          >
                            <RotateCcw size={11} />
                          </Button>
                        )}
                      </div>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
