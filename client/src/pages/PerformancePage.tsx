/**
 * PerformancePage.tsx
 * "Bid vs Actual" view for closed projects.
 *
 * Layout:
 *   1. Header band — title, subtitle, Add Closed Project button, Import from Estimate button
 *   2. KPI strip — 4 stat cards (total projects, actuals entered, avg labor Δ%, avg material Δ%)
 *   3. Projects table — per-row expansion with Bid vs Actual breakdown
 *   4. Edit Actuals dialog — PATCH form for updating any CompletedProject field
 *   5. Add Closed Project dialog — POST /api/completed-projects (manual entry)
 *   6. Import from Estimate dialog — POST /api/completed-projects/from-estimate
 *
 * Sign convention: positive variance = over budget = bad (actual > bid).
 * Variance % = (actual - bid) / bid * 100
 */

import { useState, useMemo } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import {
  TrendingUp,
  Plus,
  Upload,
  ChevronDown,
  ChevronUp,
  Edit2,
  ExternalLink,
  AlertCircle,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import AppLayout from "@/components/AppLayout";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";
import type { CompletedProject, EstimateProject } from "@shared/schema";

// ---------------------------------------------------------------------------
// Formatting helpers
// ---------------------------------------------------------------------------

/** Format a number as dollars with 2 decimal places. */
function fmt$(n: number | null | undefined): string {
  if (n == null || isNaN(n)) return "—";
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** Format a manhour integer with thousands commas. */
function fmtMH(n: number | null | undefined): string {
  if (n == null || isNaN(n)) return "—";
  return Math.round(n).toLocaleString("en-US");
}

/** Compute variance percent: (actual - bid) / bid * 100. Returns null when not computable. */
function variancePct(bid: number | null | undefined, actual: number | null | undefined): number | null {
  if (bid == null || actual == null || bid === 0) return null;
  return ((actual - bid) / Math.abs(bid)) * 100;
}

/** Format a variance percentage string with sign. */
function fmtVariance(pct: number | null): string {
  if (pct == null) return "—";
  const sign = pct >= 0 ? "+" : "";
  return `${sign}${pct.toFixed(1)}%`;
}

/**
 * Tailwind class for a variance cell.
 * green  = abs < 5%   (on budget / under budget)
 * amber  = abs 5–15%
 * red    = abs > 15%  (significantly over budget)
 */
function varianceClass(pct: number | null): string {
  if (pct == null) return "text-muted-foreground";
  const abs = Math.abs(pct);
  if (abs < 5) return "text-green-600 dark:text-green-400 font-semibold";
  if (abs < 15) return "text-amber-600 dark:text-amber-400 font-semibold";
  return "text-red-600 dark:text-red-400 font-semibold";
}

// ---------------------------------------------------------------------------
// Helpers: project status label
// ---------------------------------------------------------------------------

type ProjectStatus = "bid-only" | "actuals-entered" | "linked";

function getStatus(p: CompletedProject): ProjectStatus {
  if (p.sourceTakeoffId || p.sourceEstimateId) return "linked";
  // "Actuals entered" = the project has real labor hours or costs beyond the bid snapshot
  if (p.totalManhours > 0 || p.materialCost > 0 || p.laborCost > 0) return "actuals-entered";
  return "bid-only";
}

function StatusBadge({ status }: { status: ProjectStatus }) {
  const map: Record<ProjectStatus, { label: string; variant: "default" | "secondary" | "outline" }> = {
    "bid-only": { label: "Bid only", variant: "outline" },
    "actuals-entered": { label: "Actuals entered", variant: "secondary" },
    "linked": { label: "Linked to Takeoff/Estimate", variant: "default" },
  };
  const { label, variant } = map[status];
  return (
    <Badge variant={variant} className="text-[10px] whitespace-nowrap">
      {label}
    </Badge>
  );
}

// ---------------------------------------------------------------------------
// KPI calculation helper
// ---------------------------------------------------------------------------

interface KpiData {
  totalProjects: number;
  withActuals: number;
  avgLaborVariancePct: number | null;
  avgMaterialVariancePct: number | null;
}

function computeKpis(projects: CompletedProject[]): KpiData {
  const total = projects.length;
  let withActuals = 0;
  const laborVariances: number[] = [];
  const materialVariances: number[] = [];

  for (const p of projects) {
    // "Has actuals" means actual labor hours or actual costs have been entered
    const hasActualMH = p.totalManhours > 0;
    const hasActualCost = p.materialCost > 0 || p.laborCost > 0;
    if (hasActualMH || hasActualCost) withActuals++;

    const lv = variancePct(p.bidLaborHours, p.totalManhours);
    if (lv != null) laborVariances.push(lv);

    const mv = variancePct(p.bidMaterialCost, p.materialCost);
    if (mv != null) materialVariances.push(mv);
  }

  const avg = (arr: number[]) =>
    arr.length === 0 ? null : arr.reduce((s, x) => s + x, 0) / arr.length;

  return {
    totalProjects: total,
    withActuals,
    avgLaborVariancePct: avg(laborVariances),
    avgMaterialVariancePct: avg(materialVariances),
  };
}

// ---------------------------------------------------------------------------
// Empty form templates
// ---------------------------------------------------------------------------

const EMPTY_ADD_FORM = {
  name: "",
  client: "",
  location: "",
  discipline: "" as "" | "mechanical" | "structural" | "civil",
  scopeDescription: "",
  bidLaborHours: "" as string | number,
  bidMaterialCost: "" as string | number,
  bidTotalCost: "" as string | number,
  actualLaborHours: "" as string | number, // maps to totalManhours
  totalManhours: "" as string | number,
  materialCost: "" as string | number,
  laborCost: "" as string | number,
  totalCost: "" as string | number,
  durationDays: "" as string | number,
  tags: "",
  notes: "",
};

type AddFormState = typeof EMPTY_ADD_FORM;

// ---------------------------------------------------------------------------
// Per-category labor editor (for actualLaborByCategoryJson)
// ---------------------------------------------------------------------------

interface CategoryRow { category: string; hours: string }

function parseCategoryJson(raw: string | undefined): CategoryRow[] {
  if (!raw) return [];
  try {
    const obj = JSON.parse(raw);
    return Object.entries(obj).map(([category, hours]) => ({
      category,
      hours: String(hours),
    }));
  } catch {
    return [];
  }
}

function serializeCategoryRows(rows: CategoryRow[]): string {
  const obj: Record<string, number> = {};
  for (const r of rows) {
    if (r.category.trim()) {
      const h = parseFloat(r.hours);
      if (!isNaN(h)) obj[r.category.trim()] = h;
    }
  }
  return JSON.stringify(obj);
}

function CategoryEditor({
  rows,
  onChange,
}: {
  rows: CategoryRow[];
  onChange: (rows: CategoryRow[]) => void;
}) {
  const addRow = () => onChange([...rows, { category: "", hours: "" }]);
  const removeRow = (i: number) => onChange(rows.filter((_, idx) => idx !== i));
  const updateRow = (i: number, field: keyof CategoryRow, val: string) => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, [field]: val } : r));
    onChange(next);
  };

  return (
    <div className="space-y-1.5">
      {rows.map((r, i) => (
        <div key={i} className="flex gap-2 items-center">
          <Input
            className="h-7 text-xs flex-1"
            placeholder="Category (e.g. pipe)"
            value={r.category}
            onChange={(e) => updateRow(i, "category", e.target.value)}
          />
          <Input
            className="h-7 text-xs w-24"
            placeholder="Hours"
            type="number"
            value={r.hours}
            onChange={(e) => updateRow(i, "hours", e.target.value)}
          />
          <Button
            size="icon"
            variant="ghost"
            className="h-7 w-7 text-destructive hover:text-destructive shrink-0"
            onClick={() => removeRow(i)}
          >
            ×
          </Button>
        </div>
      ))}
      <Button size="sm" variant="outline" className="h-7 text-xs" onClick={addRow}>
        <Plus size={11} className="mr-1" /> Add category
      </Button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Add Closed Project Dialog
// ---------------------------------------------------------------------------

interface AddProjectDialogProps {
  open: boolean;
  onClose: () => void;
}

function AddProjectDialog({ open, onClose }: AddProjectDialogProps) {
  const { toast } = useToast();
  const [form, setForm] = useState<AddFormState>(EMPTY_ADD_FORM);
  const upd = (field: keyof AddFormState, val: string | number) =>
    setForm((prev) => ({ ...prev, [field]: val }));

  const mutation = useMutation({
    mutationFn: async (payload: Record<string, unknown>) => {
      const res = await apiRequest("POST", "/api/completed-projects", payload);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/completed-projects"] });
      toast({ title: "Project added", description: "Closed project saved to Performance." });
      setForm(EMPTY_ADD_FORM);
      onClose();
    },
    onError: (err: Error) => {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    },
  });

  function handleSave() {
    if (!form.name.trim() || !form.scopeDescription.trim()) {
      toast({
        title: "Required fields missing",
        description: "Project name and scope description are required.",
        variant: "destructive",
      });
      return;
    }
    const n = (v: string | number) => (v === "" ? undefined : Number(v));
    mutation.mutate({
      name: form.name.trim(),
      client: form.client.trim() || undefined,
      location: form.location.trim() || undefined,
      discipline: form.discipline || undefined,
      scopeDescription: form.scopeDescription.trim(),
      bidLaborHours: n(form.bidLaborHours),
      bidMaterialCost: n(form.bidMaterialCost),
      bidTotalCost: n(form.bidTotalCost),
      totalManhours: n(form.actualLaborHours) ?? 0,
      materialCost: n(form.materialCost) ?? 0,
      laborCost: n(form.laborCost) ?? 0,
      totalCost: n(form.totalCost) ?? 0,
      durationDays: n(form.durationDays),
      tags: form.tags.trim() || undefined,
      notes: form.notes.trim() || undefined,
    });
  }

  return (
    <Dialog open={open} onOpenChange={(v) => !v && onClose()}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-sm font-semibold">Add Closed Project</DialogTitle>
        </DialogHeader>

        <div className="space-y-4 py-2">
          {/* Row 1: name / client / location */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <Label className="text-xs">Project Name *</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={form.name}
                onChange={(e) => upd("name", e.target.value)}
                placeholder="e.g. Tank Farm Piping Phase 2"
              />
            </div>
            <div>
              <Label className="text-xs">Client</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={form.client}
                onChange={(e) => upd("client", e.target.value)}
                placeholder="e.g. ExxonMobil"
              />
            </div>
            <div>
              <Label className="text-xs">Location</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={form.location}
                onChange={(e) => upd("location", e.target.value)}
                placeholder="e.g. Baton Rouge, LA"
              />
            </div>
          </div>

          {/* Discipline */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <Label className="text-xs">Discipline</Label>
              <Select
                value={form.discipline}
                onValueChange={(v) => upd("discipline", v)}
              >
                <SelectTrigger className="h-8 text-xs mt-1">
                  <SelectValue placeholder="Select discipline" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="mechanical">Mechanical</SelectItem>
                  <SelectItem value="structural">Structural</SelectItem>
                  <SelectItem value="civil">Civil</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label className="text-xs">Duration (days)</Label>
              <Input
                className="h-8 text-xs mt-1"
                type="number"
                value={form.durationDays}
                onChange={(e) => upd("durationDays", e.target.value)}
                placeholder="0"
              />
            </div>
          </div>

          {/* Scope */}
          <div>
            <Label className="text-xs">Scope Description *</Label>
            <textarea
              className="w-full h-16 text-xs mt-1 border border-input rounded-md px-3 py-2 bg-background resize-none focus:outline-none focus:ring-2 focus:ring-ring"
              value={form.scopeDescription}
              onChange={(e) => upd("scopeDescription", e.target.value)}
              placeholder="e.g. Install 6-inch screw conveyor, run 500' of 8-inch SS pipe"
            />
          </div>

          {/* Bid numbers */}
          <div>
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1.5">
              Bid (at time of award)
            </p>
            <div className="grid grid-cols-3 gap-3">
              <div>
                <Label className="text-xs">Bid Labor MH</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.bidLaborHours}
                  onChange={(e) => upd("bidLaborHours", e.target.value)}
                  placeholder="0"
                />
              </div>
              <div>
                <Label className="text-xs">Bid Material ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.bidMaterialCost}
                  onChange={(e) => upd("bidMaterialCost", e.target.value)}
                  placeholder="0"
                />
              </div>
              <div>
                <Label className="text-xs">Bid Total ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.bidTotalCost}
                  onChange={(e) => upd("bidTotalCost", e.target.value)}
                  placeholder="0"
                />
              </div>
            </div>
          </div>

          {/* Actual numbers */}
          <div>
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1.5">
              Actuals (from closeout)
            </p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <div>
                <Label className="text-xs">Actual Labor MH</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.actualLaborHours}
                  onChange={(e) => upd("actualLaborHours", e.target.value)}
                  placeholder="0"
                />
              </div>
              <div>
                <Label className="text-xs">Material ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.materialCost}
                  onChange={(e) => upd("materialCost", e.target.value)}
                  placeholder="0"
                />
              </div>
              <div>
                <Label className="text-xs">Labor ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.laborCost}
                  onChange={(e) => upd("laborCost", e.target.value)}
                  placeholder="0"
                />
              </div>
              <div>
                <Label className="text-xs">Total ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={form.totalCost}
                  onChange={(e) => upd("totalCost", e.target.value)}
                  placeholder="0"
                />
              </div>
            </div>
          </div>

          {/* Tags / Notes */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <Label className="text-xs">Tags (comma-separated)</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={form.tags}
                onChange={(e) => upd("tags", e.target.value)}
                placeholder="piping, stainless, tank-farm"
              />
            </div>
            <div>
              <Label className="text-xs">Notes</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={form.notes}
                onChange={(e) => upd("notes", e.target.value)}
                placeholder="Any additional notes..."
              />
            </div>
          </div>
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={onClose}>
            Cancel
          </Button>
          <Button size="sm" onClick={handleSave} disabled={mutation.isPending}>
            {mutation.isPending ? "Saving..." : "Save Project"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

// ---------------------------------------------------------------------------
// Import from Estimate Dialog
// ---------------------------------------------------------------------------

interface ImportEstimateDialogProps {
  open: boolean;
  onClose: () => void;
}

function ImportEstimateDialog({ open, onClose }: ImportEstimateDialogProps) {
  const { toast } = useToast();
  const [selectedId, setSelectedId] = useState<string>("");

  const { data: estimates = [], isLoading: estimatesLoading } = useQuery<EstimateProject[]>({
    queryKey: ["/api/estimates"],
    enabled: open,
  });

  const mutation = useMutation({
    mutationFn: async (estimateId: string) => {
      const res = await apiRequest("POST", "/api/completed-projects/from-estimate", { estimateId });
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/completed-projects"] });
      toast({
        title: "Snapshot created",
        description: "Bid snapshot imported from estimate. Add actuals when available.",
      });
      setSelectedId("");
      onClose();
    },
    onError: (err: Error) => {
      toast({ title: "Import failed", description: err.message, variant: "destructive" });
    },
  });

  return (
    <Dialog open={open} onOpenChange={(v) => !v && onClose()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle className="text-sm font-semibold">Import from Estimate</DialogTitle>
        </DialogHeader>

        <div className="space-y-4 py-2">
          <p className="text-xs text-muted-foreground">
            Select an existing estimate to snapshot as a closed project. The bid labor
            hours, material cost, and total cost will be captured from the estimate at
            this moment, so subsequent changes to factors won't affect the historical record.
          </p>

          <div>
            <Label className="text-xs">Estimate</Label>
            {estimatesLoading ? (
              <p className="text-xs text-muted-foreground mt-2">Loading estimates…</p>
            ) : estimates.length === 0 ? (
              <p className="text-xs text-muted-foreground mt-2">
                No estimates found. Create an estimate on the Estimating page first.
              </p>
            ) : (
              <Select value={selectedId} onValueChange={setSelectedId}>
                <SelectTrigger className="mt-1 h-8 text-xs">
                  <SelectValue placeholder="Select an estimate…" />
                </SelectTrigger>
                <SelectContent>
                  {estimates.map((e) => (
                    <SelectItem key={e.id} value={e.id}>
                      {e.name || e.id}
                      {e.client ? ` — ${e.client}` : ""}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
          </div>
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={onClose}>
            Cancel
          </Button>
          <Button
            size="sm"
            disabled={!selectedId || mutation.isPending}
            onClick={() => mutation.mutate(selectedId)}
          >
            {mutation.isPending ? "Importing…" : "Import Snapshot"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

// ---------------------------------------------------------------------------
// Edit Actuals Dialog
// ---------------------------------------------------------------------------

interface EditActualsDialogProps {
  project: CompletedProject | null;
  onClose: () => void;
}

function EditActualsDialog({ project, onClose }: EditActualsDialogProps) {
  const { toast } = useToast();

  // Local form state — initialized when the dialog mounts with a project
  const [name, setName] = useState(project?.name ?? "");
  const [client, setClient] = useState(project?.client ?? "");
  const [location, setLocation] = useState(project?.location ?? "");
  const [discipline, setDiscipline] = useState<"" | "mechanical" | "structural" | "civil">(
    (project?.discipline as "mechanical" | "structural" | "civil") ?? ""
  );
  const [scopeDescription, setScopeDescription] = useState(project?.scopeDescription ?? "");
  const [bidLaborHours, setBidLaborHours] = useState(String(project?.bidLaborHours ?? ""));
  const [bidMaterialCost, setBidMaterialCost] = useState(String(project?.bidMaterialCost ?? ""));
  const [bidTotalCost, setBidTotalCost] = useState(String(project?.bidTotalCost ?? ""));
  const [totalManhours, setTotalManhours] = useState(String(project?.totalManhours ?? ""));
  const [materialCost, setMaterialCost] = useState(String(project?.materialCost ?? ""));
  const [laborCost, setLaborCost] = useState(String(project?.laborCost ?? ""));
  const [totalCost, setTotalCost] = useState(String(project?.totalCost ?? ""));
  const [durationDays, setDurationDays] = useState(String(project?.durationDays ?? ""));
  const [tags, setTags] = useState(project?.tags ?? "");
  const [notes, setNotes] = useState(project?.notes ?? "");
  const [categoryRows, setCategoryRows] = useState<CategoryRow[]>(
    parseCategoryJson(project?.actualLaborByCategoryJson)
  );

  const mutation = useMutation({
    mutationFn: async (payload: Record<string, unknown>) => {
      const res = await apiRequest("PATCH", `/api/completed-projects/${project!.id}`, payload);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/completed-projects"] });
      toast({ title: "Saved", description: "Project actuals updated." });
      onClose();
    },
    onError: (err: Error) => {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    },
  });

  function handleSave() {
    if (!project) return;
    const n = (v: string) => (v === "" ? undefined : Number(v));
    const categoryJson = categoryRows.length > 0
      ? serializeCategoryRows(categoryRows)
      : undefined;

    mutation.mutate({
      name: name.trim() || undefined,
      client: client.trim() || undefined,
      location: location.trim() || undefined,
      discipline: discipline || undefined,
      scopeDescription: scopeDescription.trim() || undefined,
      bidLaborHours: n(bidLaborHours),
      bidMaterialCost: n(bidMaterialCost),
      bidTotalCost: n(bidTotalCost),
      totalManhours: n(totalManhours),
      materialCost: n(materialCost),
      laborCost: n(laborCost),
      totalCost: n(totalCost),
      durationDays: n(durationDays),
      tags: tags.trim() || undefined,
      notes: notes.trim() || undefined,
      actualLaborByCategoryJson: categoryJson,
    });
  }

  if (!project) return null;

  return (
    <Dialog open={!!project} onOpenChange={(v) => !v && onClose()}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-sm font-semibold">
            Edit Actuals — {project.name}
          </DialogTitle>
        </DialogHeader>

        <div className="space-y-4 py-2">
          {/* Name / client / location */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <Label className="text-xs">Project Name</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </div>
            <div>
              <Label className="text-xs">Client</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={client}
                onChange={(e) => setClient(e.target.value)}
              />
            </div>
            <div>
              <Label className="text-xs">Location</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={location}
                onChange={(e) => setLocation(e.target.value)}
              />
            </div>
          </div>

          {/* Discipline / duration */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <Label className="text-xs">Discipline</Label>
              <Select
                value={discipline}
                onValueChange={(v) => setDiscipline(v as typeof discipline)}
              >
                <SelectTrigger className="h-8 text-xs mt-1">
                  <SelectValue placeholder="Select discipline" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="mechanical">Mechanical</SelectItem>
                  <SelectItem value="structural">Structural</SelectItem>
                  <SelectItem value="civil">Civil</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label className="text-xs">Duration (days)</Label>
              <Input
                className="h-8 text-xs mt-1"
                type="number"
                value={durationDays}
                onChange={(e) => setDurationDays(e.target.value)}
              />
            </div>
          </div>

          {/* Scope */}
          <div>
            <Label className="text-xs">Scope Description</Label>
            <textarea
              className="w-full h-14 text-xs mt-1 border border-input rounded-md px-3 py-2 bg-background resize-none focus:outline-none focus:ring-2 focus:ring-ring"
              value={scopeDescription}
              onChange={(e) => setScopeDescription(e.target.value)}
            />
          </div>

          {/* Bid numbers */}
          <div>
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1.5">
              Bid Snapshot
            </p>
            <div className="grid grid-cols-3 gap-3">
              <div>
                <Label className="text-xs">Bid Labor MH</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={bidLaborHours}
                  onChange={(e) => setBidLaborHours(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-xs">Bid Material ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={bidMaterialCost}
                  onChange={(e) => setBidMaterialCost(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-xs">Bid Total ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={bidTotalCost}
                  onChange={(e) => setBidTotalCost(e.target.value)}
                />
              </div>
            </div>
          </div>

          {/* Actual numbers */}
          <div>
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1.5">
              Actuals
            </p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <div>
                <Label className="text-xs">Actual Labor MH</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={totalManhours}
                  onChange={(e) => setTotalManhours(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-xs">Material ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={materialCost}
                  onChange={(e) => setMaterialCost(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-xs">Labor ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={laborCost}
                  onChange={(e) => setLaborCost(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-xs">Total ($)</Label>
                <Input
                  className="h-8 text-xs mt-1"
                  type="number"
                  value={totalCost}
                  onChange={(e) => setTotalCost(e.target.value)}
                />
              </div>
            </div>
          </div>

          {/* Per-category labor breakdown */}
          <div>
            <Label className="text-xs font-semibold">Per-Category Labor Breakdown</Label>
            <p className="text-[10px] text-muted-foreground mt-0.5 mb-2">
              Optional — enter actual manhours by work category (e.g. pipe, structural,
              supports). Used for calibrating future estimates.
            </p>
            <CategoryEditor rows={categoryRows} onChange={setCategoryRows} />
          </div>

          {/* Tags / Notes */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <Label className="text-xs">Tags (comma-separated)</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={tags}
                onChange={(e) => setTags(e.target.value)}
              />
            </div>
            <div>
              <Label className="text-xs">Notes</Label>
              <Input
                className="h-8 text-xs mt-1"
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
              />
            </div>
          </div>
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={onClose}>
            Cancel
          </Button>
          <Button size="sm" onClick={handleSave} disabled={mutation.isPending}>
            {mutation.isPending ? "Saving…" : "Save Changes"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

// ---------------------------------------------------------------------------
// Variance cell — renders a right-aligned colored percentage
// ---------------------------------------------------------------------------

function VCell({ bid, actual }: { bid: number | null | undefined; actual: number | null | undefined }) {
  const pct = variancePct(bid, actual);
  return (
    <td className={`px-3 py-2 text-right tabular-nums text-xs ${varianceClass(pct)}`}>
      {fmtVariance(pct)}
    </td>
  );
}

// ---------------------------------------------------------------------------
// Row expansion: Bid-vs-Actual breakdown panel
// ---------------------------------------------------------------------------

function ExpandedRow({ project, colSpan }: { project: CompletedProject; colSpan: number }) {
  const laborPct = variancePct(project.bidLaborHours, project.totalManhours);
  const materialPct = variancePct(project.bidMaterialCost, project.materialCost);
  const totalPct = variancePct(project.bidTotalCost, project.totalCost);

  const categoryData = useMemo(() => {
    if (!project.actualLaborByCategoryJson) return [];
    try {
      const obj = JSON.parse(project.actualLaborByCategoryJson) as Record<string, number>;
      return Object.entries(obj);
    } catch {
      return [];
    }
  }, [project.actualLaborByCategoryJson]);

  const disciplinePath =
    project.discipline === "structural"
      ? "/structural"
      : project.discipline === "civil"
      ? "/civil"
      : "/mechanical";

  return (
    <tr className="bg-muted/30">
      <td colSpan={colSpan} className="px-4 py-3">
        <div className="space-y-3">
          {/* Bid vs Actual comparison */}
          <div>
            <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-2">
              Bid vs Actual Breakdown
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="text-muted-foreground text-[10px] uppercase tracking-wider">
                    <th className="text-left pb-1 pr-4 font-medium">Metric</th>
                    <th className="text-right pb-1 pr-4 font-medium tabular-nums">Bid</th>
                    <th className="text-right pb-1 pr-4 font-medium tabular-nums">Actual</th>
                    <th className="text-right pb-1 font-medium">Variance</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border/50">
                  <tr>
                    <td className="py-1 pr-4 text-muted-foreground">Total Labor MH</td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmtMH(project.bidLaborHours)}
                    </td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmtMH(project.totalManhours || undefined)}
                    </td>
                    <td className={`py-1 text-right tabular-nums font-mono ${varianceClass(laborPct)}`}>
                      {fmtVariance(laborPct)}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-1 pr-4 text-muted-foreground">Material $</td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmt$(project.bidMaterialCost)}
                    </td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmt$(project.materialCost || undefined)}
                    </td>
                    <td className={`py-1 text-right tabular-nums font-mono ${varianceClass(materialPct)}`}>
                      {fmtVariance(materialPct)}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-1 pr-4 text-muted-foreground">Total $</td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmt$(project.bidTotalCost)}
                    </td>
                    <td className="py-1 pr-4 text-right tabular-nums font-mono">
                      {fmt$(project.totalCost || undefined)}
                    </td>
                    <td className={`py-1 text-right tabular-nums font-mono ${varianceClass(totalPct)}`}>
                      {fmtVariance(totalPct)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Per-category breakdown */}
          {categoryData.length > 0 && (
            <div>
              <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-2">
                Actual Labor by Category
              </p>
              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2">
                {categoryData.map(([cat, hrs]) => (
                  <div key={cat} className="bg-background border border-border rounded px-2 py-1.5">
                    <p className="text-[10px] text-muted-foreground capitalize">{cat}</p>
                    <p className="text-xs font-semibold font-mono">{fmtMH(hrs)} MH</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Notes */}
          {project.notes && (
            <div>
              <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1">
                Notes
              </p>
              <p className="text-xs text-muted-foreground italic">{project.notes}</p>
            </div>
          )}

          {/* Source links */}
          {(project.sourceTakeoffId || project.sourceEstimateId) && (
            <div className="flex items-center gap-3 flex-wrap">
              {project.sourceTakeoffId && (
                <a
                  href={`#${disciplinePath}?project=${project.sourceTakeoffId}`}
                  className="inline-flex items-center gap-1 text-xs text-primary hover:underline"
                >
                  <ExternalLink size={11} />
                  View Source Takeoff
                </a>
              )}
              {project.sourceEstimateId && (
                <a
                  href={`#/estimating?id=${project.sourceEstimateId}`}
                  className="inline-flex items-center gap-1 text-xs text-primary hover:underline"
                >
                  <ExternalLink size={11} />
                  View Source Estimate
                </a>
              )}
            </div>
          )}
        </div>
      </td>
    </tr>
  );
}

// ---------------------------------------------------------------------------
// Main page component
// ---------------------------------------------------------------------------

export default function PerformancePage() {
  const [showAddDialog, setShowAddDialog] = useState(false);
  const [showImportDialog, setShowImportDialog] = useState(false);
  const [editProject, setEditProject] = useState<CompletedProject | null>(null);
  const [expandedId, setExpandedId] = useState<string | null>(null);

  // Fetch all completed projects
  const { data: projects = [], isLoading } = useQuery<CompletedProject[]>({
    queryKey: ["/api/completed-projects"],
  });

  // KPI computations — memoized so they don't recalculate on every render
  const kpis = useMemo(() => computeKpis(projects), [projects]);

  const toggleExpand = (id: string) => {
    setExpandedId((prev) => (prev === id ? null : id));
  };

  // Number of columns in the table (used for colSpan on expansion rows)
  const COL_COUNT = 12;

  return (
    <AppLayout subtitle="Performance — Bid vs Actual">
      <div className="p-5 space-y-5 max-w-screen-2xl mx-auto">

        {/* ---------------------------------------------------------------- */}
        {/* 1. Header band                                                   */}
        {/* ---------------------------------------------------------------- */}
        <div className="flex items-start justify-between flex-wrap gap-3">
          <div>
            <h2 className="text-lg font-bold text-foreground flex items-center gap-2">
              <TrendingUp size={20} className="text-primary" />
              Performance
            </h2>
            <p className="text-xs text-muted-foreground mt-0.5">
              Bid vs Actual across closed projects
            </p>
          </div>
          <div className="flex items-center gap-2">
            <Button
              size="sm"
              variant="outline"
              onClick={() => setShowImportDialog(true)}
              data-testid="btn-import-estimate"
            >
              <Upload size={13} className="mr-1.5" />
              Import from Estimate
            </Button>
            <Button
              size="sm"
              onClick={() => setShowAddDialog(true)}
              data-testid="btn-add-closed-project"
            >
              <Plus size={13} className="mr-1.5" />
              Add Closed Project
            </Button>
          </div>
        </div>

        {/* ---------------------------------------------------------------- */}
        {/* 2. KPI strip                                                     */}
        {/* ---------------------------------------------------------------- */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <Card className="border-card-border shadow-sm">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">
                Closed Projects
              </p>
              <p className="text-2xl font-bold text-foreground tabular-nums mt-0.5">
                {kpis.totalProjects}
              </p>
            </CardContent>
          </Card>

          <Card className="border-card-border shadow-sm">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">
                With Actuals
              </p>
              <p className="text-2xl font-bold text-foreground tabular-nums mt-0.5">
                {kpis.withActuals}
              </p>
              {kpis.totalProjects > 0 && (
                <p className="text-[10px] text-muted-foreground">
                  {Math.round((kpis.withActuals / kpis.totalProjects) * 100)}% of total
                </p>
              )}
            </CardContent>
          </Card>

          <Card className="border-card-border shadow-sm">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">
                Avg Labor Variance
              </p>
              <p
                className={`text-2xl font-bold tabular-nums mt-0.5 ${
                  kpis.avgLaborVariancePct == null
                    ? "text-muted-foreground"
                    : varianceClass(kpis.avgLaborVariancePct)
                }`}
              >
                {kpis.avgLaborVariancePct == null
                  ? "—"
                  : fmtVariance(kpis.avgLaborVariancePct)}
              </p>
              <p className="text-[10px] text-muted-foreground">vs bid MH</p>
            </CardContent>
          </Card>

          <Card className="border-card-border shadow-sm">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">
                Avg Material Variance
              </p>
              <p
                className={`text-2xl font-bold tabular-nums mt-0.5 ${
                  kpis.avgMaterialVariancePct == null
                    ? "text-muted-foreground"
                    : varianceClass(kpis.avgMaterialVariancePct)
                }`}
              >
                {kpis.avgMaterialVariancePct == null
                  ? "—"
                  : fmtVariance(kpis.avgMaterialVariancePct)}
              </p>
              <p className="text-[10px] text-muted-foreground">vs bid material $</p>
            </CardContent>
          </Card>
        </div>

        {/* ---------------------------------------------------------------- */}
        {/* 3. Projects table                                                */}
        {/* ---------------------------------------------------------------- */}
        <Card className="border-card-border shadow-sm">
          <CardHeader className="p-4 pb-0">
            <CardTitle className="text-sm">Closed Projects</CardTitle>
          </CardHeader>
          <CardContent className="p-0">
            {isLoading ? (
              <div className="p-8 text-sm text-muted-foreground text-center">
                Loading projects…
              </div>
            ) : projects.length === 0 ? (
              /* Empty state */
              <div className="p-10 text-center">
                <AlertCircle
                  size={32}
                  className="mx-auto text-muted-foreground/40 mb-3"
                />
                <p className="text-sm text-muted-foreground max-w-md mx-auto">
                  No closed projects yet. Click{" "}
                  <button
                    className="underline text-foreground hover:text-primary"
                    onClick={() => setShowAddDialog(true)}
                  >
                    Add Closed Project
                  </button>{" "}
                  to log one, or{" "}
                  <button
                    className="underline text-foreground hover:text-primary"
                    onClick={() => setShowImportDialog(true)}
                  >
                    Import from Estimate
                  </button>{" "}
                  to snapshot a bid you've already prepared.
                </p>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-border bg-muted/40">
                      {/* expand toggle */}
                      <th className="w-8 px-2 py-2" />
                      <th className="text-left px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap">
                        Project
                      </th>
                      <th className="text-left px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap">
                        Discipline
                      </th>
                      <th className="text-left px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap">
                        Client
                      </th>
                      <th className="text-left px-3 py-2 font-semibold text-muted-foreground">
                        Tags
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap tabular-nums">
                        Bid MH
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap tabular-nums">
                        Actual MH
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap">
                        Δ MH %
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap tabular-nums">
                        Bid $
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap tabular-nums">
                        Actual $
                      </th>
                      <th className="text-right px-3 py-2 font-semibold text-muted-foreground whitespace-nowrap">
                        Δ $ %
                      </th>
                      <th className="text-left px-3 py-2 font-semibold text-muted-foreground">
                        Status
                      </th>
                      {/* actions */}
                      <th className="w-16 px-2 py-2" />
                    </tr>
                  </thead>
                  <tbody>
                    {projects.map((p) => {
                      const expanded = expandedId === p.id;
                      const tags = (p.tags || "")
                        .split(",")
                        .map((t) => t.trim())
                        .filter(Boolean);
                      const status = getStatus(p);

                      return (
                        <>
                          <tr
                            key={p.id}
                            className={`border-b border-border cursor-pointer hover:bg-muted/40 transition-colors ${
                              expanded ? "bg-muted/20" : ""
                            }`}
                            onClick={() => toggleExpand(p.id)}
                          >
                            {/* expand icon */}
                            <td className="px-2 py-2 text-muted-foreground">
                              {expanded ? (
                                <ChevronUp size={13} />
                              ) : (
                                <ChevronDown size={13} />
                              )}
                            </td>
                            {/* name */}
                            <td className="px-3 py-2 max-w-[200px]">
                              <p className="font-medium text-foreground truncate">{p.name}</p>
                              {p.location && (
                                <p className="text-[10px] text-muted-foreground truncate">
                                  {p.location}
                                </p>
                              )}
                            </td>
                            {/* discipline */}
                            <td className="px-3 py-2 capitalize text-muted-foreground whitespace-nowrap">
                              {p.discipline ?? "—"}
                            </td>
                            {/* client */}
                            <td className="px-3 py-2 text-muted-foreground max-w-[120px]">
                              <span className="truncate block">{p.client ?? "—"}</span>
                            </td>
                            {/* tags */}
                            <td className="px-3 py-2 max-w-[160px]">
                              <div className="flex flex-wrap gap-1">
                                {tags.length === 0 ? (
                                  <span className="text-muted-foreground">—</span>
                                ) : (
                                  tags.slice(0, 3).map((tag) => (
                                    <Badge
                                      key={tag}
                                      variant="secondary"
                                      className="text-[9px] px-1 py-0"
                                    >
                                      {tag}
                                    </Badge>
                                  ))
                                )}
                                {tags.length > 3 && (
                                  <span className="text-[9px] text-muted-foreground">
                                    +{tags.length - 3}
                                  </span>
                                )}
                              </div>
                            </td>
                            {/* Bid MH */}
                            <td className="px-3 py-2 text-right tabular-nums font-mono text-muted-foreground whitespace-nowrap">
                              {fmtMH(p.bidLaborHours)}
                            </td>
                            {/* Actual MH */}
                            <td className="px-3 py-2 text-right tabular-nums font-mono whitespace-nowrap">
                              {fmtMH(p.totalManhours || undefined)}
                            </td>
                            {/* Δ MH % */}
                            <VCell bid={p.bidLaborHours} actual={p.totalManhours || undefined} />
                            {/* Bid $ */}
                            <td className="px-3 py-2 text-right tabular-nums font-mono text-muted-foreground whitespace-nowrap">
                              {fmt$(p.bidTotalCost)}
                            </td>
                            {/* Actual $ */}
                            <td className="px-3 py-2 text-right tabular-nums font-mono whitespace-nowrap">
                              {fmt$(p.totalCost || undefined)}
                            </td>
                            {/* Δ $ % */}
                            <VCell bid={p.bidTotalCost} actual={p.totalCost || undefined} />
                            {/* Status */}
                            <td className="px-3 py-2">
                              <StatusBadge status={status} />
                            </td>
                            {/* Edit action — stopPropagation so row expand doesn't fire */}
                            <td
                              className="px-2 py-2"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <Button
                                size="icon"
                                variant="ghost"
                                className="h-7 w-7"
                                title="Edit actuals"
                                onClick={() => setEditProject(p)}
                              >
                                <Edit2 size={13} />
                              </Button>
                            </td>
                          </tr>

                          {/* Inline expansion row */}
                          {expanded && (
                            <ExpandedRow
                              key={`${p.id}-expanded`}
                              project={p}
                              colSpan={COL_COUNT + 1}
                            />
                          )}
                        </>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* Dialogs                                                             */}
      {/* ------------------------------------------------------------------ */}
      <AddProjectDialog
        open={showAddDialog}
        onClose={() => setShowAddDialog(false)}
      />

      <ImportEstimateDialog
        open={showImportDialog}
        onClose={() => setShowImportDialog(false)}
      />

      {editProject && (
        <EditActualsDialog
          project={editProject}
          onClose={() => setEditProject(null)}
        />
      )}
    </AppLayout>
  );
}
