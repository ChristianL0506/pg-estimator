// Review Mode: walks the user through extracted items that need verification.
// Displays each item with a cropped page image, the extracted values, and any
// multi-pass voting disagreements. User can Accept, Edit, or Delete.

import { useMemo, useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";
import type { TakeoffItem, TakeoffProject } from "@shared/schema";
import { Check, X, Edit3, ChevronLeft, ChevronRight, AlertTriangle, CheckCircle2, HelpCircle } from "lucide-react";

interface ReviewModeProps {
  project: TakeoffProject;
}

type FilterMode = "needs_review" | "all" | "reviewed";

export function ReviewMode({ project }: ReviewModeProps) {
  const { toast } = useToast();
  const [filterMode, setFilterMode] = useState<FilterMode>("needs_review");
  const [currentIdx, setCurrentIdx] = useState(0);
  const [editing, setEditing] = useState(false);
  const [editValues, setEditValues] = useState<Partial<TakeoffItem>>({});

  // Compute the filtered list of items to review.
  // Priority: split > single > majority > unanimous-but-unreviewed.
  const filteredItems = useMemo(() => {
    const items = project.items.filter(i => !(i as any)._dedupCandidate);
    if (filterMode === "all") {
      // Sort: needs_review first
      return [...items].sort((a, b) => {
        const score = (it: any) => {
          if (it.votingStatus === "split") return 0;
          if (it.votingStatus === "single") return 1;
          if (it.votingStatus === "majority") return 2;
          return 3;
        };
        return score(a) - score(b);
      });
    }
    if (filterMode === "reviewed") {
      return items.filter(i => i.reviewStatus === "edited" || (i.reviewStatus === "accepted" && i.reviewedBy && i.reviewedBy !== "auto"));
    }
    // needs_review (default): items that haven't been manually reviewed
    return items.filter(i => {
      const rs = i.reviewStatus || "unreviewed";
      if (rs === "unreviewed") return true;
      if (rs === "rejected") return true;
      // Auto-accepted items don't need review
      return false;
    }).sort((a: any, b: any) => {
      const score = (it: any) => {
        if (it.votingStatus === "split") return 0;
        if (it.votingStatus === "single") return 1;
        if (it.votingStatus === "majority") return 2;
        if ((it._validationNotes || []).length > 0) return 2.5;
        return 3;
      };
      return score(a) - score(b);
    });
  }, [project.items, filterMode]);

  const currentItem = filteredItems[currentIdx];

  // Reset edit state when item changes
  useEffect(() => {
    setEditing(false);
    setEditValues({});
  }, [currentItem?.id]);

  // Stats for header
  const stats = useMemo(() => {
    const items = project.items.filter(i => !(i as any)._dedupCandidate);
    const reviewed = items.filter(i => {
      const rs = i.reviewStatus || "unreviewed";
      return rs === "accepted" || rs === "edited";
    }).length;
    const splits = items.filter((i: any) => i.votingStatus === "split").length;
    const singles = items.filter((i: any) => i.votingStatus === "single").length;
    const majorities = items.filter((i: any) => i.votingStatus === "majority").length;
    const unanimous = items.filter((i: any) => i.votingStatus === "unanimous").length;
    const noVoting = items.filter((i: any) => !i.votingStatus).length;
    return { total: items.length, reviewed, splits, singles, majorities, unanimous, noVoting };
  }, [project.items]);

  const refresh = () => {
    queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects", project.id] });
    queryClient.invalidateQueries({ queryKey: ["/api/takeoff-projects", project.id] });
    queryClient.invalidateQueries({ queryKey: ["/api/takeoff-projects"] });
  };

  async function handleAccept() {
    if (!currentItem) return;
    try {
      await apiRequest("PATCH", `/api/takeoff-items/${currentItem.id}`, {
        reviewStatus: "accepted",
        reviewedBy: "user",
        reviewedAt: new Date().toISOString(),
      });
      toast({ title: "Accepted" });
      refresh();
      goNext();
    } catch (e: any) {
      toast({ title: "Failed to accept", description: e?.message, variant: "destructive" });
    }
  }

  async function handleSaveEdit() {
    if (!currentItem) return;
    try {
      const update: any = {
        ...editValues,
        reviewStatus: "edited",
        reviewedBy: "user",
        reviewedAt: new Date().toISOString(),
      };
      await apiRequest("PATCH", `/api/takeoff-items/${currentItem.id}`, update);
      toast({ title: "Saved" });
      setEditing(false);
      refresh();
      goNext();
    } catch (e: any) {
      toast({ title: "Failed to save", description: e?.message, variant: "destructive" });
    }
  }

  async function handleDelete() {
    if (!currentItem) return;
    if (!window.confirm("Delete this item from the takeoff?")) return;
    try {
      await apiRequest("DELETE", `/api/takeoff-items/${currentItem.id}`);
      toast({ title: "Deleted" });
      refresh();
      // Don't increment currentIdx because the list will be one shorter
    } catch (e: any) {
      toast({ title: "Failed to delete", description: e?.message, variant: "destructive" });
    }
  }

  function goNext() {
    if (currentIdx < filteredItems.length - 1) {
      setCurrentIdx(currentIdx + 1);
    }
  }
  function goPrev() {
    if (currentIdx > 0) setCurrentIdx(currentIdx - 1);
  }

  // Apply a voting candidate as the current value
  function applyCandidate(passNum: number) {
    if (!currentItem) return;
    const vd: any = (currentItem as any).votingDetails;
    if (!vd?.passes) return;
    const pass = vd.passes.find((p: any) => p.passNum === passNum);
    if (!pass) return;
    setEditing(true);
    setEditValues({
      ...editValues,
      size: pass.size,
      quantity: typeof pass.quantity === "number" ? pass.quantity : parseFloat(String(pass.quantity)) || 0,
      category: pass.category,
    });
  }

  if (filteredItems.length === 0) {
    return (
      <div className="space-y-3">
        <ReviewHeader stats={stats} filterMode={filterMode} setFilterMode={setFilterMode} />
        <Card className="p-8 text-center">
          <CheckCircle2 className="h-12 w-12 mx-auto text-green-600 mb-3" />
          <p className="font-semibold">All clear</p>
          <p className="text-xs text-muted-foreground mt-1">
            {filterMode === "needs_review"
              ? "No items need review. Switch the filter to see all items."
              : "No items match this filter."}
          </p>
        </Card>
      </div>
    );
  }

  const vd: any = (currentItem as any).votingDetails;
  const validationNotes: string[] = (currentItem as any)._validationNotes || [];

  return (
    <div className="space-y-3">
      <ReviewHeader stats={stats} filterMode={filterMode} setFilterMode={setFilterMode} />

      {/* Progress bar */}
      <div className="flex items-center gap-2 text-xs text-muted-foreground">
        <span>Item {currentIdx + 1} of {filteredItems.length}</span>
        <div className="flex-1 h-1.5 bg-muted rounded-full overflow-hidden">
          <div
            className="h-full bg-primary transition-all"
            style={{ width: `${((currentIdx + 1) / filteredItems.length) * 100}%` }}
          />
        </div>
        <Button size="sm" variant="ghost" onClick={goPrev} disabled={currentIdx === 0}>
          <ChevronLeft className="h-4 w-4" />
        </Button>
        <Button size="sm" variant="ghost" onClick={goNext} disabled={currentIdx >= filteredItems.length - 1}>
          <ChevronRight className="h-4 w-4" />
        </Button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
        {/* Source page image */}
        <Card className="p-3 bg-muted/30">
          <p className="text-xs font-medium mb-2">
            Source: page {currentItem.sourcePage ?? "?"}
            {(currentItem as any).drawingNumber && (
              <span className="text-muted-foreground ml-1">· {(currentItem as any).drawingNumber}</span>
            )}
          </p>
          {currentItem.sourcePage ? (
            <img
              src={`/api/takeoff/projects/${project.id}/page/${currentItem.sourcePage}`}
              alt={`Page ${currentItem.sourcePage}`}
              className="w-full max-h-[500px] object-contain border rounded"
              onError={e => { (e.target as HTMLImageElement).style.display = "none"; }}
            />
          ) : (
            <div className="flex items-center justify-center h-32 text-xs text-muted-foreground">
              No source page
            </div>
          )}
        </Card>

        {/* Item details */}
        <Card className="p-3 space-y-3">
          <div className="flex items-start justify-between gap-2">
            <div className="flex-1 min-w-0">
              <p className="text-xs text-muted-foreground">Line #{currentItem.lineNumber}</p>
              <p className="text-sm font-medium break-words">{currentItem.description}</p>
            </div>
            <ConfidenceBadge item={currentItem} />
          </div>

          {/* Voting disagreement panel */}
          {vd?.passes && vd.passes.length > 1 && (
            <div className="border border-amber-300 dark:border-amber-700 bg-amber-50 dark:bg-amber-900/30 rounded p-2">
              <p className="text-xs font-semibold text-amber-900 dark:text-amber-200 mb-1.5 flex items-center gap-1">
                <AlertTriangle className="h-3 w-3" />
                3-pass voting: {vd.agreements?.size}/{vd.passes.length} agreed on size, {vd.agreements?.quantity}/{vd.passes.length} on qty
              </p>
              <div className="space-y-1">
                {vd.passes.map((p: any) => (
                  <div key={p.passNum} className="flex items-center gap-1.5 text-xs">
                    <span className="font-mono w-6 text-muted-foreground">P{p.passNum}:</span>
                    <span className="font-medium">{p.size}</span>
                    <span className="text-muted-foreground">·</span>
                    <span>qty {String(p.quantity)}</span>
                    <span className="text-muted-foreground">·</span>
                    <span className="text-muted-foreground">{p.category}</span>
                    <Button size="sm" variant="ghost" className="ml-auto h-6 px-2 text-xs" onClick={() => applyCandidate(p.passNum)}>
                      Use
                    </Button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Validation notes */}
          {validationNotes.length > 0 && (
            <div className="border border-orange-300 dark:border-orange-700 bg-orange-50 dark:bg-orange-900/30 rounded p-2 text-xs">
              <p className="font-semibold text-orange-900 dark:text-orange-200 mb-1">Flags:</p>
              <ul className="list-disc list-inside space-y-0.5 text-orange-800 dark:text-orange-300">
                {validationNotes.map((n, i) => <li key={i}>{n}</li>)}
              </ul>
            </div>
          )}

          {/* Editable fields */}
          <div className="grid grid-cols-2 gap-2">
            <FieldRow label="Size" editing={editing}
              value={editing ? (editValues.size ?? currentItem.size) : currentItem.size}
              onChange={v => setEditValues({ ...editValues, size: v as string })}
            />
            <FieldRow label="Qty" editing={editing}
              type="number"
              value={editing ? (editValues.quantity ?? currentItem.quantity) : currentItem.quantity}
              onChange={v => setEditValues({ ...editValues, quantity: parseFloat(String(v)) || 0 })}
            />
            <FieldRow label="Unit" editing={editing}
              value={editing ? (editValues.unit ?? currentItem.unit) : currentItem.unit}
              onChange={v => setEditValues({ ...editValues, unit: v as string })}
            />
            <FieldRow label="Category" editing={editing}
              value={editing ? (editValues.category ?? currentItem.category) : currentItem.category}
              onChange={v => setEditValues({ ...editValues, category: v as string })}
            />
          </div>

          {editing && (
            <FieldRow label="Description" editing={true} fullWidth
              value={editValues.description ?? currentItem.description}
              onChange={v => setEditValues({ ...editValues, description: v as string })}
            />
          )}

          <div className="flex items-center gap-2 pt-2 border-t">
            {editing ? (
              <>
                <Button size="sm" onClick={handleSaveEdit} className="bg-blue-600 hover:bg-blue-700">
                  <Check className="h-3.5 w-3.5 mr-1.5" /> Save & Next
                </Button>
                <Button size="sm" variant="outline" onClick={() => { setEditing(false); setEditValues({}); }}>
                  Cancel
                </Button>
              </>
            ) : (
              <>
                <Button size="sm" onClick={handleAccept} className="bg-green-600 hover:bg-green-700">
                  <Check className="h-3.5 w-3.5 mr-1.5" /> Accept
                </Button>
                <Button size="sm" variant="outline" onClick={() => setEditing(true)}>
                  <Edit3 className="h-3.5 w-3.5 mr-1.5" /> Edit
                </Button>
                <Button size="sm" variant="outline" className="text-red-700 border-red-300 hover:bg-red-50 dark:text-red-400 dark:border-red-800 dark:hover:bg-red-900/30" onClick={handleDelete}>
                  <X className="h-3.5 w-3.5 mr-1.5" /> Delete
                </Button>
                <span className="ml-auto text-xs text-muted-foreground">
                  {currentItem.reviewStatus === "accepted" && currentItem.reviewedBy === "auto" && "Auto-accepted (unanimous voting)"}
                  {currentItem.reviewStatus === "accepted" && currentItem.reviewedBy && currentItem.reviewedBy !== "auto" && "Already reviewed"}
                </span>
              </>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
}

function ReviewHeader({ stats, filterMode, setFilterMode }: { stats: any; filterMode: FilterMode; setFilterMode: (m: FilterMode) => void }) {
  const reviewedPct = stats.total > 0 ? Math.round((stats.reviewed / stats.total) * 100) : 0;
  return (
    <div className="space-y-2">
      <div className="flex items-center gap-2 flex-wrap">
        <p className="text-sm font-medium">Review Mode</p>
        <Badge variant="outline" className="text-xs">
          {stats.reviewed} of {stats.total} reviewed ({reviewedPct}%)
        </Badge>
        {stats.splits > 0 && <Badge className="bg-red-100 text-red-800 border-red-300 text-xs">{stats.splits} split</Badge>}
        {stats.singles > 0 && <Badge className="bg-orange-100 text-orange-800 border-orange-300 text-xs">{stats.singles} single-pass</Badge>}
        {stats.majorities > 0 && <Badge className="bg-yellow-100 text-yellow-800 border-yellow-300 text-xs">{stats.majorities} majority</Badge>}
        {stats.unanimous > 0 && <Badge className="bg-green-100 text-green-800 border-green-300 text-xs">{stats.unanimous} unanimous</Badge>}
        {stats.noVoting > 0 && <Badge variant="outline" className="text-xs"><HelpCircle className="h-3 w-3 mr-0.5" />{stats.noVoting} no voting</Badge>}
      </div>
      <div className="flex items-center gap-1">
        <Button size="sm" variant={filterMode === "needs_review" ? "default" : "outline"} onClick={() => setFilterMode("needs_review")} className="text-xs h-7">
          Needs Review
        </Button>
        <Button size="sm" variant={filterMode === "all" ? "default" : "outline"} onClick={() => setFilterMode("all")} className="text-xs h-7">
          All
        </Button>
        <Button size="sm" variant={filterMode === "reviewed" ? "default" : "outline"} onClick={() => setFilterMode("reviewed")} className="text-xs h-7">
          Reviewed
        </Button>
      </div>
    </div>
  );
}

function ConfidenceBadge({ item }: { item: any }) {
  const status = item.votingStatus;
  if (status === "unanimous") return <Badge className="bg-green-100 text-green-800 border-green-300">Unanimous</Badge>;
  if (status === "majority") return <Badge className="bg-yellow-100 text-yellow-800 border-yellow-300">Majority</Badge>;
  if (status === "split") return <Badge className="bg-red-100 text-red-800 border-red-300">Split</Badge>;
  if (status === "single") return <Badge className="bg-orange-100 text-orange-800 border-orange-300">Single Pass</Badge>;
  if (item.confidence === "high") return <Badge variant="outline">High</Badge>;
  if (item.confidence === "medium") return <Badge variant="outline">Medium</Badge>;
  if (item.confidence === "low") return <Badge variant="outline">Low</Badge>;
  return null;
}

function FieldRow({ label, value, editing, type = "text", onChange, fullWidth }: { label: string; value: any; editing: boolean; type?: string; onChange?: (v: any) => void; fullWidth?: boolean }) {
  return (
    <div className={fullWidth ? "col-span-2" : ""}>
      <p className="text-[10px] uppercase tracking-wide text-muted-foreground mb-0.5">{label}</p>
      {editing && onChange ? (
        <Input
          type={type}
          value={value ?? ""}
          onChange={e => onChange(e.target.value)}
          className="h-8 text-sm"
        />
      ) : (
        <p className="text-sm font-medium font-mono break-all">{String(value ?? "")}</p>
      )}
    </div>
  );
}
