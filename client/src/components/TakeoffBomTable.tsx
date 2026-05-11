import { useState, useMemo, useRef, useEffect } from "react";
import type { TakeoffItem } from "@shared/schema";
import {
  Table, TableBody, TableCell, TableHead, TableHeader, TableRow,
} from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { CheckCircle2, Pencil, Settings2 } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { apiRequest } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";
import SheetDetailPanel from "./SheetDetailPanel";

interface TakeoffBomTableProps {
  items: TakeoffItem[];
  discipline: "mechanical" | "structural" | "civil";
  onItemUpdated?: () => void;
}

const CATEGORY_COLORS: Record<string, string> = {
  pipe: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  elbow: "bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300",
  tee: "bg-indigo-100 text-indigo-800 dark:bg-indigo-900/40 dark:text-indigo-300",
  reducer: "bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300",
  valve: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300",
  flange: "bg-orange-100 text-orange-800 dark:bg-orange-900/40 dark:text-orange-300",
  bolt: "bg-gray-100 text-gray-800 dark:bg-gray-800 dark:text-gray-300",
  gasket: "bg-pink-100 text-pink-800 dark:bg-pink-900/40 dark:text-pink-300",
  coupling: "bg-sky-100 text-sky-800 dark:bg-sky-900/40 dark:text-sky-300",
  cap: "bg-stone-100 text-stone-800 dark:bg-stone-800 dark:text-stone-300",
  union: "bg-lime-100 text-lime-800 dark:bg-lime-900/40 dark:text-lime-300",
  weld: "bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300",
  support: "bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300",
  wide_flange: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  hss_tube: "bg-cyan-100 text-cyan-800 dark:bg-cyan-900/40 dark:text-cyan-300",
  rebar: "bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300",
  footing: "bg-stone-100 text-stone-800 dark:bg-stone-800 dark:text-stone-300",
  slab: "bg-stone-100 text-stone-800 dark:bg-stone-800 dark:text-stone-300",
  storm_pipe: "bg-teal-100 text-teal-800 dark:bg-teal-900/40 dark:text-teal-300",
  sewer_pipe: "bg-yellow-100 text-yellow-800 dark:bg-yellow-900/40 dark:text-yellow-300",
  water_pipe: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  earthwork: "bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300",
  paving: "bg-slate-100 text-slate-800 dark:bg-slate-800 dark:text-slate-300",
};

const CONFIDENCE_DOT: Record<string, { color: string; label: string }> = {
  high: { color: "bg-green-500", label: "High confidence" },
  medium: { color: "bg-yellow-500", label: "Medium confidence — verify" },
  low: { color: "bg-red-500", label: "Low confidence — needs review" },
};

type EditableField = "size" | "quantity" | "description";

function InlineEditCell({
  value,
  field,
  itemId,
  onSave,
  className,
  children,
}: {
  value: string;
  field: EditableField;
  itemId: string;
  onSave: (itemId: string, field: string, value: string) => void;
  className?: string;
  children: React.ReactNode;
}) {
  const [editing, setEditing] = useState(false);
  const [editValue, setEditValue] = useState(value);
  const [saved, setSaved] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [editing]);

  useEffect(() => {
    if (saved) {
      const t = setTimeout(() => setSaved(false), 800);
      return () => clearTimeout(t);
    }
  }, [saved]);

  const commit = () => {
    setEditing(false);
    if (editValue !== value) {
      onSave(itemId, field, editValue);
      setSaved(true);
    }
  };

  if (editing) {
    return (
      <input
        ref={inputRef}
        className="bg-background border border-primary/60 rounded px-1 py-0.5 text-xs w-full outline-none focus:ring-1 focus:ring-primary/40"
        value={editValue}
        onChange={(e) => setEditValue(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") commit();
          if (e.key === "Escape") { setEditValue(value); setEditing(false); }
        }}
        onBlur={commit}
      />
    );
  }

  return (
    <span
      className={`group/edit cursor-pointer inline-flex items-center gap-1 ${saved ? "bg-green-100 dark:bg-green-900/30 rounded px-0.5 transition-colors" : ""} ${className || ""}`}
      onClick={() => { setEditValue(value); setEditing(true); }}
      title="Click to edit"
    >
      {children}
      <Pencil size={10} className="text-muted-foreground/0 group-hover/edit:text-muted-foreground/60 transition-opacity shrink-0" />
    </span>
  );
}

export default function TakeoffBomTable({ items, discipline, onItemUpdated }: TakeoffBomTableProps) {
  const { toast } = useToast();
  const [categoryFilter, setCategoryFilter] = useState<string>("all");
  const [sizeFilter, setSizeFilter] = useState<string>("all");
  const [sheetFilter, setSheetFilter] = useState<string>("all");
  const [confidenceFilter, setConfidenceFilter] = useState<"all" | "low" | "medium">("all");
  const [cloudFilter, setCloudFilter] = useState<"all" | "clouded" | "non-clouded">("all");
  const [verifyingAll, setVerifyingAll] = useState(false);
  const [sheetPanelPage, setSheetPanelPage] = useState<number | null>(null);

  const editable = !!onItemUpdated;

  const saveEdit = async (itemId: string, field: string, value: string) => {
    if (!onItemUpdated) return;
    const updates: Record<string, any> = { [field]: field === "quantity" ? (parseFloat(value) || 0) : value };
    try {
      await apiRequest("PATCH", `/api/takeoff-items/${itemId}`, updates);
      onItemUpdated();
    } catch {
      toast({ title: "Save failed", variant: "destructive" });
    }
  };

  // Toggle a per-line scope/inclusion flag on a takeoff item. Updates the
  // server immediately so the next downstream action (BOM/RFQ/Estimate) sees
  // the latest state. Also handles revisionClouded toggling.
  const toggleFlag = async (itemId: string, flag: string, value: boolean) => {
    if (!onItemUpdated) return;
    try {
      await apiRequest("PATCH", `/api/takeoff-items/${itemId}`, { [flag]: value });
      onItemUpdated();
    } catch {
      toast({ title: "Update failed", variant: "destructive" });
    }
  };

  const verifyItem = async (itemId: string) => {
    if (!onItemUpdated) return;
    try {
      await apiRequest("PATCH", `/api/takeoff-items/${itemId}`, { verified: true });
      onItemUpdated();
    } catch {
      toast({ title: "Verify failed", variant: "destructive" });
    }
  };

  // Derive filter options from data
  const categories = useMemo(() => {
    const cats: Record<string, number> = {};
    for (const item of items) {
      cats[item.category] = (cats[item.category] || 0) + 1;
    }
    return Object.entries(cats).sort((a, b) => b[1] - a[1]);
  }, [items]);

  const sizes = useMemo(() => {
    const s: Record<string, number> = {};
    for (const item of items) {
      const sz = item.size || "N/A";
      s[sz] = (s[sz] || 0) + 1;
    }
    return Object.entries(s).sort((a, b) => {
      const na = parseFloat(a[0]); const nb = parseFloat(b[0]);
      return isNaN(na) || isNaN(nb) ? a[0].localeCompare(b[0]) : na - nb;
    });
  }, [items]);

  const sheets = useMemo(() => {
    const s: Record<string, number> = {};
    for (const item of items) {
      const pg = (item as any).sourcePage;
      const sheet = pg != null ? `Sheet ${pg}` : (item.notes?.match(/Sheet\s+(\d+)/i)?.[0] || "Unknown");
      s[sheet] = (s[sheet] || 0) + 1;
    }
    return Object.entries(s).sort((a, b) => {
      const na = parseInt(a[0].replace(/\D/g, "")); const nb = parseInt(b[0].replace(/\D/g, ""));
      return (isNaN(na) ? 999 : na) - (isNaN(nb) ? 999 : nb);
    });
  }, [items]);

  if (!items.length) {
    return (
      <div className="flex flex-col items-center justify-center h-32 text-center">
        <p className="text-sm font-medium text-foreground">No items in this project</p>
        <p className="text-xs text-muted-foreground mt-1">Upload a PDF to extract items.</p>
      </div>
    );
  }

  const isMechanical = discipline === "mechanical";
  const isStructural = discipline === "structural";
  const isCivil = discipline === "civil";

  const lowCount = items.filter(i => (i as any).confidence === "low").length;
  const medCount = items.filter(i => (i as any).confidence === "medium").length;
  const cloudedCount = items.filter(i => i.revisionClouded).length;
  const hasCloudedItems = cloudedCount > 0;

  // Apply all filters
  let filteredItems = items;
  if (categoryFilter !== "all") filteredItems = filteredItems.filter(i => i.category === categoryFilter);
  if (sizeFilter !== "all") filteredItems = filteredItems.filter(i => (i.size || "N/A") === sizeFilter);
  if (sheetFilter !== "all") {
    filteredItems = filteredItems.filter(i => {
      const pg = (i as any).sourcePage;
      const sheet = pg != null ? `Sheet ${pg}` : (i.notes?.match(/Sheet\s+(\d+)/i)?.[0] || "Unknown");
      return sheet === sheetFilter;
    });
  }
  if (confidenceFilter !== "all") filteredItems = filteredItems.filter(i => (i as any).confidence === confidenceFilter);
  if (cloudFilter === "clouded") filteredItems = filteredItems.filter(i => i.revisionClouded);
  else if (cloudFilter === "non-clouded") filteredItems = filteredItems.filter(i => !i.revisionClouded);

  const activeFilterCount = [categoryFilter !== "all", sizeFilter !== "all", sheetFilter !== "all", confidenceFilter !== "all", cloudFilter !== "all"].filter(Boolean).length;

  const unverifiedVisible = filteredItems.filter(i => (i as any).confidence !== "high").length;

  const verifyAllVisible = async () => {
    if (!onItemUpdated) return;
    setVerifyingAll(true);
    const toVerify = filteredItems.filter(i => (i as any).confidence !== "high");
    try {
      for (const item of toVerify) {
        await apiRequest("PATCH", `/api/takeoff-items/${item.id}`, { verified: true });
      }
      onItemUpdated();
      toast({ title: `Verified ${toVerify.length} items` });
    } catch {
      toast({ title: "Bulk verify failed", variant: "destructive" });
    }
    setVerifyingAll(false);
  };

  return (
    <div className="space-y-2">
      {/* Filter bar */}
      <div className="flex items-start gap-2 flex-wrap border border-border bg-muted/30 rounded-lg px-3 py-2.5 shadow-sm">
        <span className="text-xs text-muted-foreground font-medium shrink-0">Filters:</span>

        {/* Category filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={categoryFilter}
          onChange={e => setCategoryFilter(e.target.value)}
          aria-label="Filter by category"
        >
          <option value="all">All Categories ({items.length})</option>
          {categories.map(([cat, count]) => (
            <option key={cat} value={cat}>{cat.replace(/_/g, " ")} ({count})</option>
          ))}
        </select>

        {/* Size filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={sizeFilter}
          onChange={e => setSizeFilter(e.target.value)}
          aria-label="Filter by size"
        >
          <option value="all">All Sizes</option>
          {sizes.map(([sz, count]) => (
            <option key={sz} value={sz}>{sz} ({count})</option>
          ))}
        </select>

        {/* Sheet filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={sheetFilter}
          onChange={e => setSheetFilter(e.target.value)}
          aria-label="Filter by sheet"
        >
          <option value="all">All Sheets</option>
          {sheets.map(([sheet, count]) => (
            <option key={sheet} value={sheet}>{sheet} ({count})</option>
          ))}
        </select>

        {/* Confidence filter */}
        {(lowCount > 0 || medCount > 0) && (
          <select
            className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
            value={confidenceFilter}
            onChange={e => setConfidenceFilter(e.target.value as any)}
            aria-label="Filter by confidence"
          >
            <option value="all">All Confidence</option>
            {lowCount > 0 && <option value="low">Low ({lowCount})</option>}
            {medCount > 0 && <option value="medium">Medium ({medCount})</option>}
          </select>
        )}

        {/* Cloud filter — always visible so the estimator can start clouding
            even when the extractor caught nothing. */}
        <button
          className={`h-7 text-xs px-2.5 rounded border font-medium transition-colors ${cloudFilter === "clouded" ? "bg-amber-100 text-amber-800 border-amber-300 dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700" : "bg-background text-muted-foreground border-border hover:bg-amber-50 dark:hover:bg-amber-950/20"}`}
          onClick={() => setCloudFilter(cloudFilter === "clouded" ? "all" : "clouded")}
          data-testid="btn-revisions-only"
        >
          Revisions Only ({cloudedCount})
        </button>
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={cloudFilter}
          onChange={e => setCloudFilter(e.target.value as any)}
          aria-label="Filter by revision status"
        >
          <option value="all">All Revisions</option>
          <option value="clouded">Clouded ({cloudedCount})</option>
          <option value="non-clouded">Non-Clouded ({items.length - cloudedCount})</option>
        </select>

        {/* Bulk-mark visible rows. Active filter narrows the scope so the
            estimator can sweep "all 8\" items in zone B" with one click. */}
        {editable && filteredItems.length > 0 && (
          <BulkCloudActions
            items={filteredItems}
            onBulkSet={async (clouded) => {
              try {
                await Promise.all(filteredItems.map(it =>
                  apiRequest("PATCH", `/api/takeoff-items/${it.id}`, { revisionClouded: clouded })
                ));
                onItemUpdated?.();
                toast({ title: `${clouded ? "Marked" : "Unmarked"} ${filteredItems.length} row${filteredItems.length === 1 ? "" : "s"}` });
              } catch {
                toast({ title: "Bulk update failed", variant: "destructive" });
              }
            }}
          />
        )}

        {activeFilterCount > 0 && (
          <button
            className="h-7 text-xs px-2 rounded border border-border hover:bg-accent text-muted-foreground"
            onClick={() => { setCategoryFilter("all"); setSizeFilter("all"); setSheetFilter("all"); setConfidenceFilter("all"); setCloudFilter("all"); }}
          >
            Clear ({activeFilterCount})
          </button>
        )}

        {/* Verify All Visible button */}
        {editable && unverifiedVisible > 0 && (
          <Button
            variant="outline"
            size="sm"
            className="h-7 text-xs text-green-700 border-green-300 hover:bg-green-50 dark:text-green-400 dark:border-green-700 dark:hover:bg-green-900/30"
            onClick={verifyAllVisible}
            disabled={verifyingAll}
          >
            <CheckCircle2 size={12} className="mr-1" />
            {verifyingAll ? "Verifying..." : `Verify All (${unverifiedVisible})`}
          </Button>
        )}

        <span className="text-xs text-muted-foreground ml-auto shrink-0">
          {filteredItems.length} of {items.length} items
        </span>
      </div>

      {/* Mobile card layout (< md) */}
      <div className="md:hidden space-y-2">
        {filteredItems.length === 0 ? (
          <div className="text-center text-xs text-muted-foreground py-8">No items match the selected filters.</div>
        ) : filteredItems.map((item) => {
          const conf = CONFIDENCE_DOT[(item as any).confidence || "high"] || CONFIDENCE_DOT.high;
          const isClouded = item.revisionClouded;
          const sourcePageNum = (item as any).sourcePage as number | undefined;
          const isVerified = (item as any).confidence === "high";
          const isManuallyVerified = !!(item as any).manuallyVerified;
          return (
            <div
              key={item.id}
              className={`rounded-md border border-border p-3 space-y-1.5 ${isClouded ? "bg-amber-50/50 dark:bg-amber-950/20 border-amber-200 dark:border-amber-800" : ""}`}
            >
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 min-w-0">
                  <span className={`inline-block w-2 h-2 rounded-full shrink-0 ${conf.color}`} />
                  <Badge variant="outline" className={`text-[10px] px-1.5 py-0 shrink-0 ${CATEGORY_COLORS[item.category] || ""}`}>
                    {item.category.replace(/_/g, " ")}
                  </Badge>
                  <span className="font-mono text-xs text-muted-foreground shrink-0">{item.size}</span>
                  {isClouded && <Badge variant="outline" className="text-[8px] px-1 py-0 bg-amber-100 text-amber-700 dark:bg-amber-900/40 dark:text-amber-400 shrink-0">REV</Badge>}
                </div>
                <div className="flex items-center gap-1.5 shrink-0">
                  {editable && !isVerified && (
                    <button className="text-muted-foreground hover:text-green-600 transition-colors" title="Mark as verified" onClick={() => verifyItem(item.id)}>
                      <CheckCircle2 size={14} />
                    </button>
                  )}
                  {isManuallyVerified && <CheckCircle2 size={14} className="text-green-500" />}
                  <span className="font-mono text-xs font-semibold">
                    {item.unit === "LF" || item.unit === "CY" || item.unit === "SF" || item.unit === "SY"
                      ? item.quantity.toFixed(2)
                      : item.quantity.toLocaleString()} {item.unit}
                  </span>
                </div>
              </div>
              <p className="text-xs leading-snug">{item.description}</p>
              <div className="flex items-center gap-2 text-[10px] text-muted-foreground flex-wrap">
                <span>#{item.lineNumber}</span>
                {sourcePageNum != null && <span className="underline decoration-dotted cursor-pointer hover:text-primary transition-colors" onClick={() => setSheetPanelPage(sourcePageNum)}>Pg {sourcePageNum}</span>}
                {isMechanical && item.schedule && <span>Sch: {item.schedule}</span>}
                {isMechanical && item.rating && <span>Rating: {item.rating}</span>}
                {isStructural && item.mark && <span>Mark: {item.mark}</span>}
                {isStructural && item.grade && <span>Grade: {item.grade}</span>}
                {isCivil && item.material && <span>{item.material}</span>}
                {item.notes && <span className="italic">{item.notes}</span>}
              </div>
            </div>
          );
        })}
      </div>

      {/* Desktop table layout (>= md) */}
      <div className="overflow-auto rounded-lg border border-border hidden md:block shadow-sm table-zebra table-hover-smooth">
        <TooltipProvider delayDuration={200}>
        <Table>
          <TableHeader>
            <TableRow className="bg-muted/50">
              <TableHead className="w-10 text-xs">#</TableHead>
              {editable && <TableHead className="w-6 text-xs px-1" title="Click the icon on a row to toggle revision-cloud status">Rev</TableHead>}
              {!editable && hasCloudedItems && <TableHead className="w-6 text-xs px-1">Rev</TableHead>}
              <TableHead className="w-6 text-xs"></TableHead>
              <TableHead className="text-xs">Category</TableHead>
              <TableHead className="text-xs">Size</TableHead>
              <TableHead className="text-xs max-w-xs">Description</TableHead>
              {isStructural && <TableHead className="text-xs">Mark</TableHead>}
              {isStructural && <TableHead className="text-xs">Grade</TableHead>}
              {isStructural && <TableHead className="text-xs">Weight</TableHead>}
              {isMechanical && <TableHead className="text-xs">Schedule</TableHead>}
              {isMechanical && <TableHead className="text-xs">Rating</TableHead>}
              {isCivil && <TableHead className="text-xs">Material</TableHead>}
              {isCivil && <TableHead className="text-xs">Depth</TableHead>}
              <TableHead className="text-xs text-right">Qty</TableHead>
              <TableHead className="text-xs">Unit</TableHead>
              <TableHead className="text-xs text-muted-foreground">Notes</TableHead>
              {editable && <TableHead className="w-8 text-xs"></TableHead>}
            </TableRow>
          </TableHeader>
          <TableBody>
            {filteredItems.length === 0 ? (
              <TableRow>
                <TableCell colSpan={20} className="text-center text-xs text-muted-foreground py-8">
                  No items match the selected filters.
                </TableCell>
              </TableRow>
            ) : filteredItems.map((item) => {
              const conf = CONFIDENCE_DOT[(item as any).confidence || "high"] || CONFIDENCE_DOT.high;
              const isClouded = item.revisionClouded;
              const isDedupCandidate = !!(item as any)._dedupCandidate;
              const sourcePageNum = (item as any).sourcePage as number | undefined;
              const isVerified = (item as any).confidence === "high";
              const isManuallyVerified = !!(item as any).manuallyVerified;
              // A row is excluded from some downstream view when any scope flag
              // is false. Default behavior (flag absent / true) keeps full opacity.
              const inBom = (item as any).includeInBom !== false;
              const inTakeoff = (item as any).includeInTakeoff !== false;
              const inEstimate = (item as any).includeInEstimate !== false;
              const anyExcluded = !inBom || !inTakeoff || !inEstimate;

              const qtyDisplay = item.unit === "LF" || item.unit === "CY" || item.unit === "SF" || item.unit === "SY"
                ? item.quantity.toFixed(2)
                : item.quantity.toLocaleString();

              return (
                <TableRow
                  key={item.id}
                  className={`group hover:bg-muted/30 text-xs ${isClouded ? "bg-amber-50/50 dark:bg-amber-950/20" : ""} ${isDedupCandidate || anyExcluded ? "opacity-50" : ""}`}
                  data-testid={`row-item-${item.id}`}
                >
                  <TableCell className="text-muted-foreground">{item.lineNumber}</TableCell>
                  {editable ? (
                    <TableCell className="px-1">
                      <Tooltip>
                        <TooltipTrigger asChild>
                          <button
                            onClick={() => toggleFlag(item.id, "revisionClouded", !isClouded)}
                            className={`inline-flex items-center justify-center w-5 h-5 rounded-full transition-colors ${
                              isClouded
                                ? "bg-amber-100 dark:bg-amber-900/40 hover:bg-amber-200 dark:hover:bg-amber-900/60"
                                : "opacity-30 hover:opacity-100 hover:bg-amber-50 dark:hover:bg-amber-950/40"
                            }`}
                            data-testid={`btn-toggle-cloud-${item.id}`}
                            aria-label={isClouded ? "Unmark as revision" : "Mark as revision"}
                          >
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" className={`w-3 h-3 ${isClouded ? "text-amber-600 dark:text-amber-400" : "text-muted-foreground"}`}>
                              <path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z" />
                            </svg>
                          </button>
                        </TooltipTrigger>
                        <TooltipContent side="right" className="text-[10px]">
                          {isClouded ? "Click to remove from revision" : "Click to mark as revision"}
                        </TooltipContent>
                      </Tooltip>
                    </TableCell>
                  ) : hasCloudedItems && (
                    <TableCell className="px-1">
                      {isClouded ? (
                        <Tooltip>
                          <TooltipTrigger asChild>
                            <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-100 dark:bg-amber-900/40">
                              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" className="w-3 h-3 text-amber-600 dark:text-amber-400">
                                <path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z" />
                              </svg>
                            </span>
                          </TooltipTrigger>
                          <TooltipContent side="right" className="text-[10px]">Inside revision cloud</TooltipContent>
                        </Tooltip>
                      ) : (
                        <span className="inline-block w-5" />
                      )}
                    </TableCell>
                  )}
                  <TableCell className="px-1">
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <span className={`inline-block w-2.5 h-2.5 rounded-full ${conf.color} ${isManuallyVerified ? "ring-2 ring-green-300 dark:ring-green-700" : ""}`} />
                      </TooltipTrigger>
                      <TooltipContent side="right" className="text-[10px]">
                        {conf.label}{isManuallyVerified ? " (manually verified)" : ""}{(item as any).confidenceNotes ? `: ${(item as any).confidenceNotes}` : ""}
                      </TooltipContent>
                    </Tooltip>
                  </TableCell>
                  <TableCell>
                    <Badge variant="outline" className={`text-[10px] px-1.5 py-0 ${CATEGORY_COLORS[item.category] || ""}`}>
                      {item.category.replace(/_/g, " ")}
                    </Badge>
                  </TableCell>
                  <TableCell className="font-mono text-xs">
                    {editable ? (
                      <InlineEditCell value={item.size} field="size" itemId={item.id} onSave={saveEdit}>
                        <span>{item.size}</span>
                      </InlineEditCell>
                    ) : item.size}
                  </TableCell>
                  <TableCell className="max-w-xs">
                    {editable ? (
                      <InlineEditCell value={item.description} field="description" itemId={item.id} onSave={saveEdit}>
                        <span className="line-clamp-2 text-xs">{item.description}</span>
                      </InlineEditCell>
                    ) : (
                      <span className="line-clamp-2 text-xs">{item.description}</span>
                    )}
                  </TableCell>
                  {isStructural && <TableCell className="font-mono text-xs">{item.mark || ""}</TableCell>}
                  {isStructural && <TableCell className="text-xs">{item.grade || ""}</TableCell>}
                  {isStructural && (
                    <TableCell className="text-xs text-right">
                      {item.weight ? item.weight.toLocaleString() + " lbs" : ""}
                    </TableCell>
                  )}
                  {isMechanical && <TableCell className="text-xs text-muted-foreground">{item.schedule || ""}</TableCell>}
                  {isMechanical && <TableCell className="text-xs text-muted-foreground">{item.rating || ""}</TableCell>}
                  {isCivil && <TableCell className="text-xs">{item.material || ""}</TableCell>}
                  {isCivil && <TableCell className="text-xs">{item.depth || ""}</TableCell>}
                  <TableCell className="text-right font-mono">
                    {editable ? (
                      <InlineEditCell value={String(item.quantity)} field="quantity" itemId={item.id} onSave={saveEdit}>
                        <span>{qtyDisplay}</span>
                      </InlineEditCell>
                    ) : qtyDisplay}
                  </TableCell>
                  <TableCell className="text-muted-foreground">{item.unit}</TableCell>
                  <TableCell className="text-xs text-muted-foreground whitespace-nowrap">
                    <span className="flex items-center gap-1">
                      {isDedupCandidate && (
                        <Tooltip>
                          <TooltipTrigger asChild>
                            <Badge variant="outline" className="text-[8px] px-1 py-0 bg-orange-50 text-orange-600 border-orange-200 dark:bg-orange-950/40 dark:text-orange-400 dark:border-orange-800 shrink-0">
                              dup?
                            </Badge>
                          </TooltipTrigger>
                          <TooltipContent side="top" className="text-[10px] max-w-xs">{(item as any).dedupNote || "Possible duplicate"}</TooltipContent>
                        </Tooltip>
                      )}
                      {sourcePageNum != null && (
                        <Badge
                          variant="outline"
                          className="text-[8px] px-1 py-0 bg-slate-50 text-slate-600 border-slate-200 dark:bg-slate-900/40 dark:text-slate-400 dark:border-slate-700 shrink-0 cursor-pointer underline decoration-dotted hover:text-primary hover:border-primary/50 transition-colors"
                          onClick={() => setSheetPanelPage(sourcePageNum)}
                        >
                          Pg {sourcePageNum}
                        </Badge>
                      )}
                      {item.notes || ""}
                    </span>
                  </TableCell>
                  {editable && (
                    <TableCell className="px-1">
                      <div className="flex items-center gap-1">
                        {!isVerified ? (
                          <Tooltip>
                            <TooltipTrigger asChild>
                              <button
                                className="text-muted-foreground/40 group-hover:text-muted-foreground hover:!text-green-600 transition-colors p-0.5"
                                onClick={() => verifyItem(item.id)}
                              >
                                <CheckCircle2 size={14} />
                              </button>
                            </TooltipTrigger>
                            <TooltipContent side="left" className="text-[10px]">Mark as verified</TooltipContent>
                          </Tooltip>
                        ) : isManuallyVerified ? (
                          <Tooltip>
                            <TooltipTrigger asChild>
                              <CheckCircle2 size={14} className="text-green-500" />
                            </TooltipTrigger>
                            <TooltipContent side="left" className="text-[10px]">Manually verified</TooltipContent>
                          </Tooltip>
                        ) : null}
                        <Popover>
                          <PopoverTrigger asChild>
                            <button
                              className={`text-muted-foreground/40 group-hover:text-muted-foreground hover:!text-primary transition-colors p-0.5 ${anyExcluded ? "!text-amber-600" : ""}`}
                              title="Scope flags & revision"
                              data-testid={`btn-scope-${item.id}`}
                            >
                              <Settings2 size={14} />
                            </button>
                          </PopoverTrigger>
                          <PopoverContent className="w-56 p-3 space-y-2" side="left">
                            <div>
                              <p className="text-[10px] font-semibold text-muted-foreground uppercase mb-1">Include in</p>
                              <div className="space-y-1">
                                {([
                                  { field: "includeInBom", label: "BOM / RFQ", checked: inBom },
                                  { field: "includeInTakeoff", label: "Takeoff", checked: inTakeoff },
                                  { field: "includeInEstimate", label: "Estimate", checked: inEstimate },
                                ] as const).map(({ field, label, checked }) => (
                                  <label key={field} className="flex items-center gap-2 text-[11px] cursor-pointer hover:bg-muted/40 rounded px-1 py-0.5">
                                    <input
                                      type="checkbox"
                                      className="h-3 w-3"
                                      checked={checked}
                                      onChange={e => toggleFlag(item.id, field, e.target.checked)}
                                      data-testid={`scope-${field}-${item.id}`}
                                    />
                                    <span className="flex-1">{label}</span>
                                    {!checked && <span className="text-[9px] text-amber-600 dark:text-amber-400">excluded</span>}
                                  </label>
                                ))}
                              </div>
                            </div>
                            <div className="border-t pt-2">
                              <label className="flex items-center gap-2 text-[11px] cursor-pointer hover:bg-muted/40 rounded px-1 py-0.5">
                                <input
                                  type="checkbox"
                                  className="h-3 w-3"
                                  checked={!!isClouded}
                                  onChange={e => toggleFlag(item.id, "revisionClouded", e.target.checked)}
                                />
                                <span className="flex-1">Inside revision cloud</span>
                              </label>
                              <p className="text-[9px] text-muted-foreground mt-1 leading-tight">Mark clouded items so they can be priced separately as a revision estimate.</p>
                            </div>
                          </PopoverContent>
                        </Popover>
                      </div>
                    </TableCell>
                  )}
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
        </TooltipProvider>
      </div>

      {/* Sheet Detail Side Panel */}
      {sheetPanelPage !== null && (
        <SheetDetailPanel
          pageNum={sheetPanelPage}
          items={items}
          onClose={() => setSheetPanelPage(null)}
        />
      )}
    </div>
  );
}

// Bulk-cloud action chip. Shows a small two-button cluster when at least one
// row is filtered (so the action's scope is obvious from the visible rows).
// Wired through the parent's filtered item list — estimator can filter to a
// size, sheet, or category and then sweep the whole subset with one click.
function BulkCloudActions({ items, onBulkSet }: { items: any[]; onBulkSet: (clouded: boolean) => void | Promise<void> }) {
  const cloudedInScope = items.filter((it: any) => it.revisionClouded).length;
  const allClouded = cloudedInScope === items.length;
  const noneClouded = cloudedInScope === 0;
  return (
    <div className="flex items-center gap-1">
      <button
        className="h-7 text-xs px-2 rounded border border-amber-300 text-amber-700 bg-amber-50 hover:bg-amber-100 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700 dark:hover:bg-amber-900/50 transition-colors disabled:opacity-40"
        disabled={allClouded}
        onClick={() => onBulkSet(true)}
        title={`Mark all ${items.length} visible row${items.length === 1 ? "" : "s"} as inside a revision cloud`}
        data-testid="btn-bulk-cloud"
      >
        Mark visible as revision ({items.length})
      </button>
      {!noneClouded && (
        <button
          className="h-7 text-xs px-2 rounded border border-border text-muted-foreground hover:bg-accent transition-colors"
          onClick={() => onBulkSet(false)}
          title={`Unmark all ${items.length} visible row${items.length === 1 ? "" : "s"} from revision cloud`}
          data-testid="btn-bulk-uncloud"
        >
          Clear revision
        </button>
      )}
    </div>
  );
}
