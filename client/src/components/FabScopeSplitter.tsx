import { useState, useMemo } from "react";
import { Factory, Wrench, HardHat, ChevronDown, ChevronRight, Search, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Input } from "@/components/ui/input";

// ============================================================
// Types
// ============================================================

interface FabScopeSplitterProps {
  items: any[];
}

interface WeldCounts {
  buttWelds: number;
  socketWelds: number;
  boltUps: number;
  threaded: number;
}

interface DrawingInfo {
  page: number;
  label: string;
  itemCount: number;
  pipeFootage: number;
  fittingCount: number;
}

interface ScopeBreakdown {
  items: any[];
  welds: WeldCounts;
  pipeFootage: number;
  fittingCount: number;
  materialSummary: Map<string, number>;
}

// ============================================================
// Weld inference (same rules as ConnectionsSummary)
// ============================================================

function inferWelds(item: any): WeldCounts {
  const cat = (item.category || "").toLowerCase();
  const desc = (item.description || "").toLowerCase();
  const qty = item.quantity ?? 0;
  const isThreaded = desc.includes("threaded") || desc.includes("npt") || desc.includes("screw");
  const isSW = desc.includes("socket") || desc.includes(" sw ") || desc.includes("sw-") || desc.includes(",sw,");
  const isFlanged = desc.includes("flanged") || desc.includes("flg") || desc.includes("rf ") || desc.includes("raised face");

  // === PIPE LENGTH WELDS ===
  // Pipe is purchased in 40' standard lengths. Every 40' of run requires a field weld.
  // 160' run = 3 welds (at 40, 80, 120 ft)
  if (cat === "pipe") {
    if (qty >= 40) {
      const pipeJointWelds = Math.floor(qty / 40);
      return { buttWelds: pipeJointWelds, socketWelds: 0, boltUps: 0, threaded: 0 };
    }
    return { buttWelds: 0, socketWelds: 0, boltUps: 0, threaded: 0 };
  }

  if (isThreaded) return { buttWelds: 0, socketWelds: 0, boltUps: 0, threaded: qty * 2 };

  const r: WeldCounts = { buttWelds: 0, socketWelds: 0, boltUps: 0, threaded: 0 };

  switch (cat) {
    case "elbow": case "ell":
      if (isSW) r.socketWelds = 2 * qty; else r.buttWelds = 2 * qty; break;
    case "tee":
      if (isSW) r.socketWelds = 3 * qty; else r.buttWelds = 3 * qty; break;
    case "reducer": case "reducing":
      if (isSW) r.socketWelds = 2 * qty; else r.buttWelds = 2 * qty; break;
    case "cap":
      r.buttWelds = 1 * qty; break;
    case "coupling": case "union":
      // Threaded already handled above (returns 0 welds). SW couplings: 2 SW.
      r.socketWelds = 2 * qty; break;
    case "valve":
      // ONLY socket-weld valves generate welds. All other valves (flanged,
      // threaded, butterfly, butt-weld end) connect via the surrounding flanges.
      if (isSW) r.socketWelds = 2 * qty;
      // Other valve types: 0 welds, no entry
      break;
    case "flange":
      r.buttWelds = 1 * qty; r.boltUps = Math.ceil(qty / 2); break;
    case "gasket":
    case "bolt":
      // Gaskets and stud bolts are supplied hardware, not connections.
      // The flange they pair with is what counts as the bolt-up.
      break;
    case "sockolet":
      // Sockolet: 2 SW (header bore + branch)
      r.socketWelds = 2 * qty; break;
    case "weldolet": case "olet":
      // Weldolet: 1 BW to header
      r.buttWelds = 1 * qty; break;
    default:
      // Check description for olet types
      if (desc.includes("sockolet")) r.socketWelds = 2 * qty;
      else if (desc.includes("weldolet")) r.buttWelds = 1 * qty;
      else if (desc.includes("threadolet")) r.buttWelds = 1 * qty;
      break;
  }
  return r;
}

function addWelds(a: WeldCounts, b: WeldCounts): WeldCounts {
  return {
    buttWelds: a.buttWelds + b.buttWelds,
    socketWelds: a.socketWelds + b.socketWelds,
    boltUps: a.boltUps + b.boltUps,
    threaded: a.threaded + b.threaded,
  };
}

function totalWelds(w: WeldCounts): number {
  return w.buttWelds + w.socketWelds + w.boltUps + w.threaded;
}

const ZERO_WELDS: WeldCounts = { buttWelds: 0, socketWelds: 0, boltUps: 0, threaded: 0 };

function isFitting(cat: string): boolean {
  const c = cat.toLowerCase();
  return ["elbow", "ell", "tee", "reducer", "reducing", "cap", "coupling", "union", "valve", "flange", "sockolet", "weldolet", "olet", "gasket", "bolt"].includes(c);
}

function isPipe(cat: string): boolean {
  return cat.toLowerCase() === "pipe";
}

// Basic MH estimate per weld (rough)
const WELD_MH: Record<string, number> = {
  '0.5"': 1.2, '0.75"': 1.4, '1"': 1.8, '1.5"': 2.4, '2"': 3.0,
  '3"': 4.68, '4"': 5.56, '6"': 7.5, '8"': 10.0, '10"': 13.0, '12"': 16.0,
};

function getWeldMH(size: string): number {
  const s = (size || "").replace(/\s/g, "");
  if (WELD_MH[s]) return WELD_MH[s];
  const num = parseFloat(s);
  if (isNaN(num)) return 3.5;
  if (num <= 1) return 1.8;
  if (num <= 2) return 3.0;
  if (num <= 4) return 5.0;
  if (num <= 8) return 9.0;
  return 14.0;
}

function buildScope(items: any[]): ScopeBreakdown {
  let welds = { ...ZERO_WELDS };
  let pipeFootage = 0;
  let fittingCount = 0;
  const materialSummary = new Map<string, number>();

  for (const item of items) {
    const w = inferWelds(item);
    welds = addWelds(welds, w);
    const cat = (item.category || "").toLowerCase();
    if (isPipe(cat)) pipeFootage += item.quantity ?? 0;
    if (isFitting(cat)) fittingCount += item.quantity ?? 0;
    const matKey = `${cat} - ${item.size || "N/A"}`;
    materialSummary.set(matKey, (materialSummary.get(matKey) || 0) + (item.quantity ?? 0));
  }

  return { items, welds, pipeFootage, fittingCount, materialSummary };
}

// ============================================================
// Scope Detail Table
// ============================================================

function ScopeDetail({ scope }: { scope: ScopeBreakdown }) {
  if (scope.items.length === 0) return <p className="text-xs text-muted-foreground italic">No items</p>;

  return (
    <div className="max-h-64 overflow-auto">
      <table className="w-full text-xs">
        <thead>
          <tr className="border-b border-border text-left text-muted-foreground">
            <th className="py-1.5 pr-2">Category</th>
            <th className="py-1.5 pr-2">Size</th>
            <th className="py-1.5 pr-2">Description</th>
            <th className="py-1.5 pr-2 text-right">Qty</th>
            <th className="py-1.5 pr-2 text-right">Welds</th>
            <th className="py-1.5">Location</th>
          </tr>
        </thead>
        <tbody>
          {scope.items.map((item, i) => {
            const w = inferWelds(item);
            const tw = totalWelds(w);
            const loc = item.installLocation || ((item.notes || "").toLowerCase().includes("field") ? "field" : "shop");
            return (
              <tr key={item.id || i} className="border-b border-border/40 hover:bg-muted/30">
                <td className="py-1 pr-2 capitalize">{item.category}</td>
                <td className="py-1 pr-2 font-mono">{item.size || "-"}</td>
                <td className="py-1 pr-2 truncate max-w-[200px]">{item.description}</td>
                <td className="py-1 pr-2 text-right font-mono">{item.quantity ?? 0}</td>
                <td className="py-1 pr-2 text-right font-mono">{tw > 0 ? tw : "-"}</td>
                <td className="py-1">
                  <Badge variant="outline" className={`text-[9px] ${loc === "field" ? "text-orange-600 border-orange-300 dark:text-orange-400" : "text-green-600 border-green-300 dark:text-green-400"}`}>
                    {loc}
                  </Badge>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ============================================================
// Component
// ============================================================

export default function FabScopeSplitter({ items }: FabScopeSplitterProps) {
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set());
  const [searchFilter, setSearchFilter] = useState("");
  const [expandedCards, setExpandedCards] = useState<Set<string>>(new Set());
  const [showRevisionsOnly, setShowRevisionsOnly] = useState(false);

  const cloudedCount = items.filter(i => i.revisionClouded).length;
  const hasCloudedItems = cloudedCount > 0;
  const activeItems = showRevisionsOnly ? items.filter(i => i.revisionClouded) : items;

  // Build drawing list from unique source pages
  const drawings = useMemo((): DrawingInfo[] => {
    const pageMap = new Map<number, { items: any[] }>();
    for (const item of activeItems) {
      const page = item.sourcePage ?? 0;
      if (!pageMap.has(page)) pageMap.set(page, { items: [] });
      pageMap.get(page)!.items.push(item);
    }

    return Array.from(pageMap.entries())
      .map(([page, data]) => {
        const pipeFootage = data.items
          .filter((i: any) => isPipe((i.category || "").toLowerCase()))
          .reduce((s: number, i: any) => s + (i.quantity ?? 0), 0);
        const fittingCount = data.items
          .filter((i: any) => isFitting((i.category || "").toLowerCase()))
          .reduce((s: number, i: any) => s + (i.quantity ?? 0), 0);

        // Use ISO drawing number from title block if available, fall back to page number
        let label = `Page ${page || "?"}`;
        for (const it of data.items) {
          if (it.drawingNumber) {
            label = it.drawingNumber;
            break;
          }
          const notes = it.notes || "";
          const match = notes.match(/\|\s*([^(]+?)\s*\(/) || notes.match(/Sheet\s+(\d+)/i);
          if (match) { label = match[1]?.trim() || match[0]; break; }
        }

        return { page, label, itemCount: data.items.length, pipeFootage, fittingCount };
      })
      .sort((a, b) => a.page - b.page);
  }, [activeItems]);

  // Filter drawings by search
  const filteredDrawings = useMemo(() => {
    if (!searchFilter.trim()) return drawings;
    const q = searchFilter.toLowerCase();
    return drawings.filter(d => d.label.toLowerCase().includes(q) || String(d.page).includes(q));
  }, [drawings, searchFilter]);

  // Scope calculation
  const { subShopScope, yourFieldScope, yourFullScope, totalProjectWelds } = useMemo(() => {
    const subPages = selectedPages;

    const subShopItems: any[] = [];
    const yourFieldItems: any[] = [];
    const yourFullItems: any[] = [];

    for (const item of activeItems) {
      const page = item.sourcePage ?? 0;
      const loc = item.installLocation || ((item.notes || "").toLowerCase().includes("field") ? "field" : "shop");
      const isSubbedPage = subPages.has(page);

      if (isSubbedPage) {
        if (loc === "shop") subShopItems.push(item);
        else yourFieldItems.push(item);
      } else {
        yourFullItems.push(item);
      }
    }

    const subShopScope = buildScope(subShopItems);
    const yourFieldScope = buildScope(yourFieldItems);
    const yourFullScope = buildScope(yourFullItems);

    let totalProjectWelds = { ...ZERO_WELDS };
    for (const item of activeItems) totalProjectWelds = addWelds(totalProjectWelds, inferWelds(item));

    return { subShopScope, yourFieldScope, yourFullScope, totalProjectWelds };
  }, [activeItems, selectedPages]);

  // MH estimate
  const estimateMH = useMemo(() => {
    let mh = 0;
    const yourItems = [...yourFieldScope.items, ...yourFullScope.items];
    for (const item of yourItems) {
      const w = inferWelds(item);
      const tw = totalWelds(w);
      if (tw > 0) mh += tw * getWeldMH(item.size || "");
    }
    return Math.round(mh);
  }, [yourFieldScope, yourFullScope]);

  const yourTotalWelds = totalWelds(yourFieldScope.welds) + totalWelds(yourFullScope.welds);
  const subTotalWelds = totalWelds(subShopScope.welds);
  const projectTotalWelds = totalWelds(totalProjectWelds);
  const yourPct = projectTotalWelds > 0 ? Math.round((yourTotalWelds / projectTotalWelds) * 100) : 0;
  const subPct = projectTotalWelds > 0 ? Math.round((subTotalWelds / projectTotalWelds) * 100) : 0;

  function togglePage(page: number) {
    setSelectedPages(prev => {
      const next = new Set(prev);
      if (next.has(page)) next.delete(page); else next.add(page);
      return next;
    });
  }

  function selectAll() { setSelectedPages(new Set(drawings.map(d => d.page))); }
  function deselectAll() { setSelectedPages(new Set()); }

  function toggleCard(key: string) {
    setExpandedCards(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }

  // Export scope split
  async function handleExport() {
    try {
      const body = {
        subShop: subShopScope.items.map(i => ({
          category: i.category, size: i.size, description: i.description,
          quantity: i.quantity, unit: i.unit, welds: totalWelds(inferWelds(i)),
          location: i.installLocation || "shop", sourcePage: i.sourcePage,
        })),
        yourField: yourFieldScope.items.map(i => ({
          category: i.category, size: i.size, description: i.description,
          quantity: i.quantity, unit: i.unit, welds: totalWelds(inferWelds(i)),
          location: i.installLocation || "field", sourcePage: i.sourcePage,
        })),
        yourFull: yourFullScope.items.map(i => ({
          category: i.category, size: i.size, description: i.description,
          quantity: i.quantity, unit: i.unit, welds: totalWelds(inferWelds(i)),
          location: i.installLocation || ((i.notes || "").toLowerCase().includes("field") ? "field" : "shop"),
          sourcePage: i.sourcePage,
        })),
        summary: {
          totalWelds: projectTotalWelds,
          subWelds: subTotalWelds,
          yourWelds: yourTotalWelds,
          yourMH: estimateMH,
        },
      };
      const res = await fetch("/api/export-scope-split", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "Scope Split.xlsx";
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (e) { console.error("Scope split export failed", e); }
  }

  if (items.length === 0) {
    return <p className="text-sm text-muted-foreground py-8 text-center">No items to analyze. Upload a takeoff first.</p>;
  }

  return (
    <div className="space-y-4">
      {/* Revision filter toggle */}
      {hasCloudedItems && (
        <div className="flex items-center gap-2">
          <button
            className={`h-7 text-xs px-2.5 rounded border font-medium transition-colors ${showRevisionsOnly ? "bg-amber-100 text-amber-800 border-amber-300 dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700" : "bg-background text-muted-foreground border-border hover:bg-amber-50 dark:hover:bg-amber-950/20"}`}
            onClick={() => setShowRevisionsOnly(!showRevisionsOnly)}
          >
            Revisions Only ({cloudedCount})
          </button>
          {showRevisionsOnly && (
            <span className="text-xs text-amber-600 dark:text-amber-400">Showing only revision-clouded items</span>
          )}
        </div>
      )}

      {/* Step 1: Drawing selection */}
      <Card className="border-card-border">
        <CardContent className="p-4 space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground flex items-center gap-1.5">
              <Factory size={13} /> Step 1: Select Drawings for Sub-Fab
            </h3>
            <div className="flex items-center gap-2">
              <Button variant="ghost" size="sm" className="h-6 text-[10px]" onClick={selectAll}>Select All</Button>
              <Button variant="ghost" size="sm" className="h-6 text-[10px]" onClick={deselectAll}>Deselect All</Button>
              <Badge variant="outline" className="text-[10px]">{selectedPages.size} / {drawings.length} selected</Badge>
            </div>
          </div>

          <div className="relative">
            <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground" />
            <Input
              placeholder="Search drawings..."
              value={searchFilter}
              onChange={e => setSearchFilter(e.target.value)}
              className="h-8 text-xs pl-8"
            />
          </div>

          <div className="max-h-52 overflow-auto space-y-0.5">
            {filteredDrawings.map(d => (
              <label
                key={d.page}
                className={`flex items-center gap-3 px-3 py-1.5 rounded cursor-pointer transition-colors hover:bg-muted/50 ${selectedPages.has(d.page) ? "bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-800" : ""}`}
              >
                <input
                  type="checkbox"
                  checked={selectedPages.has(d.page)}
                  onChange={() => togglePage(d.page)}
                  className="accent-green-600"
                />
                <span className="text-xs font-medium flex-1">{d.label}</span>
                <span className="text-[10px] text-muted-foreground">{d.itemCount} items</span>
                <span className="text-[10px] text-muted-foreground">{d.pipeFootage.toFixed(1)} ft pipe</span>
                <span className="text-[10px] text-muted-foreground">{d.fittingCount} fittings</span>
              </label>
            ))}
            {filteredDrawings.length === 0 && (
              <p className="text-xs text-muted-foreground text-center py-4">No drawings match filter</p>
            )}
          </div>
        </CardContent>
      </Card>

      {/* Step 2: Scope breakdown */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
        {/* Card 1: Sub Shop Fab */}
        <Card className="border-green-200 dark:border-green-800 bg-green-50/30 dark:bg-green-900/10">
          <CardContent className="p-4 space-y-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-green-100 dark:bg-green-900/40 flex items-center justify-center">
                <Factory size={14} className="text-green-600 dark:text-green-400" />
              </div>
              <div className="flex-1">
                <h4 className="text-xs font-semibold">Sub Shop Fab</h4>
                <p className="text-[9px] text-muted-foreground">Sub fabricates these - not your cost</p>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2 text-xs">
              <div>
                <span className="text-muted-foreground">Shop Welds</span>
                <p className="font-semibold font-mono">{subTotalWelds}</p>
              </div>
              <div>
                <span className="text-muted-foreground">BW / SW</span>
                <p className="font-semibold font-mono">{subShopScope.welds.buttWelds} / {subShopScope.welds.socketWelds}</p>
              </div>
              <div>
                <span className="text-muted-foreground">Pipe Footage</span>
                <p className="font-semibold font-mono">{subShopScope.pipeFootage.toFixed(1)} ft</p>
              </div>
              <div>
                <span className="text-muted-foreground">Fittings</span>
                <p className="font-semibold font-mono">{subShopScope.fittingCount}</p>
              </div>
            </div>
            <Separator />
            <button
              className="flex items-center gap-1 text-[10px] text-muted-foreground hover:text-foreground transition-colors w-full"
              onClick={() => toggleCard("subShop")}
            >
              {expandedCards.has("subShop") ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              {subShopScope.items.length} items
            </button>
            {expandedCards.has("subShop") && <ScopeDetail scope={subShopScope} />}
          </CardContent>
        </Card>

        {/* Card 2: Your Field Welds */}
        <Card className="border-orange-200 dark:border-orange-800 bg-orange-50/30 dark:bg-orange-900/10">
          <CardContent className="p-4 space-y-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-orange-100 dark:bg-orange-900/40 flex items-center justify-center">
                <Wrench size={14} className="text-orange-600 dark:text-orange-400" />
              </div>
              <div className="flex-1">
                <h4 className="text-xs font-semibold">Your Field Welds (Subbed Dwgs)</h4>
                <p className="text-[9px] text-muted-foreground">Field tie-ins when spools arrive</p>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2 text-xs">
              <div>
                <span className="text-muted-foreground">Field Welds</span>
                <p className="font-semibold font-mono">{totalWelds(yourFieldScope.welds)}</p>
              </div>
              <div>
                <span className="text-muted-foreground">Bolt-Ups</span>
                <p className="font-semibold font-mono">{yourFieldScope.welds.boltUps}</p>
              </div>
              <div>
                <span className="text-muted-foreground">Field Pipe</span>
                <p className="font-semibold font-mono">{yourFieldScope.pipeFootage.toFixed(1)} ft</p>
              </div>
              <div>
                <span className="text-muted-foreground">Fittings</span>
                <p className="font-semibold font-mono">{yourFieldScope.fittingCount}</p>
              </div>
            </div>
            <Separator />
            <button
              className="flex items-center gap-1 text-[10px] text-muted-foreground hover:text-foreground transition-colors w-full"
              onClick={() => toggleCard("yourField")}
            >
              {expandedCards.has("yourField") ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              {yourFieldScope.items.length} items
            </button>
            {expandedCards.has("yourField") && <ScopeDetail scope={yourFieldScope} />}
          </CardContent>
        </Card>

        {/* Card 3: Your Full Scope */}
        <Card className="border-blue-200 dark:border-blue-800 bg-blue-50/30 dark:bg-blue-900/10">
          <CardContent className="p-4 space-y-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-blue-100 dark:bg-blue-900/40 flex items-center justify-center">
                <HardHat size={14} className="text-blue-600 dark:text-blue-400" />
              </div>
              <div className="flex-1">
                <h4 className="text-xs font-semibold">Your Full Scope (Remaining)</h4>
                <p className="text-[9px] text-muted-foreground">100% your fab + installation</p>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2 text-xs">
              <div>
                <span className="text-muted-foreground">Total Welds</span>
                <p className="font-semibold font-mono">{totalWelds(yourFullScope.welds)}</p>
              </div>
              <div>
                <span className="text-muted-foreground">BW / SW / BU</span>
                <p className="font-semibold font-mono">{yourFullScope.welds.buttWelds} / {yourFullScope.welds.socketWelds} / {yourFullScope.welds.boltUps}</p>
              </div>
              <div>
                <span className="text-muted-foreground">Pipe Footage</span>
                <p className="font-semibold font-mono">{yourFullScope.pipeFootage.toFixed(1)} ft</p>
              </div>
              <div>
                <span className="text-muted-foreground">Fittings</span>
                <p className="font-semibold font-mono">{yourFullScope.fittingCount}</p>
              </div>
            </div>
            <Separator />
            <button
              className="flex items-center gap-1 text-[10px] text-muted-foreground hover:text-foreground transition-colors w-full"
              onClick={() => toggleCard("yourFull")}
            >
              {expandedCards.has("yourFull") ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              {yourFullScope.items.length} items
            </button>
            {expandedCards.has("yourFull") && <ScopeDetail scope={yourFullScope} />}
          </CardContent>
        </Card>
      </div>

      {/* Summary bar */}
      <Card className="border-card-border">
        <CardContent className="p-4">
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div className="flex items-center gap-6 text-xs flex-wrap">
              <div>
                <span className="text-muted-foreground">Total Project Welds</span>
                <p className="font-semibold font-mono text-sm">{projectTotalWelds}</p>
              </div>
              <Separator orientation="vertical" className="h-8" />
              <div>
                <span className="text-green-600 dark:text-green-400">Sub's Welds</span>
                <p className="font-semibold font-mono text-sm">{subTotalWelds} <span className="text-[10px] text-muted-foreground">({subPct}%)</span></p>
              </div>
              <Separator orientation="vertical" className="h-8" />
              <div>
                <span className="text-blue-600 dark:text-blue-400">Your Welds</span>
                <p className="font-semibold font-mono text-sm">
                  {totalWelds(yourFieldScope.welds)} <span className="text-[10px] text-muted-foreground">(field)</span>
                  {" + "}
                  {totalWelds(yourFullScope.welds)} <span className="text-[10px] text-muted-foreground">(remaining)</span>
                  {" = "}
                  {yourTotalWelds} <span className="text-[10px] text-muted-foreground">({yourPct}%)</span>
                </p>
              </div>
              <Separator orientation="vertical" className="h-8" />
              <div>
                <span className="text-muted-foreground">Your Est. Manhours</span>
                <p className="font-semibold font-mono text-sm">{estimateMH.toLocaleString()} MH</p>
              </div>
            </div>
            <Button
              variant="outline"
              size="sm"
              className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
              onClick={handleExport}
              disabled={items.length === 0}
            >
              <FileSpreadsheet size={14} className="mr-1.5" />
              Export Scope Split
            </Button>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
