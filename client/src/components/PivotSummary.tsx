import { useState, useMemo } from "react";
import type { TakeoffItem } from "@shared/schema";
import {
  Table, TableBody, TableCell, TableHead, TableHeader, TableRow,
} from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";

const PIVOT_CATEGORY_COLORS: Record<string, string> = {
  pipe: "bg-blue-100 text-blue-800",
  elbow: "bg-purple-100 text-purple-800",
  tee: "bg-indigo-100 text-indigo-800",
  valve: "bg-green-100 text-green-800",
  flange: "bg-orange-100 text-orange-800",
  bolt: "bg-gray-100 text-gray-800",
  gasket: "bg-pink-100 text-pink-800",
};

interface PivotSummaryProps {
  items: TakeoffItem[];
}

type PivotMode = "category_size" | "category_sheet" | "sheet_category" | "size_category";

export default function PivotSummary({ items }: PivotSummaryProps) {
  const [cloudFilter, setCloudFilter] = useState<"all" | "clouded" | "non-clouded">("all");
  const [categoryFilter, setCategoryFilter] = useState<string>("all");
  const [sizeFilter, setSizeFilter] = useState<string>("all");
  const [sheetFilter, setSheetFilter] = useState<string>("all");
  const [pivotMode, setPivotMode] = useState<PivotMode>("category_size");

  const cloudedCount = items.filter(i => i.revisionClouded).length;
  const hasCloudedItems = cloudedCount > 0;

  // Derive filter options
  const categories = useMemo(() => {
    const cats: Record<string, number> = {};
    for (const item of items) cats[item.category] = (cats[item.category] || 0) + 1;
    return Object.entries(cats).sort((a, b) => b[1] - a[1]);
  }, [items]);

  const allSizes = useMemo(() => {
    const s: Record<string, number> = {};
    for (const item of items) s[item.size || "N/A"] = (s[item.size || "N/A"] || 0) + 1;
    return Object.entries(s).sort((a, b) => {
      const na = parseFloat(a[0]); const nb = parseFloat(b[0]);
      return isNaN(na) || isNaN(nb) ? a[0].localeCompare(b[0]) : na - nb;
    });
  }, [items]);

  const allSheets = useMemo(() => {
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

  // Apply filters
  let filtered = items;
  if (cloudFilter === "clouded") filtered = filtered.filter(i => i.revisionClouded);
  else if (cloudFilter === "non-clouded") filtered = filtered.filter(i => !i.revisionClouded);
  if (categoryFilter !== "all") filtered = filtered.filter(i => i.category === categoryFilter);
  if (sizeFilter !== "all") filtered = filtered.filter(i => (i.size || "N/A") === sizeFilter);
  if (sheetFilter !== "all") {
    filtered = filtered.filter(i => {
      const pg = (i as any).sourcePage;
      const sheet = pg != null ? `Sheet ${pg}` : (i.notes?.match(/Sheet\s+(\d+)/i)?.[0] || "Unknown");
      return sheet === sheetFilter;
    });
  }

  const getSheet = (item: TakeoffItem) => {
    const pg = (item as any).sourcePage;
    return pg != null ? `Sheet ${pg}` : (item.notes?.match(/Sheet\s+(\d+)/i)?.[0] || "Unknown");
  };

  // Build pivot data based on mode
  const { rowLabels, colLabels, grid, rowTotals, colTotals, grandTotal } = useMemo(() => {
    const data: Record<string, Record<string, { qty: number; unit: string }>> = {};
    let rowKey: (item: TakeoffItem) => string;
    let colKey: (item: TakeoffItem) => string;

    switch (pivotMode) {
      case "category_size":
        rowKey = i => i.category;
        colKey = i => i.size || "N/A";
        break;
      case "category_sheet":
        rowKey = i => i.category;
        colKey = i => getSheet(i);
        break;
      case "sheet_category":
        rowKey = i => getSheet(i);
        colKey = i => i.category;
        break;
      case "size_category":
        rowKey = i => i.size || "N/A";
        colKey = i => i.category;
        break;
    }

    for (const item of filtered) {
      const rk = rowKey(item);
      const ck = colKey(item);
      if (!data[rk]) data[rk] = {};
      if (!data[rk][ck]) data[rk][ck] = { qty: 0, unit: item.unit };
      data[rk][ck].qty += item.quantity;
    }

    // Sort row/col labels
    const sortNumeric = (a: string, b: string) => {
      const na = parseFloat(a.replace(/[^0-9.]/g, "")); const nb = parseFloat(b.replace(/[^0-9.]/g, ""));
      return isNaN(na) || isNaN(nb) ? a.localeCompare(b) : na - nb;
    };

    const rLabels = Object.keys(data).sort(pivotMode.startsWith("sheet") ? sortNumeric : (a, b) => a.localeCompare(b));
    const cLabelsSet = new Set<string>();
    for (const row of Object.values(data)) for (const col of Object.keys(row)) cLabelsSet.add(col);
    const cLabels = Array.from(cLabelsSet).sort(pivotMode.endsWith("size") ? sortNumeric : pivotMode.endsWith("sheet") ? sortNumeric : (a, b) => a.localeCompare(b));

    // Compute totals
    const rTotals: Record<string, number> = {};
    const cTotals: Record<string, number> = {};
    let gt = 0;
    for (const rk of rLabels) {
      rTotals[rk] = 0;
      for (const ck of cLabels) {
        const val = data[rk]?.[ck]?.qty || 0;
        rTotals[rk] += val;
        cTotals[ck] = (cTotals[ck] || 0) + val;
        gt += val;
      }
    }

    return { rowLabels: rLabels, colLabels: cLabels, grid: data, rowTotals: rTotals, colTotals: cTotals, grandTotal: gt };
  }, [filtered, pivotMode]);

  const activeFilterCount = [categoryFilter !== "all", sizeFilter !== "all", sheetFilter !== "all", cloudFilter !== "all"].filter(Boolean).length;

  if (items.length === 0) return null;

  const formatQty = (qty: number) => {
    if (qty === 0) return "";
    return qty % 1 !== 0 ? qty.toFixed(2) : qty.toLocaleString();
  };

  return (
    <div className="space-y-2">
      {/* Filter bar */}
      <div className="flex items-center gap-2 flex-wrap border border-border bg-muted/20 rounded-md px-3 py-2">
        <span className="text-xs text-muted-foreground font-medium shrink-0">Pivot:</span>

        {/* Pivot mode */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border font-medium"
          value={pivotMode}
          onChange={e => setPivotMode(e.target.value as PivotMode)}
        >
          <option value="category_size">Category x Size</option>
          <option value="category_sheet">Category x Sheet</option>
          <option value="sheet_category">Sheet x Category</option>
          <option value="size_category">Size x Category</option>
        </select>

        <span className="text-border">|</span>

        {/* Category filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={categoryFilter}
          onChange={e => setCategoryFilter(e.target.value)}
        >
          <option value="all">All Categories</option>
          {categories.map(([cat, count]) => (
            <option key={cat} value={cat}>{cat.replace(/_/g, " ")} ({count})</option>
          ))}
        </select>

        {/* Size filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={sizeFilter}
          onChange={e => setSizeFilter(e.target.value)}
        >
          <option value="all">All Sizes</option>
          {allSizes.map(([sz, count]) => (
            <option key={sz} value={sz}>{sz} ({count})</option>
          ))}
        </select>

        {/* Sheet filter */}
        <select
          className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
          value={sheetFilter}
          onChange={e => setSheetFilter(e.target.value)}
        >
          <option value="all">All Sheets</option>
          {allSheets.map(([sheet, count]) => (
            <option key={sheet} value={sheet}>{sheet} ({count})</option>
          ))}
        </select>

        {/* Cloud filter */}
        {hasCloudedItems && (
          <select
            className="h-7 text-xs border rounded px-2 bg-background text-foreground border-border"
            value={cloudFilter}
            onChange={e => setCloudFilter(e.target.value as any)}
          >
            <option value="all">All Revisions</option>
            <option value="clouded">Clouded ({cloudedCount})</option>
            <option value="non-clouded">Non-Clouded ({items.length - cloudedCount})</option>
          </select>
        )}

        {activeFilterCount > 0 && (
          <button
            className="h-7 text-xs px-2 rounded border border-border hover:bg-accent text-muted-foreground"
            onClick={() => { setCategoryFilter("all"); setSizeFilter("all"); setSheetFilter("all"); setCloudFilter("all"); }}
          >
            Clear ({activeFilterCount})
          </button>
        )}

        <span className="text-xs text-muted-foreground ml-auto shrink-0">
          {filtered.length} of {items.length} items
        </span>
      </div>

      {/* Pivot table */}
      {rowLabels.length === 0 ? (
        <div className="flex items-center justify-center h-20 text-muted-foreground text-xs">
          No items match the selected filters.
        </div>
      ) : (
        <div className="overflow-auto rounded-md border border-border">
          <Table>
            <TableHeader>
              <TableRow className="bg-muted/50">
                <TableHead className="text-xs font-semibold sticky left-0 bg-muted/50 z-10 min-w-[120px]">
                  {pivotMode.split("_")[0] === "category" ? "Category" : pivotMode.split("_")[0] === "sheet" ? "Sheet" : "Size"}
                </TableHead>
                {colLabels.map(col => (
                  <TableHead key={col} className="text-xs text-right font-mono whitespace-nowrap min-w-[60px]">
                    {col.replace(/_/g, " ")}
                  </TableHead>
                ))}
                <TableHead className="text-xs text-right font-semibold bg-muted/70 min-w-[70px]">Total</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {rowLabels.map(row => (
                <TableRow key={row} className="hover:bg-muted/30">
                  <TableCell className="text-xs font-medium capitalize sticky left-0 bg-background z-10">
                    {row.replace(/_/g, " ")}
                  </TableCell>
                  {colLabels.map(col => {
                    const cell = grid[row]?.[col];
                    return (
                      <TableCell key={col} className="text-xs text-right font-mono whitespace-nowrap">
                        {cell ? (
                          <span>
                            {formatQty(cell.qty)}
                            {cell.qty > 0 && <span className="text-muted-foreground ml-1 text-[9px]">{cell.unit}</span>}
                          </span>
                        ) : (
                          <span className="text-muted-foreground/20">—</span>
                        )}
                      </TableCell>
                    );
                  })}
                  <TableCell className="text-xs text-right font-mono font-semibold bg-muted/30 whitespace-nowrap">
                    {formatQty(rowTotals[row] || 0)}
                  </TableCell>
                </TableRow>
              ))}

              {/* Totals row */}
              <TableRow className="bg-muted/50 font-semibold border-t-2 border-border">
                <TableCell className="text-xs font-semibold sticky left-0 bg-muted/50 z-10">Total</TableCell>
                {colLabels.map(col => (
                  <TableCell key={col} className="text-xs text-right font-mono whitespace-nowrap">
                    {formatQty(colTotals[col] || 0)}
                  </TableCell>
                ))}
                <TableCell className="text-xs text-right font-mono font-bold bg-primary/10 whitespace-nowrap">
                  {formatQty(grandTotal)}
                </TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </div>
      )}

      {/* Summary cards below pivot */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mt-3">
        {categories.slice(0, 8).map(([cat, count]) => {
          const catItems = filtered.filter(i => i.category === cat);
          const totalQty = catItems.reduce((s, i) => s + i.quantity, 0);
          const unit = catItems[0]?.unit || "EA";
          return (
            <div
              key={cat}
              className={`border rounded-md p-2 cursor-pointer transition-colors ${categoryFilter === cat ? "border-primary bg-primary/5" : "border-border hover:bg-muted/30"}`}
              onClick={() => setCategoryFilter(categoryFilter === cat ? "all" : cat)}
            >
              <div className="flex items-center gap-1.5 mb-1">
                <Badge variant="outline" className={`text-[9px] px-1 py-0 ${PIVOT_CATEGORY_COLORS[cat] || ""}`}>
                  {cat.replace(/_/g, " ")}
                </Badge>
              </div>
              <p className="text-sm font-semibold tabular-nums">{unit === "LF" ? totalQty.toFixed(1) : Math.round(totalQty).toLocaleString()} <span className="text-[10px] text-muted-foreground font-normal">{unit}</span></p>
              <p className="text-[10px] text-muted-foreground">{catItems.length} line items</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}
