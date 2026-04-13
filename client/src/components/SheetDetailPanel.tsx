import { useMemo } from "react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { X } from "lucide-react";

const WELD_RULES: Record<string, { bw: number; sw: number; label: string }> = {
  elbow: { bw: 2, sw: 0, label: "2 BW" },
  tee: { bw: 3, sw: 0, label: "3 BW" },
  reducer: { bw: 2, sw: 0, label: "2 BW" },
  cap: { bw: 1, sw: 0, label: "1 BW" },
  coupling: { bw: 0, sw: 2, label: "2 SW" },
  valve: { bw: 2, sw: 0, label: "2 BW" },
  flange: { bw: 1, sw: 0, label: "1 BW + 1 BU" },
  sockolet: { bw: 1, sw: 0, label: "1 BW" },
};

interface SheetDetailPanelProps {
  pageNum: number;
  items: any[];
  onClose: () => void;
}

export default function SheetDetailPanel({ pageNum, items, onClose }: SheetDetailPanelProps) {
  const pageItems = useMemo(() => items.filter(i => i.sourcePage === pageNum), [items, pageNum]);

  const drawingNumber = pageItems.find(i => i.drawingNumber)?.drawingNumber || null;

  const weldSummary = useMemo(() => {
    let totalBW = 0, totalSW = 0, totalBU = 0;
    const fittingRows: { category: string; size: string; description: string; qty: number; bw: number; sw: number; bu: number; ruleLabel: string }[] = [];

    for (const item of pageItems) {
      const cat = (item.category || "").toLowerCase();
      const qty = item.quantity || 0;
      const rule = WELD_RULES[cat];

      if (rule || cat === "bolt" || cat === "gasket") {
        const bw = rule ? qty * rule.bw : 0;
        const sw = rule ? qty * rule.sw : 0;
        const bu = cat === "flange" ? Math.ceil(qty / 2) : (cat === "bolt" || cat === "gasket" ? qty : 0);
        totalBW += bw;
        totalSW += sw;
        totalBU += bu;
        fittingRows.push({
          category: cat,
          size: item.size || "N/A",
          description: item.description || cat,
          qty,
          bw,
          sw,
          bu,
          ruleLabel: rule ? `${qty}x ${rule.label}` : `${qty}x BU`,
        });
      }
    }

    return { totalBW, totalSW, totalBU, fittingRows };
  }, [pageItems]);

  return (
    <div className="fixed inset-y-0 right-0 w-full max-w-md bg-background border-l border-border shadow-2xl z-50 flex flex-col animate-in slide-in-from-right-full duration-200">
      {/* Header */}
      <div className="flex items-center justify-between p-4 border-b border-border bg-muted/30">
        <div>
          <h3 className="font-semibold text-sm">Sheet {pageNum}</h3>
          {drawingNumber && <p className="text-xs text-muted-foreground">{drawingNumber}</p>}
        </div>
        <Button variant="ghost" size="sm" onClick={onClose} className="h-7 w-7 p-0">
          <X size={16} />
        </Button>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-4 gap-2 p-3 border-b border-border">
        <div className="text-center">
          <div className="text-lg font-bold text-primary">{pageItems.length}</div>
          <div className="text-[10px] text-muted-foreground">Items</div>
        </div>
        <div className="text-center">
          <div className="text-lg font-bold text-orange-500">{weldSummary.totalBW}</div>
          <div className="text-[10px] text-muted-foreground">Butt Welds</div>
        </div>
        <div className="text-center">
          <div className="text-lg font-bold text-blue-500">{weldSummary.totalSW}</div>
          <div className="text-[10px] text-muted-foreground">Socket Welds</div>
        </div>
        <div className="text-center">
          <div className="text-lg font-bold text-green-500">{weldSummary.totalBU}</div>
          <div className="text-[10px] text-muted-foreground">Bolt-Ups</div>
        </div>
      </div>

      {/* Weld calculation breakdown */}
      {weldSummary.fittingRows.length > 0 && (
        <div className="p-3 border-b border-border">
          <h4 className="text-xs font-semibold mb-2 text-muted-foreground">Weld Calculation</h4>
          <div className="space-y-1">
            {weldSummary.fittingRows.map((row, i) => (
              <div key={i} className="flex items-center justify-between text-[11px]">
                <span className="truncate flex-1">{row.qty}x {row.size}" {row.category}</span>
                <span className="text-muted-foreground ml-2 shrink-0">{row.ruleLabel}</span>
                <span className="font-mono ml-2 shrink-0 w-16 text-right">
                  {row.bw > 0 && <span className="text-orange-500">{row.bw} BW</span>}
                  {row.sw > 0 && <span className="text-blue-500">{row.sw} SW</span>}
                  {row.bu > 0 && <span className="text-green-500"> {row.bu} BU</span>}
                </span>
              </div>
            ))}
            <div className="flex items-center justify-between text-xs font-semibold border-t border-border pt-1 mt-1">
              <span>Total</span>
              <span className="font-mono">
                {weldSummary.totalBW > 0 && <span className="text-orange-500">{weldSummary.totalBW} BW</span>}
                {weldSummary.totalSW > 0 && <span className="text-blue-500 ml-1">{weldSummary.totalSW} SW</span>}
                {weldSummary.totalBU > 0 && <span className="text-green-500 ml-1">{weldSummary.totalBU} BU</span>}
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Items list */}
      <div className="flex-1 overflow-y-auto p-3">
        <h4 className="text-xs font-semibold mb-2 text-muted-foreground">All Items ({pageItems.length})</h4>
        <table className="w-full text-[11px]">
          <thead>
            <tr className="bg-muted/50 text-left">
              <th className="p-1.5 font-semibold">Category</th>
              <th className="p-1.5 font-semibold">Size</th>
              <th className="p-1.5 font-semibold">Description</th>
              <th className="p-1.5 font-semibold text-right">Qty</th>
              <th className="p-1.5 font-semibold">Unit</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((item, i) => (
              <tr key={item.id || i} className={i % 2 === 0 ? "bg-background" : "bg-muted/20"}>
                <td className="p-1.5">
                  <Badge variant="outline" className="text-[9px] px-1 py-0">{item.category}</Badge>
                </td>
                <td className="p-1.5 font-mono">{item.size}</td>
                <td className="p-1.5 truncate max-w-[160px]" title={item.description}>{item.description}</td>
                <td className="p-1.5 text-right font-mono">{item.quantity}</td>
                <td className="p-1.5 text-muted-foreground">{item.unit}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
