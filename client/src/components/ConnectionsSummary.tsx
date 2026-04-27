import { useMemo, useState } from "react";
import SheetDetailPanel from "./SheetDetailPanel";

interface ConnectionsSummaryProps {
  items: any[];
}

// Weld inference rules (mirrors server/routes.ts inferWeldsFromFittings)
// Note: valves, couplings, and olets use description-based logic below, not this table
const WELD_RULES: Record<string, { welds: number; type: string; label: string }> = {
  elbow: { welds: 2, type: "butt_weld", label: "Butt Welds" },
  tee: { welds: 3, type: "butt_weld", label: "Butt Welds" },
  reducer: { welds: 2, type: "butt_weld", label: "Butt Welds" },
  cap: { welds: 1, type: "butt_weld", label: "Butt Weld" },
  flange: { welds: 1, type: "butt_weld", label: "Butt Weld" },
};

type ConnectionRow = {
  size: string;
  buttWelds: number;
  socketWelds: number;
  boltUps: number;
  threadedConns: number;
  totalConnections: number;
};

export default function ConnectionsSummary({ items }: ConnectionsSummaryProps) {
  const [byPageOpen, setByPageOpen] = useState(false);
  const [expandedPages, setExpandedPages] = useState<Set<number>>(new Set());
  const [sheetPanelPage, setSheetPanelPage] = useState<number | null>(null);
  const [showRevisionsOnly, setShowRevisionsOnly] = useState(false);

  const cloudedCount = items.filter(i => i.revisionClouded).length;
  const hasCloudedItems = cloudedCount > 0;
  const filteredItems = showRevisionsOnly ? items.filter(i => i.revisionClouded) : items;

  const togglePageExpanded = (page: number) => {
    setExpandedPages(prev => {
      const next = new Set(prev);
      if (next.has(page)) next.delete(page);
      else next.add(page);
      return next;
    });
  };
  // Extract visual weld counts from AI-detected weld symbols on drawings
  const visualWeldData = useMemo(() => {
    const pageCounts: { page: number; visual: { buttWelds: number; socketWelds: number; fieldWelds: number }; inferred: { buttWelds: number; socketWelds: number } }[] = [];
    let totalVisualBW = 0, totalVisualSW = 0, totalVisualFW = 0;

    // Get visual counts from items that have _visualWeldCount
    const pagesWithVisual = new Map<number, { buttWelds: number; socketWelds: number; fieldWelds: number }>();
    for (const item of items) {
      if (item._visualWeldCount) {
        pagesWithVisual.set(item.sourcePage, item._visualWeldCount);
      }
    }

    if (pagesWithVisual.size === 0) return null;

    // Build continuation map: which pages are "to" pages (receiving continuations)
    // Welds at continuation points appear on both sheets — subtract from "to" sheet
    const continuationToPages = new Set<number>();
    const continuationWeldsPerPage: Record<number, { bw: number; sw: number }> = {};
    for (const item of items) {
      if (item._continuations && Array.isArray(item._continuations)) {
        for (const conn of item._continuations) {
          if (conn.direction === "to") {
            // This page has a "to" continuation — welds at that point are shared
            // Count continuation-point items on this page to estimate shared welds
            const page = item.sourcePage || 0;
            continuationToPages.add(page);
          }
        }
      }
      // Count welds from items marked at continuation points on "to" pages
      if (item.atContinuation && item.sourcePage && continuationToPages.has(item.sourcePage)) {
        if (!continuationWeldsPerPage[item.sourcePage]) continuationWeldsPerPage[item.sourcePage] = { bw: 0, sw: 0 };
        const cat = (item.category || "").toLowerCase();
        const qty = item.quantity || 0;
        // Estimate welds for this continuation item
        if (cat === "elbow") { continuationWeldsPerPage[item.sourcePage].bw += qty; } // 1 of the 2 welds is shared
        else if (cat === "tee") { continuationWeldsPerPage[item.sourcePage].bw += qty; }
        else if (cat === "reducer") { continuationWeldsPerPage[item.sourcePage].bw += qty; }
        else if (cat === "flange") { continuationWeldsPerPage[item.sourcePage].bw += qty; }
      }
    }

    // Calculate BOM-inferred weld counts per page
    const RULES: Record<string, { bw: number; sw: number }> = {
      elbow: { bw: 2, sw: 0 }, tee: { bw: 3, sw: 0 }, reducer: { bw: 2, sw: 0 },
      cap: { bw: 1, sw: 0 }, flange: { bw: 1, sw: 0 },
      sockolet: { bw: 0, sw: 2 }, weldolet: { bw: 1, sw: 0 },
    };

    let totalContinuationWelds = 0;
    for (const [page, visual] of pagesWithVisual) {
      const pageItems = items.filter(i => i.sourcePage === page);
      let inferredBW = 0, inferredSW = 0;
      for (const item of pageItems) {
        const cat = (item.category || "").toLowerCase();
        const qty = item.quantity || 0;
        const rule = RULES[cat];
        if (rule) { inferredBW += qty * rule.bw; inferredSW += qty * rule.sw; }
      }

      // Adjust visual count: subtract continuation-point welds on "to" pages
      // These weld dots appear on both sheets but should only be counted once
      let adjustedVisualBW = visual.buttWelds || 0;
      let adjustedVisualSW = visual.socketWelds || 0;
      const contWelds = continuationWeldsPerPage[page];
      if (contWelds) {
        adjustedVisualBW = Math.max(0, adjustedVisualBW - contWelds.bw);
        adjustedVisualSW = Math.max(0, adjustedVisualSW - contWelds.sw);
        totalContinuationWelds += contWelds.bw + contWelds.sw;
      }

      totalVisualBW += adjustedVisualBW;
      totalVisualSW += adjustedVisualSW;
      totalVisualFW += visual.fieldWelds || 0;
      pageCounts.push({ page, visual: { buttWelds: adjustedVisualBW, socketWelds: adjustedVisualSW, fieldWelds: visual.fieldWelds || 0 }, inferred: { buttWelds: inferredBW, socketWelds: inferredSW } });
    }

    const totalInferredBW = pageCounts.reduce((s, p) => s + p.inferred.buttWelds, 0);
    const totalInferredSW = pageCounts.reduce((s, p) => s + p.inferred.socketWelds, 0);
    const bwDiff = totalVisualBW - totalInferredBW;
    const swDiff = totalVisualSW - totalInferredSW;

    return {
      pageCounts,
      totals: { visualBW: totalVisualBW, visualSW: totalVisualSW, visualFW: totalVisualFW, inferredBW: totalInferredBW, inferredSW: totalInferredSW },
      bwDiff, swDiff,
      isVerified: Math.abs(bwDiff) <= 2 && Math.abs(swDiff) <= 2,
    };
  }, [items]);

  const { rows, totals, details } = useMemo(() => {
    const sizeMap: Record<string, { buttWelds: number; socketWelds: number; boltUps: number; threadedConns: number }> = {};
    const detailsList: { size: string; fitting: string; qty: number; connectionType: string; connectionCount: number; section: string }[] = [];

    const ensureSize = (size: string) => {
      if (!sizeMap[size]) sizeMap[size] = { buttWelds: 0, socketWelds: 0, boltUps: 0, threadedConns: 0 };
    };

    // Detect butterfly valves per page — when a butterfly valve sits between two
    // flanges, the entire joint counts as ONE bolt-up (single bolt set), not two.
    const butterflyPages = new Set<number>();
    for (const item of filteredItems) {
      const desc = (item.description || "").toLowerCase();
      if (desc.includes("butterfly") && item.sourcePage) {
        butterflyPages.add(item.sourcePage);
      }
    }

    for (const item of filteredItems) {
      const cat = (item.category || "").toLowerCase();
      const size = item.size || "N/A";
      const qty = item.quantity || 0;
      const desc = (item.description || "").toLowerCase();
      const section = item.installLocation || ((item.notes || "").toLowerCase().includes("field") ? "field" : "shop");

      ensureSize(size);

      // Determine connection type based on description keywords
      const isThreaded = desc.includes("threaded") || desc.includes("screw") || desc.includes("npt") || desc.includes("fnpt") || desc.includes("mnpt");
      const isSocketWeld = desc.includes("socket weld") || desc.includes("sw ") || desc.includes(",sw,") || desc.includes(" sw,") || /\bsw\b/i.test(desc);

      // === PIPE LENGTH WELDS ===
      // Every 40' of pipe run requires a field weld where 40' standard lengths are joined.
      if (cat === "pipe" && qty >= 40) {
        const pipeJointWelds = Math.floor(qty / 40);
        if (pipeJointWelds > 0) {
          sizeMap[size].buttWelds += pipeJointWelds;
          detailsList.push({ size, fitting: `${qty.toFixed(1)} LF Pipe (40' joints)`, qty: pipeJointWelds, connectionType: "Butt Welds", connectionCount: pipeJointWelds, section });
        }
        continue;
      }
      if (cat === "pipe") continue;

      // Gaskets and stud bolts do NOT generate bolt-up counts — only flanges do.
      // The flange is the connection that requires the bolt-up; gaskets and bolts
      // are just the supplied hardware for that connection.
      if (cat === "gasket" || desc.includes("gasket") || desc.includes("stud")) {
        continue;
      }
      if (cat === "bolt" || desc.includes("bolt")) {
        continue;
      }

      if (cat === "flange") {
        // Each flange has 1 weld to pipe + contributes to a bolt-up connection.
        // Two flanges = 1 bolt-up joint.
        // Butterfly valve adjustment: when a butterfly valve is on the same page,
        // the two flanges around it form a single butterfly joint = 1 bolt-up.
        // Without butterfly: 2 flanges = 1 bolt-up. With butterfly: same math holds.
        sizeMap[size].buttWelds += qty;
        // Butterfly joint is 2 flanges + 1 valve, still counts as 1 bolt-up per joint.
        sizeMap[size].boltUps += Math.ceil(qty / 2);
        const noteSuffix = butterflyPages.has(item.sourcePage) ? " (butterfly joints on this page)" : "";
        detailsList.push({ size, fitting: (item.description || "Flange") + noteSuffix, qty, connectionType: "Butt Weld + Bolt-Up", connectionCount: qty, section });
      } else if (isThreaded) {
        // Threaded couplings/unions/valves: 0 welds, just threaded connections
        sizeMap[size].threadedConns += qty;
        detailsList.push({ size, fitting: item.description || cat, qty, connectionType: "Threaded", connectionCount: qty, section });
      } else if (cat === "coupling" || cat === "union") {
        // Socket weld couplings/unions: 2 SW each (threaded already handled above)
        const weldCount = qty * 2;
        if (isSocketWeld) {
          sizeMap[size].socketWelds += weldCount;
          detailsList.push({ size, fitting: item.description || cat, qty, connectionType: "Socket Welds", connectionCount: weldCount, section });
        } else {
          sizeMap[size].socketWelds += weldCount;
          detailsList.push({ size, fitting: item.description || cat, qty, connectionType: "Socket Welds", connectionCount: weldCount, section });
        }
      } else if (cat === "valve") {
        // ONLY socket-weld valves generate welds. All other valves (flanged,
        // threaded, butterfly, butt-weld end) connect via the surrounding
        // flanges/joints and have no welds of their own.
        if (isSocketWeld) {
          sizeMap[size].socketWelds += qty * 2;
          detailsList.push({ size, fitting: item.description || "Valve (SW)", qty, connectionType: "Socket Welds", connectionCount: qty * 2, section });
        }
        // Flanged, threaded, butterfly, BW valves: skipped (no weld entry)
      } else if (cat === "sockolet" || desc.includes("sockolet")) {
        // Sockolet: 2 SW (header bore + branch)
        sizeMap[size].socketWelds += qty * 2;
        detailsList.push({ size, fitting: item.description || "Sockolet", qty, connectionType: "Socket Welds", connectionCount: qty * 2, section });
      } else if (cat === "weldolet" || desc.includes("weldolet")) {
        // Weldolet: 1 BW to header
        sizeMap[size].buttWelds += qty * 1;
        detailsList.push({ size, fitting: item.description || "Weldolet", qty, connectionType: "Butt Weld", connectionCount: qty, section });
      } else if (desc.includes("threadolet")) {
        // Threadolet: 1 weld to header
        sizeMap[size].buttWelds += qty * 1;
        detailsList.push({ size, fitting: item.description || "Threadolet", qty, connectionType: "Butt Weld", connectionCount: qty, section });
      } else if (WELD_RULES[cat]) {
        const rule = WELD_RULES[cat];
        const weldCount = qty * rule.welds;
        if (rule.type === "socket_weld" || isSocketWeld) {
          sizeMap[size].socketWelds += weldCount;
        } else {
          sizeMap[size].buttWelds += weldCount;
        }
        detailsList.push({ size, fitting: item.description || cat, qty, connectionType: rule.label, connectionCount: weldCount, section });
      }
    }

    // Sort sizes numerically
    const sortedSizes = Object.keys(sizeMap).sort((a, b) => {
      const na = parseFloat(a) || 0;
      const nb = parseFloat(b) || 0;
      return na - nb;
    });

    const rows: ConnectionRow[] = sortedSizes.map(size => ({
      size,
      buttWelds: sizeMap[size].buttWelds,
      socketWelds: sizeMap[size].socketWelds,
      boltUps: sizeMap[size].boltUps,
      threadedConns: sizeMap[size].threadedConns,
      totalConnections: sizeMap[size].buttWelds + sizeMap[size].socketWelds + sizeMap[size].boltUps + sizeMap[size].threadedConns,
    })).filter(r => r.totalConnections > 0);

    const totals = rows.reduce((acc, r) => ({
      buttWelds: acc.buttWelds + r.buttWelds,
      socketWelds: acc.socketWelds + r.socketWelds,
      boltUps: acc.boltUps + r.boltUps,
      threadedConns: acc.threadedConns + r.threadedConns,
      totalConnections: acc.totalConnections + r.totalConnections,
    }), { buttWelds: 0, socketWelds: 0, boltUps: 0, threadedConns: 0, totalConnections: 0 });

    return { rows, totals, details: detailsList };
  }, [filteredItems]);

  // Group connections by page
  const byPage = useMemo(() => {
    const RULES: Record<string, { bw: number; sw: number; label: string }> = {
      elbow: { bw: 2, sw: 0, label: "2 BW" }, tee: { bw: 3, sw: 0, label: "3 BW" }, reducer: { bw: 2, sw: 0, label: "2 BW" },
      cap: { bw: 1, sw: 0, label: "1 BW" }, flange: { bw: 1, sw: 0, label: "1 BW + BU" },
      sockolet: { bw: 0, sw: 2, label: "2 SW" }, weldolet: { bw: 1, sw: 0, label: "1 BW" },
    };
    const pageMap: Record<number, { drawingNumber: string | null; fittings: { category: string; size: string; description: string; qty: number; bw: number; sw: number; bu: number; ruleLabel: string }[]; totalBW: number; totalSW: number; totalBU: number }> = {};

    for (const item of filteredItems) {
      const page = item.sourcePage || 0;
      if (!pageMap[page]) pageMap[page] = { drawingNumber: item.drawingNumber || null, fittings: [], totalBW: 0, totalSW: 0, totalBU: 0 };
      if (!pageMap[page].drawingNumber && item.drawingNumber) pageMap[page].drawingNumber = item.drawingNumber;

      const cat = (item.category || "").toLowerCase();
      const desc = (item.description || "").toLowerCase();
      const qty = item.quantity || 0;
      const isThreaded = desc.includes("threaded") || desc.includes("npt") || desc.includes("screw");
      const isSW = desc.includes("socket") || desc.includes(" sw ") || desc.includes(",sw,") || /\bsw\b/i.test(desc);
      const isFlanged = desc.includes("flanged") || desc.includes("flg") || desc.includes("rf ") || desc.includes("raised face");

      let bw = 0, sw = 0, bu = 0, ruleLabel = "";

      if (cat === "bolt" || cat === "gasket") {
        bu = qty; ruleLabel = `${qty}x BU`;
      } else if (cat === "flange") {
        bw = qty; bu = Math.ceil(qty / 2); ruleLabel = `${qty}x 1 BW + BU`;
      } else if (cat === "valve") {
        if (isFlanged || isThreaded) { ruleLabel = `${qty}x 0 welds`; }
        else if (isSW) { sw = qty * 2; ruleLabel = `${qty}x 2 SW`; }
        else { bw = qty * 2; ruleLabel = `${qty}x 2 BW`; }
      } else if (cat === "coupling" || cat === "union") {
        if (isThreaded) { ruleLabel = `${qty}x threaded`; }
        else { sw = qty * 2; ruleLabel = `${qty}x 2 SW`; }
      } else {
        const rule = RULES[cat as keyof typeof RULES];
        if (rule) {
          bw = qty * rule.bw; sw = qty * rule.sw; ruleLabel = `${qty}x ${rule.label}`;
        }
      }

      if (bw > 0 || sw > 0 || bu > 0 || ruleLabel) {
        pageMap[page].fittings.push({
          category: cat, size: item.size || "N/A", description: item.description || cat, qty, bw, sw, bu, ruleLabel,
        });
        pageMap[page].totalBW += bw;
        pageMap[page].totalSW += sw;
        pageMap[page].totalBU += bu;
      }
    }

    return Object.entries(pageMap)
      .filter(([, data]) => data.fittings.length > 0)
      .map(([page, data]) => ({ page: parseInt(page), ...data }))
      .sort((a, b) => a.page - b.page);
  }, [filteredItems]);

  if (rows.length === 0) {
    return (
      <div className="text-center py-8 text-muted-foreground text-sm">
        No connections found. Run a takeoff with fittings, flanges, or valves to see weld and bolt-up counts.
      </div>
    );
  }

  return (
    <div className="space-y-6">
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

      {/* Summary Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div className="bg-card border rounded-lg p-3 text-center">
          <div className="text-2xl font-bold text-primary">{totals.buttWelds}</div>
          <div className="text-xs text-muted-foreground">Butt Welds</div>
        </div>
        <div className="bg-card border rounded-lg p-3 text-center">
          <div className="text-2xl font-bold text-primary">{totals.socketWelds}</div>
          <div className="text-xs text-muted-foreground">Socket Welds</div>
        </div>
        <div className="bg-card border rounded-lg p-3 text-center">
          <div className="text-2xl font-bold text-primary">{totals.boltUps}</div>
          <div className="text-xs text-muted-foreground">Bolt-Ups</div>
        </div>
        <div className="bg-card border rounded-lg p-3 text-center">
          <div className="text-2xl font-bold text-primary">{totals.threadedConns}</div>
          <div className="text-xs text-muted-foreground">Threaded</div>
        </div>
      </div>

      {/* Connections by Size Table */}
      <div>
        <h3 className="text-sm font-semibold mb-2">Connections by Size</h3>
        <div className="border rounded-lg overflow-hidden">
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-muted/50">
                <th className="text-left p-2 font-semibold">Size</th>
                <th className="text-right p-2 font-semibold">Butt Welds</th>
                <th className="text-right p-2 font-semibold">Socket Welds</th>
                <th className="text-right p-2 font-semibold">Bolt-Ups</th>
                <th className="text-right p-2 font-semibold">Threaded</th>
                <th className="text-right p-2 font-semibold">Total</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row, i) => (
                <tr key={row.size} className={i % 2 === 0 ? "bg-background" : "bg-muted/20"}>
                  <td className="p-2 font-medium">{row.size}"</td>
                  <td className="p-2 text-right">{row.buttWelds || "—"}</td>
                  <td className="p-2 text-right">{row.socketWelds || "—"}</td>
                  <td className="p-2 text-right">{row.boltUps || "—"}</td>
                  <td className="p-2 text-right">{row.threadedConns || "—"}</td>
                  <td className="p-2 text-right font-semibold">{row.totalConnections}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-primary/10 font-semibold border-t">
                <td className="p-2">TOTAL</td>
                <td className="p-2 text-right">{totals.buttWelds}</td>
                <td className="p-2 text-right">{totals.socketWelds}</td>
                <td className="p-2 text-right">{totals.boltUps}</td>
                <td className="p-2 text-right">{totals.threadedConns}</td>
                <td className="p-2 text-right">{totals.totalConnections}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Weld Verification (visual dots vs BOM-inferred) */}
      {visualWeldData && (
        <div>
          <h3 className="text-sm font-semibold mb-2">Weld Verification (Drawing Symbols vs. BOM)</h3>
          <div className={`border rounded-lg p-3 ${visualWeldData.isVerified ? 'border-green-500/30 bg-green-500/5' : 'border-yellow-500/30 bg-yellow-500/5'}`}>
            <div className="grid grid-cols-4 gap-3 text-xs mb-3">
              <div className="text-center">
                <div className="text-muted-foreground">Drawing Symbols</div>
                <div className="font-bold text-lg">{visualWeldData.totals.visualBW} BW</div>
                <div className="font-bold">{visualWeldData.totals.visualSW} SW</div>
                {visualWeldData.totals.visualFW > 0 && <div className="font-bold">{visualWeldData.totals.visualFW} FW</div>}
              </div>
              <div className="text-center">
                <div className="text-muted-foreground">BOM Inferred</div>
                <div className="font-bold text-lg">{visualWeldData.totals.inferredBW} BW</div>
                <div className="font-bold">{visualWeldData.totals.inferredSW} SW</div>
              </div>
              <div className="text-center">
                <div className="text-muted-foreground">Difference</div>
                <div className={`font-bold text-lg ${visualWeldData.bwDiff === 0 ? 'text-green-400' : Math.abs(visualWeldData.bwDiff) <= 2 ? 'text-yellow-400' : 'text-red-400'}`}>
                  {visualWeldData.bwDiff > 0 ? '+' : ''}{visualWeldData.bwDiff} BW
                </div>
                <div className={`font-bold ${visualWeldData.swDiff === 0 ? 'text-green-400' : 'text-yellow-400'}`}>
                  {visualWeldData.swDiff > 0 ? '+' : ''}{visualWeldData.swDiff} SW
                </div>
              </div>
              <div className="text-center">
                <div className="text-muted-foreground">Status</div>
                <div className={`font-bold text-sm ${visualWeldData.isVerified ? 'text-green-400' : 'text-yellow-400'}`}>
                  {visualWeldData.isVerified ? '\u2713 Verified' : '\u26a0 Review'}
                </div>
              </div>
            </div>
            <p className="text-[10px] text-muted-foreground">
              Weld symbols (● butt, ○ socket) counted from the isometric drawings are compared against welds inferred from the BOM fittings.
              {!visualWeldData.isVerified && " Differences may indicate missed fittings, continuation welds, or field welds not in the BOM."}
            </p>
          </div>
        </div>
      )}

      {/* Detail Breakdown */}
      <div>
        <h3 className="text-sm font-semibold mb-2">Connection Detail</h3>
        <div className="border rounded-lg overflow-hidden max-h-[400px] overflow-y-auto">
          <table className="w-full text-xs">
            <thead className="sticky top-0">
              <tr className="bg-muted/50">
                <th className="text-left p-2 font-semibold">Size</th>
                <th className="text-left p-2 font-semibold">Fitting</th>
                <th className="text-right p-2 font-semibold">Qty</th>
                <th className="text-left p-2 font-semibold">Connection Type</th>
                <th className="text-right p-2 font-semibold">Connections</th>
                <th className="text-left p-2 font-semibold">Location</th>
              </tr>
            </thead>
            <tbody>
              {details
                .sort((a, b) => (parseFloat(a.size) || 0) - (parseFloat(b.size) || 0))
                .map((d, i) => (
                <tr key={i} className={i % 2 === 0 ? "bg-background" : "bg-muted/20"}>
                  <td className="p-2">{d.size}"</td>
                  <td className="p-2 truncate max-w-[200px]" title={d.fitting}>{d.fitting}</td>
                  <td className="p-2 text-right">{d.qty}</td>
                  <td className="p-2">
                    <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${
                      d.connectionType.includes("Butt") ? "bg-orange-500/20 text-orange-400" :
                      d.connectionType.includes("Socket") ? "bg-blue-500/20 text-blue-400" :
                      d.connectionType.includes("Bolt") ? "bg-green-500/20 text-green-400" :
                      "bg-purple-500/20 text-purple-400"
                    }`}>
                      {d.connectionType}
                    </span>
                  </td>
                  <td className="p-2 text-right font-medium">{d.connectionCount}</td>
                  <td className="p-2 capitalize text-muted-foreground">{d.section}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Connections by Page — collapsible accordion */}
      {byPage.length > 0 && (
        <div>
          <button
            className="flex items-center gap-2 text-sm font-semibold mb-2 hover:text-primary transition-colors w-full text-left"
            onClick={() => setByPageOpen(!byPageOpen)}
          >
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={`w-4 h-4 transition-transform ${byPageOpen ? "rotate-90" : ""}`}>
              <polyline points="9 18 15 12 9 6" />
            </svg>
            Connections by Page ({byPage.length} pages)
          </button>

          {byPageOpen && (
            <div className="border rounded-lg overflow-hidden">
              {byPage.map((pageData) => {
                const isExpanded = expandedPages.has(pageData.page);
                const totalConns = pageData.totalBW + pageData.totalSW + pageData.totalBU;
                return (
                  <div key={pageData.page} className="border-b border-border last:border-b-0">
                    <button
                      className="w-full flex items-center justify-between p-2.5 text-xs hover:bg-muted/30 transition-colors"
                      onClick={() => togglePageExpanded(pageData.page)}
                    >
                      <div className="flex items-center gap-2">
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={`w-3 h-3 transition-transform ${isExpanded ? "rotate-90" : ""}`}>
                          <polyline points="9 18 15 12 9 6" />
                        </svg>
                        <span
                          className="font-semibold text-primary underline decoration-dotted cursor-pointer hover:text-primary/80"
                          onClick={(e) => { e.stopPropagation(); setSheetPanelPage(pageData.page); }}
                        >
                          Page {pageData.page}
                        </span>
                        {pageData.drawingNumber && (
                          <span className="text-muted-foreground">| {pageData.drawingNumber}</span>
                        )}
                      </div>
                      <div className="flex items-center gap-3">
                        {pageData.totalBW > 0 && <span className="text-orange-500 font-medium">{pageData.totalBW} BW</span>}
                        {pageData.totalSW > 0 && <span className="text-blue-500 font-medium">{pageData.totalSW} SW</span>}
                        {pageData.totalBU > 0 && <span className="text-green-500 font-medium">{pageData.totalBU} BU</span>}
                        <span className="text-muted-foreground">= {totalConns}</span>
                      </div>
                    </button>

                    {isExpanded && (
                      <div className="px-4 pb-3 bg-muted/10">
                        <table className="w-full text-[11px]">
                          <thead>
                            <tr className="text-left text-muted-foreground">
                              <th className="py-1 pr-2">Fitting</th>
                              <th className="py-1 pr-2">Size</th>
                              <th className="py-1 pr-2 text-right">Qty</th>
                              <th className="py-1 pr-2">Rule</th>
                              <th className="py-1 text-right">Welds</th>
                            </tr>
                          </thead>
                          <tbody>
                            {pageData.fittings.map((f, i) => (
                              <tr key={i} className="border-t border-border/30">
                                <td className="py-1 pr-2 capitalize">{f.category}</td>
                                <td className="py-1 pr-2 font-mono">{f.size}"</td>
                                <td className="py-1 pr-2 text-right">{f.qty}</td>
                                <td className="py-1 pr-2 text-muted-foreground">{f.ruleLabel}</td>
                                <td className="py-1 text-right font-mono">
                                  {f.bw > 0 && <span className="text-orange-500">{f.bw} BW </span>}
                                  {f.sw > 0 && <span className="text-blue-500">{f.sw} SW </span>}
                                  {f.bu > 0 && <span className="text-green-500">{f.bu} BU</span>}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

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
