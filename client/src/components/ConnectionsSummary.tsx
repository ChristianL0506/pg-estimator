import { useMemo } from "react";

interface ConnectionsSummaryProps {
  items: any[];
}

// Weld inference rules (mirrors server/routes.ts inferWeldsFromFittings)
const WELD_RULES: Record<string, { welds: number; type: string; label: string }> = {
  elbow: { welds: 2, type: "butt_weld", label: "Butt Welds" },
  tee: { welds: 3, type: "butt_weld", label: "Butt Welds" },
  reducer: { welds: 2, type: "butt_weld", label: "Butt Welds" },
  cap: { welds: 1, type: "butt_weld", label: "Butt Weld" },
  coupling: { welds: 2, type: "socket_weld", label: "Socket Welds" },
  union: { welds: 2, type: "socket_weld", label: "Socket Welds" },
  valve: { welds: 2, type: "butt_weld", label: "Butt Welds" },
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
  const { rows, totals, details } = useMemo(() => {
    const sizeMap: Record<string, { buttWelds: number; socketWelds: number; boltUps: number; threadedConns: number }> = {};
    const detailsList: { size: string; fitting: string; qty: number; connectionType: string; connectionCount: number; section: string }[] = [];

    const ensureSize = (size: string) => {
      if (!sizeMap[size]) sizeMap[size] = { buttWelds: 0, socketWelds: 0, boltUps: 0, threadedConns: 0 };
    };

    for (const item of items) {
      const cat = (item.category || "").toLowerCase();
      const size = item.size || "N/A";
      const qty = item.quantity || 0;
      const desc = (item.description || "").toLowerCase();
      const section = (item.installLocation || (item.notes || "").toLowerCase().includes("field") ? "field" : "shop");

      ensureSize(size);

      // Determine connection type based on description keywords
      const isThreaded = desc.includes("threaded") || desc.includes("screw") || desc.includes("npt") || desc.includes("fnpt") || desc.includes("mnpt");
      const isSocketWeld = desc.includes("socket weld") || desc.includes("sw ") || desc.includes(",sw,") || desc.includes(" sw,") || /\bsw\b/i.test(desc);

      if (cat === "bolt" || cat === "gasket" || desc.includes("bolt") || desc.includes("stud")) {
        // Bolts/studs = bolt-up connections
        // A set of bolts for one flange pair = 1 bolt-up
        // Estimate: bolt qty / typical bolt count per flange (use raw count)
        sizeMap[size].boltUps += qty;
        detailsList.push({ size, fitting: item.description || cat, qty, connectionType: "Bolt-Up", connectionCount: qty, section });
      } else if (cat === "flange") {
        // Each flange has 1 weld to pipe + contributes to a bolt-up connection
        sizeMap[size].buttWelds += qty; // 1 weld per flange
        sizeMap[size].boltUps += Math.ceil(qty / 2); // 2 flanges = 1 bolt-up joint
        detailsList.push({ size, fitting: item.description || "Flange", qty, connectionType: "Butt Weld + Bolt-Up", connectionCount: qty, section });
      } else if (isThreaded) {
        sizeMap[size].threadedConns += qty;
        detailsList.push({ size, fitting: item.description || cat, qty, connectionType: "Threaded", connectionCount: qty, section });
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
  }, [items]);

  if (rows.length === 0) {
    return (
      <div className="text-center py-8 text-muted-foreground text-sm">
        No connections found. Run a takeoff with fittings, flanges, or valves to see weld and bolt-up counts.
      </div>
    );
  }

  return (
    <div className="space-y-6">
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
    </div>
  );
}
