import type { TakeoffItem } from "@shared/schema";
import { Card, CardContent } from "@/components/ui/card";
import { Package, Ruler, Wrench as WrenchIcon, CircleDot, Layers, Mountain, Pickaxe, Box, Truck, Construction } from "lucide-react";

interface SummaryCardsProps {
  items: TakeoffItem[];
  discipline: "mechanical" | "structural" | "civil";
}

function StatCard({ label, value, unit, icon: Icon, accent }: { label: string; value: string; unit?: string; icon?: any; accent?: string }) {
  return (
    <Card className={`bg-card border-card-border transition-all hover:shadow-md ${accent ? `border-l-4 ${accent}` : ""}`}>
      <CardContent className="p-4">
        <div className="flex items-start justify-between gap-2 mb-1">
          <p className="text-xs text-muted-foreground font-medium">{label}</p>
          {Icon && <Icon size={14} className="text-muted-foreground/60 shrink-0 mt-0.5" />}
        </div>
        <p className="text-xl font-semibold text-foreground tabular-nums">
          {value}
          {unit && <span className="text-sm font-normal text-muted-foreground ml-1">{unit}</span>}
        </p>
      </CardContent>
    </Card>
  );
}

export default function SummaryCards({ items, discipline }: SummaryCardsProps) {
  if (discipline === "mechanical") {
    const byCategory: Record<string, number> = {};
    let totalPipeLf = 0;
    for (const item of items) {
      byCategory[item.category] = (byCategory[item.category] || 0) + item.quantity;
      if (item.category === "pipe") totalPipeLf += item.quantity;
    }
    const fittingCount = (byCategory.elbow || 0) + (byCategory.tee || 0) + (byCategory.reducer || 0) + (byCategory.coupling || 0) + (byCategory.cap || 0) + (byCategory.union || 0);
    return (
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label="Total Items" value={items.length.toLocaleString()} icon={Package} accent="border-l-blue-500" />
        <StatCard label="Pipe" value={totalPipeLf.toFixed(1)} unit="LF" icon={Ruler} accent="border-l-teal-500" />
        <StatCard label="Fittings" value={fittingCount.toLocaleString()} unit="EA" icon={WrenchIcon} accent="border-l-amber-500" />
        <StatCard label="Valves" value={(byCategory.valve || 0).toLocaleString()} unit="EA" icon={CircleDot} accent="border-l-purple-500" />
      </div>
    );
  }

  if (discipline === "structural") {
    const totalWeight = items.reduce((s, i) => s + (i.weight || 0), 0);
    const concreteCY = items.filter(i => ["footing", "grade_beam", "concrete_wall", "slab", "concrete_column"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    const rebarLbs = items.filter(i => i.category === "rebar").reduce((s, i) => s + i.quantity, 0);
    return (
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label="Total Items" value={items.length.toLocaleString()} icon={Package} accent="border-l-blue-500" />
        <StatCard label="Steel Weight" value={(totalWeight / 2000).toFixed(1)} unit="TON" icon={Layers} accent="border-l-purple-500" />
        <StatCard label="Concrete" value={concreteCY.toFixed(1)} unit="CY" icon={Box} accent="border-l-amber-500" />
        <StatCard label="Rebar" value={rebarLbs.toLocaleString()} unit="LBS" icon={Construction} accent="border-l-teal-500" />
      </div>
    );
  }

  if (discipline === "civil") {
    const pipeLf = items.filter(i => ["storm_pipe", "sewer_pipe", "water_pipe", "gas_pipe"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    const earthworkCY = items.filter(i => ["earthwork", "backfill"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    const pavingSf = items.filter(i => ["paving", "concrete_paving", "base_course"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    const structures = items.filter(i => ["manhole", "catch_basin", "fire_hydrant"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    return (
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label="Pipe" value={pipeLf.toFixed(1)} unit="LF" icon={Ruler} accent="border-l-blue-500" />
        <StatCard label="Earthwork" value={earthworkCY.toFixed(1)} unit="CY" icon={Mountain} accent="border-l-amber-500" />
        <StatCard label="Paving" value={pavingSf.toLocaleString()} unit="SF" icon={Truck} accent="border-l-teal-500" />
        <StatCard label="Structures" value={structures.toLocaleString()} unit="EA" icon={Pickaxe} accent="border-l-purple-500" />
      </div>
    );
  }
  return null;
}
