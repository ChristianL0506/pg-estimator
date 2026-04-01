import type { TakeoffItem } from "@shared/schema";
import { Card, CardContent } from "@/components/ui/card";

interface SummaryCardsProps {
  items: TakeoffItem[];
  discipline: "mechanical" | "structural" | "civil";
}

function StatCard({ label, value, unit }: { label: string; value: string; unit?: string }) {
  return (
    <Card className="bg-card border-card-border">
      <CardContent className="p-4">
        <p className="text-xs text-muted-foreground mb-1">{label}</p>
        <p className="text-xl font-semibold text-foreground">
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
    return (
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label="Total Items" value={items.length.toLocaleString()} />
        <StatCard label="Pipe" value={totalPipeLf.toFixed(1)} unit="LF" />
        <StatCard label="Fittings" value={((byCategory.elbow || 0) + (byCategory.tee || 0) + (byCategory.reducer || 0) + (byCategory.coupling || 0) + (byCategory.cap || 0) + (byCategory.union || 0)).toLocaleString()} unit="EA" />
        <StatCard label="Valves" value={(byCategory.valve || 0).toLocaleString()} unit="EA" />
      </div>
    );
  }

  if (discipline === "structural") {
    const totalWeight = items.reduce((s, i) => s + (i.weight || 0), 0);
    const concreteCY = items.filter(i => ["footing", "grade_beam", "concrete_wall", "slab", "concrete_column"].includes(i.category)).reduce((s, i) => s + i.quantity, 0);
    const rebarLbs = items.filter(i => i.category === "rebar").reduce((s, i) => s + i.quantity, 0);
    return (
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label="Total Items" value={items.length.toLocaleString()} />
        <StatCard label="Steel Weight" value={(totalWeight / 2000).toFixed(1)} unit="TON" />
        <StatCard label="Concrete" value={concreteCY.toFixed(1)} unit="CY" />
        <StatCard label="Rebar" value={rebarLbs.toLocaleString()} unit="LBS" />
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
        <StatCard label="Pipe" value={pipeLf.toFixed(1)} unit="LF" />
        <StatCard label="Earthwork" value={earthworkCY.toFixed(1)} unit="CY" />
        <StatCard label="Paving" value={pavingSf.toFixed(0)} unit="SF" />
        <StatCard label="Structures" value={structures.toLocaleString()} unit="EA" />
      </div>
    );
  }

  console.warn(`SummaryCards: unknown discipline "${discipline}"`);
  return null;
}
