import React, { useState, useMemo } from "react";
import { ChevronDown, ChevronUp, Users, Calculator, Zap, DollarSign, Flame, Info, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";

function fmt$(n: number) { return `$${n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
function fmtK$(n: number) { return n >= 10000 ? `$${(n / 1000).toFixed(1)}k` : fmt$(n); }

// ============================================================
// Picou Group Actuals (Stolthaven Phase 6)
// ============================================================

const ROLE_RATES = {
  welders:        { rate: 35, perDiem: 75,  label: "Welders",        short: "Weld", productive: true },
  fitters:        { rate: 35, perDiem: 75,  label: "Fitters",        short: "Fit",  productive: true },
  helpers:        { rate: 25, perDiem: 75,  label: "Helpers",        short: "Help", productive: true },
  firewatches:    { rate: 20, perDiem: 0,   label: "Firewatches",    short: "FW",   productive: false },
  foreman:        { rate: 35, perDiem: 100, label: "Foreman",        short: "Fore", productive: false },
  superintendent: { rate: 45, perDiem: 100, label: "Superintendent", short: "Supt", productive: false },
  safety:         { rate: 30, perDiem: 75,  label: "Safety",         short: "Safe", productive: false },
  equipmentOp:    { rate: 30, perDiem: 75,  label: "Equipment Op",   short: "Equip", productive: false },
} as const;

type RoleKey = keyof typeof ROLE_RATES;
const ALL_ROLES: RoleKey[] = ["welders", "fitters", "helpers", "firewatches", "foreman", "superintendent", "safety", "equipmentOp"];

type PipeSizeMix = "large" | "mixed" | "small";

interface CrewCounts extends Record<RoleKey, number> {
  hoursPerDay: number;
}

interface CrewPlannerProps {
  totalLaborHours: number;
  laborRate: number;
  perDiemRate: number;
}

// ============================================================
// Cost calculation helpers
// ============================================================

function calcRoleSpecificCosts(counts: CrewCounts, durationDays: number) {
  let totalLabor = 0;
  let totalPerDiem = 0;
  const breakdown: { role: string; count: number; laborCost: number; perDiemCost: number; rate: number; perDiemRate: number }[] = [];

  for (const role of ALL_ROLES) {
    const info = ROLE_RATES[role];
    const count = counts[role];
    if (count <= 0) continue;
    const laborCost = count * counts.hoursPerDay * info.rate * durationDays;
    const perDiemCost = count * info.perDiem * durationDays;
    totalLabor += laborCost;
    totalPerDiem += perDiemCost;
    breakdown.push({ role: info.label, count, laborCost, perDiemCost, rate: info.rate, perDiemRate: info.perDiem });
  }

  return { totalLabor, totalPerDiem, totalCost: totalLabor + totalPerDiem, breakdown };
}

function calcBlendedCosts(counts: CrewCounts, durationDays: number, totalLaborHours: number, laborRate: number, perDiemRate: number) {
  const totalCrew = ALL_ROLES.reduce((s, r) => s + counts[r], 0);
  const totalLaborCost = totalLaborHours * laborRate;
  const totalPerDiem = totalCrew * durationDays * perDiemRate;
  return { totalLaborCost, totalPerDiem, totalCost: totalLaborCost + totalPerDiem, totalCrew };
}

function getProductiveWorkers(c: CrewCounts): number {
  return ALL_ROLES.filter(r => ROLE_RATES[r].productive).reduce((s, r) => s + c[r], 0);
}

function getTotalCrew(c: CrewCounts): number {
  return ALL_ROLES.reduce((s, r) => s + c[r], 0);
}

function getDuration(c: CrewCounts, totalLaborHours: number): { days: number; weeks: number } {
  const productive = getProductiveWorkers(c);
  const mhPerDay = productive * c.hoursPerDay;
  if (mhPerDay <= 0) return { days: 0, weeks: 0 };
  const days = Math.ceil(totalLaborHours / mhPerDay);
  return { days, weeks: Math.ceil(days / 5) };
}

// ============================================================
// Scenario builder
// ============================================================

function buildScenario(
  name: string,
  tag: string,
  description: string,
  targetWeeks: [number, number],
  totalLaborHours: number,
  pipeMix: PipeSizeMix,
  areas: number,
  safetyRequired: boolean,
  hoursPerDay: number,
  daysPerWeek: number,
): ScenarioResult {
  const targetDays = Math.round(((targetWeeks[0] + targetWeeks[1]) / 2) * daysPerWeek);
  const mhPerDayNeeded = totalLaborHours / targetDays;

  // Welder:fitter ratio based on pipe size mix
  const welderFitterRatio = pipeMix === "large" ? 1.75 : pipeMix === "small" ? 1.0 : 1.35;

  // Solve: welders * h + (welders / ratio) * h + helpers * h = mhPerDayNeeded
  // helpers = ceil((welders + fitters) / 3)
  // Let w = welders, f = w / ratio, helper = ceil((w + f) / 3)
  // (w + f + helper) * h = mhPerDay => w + w/ratio + ceil((w + w/ratio)/3) = mhPerDay/h
  const productivePerWelderUnit = 1 + 1 / welderFitterRatio;
  const withHelpers = productivePerWelderUnit + productivePerWelderUnit / 3;
  const welders = Math.max(2, Math.round(mhPerDayNeeded / (hoursPerDay * withHelpers)));
  const fitters = Math.max(1, Math.round(welders / welderFitterRatio));
  const helpers = Math.max(1, Math.ceil((welders + fitters) / 3));
  const firewatches = areas;
  const foreman = areas;
  const equipmentOp = areas;
  const superintendent = 1;
  const safety = safetyRequired ? 1 : 0;

  const counts: CrewCounts = {
    welders, fitters, helpers, firewatches, foreman,
    superintendent, safety, equipmentOp, hoursPerDay,
  };

  const { days, weeks } = getDuration(counts, totalLaborHours);
  const roleSpecific = calcRoleSpecificCosts(counts, days);

  return {
    name, tag, description, crew: counts,
    totalCrew: getTotalCrew(counts),
    productiveMHPerDay: getProductiveWorkers(counts) * hoursPerDay,
    durationDays: days,
    durationWeeks: weeks,
    roleSpecific,
  };
}

interface ScenarioResult {
  name: string;
  tag: string;
  description: string;
  crew: CrewCounts;
  totalCrew: number;
  productiveMHPerDay: number;
  durationDays: number;
  durationWeeks: number;
  roleSpecific: ReturnType<typeof calcRoleSpecificCosts>;
}

// ============================================================
// Component
// ============================================================

export default function CrewPlanner({ totalLaborHours, laborRate, perDiemRate }: CrewPlannerProps) {
  const [expanded, setExpanded] = useState(false);

  // --- Inputs ---
  const [areas, setAreas] = useState(1);
  const [pipeMix, setPipeMix] = useState<PipeSizeMix>("mixed");
  const [safetyRequired, setSafetyRequired] = useState(true);
  const [hoursPerDay, setHoursPerDay] = useState(10);
  const [daysPerWeek, setDaysPerWeek] = useState(5);
  const [useRoleRates, setUseRoleRates] = useState(true);

  // --- Custom crew ---
  const [crew, setCrew] = useState<CrewCounts>({
    welders: 0, fitters: 0, helpers: 0, firewatches: 0,
    foreman: 1, superintendent: 1, safety: 1, equipmentOp: 0,
    hoursPerDay: 10,
  });
  const [showCustom, setShowCustom] = useState(false);

  function updateCrew(field: keyof CrewCounts, value: number) {
    setCrew(prev => ({ ...prev, [field]: Math.max(0, value) }));
  }

  // --- Custom calc ---
  const customResult = useMemo(() => {
    const productive = getProductiveWorkers(crew);
    const total = getTotalCrew(crew);
    const mhPerDay = productive * crew.hoursPerDay;
    const { days, weeks } = getDuration(crew, totalLaborHours);

    const roleSpecific = calcRoleSpecificCosts(crew, days);
    const blended = calcBlendedCosts(crew, days, totalLaborHours, laborRate, perDiemRate);

    const dailyLaborRole = days > 0 ? roleSpecific.totalLabor / days : 0;
    const dailyPerDiemRole = days > 0 ? roleSpecific.totalPerDiem / days : 0;

    return {
      productive, total, mhPerDay, days, weeks,
      roleSpecific, blended,
      dailyLaborRole, dailyPerDiemRole,
      dailyTotal: dailyLaborRole + dailyPerDiemRole,
    };
  }, [crew, totalLaborHours, laborRate, perDiemRate]);

  // --- Scenarios ---
  const scenarios = useMemo((): ScenarioResult[] => {
    if (totalLaborHours <= 0) return [];
    return [
      buildScenario("Lean", "lean", "Smaller crew, 6-8 week duration. Lower per diem, longer schedule.", [6, 8], totalLaborHours, pipeMix, areas, safetyRequired, hoursPerDay, daysPerWeek),
      buildScenario("Standard", "standard", "Balanced crew, 4-6 weeks. Recommended for most projects.", [4, 6], totalLaborHours, pipeMix, areas, safetyRequired, hoursPerDay, daysPerWeek),
      buildScenario("Aggressive", "aggressive", "Large crew, 2-3 weeks. Higher per diem but fast turnaround.", [2, 3], totalLaborHours, pipeMix, areas, safetyRequired, hoursPerDay, daysPerWeek),
    ];
  }, [totalLaborHours, pipeMix, areas, safetyRequired, hoursPerDay, daysPerWeek]);

  const handleExportCrewPlan = async () => {
    try {
      const scenarioData = scenarios.map(s => ({
        name: s.name,
        duration: s.durationDays,
        dailyBurn: s.roleSpecific.totalCost / Math.max(1, s.durationDays),
        totalCost: s.roleSpecific.totalCost,
        roles: s.roleSpecific.breakdown.map(b => ({
          name: b.role, count: b.count, rate: b.rate, perDiem: b.perDiemRate,
        })),
      }));
      const customCrewData = showCustom ? {
        roles: ALL_ROLES.map(r => ({
          name: ROLE_RATES[r].label, count: crew[r], rate: ROLE_RATES[r].rate, perDiem: ROLE_RATES[r].perDiem,
        })),
      } : undefined;
      const rateCardData = ALL_ROLES.map(r => ({
        name: ROLE_RATES[r].label, rate: ROLE_RATES[r].rate, perDiemEligible: ROLE_RATES[r].perDiem > 0,
      }));
      const res = await fetch("/api/export-crew-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ projectName: "Estimate", scenarios: scenarioData, customCrew: customCrewData, rateCard: rateCardData }),
      });
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "Crew Plan.xlsx";
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (e) { console.error("Crew plan export failed", e); }
  };

  return (
    <Card className="border-card-border">
      <CardHeader
        className="p-4 pb-2 cursor-pointer select-none"
        onClick={() => setExpanded(!expanded)}
      >
        <div className="flex items-center justify-between">
          <CardTitle className="text-sm flex items-center gap-2">
            <Users size={15} className="text-primary" />
            Crew Planner
          </CardTitle>
          <div className="flex items-center gap-2">
            {expanded && scenarios.length > 0 && (
              <Button
                variant="outline"
                size="sm"
                className="h-6 text-[10px] text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                onClick={e => { e.stopPropagation(); handleExportCrewPlan(); }}
              >
                <FileSpreadsheet size={12} className="mr-1" />
                Export
              </Button>
            )}
            {expanded ? <ChevronUp size={16} className="text-muted-foreground" /> : <ChevronDown size={16} className="text-muted-foreground" />}
          </div>
        </div>
        {!expanded && (
          <p className="text-[10px] text-muted-foreground mt-0.5">
            Plan crew size, calculate duration, and compare scenarios
          </p>
        )}
      </CardHeader>

      {expanded && (
        <CardContent className="p-4 pt-0 space-y-4">
          {/* ===== Summary bar ===== */}
          <div className="text-xs text-muted-foreground flex flex-wrap gap-x-3 gap-y-0.5">
            <span>Total estimated: <span className="font-semibold text-foreground">{totalLaborHours.toFixed(1)} MH</span></span>
            <span>Blended rate: <span className="font-semibold text-foreground">{fmt$(laborRate)}/hr</span></span>
            <span>Per diem: <span className="font-semibold text-foreground">{fmt$(perDiemRate)}/day</span></span>
          </div>

          {/* ===== Project Parameters ===== */}
          <div className="space-y-3">
            <Label className="text-xs font-semibold">Project Parameters</Label>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div>
                <Label className="text-[10px] text-muted-foreground">Number of Areas</Label>
                <Input
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  type="number" min={1} max={20}
                  value={areas}
                  onChange={e => setAreas(Math.max(1, parseInt(e.target.value) || 1))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Hours / Day</Label>
                <Input
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  type="number" min={4} max={14}
                  value={hoursPerDay}
                  onChange={e => setHoursPerDay(Math.max(4, Math.min(14, parseInt(e.target.value) || 10)))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Days / Week</Label>
                <Input
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  type="number" min={4} max={7}
                  value={daysPerWeek}
                  onChange={e => setDaysPerWeek(Math.max(4, Math.min(7, parseInt(e.target.value) || 5)))}
                />
              </div>
              <div className="flex items-end gap-2 pb-0.5">
                <div className="flex items-center gap-2">
                  <Switch
                    id="safety-toggle"
                    checked={safetyRequired}
                    onCheckedChange={setSafetyRequired}
                  />
                  <Label htmlFor="safety-toggle" className="text-xs cursor-pointer">Safety Man</Label>
                </div>
              </div>
            </div>

            {/* Pipe size mix */}
            <div>
              <Label className="text-[10px] text-muted-foreground mb-1.5 block">Pipe Size Mix (sets welder:fitter ratio)</Label>
              <RadioGroup
                value={pipeMix}
                onValueChange={(v) => setPipeMix(v as PipeSizeMix)}
                className="flex flex-wrap gap-4"
              >
                {([
                  { value: "large", label: 'Mostly Large (6"+)', sub: "1.75 welders per fitter" },
                  { value: "mixed", label: "Mixed Sizes", sub: "1.35 welders per fitter" },
                  { value: "small", label: 'Mostly Small (under 6")', sub: "1:1 ratio" },
                ] as const).map(opt => (
                  <div key={opt.value} className="flex items-center gap-2">
                    <RadioGroupItem value={opt.value} id={`pipe-${opt.value}`} />
                    <Label htmlFor={`pipe-${opt.value}`} className="text-xs cursor-pointer">
                      {opt.label}
                      <span className="text-[10px] text-muted-foreground ml-1">({opt.sub})</span>
                    </Label>
                  </div>
                ))}
              </RadioGroup>
            </div>

            {/* Role rates toggle */}
            <div className="flex items-center gap-2">
              <Switch
                id="role-rates-toggle"
                checked={useRoleRates}
                onCheckedChange={setUseRoleRates}
              />
              <Label htmlFor="role-rates-toggle" className="text-xs cursor-pointer">
                Use Role-Specific Rates
                <span className="text-[10px] text-muted-foreground ml-1">
                  {useRoleRates ? "(Picou Group actuals)" : `(blended ${fmt$(laborRate)}/hr for all)`}
                </span>
              </Label>
            </div>
          </div>

          <Separator />

          {/* ===== Optimal Crew Suggestions ===== */}
          {totalLaborHours > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Zap size={14} className="text-yellow-500" />
                <Label className="text-xs font-semibold">Crew Scenarios</Label>
                <span className="text-[10px] text-muted-foreground">
                  ({areas} area{areas > 1 ? "s" : ""}, {pipeMix} pipe, {hoursPerDay}h/day, {daysPerWeek}d/wk)
                </span>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                {scenarios.map(scenario => {
                  const blended = calcBlendedCosts(scenario.crew, scenario.durationDays, totalLaborHours, laborRate, perDiemRate);
                  const costs = useRoleRates ? scenario.roleSpecific : { totalLabor: blended.totalLaborCost, totalPerDiem: blended.totalPerDiem, totalCost: blended.totalCost, breakdown: [] };
                  const dailyLabor = scenario.durationDays > 0 ? costs.totalLabor / scenario.durationDays : 0;
                  const dailyPerDiem = scenario.durationDays > 0 ? costs.totalPerDiem / scenario.durationDays : 0;

                  return (
                    <Card
                      key={scenario.name}
                      className={`border-card-border ${scenario.tag === "standard" ? "ring-1 ring-primary/40" : ""}`}
                    >
                      <CardHeader className="p-3 pb-1">
                        <div className="flex items-center justify-between">
                          <CardTitle className="text-xs font-semibold">{scenario.name}</CardTitle>
                          {scenario.tag === "standard" && (
                            <Badge variant="default" className="text-[9px] px-1.5 py-0">Recommended</Badge>
                          )}
                        </div>
                        <p className="text-[10px] text-muted-foreground">{scenario.description}</p>
                      </CardHeader>
                      <CardContent className="p-3 pt-0 space-y-2">
                        {/* Crew grid: all 8 roles + total */}
                        <div className="grid grid-cols-3 gap-1 text-[10px]">
                          {ALL_ROLES.map(role => (
                            <div key={role} className="text-center bg-muted/50 rounded p-1">
                              <span className="text-muted-foreground">{ROLE_RATES[role].short}</span>
                              <p className={`font-semibold ${scenario.crew[role] === 0 ? "text-muted-foreground/50" : ""}`}>
                                {scenario.crew[role]}
                              </p>
                            </div>
                          ))}
                          <div className="text-center bg-primary/10 rounded p-1">
                            <span className="text-muted-foreground">Total</span>
                            <p className="font-semibold text-primary">{scenario.totalCrew}</p>
                          </div>
                        </div>

                        <Separator className="my-1" />

                        {/* Duration + Daily burn */}
                        <div className="space-y-1 text-[10px]">
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">Duration</span>
                            <span className="font-semibold">{scenario.durationDays} days ({scenario.durationWeeks} wks)</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">MH/Day</span>
                            <span className="font-mono">{scenario.productiveMHPerDay}</span>
                          </div>
                        </div>

                        <Separator className="my-1" />

                        {/* Daily burn rate */}
                        <div className="bg-muted/30 rounded p-1.5 text-[10px]">
                          <div className="flex items-center gap-1 mb-1">
                            <Flame size={10} className="text-orange-500" />
                            <span className="font-semibold">Daily Burn Rate</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">Labor</span>
                            <span className="font-mono text-orange-600 dark:text-orange-400">{fmtK$(dailyLabor)}/day</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">Per Diem</span>
                            <span className="font-mono text-blue-600 dark:text-blue-400">{fmtK$(dailyPerDiem)}/day</span>
                          </div>
                          <div className="flex justify-between font-semibold border-t border-border/50 mt-0.5 pt-0.5">
                            <span>Total</span>
                            <span className="font-mono text-primary">{fmtK$(dailyLabor + dailyPerDiem)}/day</span>
                          </div>
                        </div>

                        <Separator className="my-1" />

                        {/* Cost breakdown */}
                        {useRoleRates && scenario.roleSpecific.breakdown.length > 0 && (
                          <div className="space-y-0.5 text-[10px]">
                            <span className="font-semibold text-muted-foreground flex items-center gap-1">
                              <DollarSign size={10} /> Crew Cost Breakdown
                            </span>
                            {/* Group: Journeymen */}
                            {(() => {
                              const journeymen = scenario.roleSpecific.breakdown.filter(b => b.role === "Welders" || b.role === "Fitters");
                              const helperRows = scenario.roleSpecific.breakdown.filter(b => b.role === "Helpers");
                              const fwRows = scenario.roleSpecific.breakdown.filter(b => b.role === "Firewatches");
                              const supervision = scenario.roleSpecific.breakdown.filter(b => b.role === "Foreman" || b.role === "Superintendent");
                              const support = scenario.roleSpecific.breakdown.filter(b => b.role === "Safety" || b.role === "Equipment Op");
                              const groups = [
                                { label: "Journeymen", rows: journeymen },
                                { label: "Helpers", rows: helperRows },
                                { label: "Firewatches", rows: fwRows },
                                { label: "Supervision", rows: supervision },
                                { label: "Support", rows: support },
                              ].filter(g => g.rows.length > 0);

                              return groups.map(g => {
                                const gLabor = g.rows.reduce((s, r) => s + r.laborCost, 0);
                                const gPerDiem = g.rows.reduce((s, r) => s + r.perDiemCost, 0);
                                return (
                                  <div key={g.label} className="flex justify-between">
                                    <span className="text-muted-foreground">
                                      {g.label}
                                      <span className="opacity-60 ml-0.5">
                                        ({g.rows.map(r => `${r.count} ${r.role}`).join(" + ")})
                                      </span>
                                    </span>
                                    <span className="font-mono">{fmtK$(gLabor + gPerDiem)}</span>
                                  </div>
                                );
                              });
                            })()}

                            {/* Per diem note */}
                            <div className="flex items-start gap-1 mt-1 text-muted-foreground opacity-75">
                              <Info size={9} className="mt-0.5 shrink-0" />
                              <span>Firewatches: no per diem. Foreman/Supt: $100/day. Others: $75/day.</span>
                            </div>
                          </div>
                        )}

                        {/* Totals */}
                        <div className="space-y-1 text-[10px]">
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">Labor</span>
                            <span className="font-mono text-orange-600 dark:text-orange-400">{fmt$(costs.totalLabor)}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-muted-foreground">Per Diem</span>
                            <span className="font-mono text-blue-600 dark:text-blue-400">{fmt$(costs.totalPerDiem)}</span>
                          </div>
                          <Separator className="my-0.5" />
                          <div className="flex justify-between font-semibold">
                            <span>Total Crew Cost</span>
                            <span className="font-mono text-primary">{fmt$(costs.totalCost)}</span>
                          </div>
                        </div>

                        {/* Use this crew button */}
                        <Button
                          size="sm" variant="ghost"
                          className="w-full text-[10px] h-6 mt-1"
                          onClick={() => {
                            setCrew(scenario.crew);
                            setShowCustom(true);
                          }}
                        >
                          Use This Crew
                        </Button>
                      </CardContent>
                    </Card>
                  );
                })}
              </div>
            </div>
          )}

          <Separator />

          {/* ===== Custom Crew Configuration ===== */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <Label className="text-xs font-semibold">Custom Crew Configuration</Label>
              <Button
                size="sm" variant="ghost" className="text-[10px] h-6"
                onClick={() => setShowCustom(!showCustom)}
              >
                {showCustom ? "Hide" : "Show"}
              </Button>
            </div>

            {showCustom && (
              <div className="space-y-3">
                <div className="grid grid-cols-3 sm:grid-cols-5 md:grid-cols-9 gap-2">
                  {ALL_ROLES.map(role => (
                    <div key={role}>
                      <Label className="text-[10px] text-muted-foreground">{ROLE_RATES[role].label}</Label>
                      <Input
                        className="h-8 text-xs mt-0.5 font-mono text-center"
                        type="number" min={0}
                        value={crew[role]}
                        onChange={e => updateCrew(role, parseInt(e.target.value) || 0)}
                        data-testid={`crew-${role}`}
                      />
                      {useRoleRates && (
                        <p className="text-[9px] text-muted-foreground text-center mt-0.5">
                          ${ROLE_RATES[role].rate}/hr
                        </p>
                      )}
                    </div>
                  ))}
                  <div>
                    <Label className="text-[10px] text-muted-foreground">Hrs/Day</Label>
                    <Input
                      className="h-8 text-xs mt-0.5 font-mono text-center"
                      type="number" min={4} max={14}
                      value={crew.hoursPerDay}
                      onChange={e => updateCrew("hoursPerDay", Math.max(4, parseInt(e.target.value) || 10))}
                      data-testid="crew-hoursPerDay"
                    />
                  </div>
                </div>

                {/* Calculate */}
                <Button
                  size="sm" variant="outline"
                  onClick={() => setShowCustom(true)}
                  disabled={customResult.productive === 0}
                  className="gap-1.5"
                >
                  <Calculator size={13} />
                  Calculate
                </Button>

                {customResult.productive === 0 && (
                  <p className="text-xs text-destructive">Add at least one welder, fitter, or helper to calculate.</p>
                )}

                {/* Custom Results */}
                {customResult.mhPerDay > 0 && (
                  <Card className="bg-muted/30 border-primary/20">
                    <CardContent className="p-4 space-y-3">
                      <p className="text-sm font-semibold text-foreground">
                        Duration:{" "}
                        <span className="text-primary">{customResult.days} working days</span>
                        {" "}(<span className="text-primary">{customResult.weeks} weeks</span>)
                      </p>

                      {/* Daily burn */}
                      <div className="bg-muted/50 rounded-md p-2 text-xs">
                        <div className="flex items-center gap-1.5 mb-1">
                          <Flame size={12} className="text-orange-500" />
                          <span className="font-semibold">Daily Burn Rate</span>
                        </div>
                        <span className="font-mono text-orange-600 dark:text-orange-400">{fmtK$(customResult.dailyLaborRole)}/day</span>
                        <span className="text-muted-foreground mx-1">labor</span>
                        <span className="text-muted-foreground">+</span>
                        <span className="font-mono text-blue-600 dark:text-blue-400 mx-1">{fmtK$(customResult.dailyPerDiemRole)}/day</span>
                        <span className="text-muted-foreground">per diem</span>
                        <span className="text-muted-foreground mx-1">=</span>
                        <span className="font-mono font-semibold text-primary">{fmtK$(customResult.dailyTotal)}/day total</span>
                      </div>

                      {/* Stats grid */}
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                        <div className="bg-background rounded p-2 text-center">
                          <p className="text-[10px] text-muted-foreground">Daily MH Burn</p>
                          <p className="text-sm font-semibold font-mono">{customResult.mhPerDay.toFixed(0)} MH/day</p>
                        </div>
                        <div className="bg-background rounded p-2 text-center">
                          <p className="text-[10px] text-muted-foreground">Total Labor</p>
                          <p className="text-sm font-semibold font-mono text-orange-600 dark:text-orange-400">
                            {fmt$(useRoleRates ? customResult.roleSpecific.totalLabor : customResult.blended.totalLaborCost)}
                          </p>
                        </div>
                        <div className="bg-background rounded p-2 text-center">
                          <p className="text-[10px] text-muted-foreground">Total Per Diem</p>
                          <p className="text-sm font-semibold font-mono text-blue-600 dark:text-blue-400">
                            {fmt$(useRoleRates ? customResult.roleSpecific.totalPerDiem : customResult.blended.totalPerDiem)}
                          </p>
                        </div>
                        <div className="bg-background rounded p-2 text-center">
                          <p className="text-[10px] text-muted-foreground">Total Crew Cost</p>
                          <p className="text-sm font-semibold font-mono text-primary">
                            {fmt$(useRoleRates ? customResult.roleSpecific.totalCost : customResult.blended.totalCost)}
                          </p>
                        </div>
                      </div>

                      {/* Role-specific breakdown */}
                      {useRoleRates && customResult.roleSpecific.breakdown.length > 0 && (
                        <div className="text-[10px] space-y-0.5">
                          <p className="font-semibold text-muted-foreground flex items-center gap-1">
                            <DollarSign size={10} /> Role-Specific Breakdown ({customResult.days} days)
                          </p>
                          <div className="grid grid-cols-[1fr_auto_auto_auto] gap-x-3 gap-y-0.5">
                            <span className="font-semibold text-muted-foreground">Role</span>
                            <span className="font-semibold text-muted-foreground text-right">Labor</span>
                            <span className="font-semibold text-muted-foreground text-right">Per Diem</span>
                            <span className="font-semibold text-muted-foreground text-right">Total</span>
                            {customResult.roleSpecific.breakdown.map(row => (
                              <React.Fragment key={row.role}>
                                <span className="text-muted-foreground">
                                  {row.count} {row.role} @ ${row.rate}/hr
                                </span>
                                <span className="font-mono text-right">{fmtK$(row.laborCost)}</span>
                                <span className="font-mono text-right">{row.perDiemRate > 0 ? fmtK$(row.perDiemCost) : <span className="text-muted-foreground/50">none</span>}</span>
                                <span className="font-mono text-right">{fmtK$(row.laborCost + row.perDiemCost)}</span>
                              </React.Fragment>
                            ))}
                          </div>
                        </div>
                      )}

                      {/* Blended comparison */}
                      {useRoleRates && (
                        <div className="text-[10px] text-muted-foreground mt-2 border-t border-border/50 pt-2">
                          <span className="font-semibold">Blended rate comparison:</span>{" "}
                          Using {fmt$(laborRate)}/hr for all roles would give{" "}
                          <span className="font-mono">{fmt$(customResult.blended.totalCost)}</span> total
                          {customResult.roleSpecific.totalCost !== customResult.blended.totalCost && (
                            <span>
                              {" "}({customResult.roleSpecific.totalCost < customResult.blended.totalCost ? "saves" : "costs"}{" "}
                              <span className="font-mono">{fmt$(Math.abs(customResult.roleSpecific.totalCost - customResult.blended.totalCost))}</span>
                              {" "}vs role-specific)
                            </span>
                          )}
                        </div>
                      )}

                      <div className="text-[10px] text-muted-foreground">
                        Total crew: {customResult.total} | Productive: {customResult.productive} |{" "}
                        {daysPerWeek} day weeks
                      </div>
                    </CardContent>
                  </Card>
                )}
              </div>
            )}
          </div>
        </CardContent>
      )}
    </Card>
  );
}
