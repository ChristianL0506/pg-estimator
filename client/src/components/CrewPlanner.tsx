import { useState, useMemo } from "react";
import { ChevronDown, ChevronUp, Users, Calculator, Zap } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Label } from "@/components/ui/label";

function fmt$(n: number) { return `$${n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }

interface CrewPlannerProps {
  totalLaborHours: number;
  laborRate: number;
  perDiemRate: number;
}

interface CrewConfig {
  welders: number;
  fitters: number;
  helpers: number;
  foreman: number;
  safety: number;
  operators: number;
  hoursPerDay: number;
}

interface ScenarioResult {
  name: string;
  description: string;
  crew: CrewConfig;
  totalCrew: number;
  productiveMHPerDay: number;
  durationDays: number;
  durationWeeks: number;
  totalLaborCost: number;
  totalPerDiem: number;
  totalCrewCost: number;
}

export default function CrewPlanner({ totalLaborHours, laborRate, perDiemRate }: CrewPlannerProps) {
  const [expanded, setExpanded] = useState(false);
  const [crew, setCrew] = useState<CrewConfig>({
    welders: 0,
    fitters: 0,
    helpers: 0,
    foreman: 1,
    safety: 1,
    operators: 0,
    hoursPerDay: 10,
  });
  const [showResults, setShowResults] = useState(false);

  function updateCrew(field: keyof CrewConfig, value: number) {
    setCrew(prev => ({ ...prev, [field]: value }));
  }

  // Custom crew calculations
  const customResult = useMemo(() => {
    const totalCrew = crew.welders + crew.fitters + crew.helpers + crew.foreman + crew.safety + crew.operators;
    const productiveWorkers = crew.welders + crew.fitters + crew.helpers;
    const productiveMHPerDay = productiveWorkers * crew.hoursPerDay;
    const durationDays = productiveMHPerDay > 0 ? Math.ceil(totalLaborHours / productiveMHPerDay) : 0;
    const durationWeeks = durationDays > 0 ? Math.ceil(durationDays / 5) : 0;
    const totalLaborCost = totalLaborHours * laborRate;
    const totalPerDiem = totalCrew * durationDays * perDiemRate;
    const totalCrewCost = totalLaborCost + totalPerDiem;

    return {
      totalCrew,
      productiveWorkers,
      productiveMHPerDay,
      durationDays,
      durationWeeks,
      totalLaborCost,
      totalPerDiem,
      totalCrewCost,
    };
  }, [crew, totalLaborHours, laborRate, perDiemRate]);

  // Optimal crew scenarios
  const scenarios = useMemo((): ScenarioResult[] => {
    if (totalLaborHours <= 0) return [];

    function buildScenario(name: string, description: string, welders: number, hoursPerDay: number): ScenarioResult {
      // Ratio: 1 welder : 1 fitter : 0.5 helper
      const fitters = welders;
      const helpers = Math.max(1, Math.ceil(welders * 0.5));
      const totalWorkers = welders + fitters + helpers;
      const foreman = Math.max(1, Math.ceil(totalWorkers / 7));
      const safety = totalWorkers > 6 ? 1 : 0;
      const operators = 0;
      const totalCrew = welders + fitters + helpers + foreman + safety + operators;

      const productiveWorkers = welders + fitters + helpers;
      const productiveMHPerDay = productiveWorkers * hoursPerDay;
      const durationDays = productiveMHPerDay > 0 ? Math.ceil(totalLaborHours / productiveMHPerDay) : 0;
      const durationWeeks = durationDays > 0 ? Math.ceil(durationDays / 5) : 0;
      const totalLaborCost = totalLaborHours * laborRate;
      const totalPerDiem = totalCrew * durationDays * perDiemRate;
      const totalCrewCost = totalLaborCost + totalPerDiem;

      return {
        name,
        description,
        crew: { welders, fitters, helpers, foreman, safety, operators, hoursPerDay },
        totalCrew,
        productiveMHPerDay,
        durationDays,
        durationWeeks,
        totalLaborCost,
        totalPerDiem,
        totalCrewCost,
      };
    }

    // Calculate base welder count from total hours: target ~4-6 week project for standard
    const targetDaysStandard = 25; // ~5 weeks
    const stdHoursPerDay = 10;
    const stdWelders = Math.max(2, Math.round(totalLaborHours / (targetDaysStandard * stdHoursPerDay * 2.5))); // 2.5x: each welder + 1 fitter + 0.5 helper

    const leanWelders = Math.max(1, Math.ceil(stdWelders * 0.5));
    const aggressiveWelders = Math.max(3, Math.ceil(stdWelders * 1.8));

    return [
      buildScenario("Lean", "Minimum crew, longer duration — lower per diem but extended schedule", leanWelders, stdHoursPerDay),
      buildScenario("Standard", "Balanced crew size and duration — recommended approach", stdWelders, stdHoursPerDay),
      buildScenario("Aggressive", "Larger crew, faster completion — higher per diem but shorter schedule", aggressiveWelders, stdHoursPerDay),
    ];
  }, [totalLaborHours, laborRate, perDiemRate]);

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
          {expanded ? <ChevronUp size={16} className="text-muted-foreground" /> : <ChevronDown size={16} className="text-muted-foreground" />}
        </div>
        {!expanded && (
          <p className="text-[10px] text-muted-foreground mt-0.5">
            Plan crew size, calculate duration, and compare scenarios
          </p>
        )}
      </CardHeader>

      {expanded && (
        <CardContent className="p-4 pt-0 space-y-4">
          <div className="text-xs text-muted-foreground">
            Total estimated labor hours: <span className="font-semibold text-foreground">{totalLaborHours.toFixed(1)} MH</span>
            {" | "}Labor rate: <span className="font-semibold text-foreground">{fmt$(laborRate)}/hr</span>
            {" | "}Per diem: <span className="font-semibold text-foreground">{fmt$(perDiemRate)}/day</span>
          </div>

          {/* Crew input row */}
          <div>
            <Label className="text-xs font-semibold">Custom Crew Configuration</Label>
            <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-7 gap-2 mt-1.5">
              {([
                { field: "welders" as const, label: "Welders" },
                { field: "fitters" as const, label: "Fitters" },
                { field: "helpers" as const, label: "Helpers/FW" },
                { field: "foreman" as const, label: "Foreman" },
                { field: "safety" as const, label: "Safety" },
                { field: "operators" as const, label: "Operators" },
                { field: "hoursPerDay" as const, label: "Hrs/Day" },
              ]).map(({ field, label }) => (
                <div key={field}>
                  <Label className="text-[10px] text-muted-foreground">{label}</Label>
                  <Input
                    className="h-8 text-xs mt-0.5 font-mono text-center"
                    type="number"
                    min={0}
                    value={crew[field]}
                    onChange={e => updateCrew(field, parseFloat(e.target.value) || 0)}
                    data-testid={`crew-${field}`}
                  />
                </div>
              ))}
            </div>
          </div>

          {/* Calculate button */}
          <Button size="sm" variant="outline" onClick={() => setShowResults(true)} className="gap-1.5">
            <Calculator size={13} />
            Calculate
          </Button>

          {/* Custom Results */}
          {showResults && customResult.productiveMHPerDay > 0 && (
            <Card className="bg-muted/30 border-primary/20">
              <CardContent className="p-4 space-y-3">
                <p className="text-sm font-semibold text-foreground">
                  This project will take approximately{" "}
                  <span className="text-primary">{customResult.durationDays} working days</span>
                  {" "}(<span className="text-primary">{customResult.durationWeeks} weeks</span>) with this crew
                </p>

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  <div className="bg-background rounded p-2 text-center">
                    <p className="text-[10px] text-muted-foreground">Daily MH Burn</p>
                    <p className="text-sm font-semibold font-mono">{customResult.productiveMHPerDay.toFixed(0)} MH/day</p>
                  </div>
                  <div className="bg-background rounded p-2 text-center">
                    <p className="text-[10px] text-muted-foreground">Total Labor Cost</p>
                    <p className="text-sm font-semibold font-mono text-orange-600 dark:text-orange-400">{fmt$(customResult.totalLaborCost)}</p>
                  </div>
                  <div className="bg-background rounded p-2 text-center">
                    <p className="text-[10px] text-muted-foreground">Total Per Diem</p>
                    <p className="text-sm font-semibold font-mono text-blue-600 dark:text-blue-400">{fmt$(customResult.totalPerDiem)}</p>
                  </div>
                  <div className="bg-background rounded p-2 text-center">
                    <p className="text-[10px] text-muted-foreground">Total Crew Cost</p>
                    <p className="text-sm font-semibold font-mono text-primary">{fmt$(customResult.totalCrewCost)}</p>
                  </div>
                </div>

                <div className="text-[10px] text-muted-foreground">
                  Total crew: {customResult.totalCrew} | Productive workers: {customResult.productiveWorkers} | 
                  Per diem: {customResult.totalCrew} crew × {customResult.durationDays} days × {fmt$(perDiemRate)}/day
                </div>
              </CardContent>
            </Card>
          )}

          <Separator />

          {/* Optimal Crew Suggestions */}
          {totalLaborHours > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Zap size={14} className="text-yellow-500" />
                <Label className="text-xs font-semibold">Optimal Crew Suggestions</Label>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                {scenarios.map(scenario => (
                  <Card
                    key={scenario.name}
                    className={`border-card-border ${scenario.name === "Standard" ? "ring-1 ring-primary/40" : ""}`}
                  >
                    <CardHeader className="p-3 pb-1">
                      <div className="flex items-center justify-between">
                        <CardTitle className="text-xs font-semibold">{scenario.name}</CardTitle>
                        {scenario.name === "Standard" && (
                          <Badge variant="default" className="text-[9px] px-1.5 py-0">Recommended</Badge>
                        )}
                      </div>
                      <p className="text-[10px] text-muted-foreground">{scenario.description}</p>
                    </CardHeader>
                    <CardContent className="p-3 pt-0 space-y-2">
                      {/* Crew breakdown */}
                      <div className="grid grid-cols-3 gap-1 text-[10px]">
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Weld</span>
                          <p className="font-semibold">{scenario.crew.welders}</p>
                        </div>
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Fit</span>
                          <p className="font-semibold">{scenario.crew.fitters}</p>
                        </div>
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Help</span>
                          <p className="font-semibold">{scenario.crew.helpers}</p>
                        </div>
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Fore</span>
                          <p className="font-semibold">{scenario.crew.foreman}</p>
                        </div>
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Safety</span>
                          <p className="font-semibold">{scenario.crew.safety}</p>
                        </div>
                        <div className="text-center bg-muted/50 rounded p-1">
                          <span className="text-muted-foreground">Total</span>
                          <p className="font-semibold text-primary">{scenario.totalCrew}</p>
                        </div>
                      </div>

                      <Separator className="my-1" />

                      {/* Duration + Cost */}
                      <div className="space-y-1 text-[10px]">
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Duration</span>
                          <span className="font-semibold">{scenario.durationDays} days ({scenario.durationWeeks} wks)</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">MH/Day</span>
                          <span className="font-mono">{scenario.productiveMHPerDay}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Labor</span>
                          <span className="font-mono text-orange-600 dark:text-orange-400">{fmt$(scenario.totalLaborCost)}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Per Diem</span>
                          <span className="font-mono text-blue-600 dark:text-blue-400">{fmt$(scenario.totalPerDiem)}</span>
                        </div>
                        <Separator className="my-0.5" />
                        <div className="flex justify-between font-semibold">
                          <span>Total Crew Cost</span>
                          <span className="font-mono text-primary">{fmt$(scenario.totalCrewCost)}</span>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      )}
    </Card>
  );
}
