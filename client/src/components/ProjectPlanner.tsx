import { useState, useMemo, useCallback, useEffect } from "react";
import { Calendar, ClipboardList, AlertTriangle, CheckCircle2, ChevronDown, ChevronUp, Play, Diamond, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Label } from "@/components/ui/label";
import { Slider } from "@/components/ui/slider";

// ============================================================
// Types
// ============================================================

interface ProjectPlannerProps {
  totalManhours: number;
  crewSize?: number;
  hoursPerDay?: number;
  daysPerWeek?: number;
  projectName?: string;
  items?: any[];
}

interface ActivityDef {
  id: string;
  name: string;
  phase: Phase;
  predecessors: string[];
  manhourPct: number;       // fraction of total MH (0.0-1.0), 0 = fixed duration
  fixedDays?: number;       // fixed duration (overrides MH calc)
  lagPct?: number;          // 0-1, allows starting when predecessor is this % done
  crew: string;
  efficiency: number;
  isMilestone: boolean;
  crewFraction: number;     // fraction of total crew assigned to this activity
}

interface Activity extends ActivityDef {
  durationDays: number;
  manhours: number;
  startDay: number;
  endDay: number;
  earlyStart: number;
  earlyFinish: number;
  lateStart: number;
  lateFinish: number;
  totalFloat: number;
  isOnCriticalPath: boolean;
}

type Phase = "Mobilization" | "Fabrication" | "Installation" | "Testing" | "Closeout";

// ============================================================
// Default WBS Activities
// ============================================================

const DEFAULT_ACTIVITIES: ActivityDef[] = [
  // Phase 1: Mobilization & Setup (3-5% of MH)
  { id: "mob",              name: "Site Mobilization",               phase: "Mobilization", predecessors: [],                    manhourPct: 0,    fixedDays: 4,  crew: "All",          efficiency: 1.0,  isMilestone: false, crewFraction: 0.3, lagPct: 0 },
  { id: "safety_orient",    name: "Safety Orientation & Training",   phase: "Mobilization", predecessors: ["mob"],               manhourPct: 0,    fixedDays: 2,  crew: "All",          efficiency: 1.0,  isMilestone: false, crewFraction: 1.0, lagPct: 0 },
  { id: "material_staging", name: "Material Receiving & Staging",    phase: "Mobilization", predecessors: ["mob"],               manhourPct: 0,    fixedDays: 4,  crew: "Helpers/Op",   efficiency: 1.0,  isMilestone: false, crewFraction: 0.2, lagPct: 0 },
  { id: "scaffold",         name: "Scaffolding & Access Setup",      phase: "Mobilization", predecessors: ["safety_orient"],     manhourPct: 0.06, fixedDays: undefined, crew: "Scaffold",   efficiency: 0.70, isMilestone: false, crewFraction: 0.25, lagPct: 0 },

  // Phase 2: Fabrication (25-35% of MH)
  { id: "shop_fab",   name: "Shop Fabrication (Spool Pieces)",  phase: "Fabrication", predecessors: ["material_staging"],  manhourPct: 0.25, crew: "Fitters/Welders", efficiency: 0.85, isMilestone: false, crewFraction: 0.6, lagPct: 0 },
  { id: "qc_fab",     name: "Fabrication QC / NDE",             phase: "Fabrication", predecessors: ["shop_fab"],          manhourPct: 0.03, crew: "QC",             efficiency: 0.80, isMilestone: false, crewFraction: 0.1, lagPct: 0.5 },

  // Phase 3: Field Installation (40-50% of MH)
  { id: "support_install", name: "Pipe Support Installation",    phase: "Installation", predecessors: ["scaffold"],                      manhourPct: 0.08, crew: "Fitters/Helpers", efficiency: 0.75, isMilestone: false, crewFraction: 0.3,  lagPct: 0 },
  { id: "large_bore",      name: "Large Bore Pipe Install (6\"+)", phase: "Installation", predecessors: ["support_install", "shop_fab"],  manhourPct: 0.20, crew: "Fitters/Welders", efficiency: 0.70, isMilestone: false, crewFraction: 0.7,  lagPct: 0 },
  { id: "small_bore",      name: "Small Bore Pipe Install (<6\")", phase: "Installation", predecessors: ["large_bore"],                   manhourPct: 0.15, crew: "Fitters/Welders", efficiency: 0.65, isMilestone: false, crewFraction: 0.5,  lagPct: 0.3 },
  { id: "field_weld",      name: "Field Welding",                 phase: "Installation", predecessors: ["large_bore"],                   manhourPct: 0,    fixedDays: undefined,  crew: "Welders",         efficiency: 0.70, isMilestone: false, crewFraction: 0.4,  lagPct: 0 },
  { id: "bolt_up",         name: "Bolt-Up & Torquing",            phase: "Installation", predecessors: ["large_bore"],                   manhourPct: 0.05, crew: "Fitters",         efficiency: 0.80, isMilestone: false, crewFraction: 0.3,  lagPct: 0 },
  { id: "instruments",     name: "Instrument Connections",        phase: "Installation", predecessors: ["small_bore"],                   manhourPct: 0.03, crew: "Instrument",      efficiency: 0.70, isMilestone: false, crewFraction: 0.15, lagPct: 0 },

  // Phase 4: Testing & Commissioning (8-12% of MH)
  { id: "pressure_test", name: "Hydrostatic / Pressure Testing",  phase: "Testing", predecessors: ["bolt_up", "small_bore", "field_weld"], manhourPct: 0.05, crew: "Fitters/Helpers", efficiency: 0.75, isMilestone: false, crewFraction: 0.4, lagPct: 0 },
  { id: "punch_list",    name: "Punch List & Repairs",            phase: "Testing", predecessors: ["pressure_test"],                        manhourPct: 0.03, crew: "All Trades",      efficiency: 0.65, isMilestone: false, crewFraction: 0.3, lagPct: 0 },
  { id: "insulation",    name: "Insulation & Heat Tracing",       phase: "Testing", predecessors: ["pressure_test"],                        manhourPct: 0.05, crew: "Insulators",      efficiency: 0.70, isMilestone: false, crewFraction: 0.25, lagPct: 0 },

  // Phase 5: Closeout (2-3% of MH)
  { id: "turnover", name: "System Turnover & Documentation", phase: "Closeout", predecessors: ["punch_list", "insulation"], manhourPct: 0, fixedDays: 3,  crew: "Supervision", efficiency: 1.0, isMilestone: true,  crewFraction: 0.1, lagPct: 0 },
  { id: "demob",    name: "Demobilization",                  phase: "Closeout", predecessors: ["turnover"],                 manhourPct: 0, fixedDays: 3,  crew: "All",         efficiency: 1.0, isMilestone: false, crewFraction: 0.3, lagPct: 0 },
];

// ============================================================
// Phase colors
// ============================================================

const PHASE_COLORS: Record<Phase, { bar: string; bg: string; text: string; border: string }> = {
  Mobilization: { bar: "bg-blue-500",    bg: "bg-blue-50 dark:bg-blue-950/30",     text: "text-blue-700 dark:text-blue-300",     border: "border-blue-200 dark:border-blue-800" },
  Fabrication:  { bar: "bg-emerald-500", bg: "bg-emerald-50 dark:bg-emerald-950/30", text: "text-emerald-700 dark:text-emerald-300", border: "border-emerald-200 dark:border-emerald-800" },
  Installation: { bar: "bg-orange-500",  bg: "bg-orange-50 dark:bg-orange-950/30",  text: "text-orange-700 dark:text-orange-300",  border: "border-orange-200 dark:border-orange-800" },
  Testing:      { bar: "bg-purple-500",  bg: "bg-purple-50 dark:bg-purple-950/30",  text: "text-purple-700 dark:text-purple-300",  border: "border-purple-200 dark:border-purple-800" },
  Closeout:     { bar: "bg-gray-500",    bg: "bg-gray-50 dark:bg-gray-900/30",      text: "text-gray-700 dark:text-gray-300",      border: "border-gray-200 dark:border-gray-700" },
};

// ============================================================
// Schedule calculation
// ============================================================

function computeSchedule(
  defs: ActivityDef[],
  totalMH: number,
  crewSize: number,
  hoursPerDay: number,
  overallEfficiency: number,
): Activity[] {
  const map = new Map<string, Activity>();

  // Step 1: compute durations
  for (const def of defs) {
    let manhours = def.manhourPct * totalMH;
    let durationDays: number;

    if (def.fixedDays != null && def.fixedDays > 0) {
      durationDays = def.fixedDays;
    } else if (def.manhourPct > 0) {
      const crewForActivity = Math.max(1, Math.round(crewSize * def.crewFraction));
      const effectiveRate = crewForActivity * hoursPerDay * def.efficiency * overallEfficiency;
      durationDays = effectiveRate > 0 ? Math.max(1, Math.ceil(manhours / effectiveRate)) : 1;
    } else {
      // field_weld: concurrent with large_bore — match its duration
      durationDays = 0; // will be resolved after predecessors
      manhours = 0;
    }

    map.set(def.id, {
      ...def,
      durationDays,
      manhours,
      startDay: 0,
      endDay: 0,
      earlyStart: 0,
      earlyFinish: 0,
      lateStart: 0,
      lateFinish: 0,
      totalFloat: 0,
      isOnCriticalPath: false,
    });
  }

  // Resolve field_weld: match largest predecessor duration
  const fw = map.get("field_weld");
  if (fw && fw.durationDays === 0) {
    const predDurations = fw.predecessors.map(pid => map.get(pid)?.durationDays || 0);
    fw.durationDays = Math.max(1, ...predDurations);
  }

  // Step 2: Forward pass — earliest start/finish
  const order = topologicalSort(defs);
  for (const id of order) {
    const act = map.get(id)!;
    if (act.predecessors.length === 0) {
      act.earlyStart = 0;
    } else {
      let maxFinish = 0;
      for (const pid of act.predecessors) {
        const pred = map.get(pid);
        if (!pred) continue;
        const lagDays = (act.lagPct || 0) > 0 ? Math.floor(pred.durationDays * (1 - (act.lagPct || 0))) : pred.durationDays;
        const predFinish = pred.earlyStart + lagDays;
        if (predFinish > maxFinish) maxFinish = predFinish;
      }
      act.earlyStart = maxFinish;
    }
    act.earlyFinish = act.earlyStart + act.durationDays;
    act.startDay = act.earlyStart;
    act.endDay = act.earlyFinish;
  }

  // Step 3: Backward pass — latest start/finish
  const projectEnd = Math.max(...Array.from(map.values()).map(a => a.earlyFinish));
  // Initialize all late finishes to project end
  for (const act of map.values()) {
    act.lateFinish = projectEnd;
    act.lateStart = projectEnd - act.durationDays;
  }

  const reverseOrder = [...order].reverse();
  for (const id of reverseOrder) {
    const act = map.get(id)!;
    // Find all successors
    for (const other of map.values()) {
      if (other.predecessors.includes(id)) {
        const lagDays = (other.lagPct || 0) > 0 ? Math.floor(act.durationDays * (1 - (other.lagPct || 0))) : act.durationDays;
        const latestFinishForPred = other.lateStart - lagDays + act.durationDays;
        if (latestFinishForPred < act.lateFinish) {
          act.lateFinish = latestFinishForPred;
          act.lateStart = act.lateFinish - act.durationDays;
        }
      }
    }
    act.totalFloat = act.lateStart - act.earlyStart;
    act.isOnCriticalPath = act.totalFloat <= 0;
  }

  return order.map(id => map.get(id)!);
}

function topologicalSort(defs: ActivityDef[]): string[] {
  const result: string[] = [];
  const visited = new Set<string>();
  const visiting = new Set<string>();
  const defMap = new Map(defs.map(d => [d.id, d]));

  function visit(id: string) {
    if (visited.has(id)) return;
    if (visiting.has(id)) return; // cycle guard
    visiting.add(id);
    const def = defMap.get(id);
    if (def) {
      for (const pid of def.predecessors) visit(pid);
    }
    visiting.delete(id);
    visited.add(id);
    result.push(id);
  }

  for (const def of defs) visit(def.id);
  return result;
}

function addWorkingDays(start: Date, workDays: number, daysPerWeek: number): Date {
  const d = new Date(start);
  let added = 0;
  while (added < workDays) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (daysPerWeek >= 7 || (daysPerWeek >= 6 && dow !== 0) || (daysPerWeek <= 5 && dow !== 0 && dow !== 6)) {
      added++;
    }
  }
  return d;
}

function fmtDate(d: Date): string {
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

// ============================================================
// Component
// ============================================================

export default function ProjectPlanner({
  totalManhours,
  crewSize: propCrew,
  hoursPerDay: propHours,
  daysPerWeek: propDays,
  projectName,
}: ProjectPlannerProps) {
  const [expanded, setExpanded] = useState(false);

  // Parameters
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    return d.toISOString().split("T")[0];
  });
  const [mhInput, setMhInput] = useState(totalManhours);
  const [crewInput, setCrewInput] = useState(propCrew || 12);
  const [hpd, setHpd] = useState(propHours || 10);
  const [dpw, setDpw] = useState(propDays || 5);
  const [efficiency, setEfficiency] = useState(75);
  const [generated, setGenerated] = useState(false);

  // Sync manhours from props when user hasn't generated yet
  useEffect(() => {
    if (totalManhours > 0 && !generated) {
      setMhInput(totalManhours);
    }
  }, [totalManhours, generated]);

  const handleGenerate = useCallback(() => {
    setMhInput(totalManhours > 0 ? totalManhours : mhInput);
    setGenerated(true);
  }, [totalManhours, mhInput]);

  const schedule = useMemo((): Activity[] => {
    if (!generated) return [];
    const mh = mhInput > 0 ? mhInput : totalManhours;
    if (mh <= 0) return [];
    return computeSchedule(DEFAULT_ACTIVITIES, mh, crewInput, hpd, efficiency / 100);
  }, [generated, mhInput, totalManhours, crewInput, hpd, efficiency]);

  const projectEnd = useMemo(() => Math.max(0, ...schedule.map(a => a.endDay)), [schedule]);
  const criticalPathDays = useMemo(() => {
    const cpActs = schedule.filter(a => a.isOnCriticalPath);
    return cpActs.length > 0 ? Math.max(0, ...cpActs.map(a => a.endDay)) : projectEnd;
  }, [schedule, projectEnd]);

  const totalScheduledMH = useMemo(() => schedule.reduce((s, a) => s + a.manhours, 0), [schedule]);
  const startDateObj = useMemo(() => new Date(startDate + "T00:00:00"), [startDate]);
  const endDateObj = useMemo(() => addWorkingDays(startDateObj, projectEnd, dpw), [startDateObj, projectEnd, dpw]);

  const crewUtilization = useMemo(() => {
    if (projectEnd <= 0 || crewInput <= 0) return 0;
    const availableMH = crewInput * hpd * projectEnd;
    return availableMH > 0 ? Math.min(100, (totalScheduledMH / availableMH) * 100) : 0;
  }, [projectEnd, crewInput, hpd, totalScheduledMH]);

  // Gantt week markers
  const weekCount = Math.ceil(projectEnd / dpw);

  const handleExportPlan = async () => {
    try {
      const activitiesData = schedule.map(a => ({
        id: a.id, name: a.name, phase: a.phase, duration: a.durationDays,
        manhours: a.manhours, crew: a.crew, predecessors: a.predecessors,
        startDate: fmtDate(addWorkingDays(startDateObj, a.startDay, dpw)),
        endDate: fmtDate(addWorkingDays(startDateObj, a.endDay, dpw)),
        float: a.totalFloat, critical: a.isOnCriticalPath, milestone: a.isMilestone,
      }));
      const summaryData = {
        totalDuration: projectEnd,
        criticalPathDuration: criticalPathDays,
        totalManhours: totalScheduledMH,
        completionDate: fmtDate(endDateObj),
        utilization: Math.round(crewUtilization),
      };
      const res = await fetch("/api/export-project-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ projectName: projectName || "Project Plan", activities: activitiesData, summary: summaryData }),
      });
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${projectName || "Project Plan"} - Project Plan.xlsx`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (e) { console.error("Project plan export failed", e); }
  };

  return (
    <Card className="border-card-border">
      <CardHeader
        className="p-4 pb-2 cursor-pointer select-none"
        onClick={() => setExpanded(!expanded)}
      >
        <div className="flex items-center justify-between">
          <CardTitle className="text-sm flex items-center gap-2">
            <Calendar size={15} className="text-primary" />
            Project Planner
            {projectName && <span className="text-xs font-normal text-muted-foreground ml-1">({projectName})</span>}
          </CardTitle>
          <div className="flex items-center gap-2">
            {expanded && schedule.length > 0 && (
              <Button
                variant="outline"
                size="sm"
                className="h-6 text-[10px] text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                onClick={e => { e.stopPropagation(); handleExportPlan(); }}
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
            Build a WBS schedule with task sequencing, critical path, and timeline
          </p>
        )}
      </CardHeader>

      {expanded && (
        <CardContent className="p-4 pt-0 space-y-4">
          {/* ===== Parameters ===== */}
          <div className="space-y-3">
            <Label className="text-xs font-semibold">Project Parameters</Label>
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-3">
              <div>
                <Label className="text-[10px] text-muted-foreground">Start Date</Label>
                <Input
                  type="date"
                  className="h-8 text-xs mt-0.5"
                  value={startDate}
                  onChange={e => setStartDate(e.target.value)}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Total Manhours</Label>
                <Input
                  type="number" min={0}
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  value={mhInput}
                  onChange={e => setMhInput(Math.max(0, parseFloat(e.target.value) || 0))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Crew Size (productive)</Label>
                <Input
                  type="number" min={1}
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  value={crewInput}
                  onChange={e => setCrewInput(Math.max(1, parseInt(e.target.value) || 1))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Hours/Day</Label>
                <Input
                  type="number" min={4} max={14}
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  value={hpd}
                  onChange={e => setHpd(Math.max(4, Math.min(14, parseInt(e.target.value) || 10)))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Days/Week</Label>
                <Input
                  type="number" min={4} max={7}
                  className="h-8 text-xs mt-0.5 font-mono text-center"
                  value={dpw}
                  onChange={e => setDpw(Math.max(4, Math.min(7, parseInt(e.target.value) || 5)))}
                />
              </div>
              <div>
                <Label className="text-[10px] text-muted-foreground">Efficiency: {efficiency}%</Label>
                <Slider
                  className="mt-2.5"
                  min={50} max={100} step={5}
                  value={[efficiency]}
                  onValueChange={v => setEfficiency(v[0])}
                />
              </div>
            </div>

            <Button size="sm" onClick={handleGenerate} className="gap-1.5">
              <Play size={13} />
              Generate Schedule
            </Button>
          </div>

          {generated && schedule.length > 0 && (
            <>
              <Separator />

              {/* ===== Summary Cards ===== */}
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                <SummaryCard label="Total Duration" value={`${Math.ceil(projectEnd / dpw)} weeks`} sub={`${projectEnd} working days`} icon={<Calendar size={13} className="text-primary" />} />
                <SummaryCard label="Critical Path" value={`${Math.ceil(criticalPathDays / dpw)} weeks`} sub={`${criticalPathDays} days`} icon={<AlertTriangle size={13} className="text-red-500" />} />
                <SummaryCard label="Total Manhours" value={`${Math.round(totalScheduledMH).toLocaleString()} MH`} sub={`${Math.round(mhInput).toLocaleString()} input`} icon={<ClipboardList size={13} className="text-orange-500" />} />
                <SummaryCard label="Est. Completion" value={fmtDate(endDateObj)} sub={`Start ${fmtDate(startDateObj)}`} icon={<CheckCircle2 size={13} className="text-emerald-500" />} />
                <SummaryCard label="Crew Utilization" value={`${crewUtilization.toFixed(0)}%`} sub={`${crewInput} workers`} icon={<Play size={13} className="text-indigo-500" />} />
              </div>

              <Separator />

              {/* ===== Activity Table ===== */}
              <div>
                <Label className="text-xs font-semibold mb-2 block">Activity Schedule</Label>
                <div className="border border-border rounded-md overflow-auto max-h-[400px]">
                  <table className="w-full text-[10px]">
                    <thead className="bg-muted/50 sticky top-0">
                      <tr>
                        <th className="text-left px-2 py-1.5 font-semibold">#</th>
                        <th className="text-left px-2 py-1.5 font-semibold">Activity</th>
                        <th className="text-left px-2 py-1.5 font-semibold">Phase</th>
                        <th className="text-right px-2 py-1.5 font-semibold">Duration</th>
                        <th className="text-right px-2 py-1.5 font-semibold">MH</th>
                        <th className="text-left px-2 py-1.5 font-semibold">Crew</th>
                        <th className="text-left px-2 py-1.5 font-semibold">Predecessors</th>
                        <th className="text-right px-2 py-1.5 font-semibold">Start</th>
                        <th className="text-right px-2 py-1.5 font-semibold">End</th>
                        <th className="text-right px-2 py-1.5 font-semibold">Float</th>
                        <th className="text-center px-2 py-1.5 font-semibold">Critical</th>
                      </tr>
                    </thead>
                    <tbody>
                      {schedule.map((act, idx) => {
                        const pc = PHASE_COLORS[act.phase];
                        const startD = addWorkingDays(startDateObj, act.startDay, dpw);
                        const endD = addWorkingDays(startDateObj, act.endDay, dpw);
                        return (
                          <tr
                            key={act.id}
                            className={`border-t border-border/50 ${act.isOnCriticalPath ? "bg-red-50/50 dark:bg-red-950/20" : ""} hover:bg-muted/30`}
                          >
                            <td className="px-2 py-1.5 text-muted-foreground">{idx + 1}</td>
                            <td className="px-2 py-1.5 font-medium">
                              <div className="flex items-center gap-1.5">
                                {act.isMilestone && <Diamond size={8} className="text-yellow-500 shrink-0" />}
                                {act.name}
                              </div>
                            </td>
                            <td className="px-2 py-1.5">
                              <Badge variant="outline" className={`text-[8px] px-1 py-0 ${pc.text} ${pc.border}`}>
                                {act.phase}
                              </Badge>
                            </td>
                            <td className="px-2 py-1.5 text-right font-mono">{act.durationDays}d</td>
                            <td className="px-2 py-1.5 text-right font-mono">{act.manhours > 0 ? Math.round(act.manhours).toLocaleString() : "-"}</td>
                            <td className="px-2 py-1.5 text-muted-foreground">{act.crew}</td>
                            <td className="px-2 py-1.5 text-muted-foreground">{act.predecessors.join(", ") || "-"}</td>
                            <td className="px-2 py-1.5 text-right font-mono">{fmtDate(startD)}</td>
                            <td className="px-2 py-1.5 text-right font-mono">{fmtDate(endD)}</td>
                            <td className="px-2 py-1.5 text-right font-mono">{act.totalFloat}d</td>
                            <td className="px-2 py-1.5 text-center">
                              {act.isOnCriticalPath ? (
                                <span className="inline-block w-2 h-2 rounded-full bg-red-500" title="Critical path" />
                              ) : (
                                <span className="inline-block w-2 h-2 rounded-full bg-muted-foreground/20" />
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              <Separator />

              {/* ===== Gantt Chart ===== */}
              <div>
                <Label className="text-xs font-semibold mb-2 block">Visual Timeline</Label>
                <div className="border border-border rounded-md overflow-auto">
                  <div className="min-w-[600px]">
                    {/* Week headers */}
                    <div className="flex border-b border-border bg-muted/30">
                      <div className="w-40 shrink-0 px-2 py-1 text-[9px] font-semibold text-muted-foreground border-r border-border">
                        Activity
                      </div>
                      <div className="flex-1 flex">
                        {Array.from({ length: Math.max(1, weekCount) }, (_, i) => {
                          const weekStart = addWorkingDays(startDateObj, i * dpw, dpw);
                          return (
                            <div
                              key={i}
                              className="text-center text-[8px] text-muted-foreground py-1 border-r border-border/30"
                              style={{ width: `${100 / Math.max(1, weekCount)}%` }}
                            >
                              Wk {i + 1}
                              <div className="text-[7px] opacity-60">{fmtDate(weekStart)}</div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Activity bars */}
                    {schedule.map(act => {
                      const pc = PHASE_COLORS[act.phase];
                      const leftPct = projectEnd > 0 ? (act.startDay / projectEnd) * 100 : 0;
                      const widthPct = projectEnd > 0 ? Math.max(1, (act.durationDays / projectEnd) * 100) : 0;

                      return (
                        <div key={act.id} className="flex border-b border-border/30 hover:bg-muted/20 group">
                          <div className="w-40 shrink-0 px-2 py-1.5 text-[9px] border-r border-border truncate flex items-center gap-1">
                            {act.isMilestone && <Diamond size={7} className="text-yellow-500 shrink-0" />}
                            <span className={act.isOnCriticalPath ? "font-semibold text-red-600 dark:text-red-400" : ""}>
                              {act.name}
                            </span>
                          </div>
                          <div className="flex-1 relative py-1">
                            {act.isMilestone ? (
                              <div
                                className="absolute top-1/2 -translate-y-1/2 w-3 h-3 rotate-45 bg-yellow-500 border border-yellow-600"
                                style={{ left: `${leftPct}%` }}
                                title={`${act.name}: Day ${act.startDay}`}
                              />
                            ) : (
                              <div
                                className={`absolute top-1/2 -translate-y-1/2 h-4 rounded-sm ${pc.bar} ${
                                  act.isOnCriticalPath ? "ring-2 ring-red-500 ring-offset-1 ring-offset-background" : ""
                                } opacity-80 group-hover:opacity-100 transition-opacity`}
                                style={{ left: `${leftPct}%`, width: `${widthPct}%` }}
                                title={`${act.name}: Day ${act.startDay}-${act.endDay} (${act.durationDays}d)`}
                              >
                                {widthPct > 8 && (
                                  <span className="absolute inset-0 flex items-center justify-center text-[7px] text-white font-medium truncate px-1">
                                    {act.durationDays}d
                                  </span>
                                )}
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}

                    {/* Legend */}
                    <div className="flex flex-wrap gap-3 px-3 py-2 bg-muted/20 border-t border-border">
                      {(Object.entries(PHASE_COLORS) as [Phase, typeof PHASE_COLORS[Phase]][]).map(([phase, colors]) => (
                        <div key={phase} className="flex items-center gap-1 text-[8px] text-muted-foreground">
                          <span className={`w-2.5 h-2.5 rounded-sm ${colors.bar}`} />
                          {phase}
                        </div>
                      ))}
                      <div className="flex items-center gap-1 text-[8px] text-muted-foreground">
                        <span className="w-2.5 h-2.5 rounded-sm bg-gray-300 ring-1 ring-red-500" />
                        Critical Path
                      </div>
                      <div className="flex items-center gap-1 text-[8px] text-muted-foreground">
                        <Diamond size={8} className="text-yellow-500" />
                        Milestone
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}

          {!generated && (
            <div className="text-center py-6 text-xs text-muted-foreground">
              <ClipboardList size={24} className="mx-auto mb-2 opacity-40" />
              Set parameters above and click Generate Schedule to build a project plan
              {totalManhours > 0 && (
                <p className="mt-1 text-primary font-medium">{Math.round(totalManhours).toLocaleString()} MH available from estimate</p>
              )}
            </div>
          )}
        </CardContent>
      )}
    </Card>
  );
}

// ============================================================
// Summary Card sub-component
// ============================================================

function SummaryCard({ label, value, sub, icon }: { label: string; value: string; sub: string; icon: React.ReactNode }) {
  return (
    <div className="bg-muted/30 rounded-md p-2.5 text-center">
      <div className="flex items-center justify-center gap-1 mb-1">{icon}</div>
      <p className="text-[10px] text-muted-foreground">{label}</p>
      <p className="text-sm font-semibold font-mono">{value}</p>
      <p className="text-[9px] text-muted-foreground">{sub}</p>
    </div>
  );
}
