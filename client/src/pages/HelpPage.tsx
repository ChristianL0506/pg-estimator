// Help & How-To page.
// Single comprehensive resource for new and returning users. Three sections:
//   1. Getting Started — narrative onboarding walkthrough
//   2. How To — task-based guides for common workflows
//   3. Reference — quick-reference cards for rates, weld math, calibration data

import { useState, useMemo } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import {
  BookOpen, Rocket, Search, Wrench, Calculator, Users, FolderOpen, Download,
  Image as ImageIcon, AlertCircle, CheckCircle2, ArrowRight, Lightbulb, FileText, Settings, Database, ClipboardList, Activity
} from "lucide-react";
import AppLayout from "@/components/AppLayout";

// ── Data: Getting Started steps ─────────────────────────────────────────────
const GETTING_STARTED_STEPS = [
  {
    title: "Sign in",
    body: "Use your admin / picougroup credentials at the login screen. The session cookie keeps you signed in across page reloads.",
    icon: CheckCircle2,
  },
  {
    title: "Pick a discipline",
    body: "Mechanical for piping ISOs (most common), Structural for steel takeoffs, or Civil for site & concrete. Each discipline has its own takeoff page.",
    icon: Wrench,
  },
  {
    title: "Upload a PDF",
    body: "Drag and drop or click the upload zone. Multi-page packages are supported. Toggle 'Drawings have revisions' if your set has revision clouds. Toggle 'Dual-Model Verification' to cross-check Claude with Gemini (when available).",
    icon: FileText,
  },
  {
    title: "Wait for extraction",
    body: "Progress bar shows pages processed and items found. Large packages (>100 pages) stream chunk-by-chunk so you see results as they come in. Walk away — extraction continues in the background.",
    icon: Activity,
  },
  {
    title: "Review the BOM",
    body: "Use the View Pages section to confirm sheets were read correctly. Edit any mis-extracted rows inline. Items with ⚠ flags (e.g., qty=0 because the AI couldn't read the cell) need manual entry from the drawing.",
    icon: ClipboardList,
  },
  {
    title: "Export",
    body: "BOM export gives you the full item list. Connections export shows shop welds vs 40-foot field welds split by size. Pivot export gives you size totals. All formatted for Excel.",
    icon: Download,
  },
];

// ── Data: How-To topics ─────────────────────────────────────────────────────
type HowTo = {
  id: string;
  title: string;
  category: "Takeoff" | "Estimating" | "Project Mgmt" | "Reference";
  icon: any;
  steps: { title?: string; body: string }[];
  tips?: string[];
};

const HOW_TOS: HowTo[] = [
  {
    id: "run-takeoff",
    category: "Takeoff",
    title: "Run a takeoff from a PDF",
    icon: Rocket,
    steps: [
      { title: "Open the discipline page", body: "Click Mechanical (or Structural / Civil) in the sidebar." },
      { title: "Drop the PDF onto the upload zone", body: "You can drag-and-drop or click to browse. The file uploads immediately and chunking begins." },
      { title: "Set toggles before clicking Process", body: "If your drawings have REV C clouds, enable 'Drawings have revisions'. If you want Gemini to cross-check Claude, enable 'Dual-Model Verification' (Gemini API key required)." },
      { title: "Watch the progress", body: "Page count, items found, and any warnings update in real time. Each chunk takes 1-3 minutes depending on page density." },
      { title: "Open the project when done", body: "Once 'Complete' appears, click into the project. The BOM, Connections, Pivot, and Fab Scope tabs are all populated." },
    ],
    tips: [
      "For very large packages (200+ pages), let it run for 30-60 minutes. The streaming mode saves results after each chunk so you don't lose work.",
      "If extraction shows lots of warnings, that's a flag for sheets that need a manual look-over. Use the Page Viewer to verify.",
    ],
  },
  {
    id: "review-flagged-pipes",
    category: "Takeoff",
    title: "Manually enter pipe quantities flagged for review",
    icon: AlertCircle,
    steps: [
      { title: "Find flagged rows in the BOM table", body: "Pipe rows where the AI couldn't read the qty cell show qty=0 with a ⚠ note explaining the issue (e.g., 'QTY 1 matches pipe SIZE 1\" — AI likely read the size column')." },
      { title: "Click the qty cell to edit it inline", body: "Type the actual length from the drawing in feet-inches format (5'-3\", 22'-6\", etc.) or a decimal (5.25)." },
      { title: "Press Enter to save", body: "The cell exits edit mode and the new value persists. The ⚠ flag clears once a non-zero value is entered." },
      { title: "Use the Page Viewer to find the right value", body: "If you don't have the original PDF handy, click View Pages, find the right ISO sheet, and read the QTY column from the rendered image." },
    ],
    tips: [
      "Genuine 1\"-11\" inch spool pieces often appear as 0 LF flagged for review. That's intentional — the system can't tell a 1\" inch spool from a 1\" misread without context. Enter the value from the drawing.",
      "Flagged rows are surfaced as ⚠ in the validation column so you can filter by them.",
    ],
  },
  {
    id: "view-pages",
    category: "Takeoff",
    title: "Use the Page Viewer to verify extraction",
    icon: ImageIcon,
    steps: [
      { title: "Click 'View Pages' on the takeoff page", body: "The page-thumbnail strip appears with one button per ISO sheet. Each button shows the page number and drawing number when available." },
      { title: "Click a page button to load the rendered image", body: "The rendered ISO sheet loads in the viewer. Use the ← / → arrows in the header to navigate. The image is zoomable; click it to open in a new tab at full resolution." },
      { title: "Cross-check the BOM against the drawing", body: "Match the items extracted in the BOM tab against what you see drawn. Look for mismatches in fitting count, size, or category." },
      { title: "Edit items as needed", body: "Switch to the BOM tab and edit the affected rows inline." },
    ],
  },
  {
    id: "export-connections",
    category: "Takeoff",
    title: "Export connections / shop vs field welds",
    icon: Download,
    steps: [
      { title: "Click 'Export Connections' in the takeoff header", body: "An Excel workbook downloads with three sheets: Connections by Size, Connection Detail, and 40' Field Welds." },
      { title: "Read the 'Connections by Size' sheet", body: "Each size has columns for Shop BW / Shop SW / Shop Bolt-Ups / Shop Threaded / Shop Total / Field BW / Field SW / Field Total / Grand Total. Compare 'Shop Total' directly against your fab-shop pivot." },
      { title: "Audit field welds in the '40' Field Welds' sheet", body: "Every pipe spool >40 LF generates a field weld at the 40-foot joint. The sheet itemizes each one with size, description, length, and weld count." },
    ],
    tips: [
      "Shop welds = fittings/flanges/valves/olets (compare to fab-shop pivot).",
      "Field welds = 40-foot pipe joint welds (NOT in fab shop count, installed during erection).",
    ],
  },
  {
    id: "project-folders",
    category: "Project Mgmt",
    title: "Combine multiple takeoffs into a Project Folder",
    icon: FolderOpen,
    steps: [
      { title: "Open Project History from the sidebar", body: "All takeoff projects appear, organized by discipline." },
      { title: "Click 'New Folder'", body: "Name the folder (e.g., 'Stolthaven Phase 6 — All Areas')." },
      { title: "Add takeoffs to the folder", body: "Drag projects onto the folder, or use the project's menu to assign. A project can be in multiple folders." },
      { title: "Open the folder to see combined totals", body: "All BOM items, connections, and crew estimates are aggregated. Export folder-level reports from the folder header." },
    ],
  },
  {
    id: "crew-planner",
    category: "Estimating",
    title: "Build a crew with the Crew Planner",
    icon: Users,
    steps: [
      { title: "Open the project's Crew tab", body: "The default crew is built from the Picou Group rate sheet (firewatches, helpers, fitters/welders, foreman, etc.)." },
      { title: "Adjust crew size by role", body: "Increase/decrease firewatches, helpers, etc. directly. The total man-hours and weekly cost update live." },
      { title: "Set the workweek", body: "Toggle 4x10 or 5x8. Per-diem days follow the schedule." },
      { title: "Adjust labor rates if needed", body: "Override the default Picou rates per role. ST=$56/hr, OT=$79/hr, DT=$100/hr, Per Diem=$75/day." },
      { title: "Export crew schedule to Excel", body: "Get a per-week breakdown with hours, costs, per-diem days, and totals." },
    ],
    tips: ["Default rack factor 1.3x is applied to convert direct labor to billed labor.", "Stolthaven calibration: IPMH 0.437 (target 0.45)."],
  },
  {
    id: "project-planner",
    category: "Estimating",
    title: "Use the Project Planner / Gantt view",
    icon: Calculator,
    steps: [
      { title: "Open the project's Schedule tab", body: "A Gantt chart auto-builds from the BOM man-hours plus crew." },
      { title: "Set the start date", body: "Pick the project kickoff date. The schedule extends from that date based on man-hours and crew capacity." },
      { title: "Identify the critical path", body: "The longest task chain is highlighted. Add buffer or boost crew on critical tasks to compress the schedule." },
      { title: "Export to Excel", body: "Includes weekly milestones, manhour rollups, and cost-loaded schedule." },
    ],
  },
  {
    id: "patterns",
    category: "Estimating",
    title: "Apply pattern learning from corrections",
    icon: Lightbulb,
    steps: [
      { title: "Pattern learning is automatic", body: "When you correct an item (e.g., change qty or size), the correction is logged. After 3+ similar corrections, the pattern auto-applies on future projects." },
      { title: "View applied patterns", body: "In Settings → Pattern Learning, see all stored corrections grouped by category. Each pattern shows how many times it has been applied." },
      { title: "Disable a pattern if it's wrong", body: "Click 'Disable' on any pattern that's auto-correcting incorrectly. Future extractions ignore it." },
    ],
  },
  {
    id: "rates",
    category: "Estimating",
    title: "Manage labor rates and rack factors",
    icon: Settings,
    steps: [
      { title: "Open Settings → Labor Rates", body: "All Picou Group default rates listed: ST, OT, DT, per-diem by role." },
      { title: "Override per project if needed", body: "Project-specific rates take precedence over defaults. Useful for shop bids vs field bids with different rate structures." },
      { title: "Change rack factor", body: "Default 1.3x. The rack factor multiplies direct labor to billed labor (covers overhead, profit, taxes)." },
    ],
  },
  {
    id: "fab-scope",
    category: "Estimating",
    title: "Split fab-shop scope from your scope",
    icon: ClipboardList,
    steps: [
      { title: "Open the Fab Scope tab", body: "All BOM items appear with a checkbox for 'Sub-fabricated' (shipped pre-fab from a sub) or 'Your Scope' (you fab in your shop)." },
      { title: "Mark items as sub-fab", body: "Check the items that come pre-assembled (e.g., welded spools from a fab shop). Those welds and labor drop out of your manhour estimate." },
      { title: "Export the split", body: "Two-column Excel: Sub Scope (their welds, your install only) vs Your Scope (your full fab+install)." },
    ],
  },
  {
    id: "exports",
    category: "Reference",
    title: "What each export contains",
    icon: Download,
    steps: [
      { body: "BOM Export: full item list with size, qty, description, category, source page, drawing number, validation flags. The raw output of extraction." },
      { body: "Connections Export: shop welds + 40' field welds split by size. Compare directly to fab-shop pivot." },
      { body: "Pivot Summary Export: total welds and LF by size. Useful for high-level estimating quickly." },
      { body: "Crew Schedule Export: weekly crew, hours, per-diem days, costs." },
      { body: "Project Planner Export: Gantt schedule with critical path and weekly milestones." },
      { body: "Cost Database Export: line-item costs (parts, materials, sub costs)." },
      { body: "Folder Export: combined totals across multiple projects in a folder." },
    ],
  },
];

// ── Data: Reference cards ───────────────────────────────────────────────────
const REFERENCE_CARDS = [
  {
    title: "Picou Labor Rates",
    rows: [
      ["Straight Time", "$56/hr"],
      ["Overtime (>40 hrs/wk)", "$79/hr"],
      ["Double Time (Sundays/holidays)", "$100/hr"],
      ["Per Diem (helpers/fitters/welders/foreman)", "$75/day"],
      ["Default Rack Factor", "1.3x"],
    ],
  },
  {
    title: "Stolthaven Calibration",
    rows: [
      ["Achieved IPMH", "0.437"],
      ["Target IPMH", "0.45"],
      ["3\" SS weld factor", "4.68 MH"],
      ["4\" SS weld factor", "5.56 MH"],
      ["Phase 6 firewatch rate", "$20/hr (no per diem)"],
    ],
  },
  {
    title: "Weld Math (40-foot rule)",
    rows: [
      ["Pipe run > 40' total", "Add 1 field weld per 40 LF"],
      ["Example: 160 LF run", "3 field welds (160 / 40 = 4 segments → 3 joints)"],
      ["Pipe size ≤ 1.5\"", "Field weld = SW (socket weld)"],
      ["Pipe size > 1.5\"", "Field weld = BW (butt weld)"],
      ["These welds NOT in fab shop count", "Installed in field during erection"],
    ],
  },
  {
    title: "Bolt-Up Rules",
    rows: [
      ["Flange-to-flange connection", "1 bolt-up"],
      ["Butterfly valve between flanges", "1 bolt-up (NOT 2)"],
      ["Gaskets / stud bolts on their own", "No bolt-up (consumables, not connections)"],
      ["Field-only methodology (Justin's method)", "Bolt-ups counted at install, not in shop"],
    ],
  },
  {
    title: "Valve Weld Rules",
    rows: [
      ["Socket-weld valve", "2 SW welds (one per side)"],
      ["SW × NPT mixed valve", "1 SW weld + 1 threaded end"],
      ["Threaded valve (NPT both sides)", "0 welds"],
      ["Flanged valve", "0 welds (welds come from flanges around it)"],
      ["Butterfly valve", "0 welds (lugged or wafer)"],
      ["BW valve", "0 welds (welds from butt-weld pipe ends)"],
    ],
  },
  {
    title: "Weld Count by Fitting",
    rows: [
      ["90 / 45 elbow", "2 welds (one per end)"],
      ["Tee (equal)", "3 welds (two run + one branch)"],
      ["Tee (reducing)", "2 welds at run + 1 at branch (smaller bore)"],
      ["Reducer / swage", "1 weld at large end + 1 at small end"],
      ["Coupling", "2 welds (one per end) — usually SW"],
      ["Cap", "1 weld"],
      ["Sockolet / weldolet", "1 weld at branch (small bore)"],
      ["Threadolet", "1 weld at header (BW) + threaded branch end"],
      ["Flange (WN, RF)", "1 weld at pipe end + 1 bolt-up to mating flange"],
    ],
  },
];

// ── Component ───────────────────────────────────────────────────────────────
export default function HelpPage() {
  const [search, setSearch] = useState("");
  const [activeTab, setActiveTab] = useState<"getting-started" | "how-to" | "reference">("getting-started");

  const filteredHowTos = useMemo(() => {
    if (!search.trim()) return HOW_TOS;
    const q = search.toLowerCase();
    return HOW_TOS.filter(h =>
      h.title.toLowerCase().includes(q) ||
      h.category.toLowerCase().includes(q) ||
      h.steps.some(s => (s.title || "").toLowerCase().includes(q) || s.body.toLowerCase().includes(q))
    );
  }, [search]);

  const categorized = useMemo(() => {
    const map: Record<string, HowTo[]> = {};
    for (const h of filteredHowTos) {
      if (!map[h.category]) map[h.category] = [];
      map[h.category].push(h);
    }
    return map;
  }, [filteredHowTos]);

  return (
    <AppLayout subtitle="Help & How-To">
      <div className="max-w-5xl mx-auto p-6 space-y-6">
        {/* Header */}
        <div className="flex items-start justify-between gap-4 flex-wrap">
          <div>
            <h1 className="text-2xl font-semibold flex items-center gap-2">
              <BookOpen className="text-primary" /> Help & How-To
            </h1>
            <p className="text-sm text-muted-foreground mt-1 max-w-2xl">
              Everything you need to know to run a takeoff, build an estimate, and ship a bid using the Picou Group estimating system.
            </p>
          </div>
          <div className="relative w-full sm:w-72">
            <Search className="absolute left-2 top-1/2 -translate-y-1/2 text-muted-foreground" size={14} />
            <Input
              placeholder="Search how-to guides..."
              value={search}
              onChange={e => setSearch(e.target.value)}
              className="pl-7 h-9 text-sm"
            />
          </div>
        </div>

        <Tabs value={activeTab} onValueChange={v => setActiveTab(v as any)}>
          <TabsList>
            <TabsTrigger value="getting-started" className="gap-1.5"><Rocket size={14} /> Getting Started</TabsTrigger>
            <TabsTrigger value="how-to" className="gap-1.5"><BookOpen size={14} /> How To</TabsTrigger>
            <TabsTrigger value="reference" className="gap-1.5"><Database size={14} /> Reference</TabsTrigger>
          </TabsList>

          {/* ─────────── GETTING STARTED ─────────── */}
          <TabsContent value="getting-started" className="space-y-4 mt-4">
            <Card>
              <CardHeader>
                <CardTitle className="text-base">First-time walkthrough</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                {GETTING_STARTED_STEPS.map((step, i) => {
                  const Icon = step.icon;
                  return (
                    <div key={i} className="flex gap-3 items-start">
                      <div className="flex flex-col items-center shrink-0">
                        <div className="h-7 w-7 rounded-full bg-primary text-primary-foreground flex items-center justify-center text-xs font-semibold">
                          {i + 1}
                        </div>
                        {i < GETTING_STARTED_STEPS.length - 1 && (
                          <div className="w-px flex-1 bg-border mt-1 min-h-[20px]" />
                        )}
                      </div>
                      <div className="flex-1 pb-3">
                        <p className="text-sm font-semibold flex items-center gap-1.5">
                          <Icon size={14} className="text-primary" />
                          {step.title}
                        </p>
                        <p className="text-xs text-muted-foreground mt-1 leading-relaxed">{step.body}</p>
                      </div>
                    </div>
                  );
                })}
              </CardContent>
            </Card>

            <Card>
              <CardHeader><CardTitle className="text-base flex items-center gap-1.5"><Lightbulb size={16} /> Quick tips</CardTitle></CardHeader>
              <CardContent className="space-y-2 text-sm">
                <p>• <span className="font-medium">Login:</span> admin / picougroup</p>
                <p>• <span className="font-medium">Pipe quantity flags:</span> rows showing qty=0 with ⚠ are pages where the AI couldn't read the cell. Edit inline from the drawing.</p>
                <p>• <span className="font-medium">Big packages:</span> 200+ pages stream chunk-by-chunk. You can walk away.</p>
                <p>• <span className="font-medium">Drawing revisions:</span> if your set has REV C clouds, enable that toggle on upload so the cloud detection runs.</p>
                <p>• <span className="font-medium">Recent projects:</span> sidebar shows your last 5 takeoffs. Click any to jump back.</p>
              </CardContent>
            </Card>
          </TabsContent>

          {/* ─────────── HOW TO ─────────── */}
          <TabsContent value="how-to" className="space-y-4 mt-4">
            {Object.entries(categorized).length === 0 && (
              <p className="text-sm text-muted-foreground text-center py-8">No guides match "{search}". Try a different keyword.</p>
            )}
            {Object.entries(categorized).map(([category, guides]) => (
              <div key={category} className="space-y-3">
                <h2 className="text-xs uppercase tracking-wider text-muted-foreground font-semibold">{category}</h2>
                <div className="grid gap-3 grid-cols-1 md:grid-cols-2">
                  {guides.map(g => <HowToCard key={g.id} guide={g} />)}
                </div>
              </div>
            ))}
          </TabsContent>

          {/* ─────────── REFERENCE ─────────── */}
          <TabsContent value="reference" className="space-y-4 mt-4">
            <p className="text-sm text-muted-foreground">Quick-reference cards for daily use. Print or bookmark these.</p>
            <div className="grid gap-4 grid-cols-1 md:grid-cols-2">
              {REFERENCE_CARDS.map(card => (
                <Card key={card.title}>
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm">{card.title}</CardTitle>
                  </CardHeader>
                  <CardContent className="space-y-1.5">
                    {card.rows.map(([label, value], i) => (
                      <div key={i} className="flex items-baseline justify-between text-xs gap-2 pb-1.5 border-b border-border/60 last:border-b-0 last:pb-0">
                        <span className="text-muted-foreground flex-1">{label}</span>
                        <span className="font-mono font-medium text-foreground text-right">{value}</span>
                      </div>
                    ))}
                  </CardContent>
                </Card>
              ))}
            </div>
          </TabsContent>
        </Tabs>
      </div>
    </AppLayout>
  );
}

function HowToCard({ guide }: { guide: HowTo }) {
  const [open, setOpen] = useState(false);
  const Icon = guide.icon;
  return (
    <Card className={`transition-all ${open ? "ring-1 ring-primary/30" : ""}`}>
      <button
        onClick={() => setOpen(!open)}
        className="w-full text-left p-4 flex items-start gap-3 hover-elevate active-elevate-2 rounded-lg"
        aria-expanded={open}
      >
        <Icon size={18} className="text-primary mt-0.5 shrink-0" />
        <div className="flex-1 min-w-0">
          <p className="text-sm font-medium">{guide.title}</p>
          <p className="text-[10px] text-muted-foreground uppercase tracking-wide mt-0.5">{guide.category} · {guide.steps.length} step{guide.steps.length === 1 ? "" : "s"}</p>
        </div>
        <ArrowRight size={14} className={`text-muted-foreground transition-transform ${open ? "rotate-90" : ""}`} />
      </button>
      {open && (
        <div className="px-4 pb-4 space-y-2 border-t border-border/60 pt-3">
          {guide.steps.map((s, i) => (
            <div key={i} className="flex gap-2 text-xs">
              <span className="font-mono text-muted-foreground w-5 shrink-0">{i + 1}.</span>
              <div className="flex-1">
                {s.title && <p className="font-medium">{s.title}</p>}
                <p className="text-muted-foreground mt-0.5 leading-relaxed">{s.body}</p>
              </div>
            </div>
          ))}
          {guide.tips && guide.tips.length > 0 && (
            <div className="mt-3 pt-2 border-t border-border/60">
              <p className="text-[10px] uppercase tracking-wider text-muted-foreground mb-1.5 flex items-center gap-1"><Lightbulb size={10} /> Tips</p>
              <ul className="space-y-1">
                {guide.tips.map((t, i) => (
                  <li key={i} className="text-xs text-muted-foreground leading-relaxed flex gap-2">
                    <span>•</span><span className="flex-1">{t}</span>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </Card>
  );
}
