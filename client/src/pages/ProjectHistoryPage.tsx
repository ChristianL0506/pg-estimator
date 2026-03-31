import { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { Search, Plus, Trash2, Clock, MapPin, User, Tag, ChevronDown, ChevronUp, Calculator, TrendingUp } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Label } from "@/components/ui/label";
import AppLayout from "@/components/AppLayout";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";
import type { CompletedProject } from "@shared/schema";

function fmt$(n: number) { return `$${n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
function fmtNum(n: number) { return n.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 }); }

export default function ProjectHistoryPage() {
  const { toast } = useToast();
  const [searchQuery, setSearchQuery] = useState("");
  const [showAddForm, setShowAddForm] = useState(false);

  // Form state
  const [formName, setFormName] = useState("");
  const [formClient, setFormClient] = useState("");
  const [formLocation, setFormLocation] = useState("");
  const [formScope, setFormScope] = useState("");
  const [formStartDate, setFormStartDate] = useState("");
  const [formEndDate, setFormEndDate] = useState("");
  const [formWelderHours, setFormWelderHours] = useState(0);
  const [formFitterHours, setFormFitterHours] = useState(0);
  const [formHelperHours, setFormHelperHours] = useState(0);
  const [formForemanHours, setFormForemanHours] = useState(0);
  const [formOperatorHours, setFormOperatorHours] = useState(0);
  const [formMaterialCost, setFormMaterialCost] = useState(0);
  const [formLaborCost, setFormLaborCost] = useState(0);
  const [formPeakCrew, setFormPeakCrew] = useState(0);
  const [formDuration, setFormDuration] = useState(0);
  const [formTags, setFormTags] = useState("");
  const [formNotes, setFormNotes] = useState("");

  // Query: either search or list all
  const queryKey = searchQuery.trim()
    ? [`/api/project-history/search?q=${encodeURIComponent(searchQuery.trim())}`]
    : ["/api/project-history"];

  const { data: projects = [], isLoading } = useQuery<CompletedProject[]>({
    queryKey,
  });

  const addMutation = useMutation({
    mutationFn: async (data: any) => {
      const res = await apiRequest("POST", "/api/project-history", data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/project-history"] });
      toast({ title: "Project added", description: "Completed project saved to history." });
      resetForm();
      setShowAddForm(false);
    },
    onError: (err: any) => {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: async (id: string) => {
      await apiRequest("DELETE", `/api/project-history/${id}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/project-history"] });
      toast({ title: "Deleted", description: "Project removed from history." });
    },
  });

  function resetForm() {
    setFormName(""); setFormClient(""); setFormLocation(""); setFormScope("");
    setFormStartDate(""); setFormEndDate(""); setFormWelderHours(0); setFormFitterHours(0);
    setFormHelperHours(0); setFormForemanHours(0); setFormOperatorHours(0);
    setFormMaterialCost(0); setFormLaborCost(0); setFormPeakCrew(0);
    setFormDuration(0); setFormTags(""); setFormNotes("");
  }

  function handleSubmit() {
    if (!formName.trim() || !formScope.trim()) {
      toast({ title: "Required", description: "Name and scope description are required.", variant: "destructive" });
      return;
    }
    addMutation.mutate({
      name: formName.trim(),
      client: formClient.trim() || undefined,
      location: formLocation.trim() || undefined,
      scopeDescription: formScope.trim(),
      startDate: formStartDate || undefined,
      endDate: formEndDate || undefined,
      welderHours: formWelderHours,
      fitterHours: formFitterHours,
      helperHours: formHelperHours,
      foremanHours: formForemanHours,
      operatorHours: formOperatorHours,
      materialCost: formMaterialCost,
      laborCost: formLaborCost,
      peakCrewSize: formPeakCrew || undefined,
      durationDays: formDuration || undefined,
      tags: formTags.trim() || undefined,
      notes: formNotes.trim() || undefined,
    });
  }

  // Quick Re-Estimate
  const [quickScope, setQuickScope] = useState("");
  const [quickTags, setQuickTags] = useState("");
  const [quickResult, setQuickResult] = useState<any>(null);
  const [showQuickEstimate, setShowQuickEstimate] = useState(false);

  const quickEstimateMutation = useMutation({
    mutationFn: async (data: { scopeDescription: string; tags: string }) => {
      const res = await apiRequest("POST", "/api/project-history/quick-estimate", data);
      return res.json();
    },
    onSuccess: (data) => {
      setQuickResult(data);
    },
    onError: (err: any) => {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    },
  });

  // Summary stats
  const totalProjects = projects.length;
  const totalManhours = projects.reduce((s, p) => s + (p.totalManhours || 0), 0);
  const totalCostAll = projects.reduce((s, p) => s + (p.totalCost || 0), 0);
  const avgCostPerMH = totalManhours > 0 ? totalCostAll / totalManhours : 0;

  return (
    <AppLayout subtitle="Project History">
      <div className="p-5 space-y-5 max-w-7xl mx-auto">
        {/* Header */}
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <h2 className="text-lg font-bold text-foreground flex items-center gap-2">
              <Clock size={20} className="text-primary" />
              Completed Projects Knowledge Base
            </h2>
            <p className="text-xs text-muted-foreground mt-0.5">
              Search past projects by scope, tags, or name to inform future estimates
            </p>
          </div>
          <Button
            size="sm"
            onClick={() => setShowAddForm(!showAddForm)}
            data-testid="btn-add-project"
          >
            {showAddForm ? <ChevronUp size={14} className="mr-1.5" /> : <Plus size={14} className="mr-1.5" />}
            {showAddForm ? "Close" : "Add Project"}
          </Button>
        </div>

        {/* Summary stats */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <Card className="border-card-border">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Total Projects</p>
              <p className="text-2xl font-bold text-foreground">{totalProjects}</p>
            </CardContent>
          </Card>
          <Card className="border-card-border">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Total Manhours</p>
              <p className="text-2xl font-bold text-foreground">{fmtNum(totalManhours)}</p>
            </CardContent>
          </Card>
          <Card className="border-card-border">
            <CardContent className="p-4">
              <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Avg Cost / MH</p>
              <p className="text-2xl font-bold text-foreground">{fmt$(avgCostPerMH)}</p>
            </CardContent>
          </Card>
        </div>

        {/* Search */}
        <div className="relative max-w-lg">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
          <Input
            className="pl-9 h-9 text-sm"
            placeholder='Search projects... e.g. "screw conveyor" or "8 inch pipe"'
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            data-testid="input-search-history"
          />
        </div>

        {/* Quick Re-Estimate */}
        <Card className="border-card-border">
          <CardHeader className="p-4 pb-2 cursor-pointer" onClick={() => setShowQuickEstimate(!showQuickEstimate)}>
            <CardTitle className="text-sm flex items-center gap-2">
              <Calculator size={14} className="text-primary" />
              Quick Re-Estimate from History
              {showQuickEstimate ? <ChevronUp size={14} className="ml-auto" /> : <ChevronDown size={14} className="ml-auto" />}
            </CardTitle>
          </CardHeader>
          {showQuickEstimate && (
            <CardContent className="p-4 pt-0 space-y-3">
              <p className="text-xs text-muted-foreground">Describe a scope and we'll find similar past projects to estimate cost and manhour ranges.</p>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div>
                  <Label className="text-xs">Scope Description</Label>
                  <textarea
                    className="w-full h-16 text-xs mt-1 border border-input rounded-md px-3 py-2 bg-background resize-none focus:outline-none focus:ring-2 focus:ring-ring"
                    value={quickScope}
                    onChange={e => setQuickScope(e.target.value)}
                    placeholder='e.g. "Install 6 inch stainless steel screw conveyor"'
                  />
                </div>
                <div>
                  <Label className="text-xs">Tags (optional)</Label>
                  <Input className="h-8 text-xs mt-1" value={quickTags} onChange={e => setQuickTags(e.target.value)} placeholder="stainless, conveyor, 6-inch" />
                  <Button
                    size="sm"
                    className="mt-2"
                    disabled={quickEstimateMutation.isPending || !quickScope.trim()}
                    onClick={() => quickEstimateMutation.mutate({ scopeDescription: quickScope, tags: quickTags })}
                  >
                    <TrendingUp size={14} className="mr-1" />
                    {quickEstimateMutation.isPending ? "Searching..." : "Estimate"}
                  </Button>
                </div>
              </div>

              {quickResult && (
                <div className="space-y-3 pt-2 border-t border-border">
                  {quickResult.matches?.length === 0 ? (
                    <p className="text-xs text-muted-foreground">No matching projects found. Add more projects to your history for better estimates.</p>
                  ) : (
                    <>
                      <p className="text-xs text-muted-foreground">Based on {quickResult.estimate?.basedOn || 0} similar project(s):</p>
                      <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                        {quickResult.estimate?.manhours && (
                          <div className="bg-blue-50 dark:bg-blue-900/20 rounded p-2 text-center">
                            <p className="text-[10px] text-blue-600 dark:text-blue-400">Manhours</p>
                            <p className="text-xs font-semibold font-mono">{fmtNum(quickResult.estimate.manhours.min)} — {fmtNum(quickResult.estimate.manhours.max)}</p>
                            <p className="text-[9px] text-muted-foreground">avg {fmtNum(quickResult.estimate.manhours.avg)}</p>
                          </div>
                        )}
                        {quickResult.estimate?.totalCost && (
                          <div className="bg-primary/10 rounded p-2 text-center">
                            <p className="text-[10px] text-primary">Total Cost</p>
                            <p className="text-xs font-semibold font-mono">{fmt$(quickResult.estimate.totalCost.min)} — {fmt$(quickResult.estimate.totalCost.max)}</p>
                            <p className="text-[9px] text-muted-foreground">avg {fmt$(quickResult.estimate.totalCost.avg)}</p>
                          </div>
                        )}
                        {quickResult.estimate?.materialCost && (
                          <div className="bg-green-50 dark:bg-green-900/20 rounded p-2 text-center">
                            <p className="text-[10px] text-green-600 dark:text-green-400">Material</p>
                            <p className="text-xs font-semibold font-mono">{fmt$(quickResult.estimate.materialCost.min)} — {fmt$(quickResult.estimate.materialCost.max)}</p>
                          </div>
                        )}
                        {quickResult.estimate?.laborCost && (
                          <div className="bg-orange-50 dark:bg-orange-900/20 rounded p-2 text-center">
                            <p className="text-[10px] text-orange-600 dark:text-orange-400">Labor</p>
                            <p className="text-xs font-semibold font-mono">{fmt$(quickResult.estimate.laborCost.min)} — {fmt$(quickResult.estimate.laborCost.max)}</p>
                          </div>
                        )}
                        {quickResult.estimate?.duration && (
                          <div className="bg-purple-50 dark:bg-purple-900/20 rounded p-2 text-center">
                            <p className="text-[10px] text-purple-600 dark:text-purple-400">Duration</p>
                            <p className="text-xs font-semibold font-mono">{quickResult.estimate.duration.min} — {quickResult.estimate.duration.max} days</p>
                          </div>
                        )}
                        {quickResult.estimate?.crewSize && (
                          <div className="bg-slate-50 dark:bg-slate-900/20 rounded p-2 text-center">
                            <p className="text-[10px] text-slate-600 dark:text-slate-400">Crew Size</p>
                            <p className="text-xs font-semibold font-mono">{quickResult.estimate.crewSize.min} — {quickResult.estimate.crewSize.max}</p>
                          </div>
                        )}
                      </div>

                      {/* Matching projects */}
                      <div>
                        <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1">Matching Projects</p>
                        <div className="space-y-1">
                          {quickResult.matches?.slice(0, 5).map((m: any) => (
                            <div key={m.id} className="flex items-center justify-between text-xs px-2 py-1 rounded bg-muted/30">
                              <span className="truncate max-w-[250px] font-medium">{m.name}</span>
                              <div className="flex items-center gap-3 text-muted-foreground shrink-0">
                                <span>{fmtNum(m.totalManhours || 0)} MH</span>
                                <span>{fmt$(m.totalCost || 0)}</span>
                                <Badge variant="outline" className="text-[8px] px-1 py-0">Score: {m.matchScore}</Badge>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </>
                  )}
                </div>
              )}
            </CardContent>
          )}
        </Card>

        {/* Add Project Form */}
        {showAddForm && (
          <Card className="border-primary/30 bg-card">
            <CardHeader className="p-4 pb-2">
              <CardTitle className="text-sm">Add Completed Project</CardTitle>
            </CardHeader>
            <CardContent className="p-4 pt-0 space-y-4">
              {/* Row 1: Name, Client, Location */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                <div>
                  <Label className="text-xs">Project Name *</Label>
                  <Input className="h-8 text-xs mt-1" value={formName} onChange={e => setFormName(e.target.value)} placeholder="e.g. Tank Farm Piping Phase 2" data-testid="input-ph-name" />
                </div>
                <div>
                  <Label className="text-xs">Client</Label>
                  <Input className="h-8 text-xs mt-1" value={formClient} onChange={e => setFormClient(e.target.value)} placeholder="e.g. ExxonMobil" />
                </div>
                <div>
                  <Label className="text-xs">Location</Label>
                  <Input className="h-8 text-xs mt-1" value={formLocation} onChange={e => setFormLocation(e.target.value)} placeholder="e.g. Baton Rouge, LA" />
                </div>
              </div>

              {/* Scope Description */}
              <div>
                <Label className="text-xs">Scope Description *</Label>
                <textarea
                  className="w-full h-20 text-xs mt-1 border border-input rounded-md px-3 py-2 bg-background resize-none focus:outline-none focus:ring-2 focus:ring-ring"
                  value={formScope}
                  onChange={e => setFormScope(e.target.value)}
                  placeholder={"e.g. Install 6\" screw conveyor, run 500' of 8\" SS pipe, fabricate & install 3 process skids"}
                  data-testid="input-ph-scope"
                />
              </div>

              {/* Dates + Duration */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <div>
                  <Label className="text-xs">Start Date</Label>
                  <Input className="h-8 text-xs mt-1" type="date" value={formStartDate} onChange={e => setFormStartDate(e.target.value)} />
                </div>
                <div>
                  <Label className="text-xs">End Date</Label>
                  <Input className="h-8 text-xs mt-1" type="date" value={formEndDate} onChange={e => setFormEndDate(e.target.value)} />
                </div>
                <div>
                  <Label className="text-xs">Duration (days)</Label>
                  <Input className="h-8 text-xs mt-1" type="number" value={formDuration || ""} onChange={e => setFormDuration(parseFloat(e.target.value) || 0)} />
                </div>
                <div>
                  <Label className="text-xs">Peak Crew Size</Label>
                  <Input className="h-8 text-xs mt-1" type="number" value={formPeakCrew || ""} onChange={e => setFormPeakCrew(parseFloat(e.target.value) || 0)} />
                </div>
              </div>

              {/* Manhours by trade */}
              <div>
                <Label className="text-xs font-semibold">Manhours by Trade</Label>
                <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mt-1">
                  {[
                    { label: "Welder", val: formWelderHours, set: setFormWelderHours },
                    { label: "Fitter", val: formFitterHours, set: setFormFitterHours },
                    { label: "Helper", val: formHelperHours, set: setFormHelperHours },
                    { label: "Foreman", val: formForemanHours, set: setFormForemanHours },
                    { label: "Operator", val: formOperatorHours, set: setFormOperatorHours },
                  ].map(({ label, val, set }) => (
                    <div key={label}>
                      <Label className="text-[10px] text-muted-foreground">{label} hrs</Label>
                      <Input className="h-8 text-xs mt-0.5" type="number" value={val || ""} onChange={e => set(parseFloat(e.target.value) || 0)} />
                    </div>
                  ))}
                </div>
              </div>

              {/* Costs */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label className="text-xs">Material Cost ($)</Label>
                  <Input className="h-8 text-xs mt-1" type="number" value={formMaterialCost || ""} onChange={e => setFormMaterialCost(parseFloat(e.target.value) || 0)} />
                </div>
                <div>
                  <Label className="text-xs">Labor Cost ($)</Label>
                  <Input className="h-8 text-xs mt-1" type="number" value={formLaborCost || ""} onChange={e => setFormLaborCost(parseFloat(e.target.value) || 0)} />
                </div>
              </div>

              {/* Tags + Notes */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div>
                  <Label className="text-xs">Tags (comma-separated)</Label>
                  <Input className="h-8 text-xs mt-1" value={formTags} onChange={e => setFormTags(e.target.value)} placeholder="conveyor, stainless, 8-inch, tank farm" />
                </div>
                <div>
                  <Label className="text-xs">Notes</Label>
                  <Input className="h-8 text-xs mt-1" value={formNotes} onChange={e => setFormNotes(e.target.value)} placeholder="Any additional notes..." />
                </div>
              </div>

              <div className="flex gap-2 pt-1">
                <Button size="sm" onClick={handleSubmit} disabled={addMutation.isPending} data-testid="btn-save-project">
                  {addMutation.isPending ? "Saving..." : "Save Project"}
                </Button>
                <Button size="sm" variant="outline" onClick={() => { resetForm(); setShowAddForm(false); }}>Cancel</Button>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Project Cards */}
        {isLoading ? (
          <div className="text-sm text-muted-foreground">Loading projects...</div>
        ) : projects.length === 0 ? (
          <Card className="border-dashed border-2 border-muted-foreground/20">
            <CardContent className="p-8 text-center">
              <Clock size={32} className="mx-auto text-muted-foreground/40 mb-2" />
              <p className="text-sm text-muted-foreground">
                {searchQuery ? "No projects match your search." : "No completed projects yet. Add your first project to build your knowledge base."}
              </p>
            </CardContent>
          </Card>
        ) : (
          <div className="space-y-3">
            {projects.map(project => (
              <ProjectCard key={project.id} project={project} onDelete={(id) => deleteMutation.mutate(id)} />
            ))}
          </div>
        )}
      </div>
    </AppLayout>
  );
}

function ProjectCard({ project, onDelete }: { project: CompletedProject; onDelete: (id: string) => void }) {
  const [expanded, setExpanded] = useState(false);
  const totalMH = project.totalManhours || 0;
  const tags = (project.tags || "").split(",").map(t => t.trim()).filter(Boolean);

  return (
    <Card className="border-card-border hover:border-primary/30 transition-colors">
      <CardContent className="p-4">
        {/* Header row */}
        <div className="flex items-start justify-between gap-3">
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <h3 className="text-sm font-semibold text-foreground">{project.name}</h3>
              {project.client && (
                <span className="text-xs text-muted-foreground flex items-center gap-1">
                  <User size={10} /> {project.client}
                </span>
              )}
              {project.location && (
                <span className="text-xs text-muted-foreground flex items-center gap-1">
                  <MapPin size={10} /> {project.location}
                </span>
              )}
            </div>
            <p className="text-xs text-muted-foreground mt-1 leading-relaxed">{project.scopeDescription}</p>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <Button variant="ghost" size="icon" className="h-7 w-7" onClick={() => setExpanded(!expanded)}>
              {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </Button>
            <Button variant="ghost" size="icon" className="h-7 w-7 text-destructive hover:text-destructive" onClick={() => onDelete(project.id)}>
              <Trash2 size={14} />
            </Button>
          </div>
        </div>

        {/* Quick stats row */}
        <div className="flex items-center gap-4 mt-2 flex-wrap">
          <div className="text-xs">
            <span className="text-muted-foreground">Total MH:</span>{" "}
            <span className="font-semibold text-foreground">{fmtNum(totalMH)}</span>
          </div>
          {project.durationDays && (
            <div className="text-xs">
              <span className="text-muted-foreground">Duration:</span>{" "}
              <span className="font-semibold text-foreground">{project.durationDays} days</span>
            </div>
          )}
          {project.peakCrewSize && (
            <div className="text-xs">
              <span className="text-muted-foreground">Peak Crew:</span>{" "}
              <span className="font-semibold text-foreground">{project.peakCrewSize}</span>
            </div>
          )}
          <div className="text-xs">
            <span className="text-muted-foreground">Total Cost:</span>{" "}
            <span className="font-semibold text-primary">{fmt$(project.totalCost || 0)}</span>
          </div>
        </div>

        {/* Tags */}
        {tags.length > 0 && (
          <div className="flex items-center gap-1.5 mt-2 flex-wrap">
            <Tag size={10} className="text-muted-foreground" />
            {tags.map(tag => (
              <Badge key={tag} variant="secondary" className="text-[10px] px-1.5 py-0">
                {tag}
              </Badge>
            ))}
          </div>
        )}

        {/* Expanded details */}
        {expanded && (
          <div className="mt-3 pt-3 border-t border-border space-y-3">
            {/* Manhour breakdown */}
            <div>
              <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider mb-1.5">Manhour Breakdown</p>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                {[
                  { label: "Welder", val: project.welderHours },
                  { label: "Fitter", val: project.fitterHours },
                  { label: "Helper", val: project.helperHours },
                  { label: "Foreman", val: project.foremanHours },
                  { label: "Operator", val: project.operatorHours },
                ].map(({ label, val }) => (
                  <div key={label} className="bg-muted/50 rounded p-2 text-center">
                    <p className="text-[10px] text-muted-foreground">{label}</p>
                    <p className="text-xs font-semibold font-mono">{fmtNum(val || 0)}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* Cost breakdown */}
            <div className="grid grid-cols-3 gap-2">
              <div className="bg-blue-500/10 rounded p-2 text-center">
                <p className="text-[10px] text-blue-600 dark:text-blue-400">Material</p>
                <p className="text-xs font-semibold font-mono">{fmt$(project.materialCost || 0)}</p>
              </div>
              <div className="bg-orange-500/10 rounded p-2 text-center">
                <p className="text-[10px] text-orange-600 dark:text-orange-400">Labor</p>
                <p className="text-xs font-semibold font-mono">{fmt$(project.laborCost || 0)}</p>
              </div>
              <div className="bg-primary/10 rounded p-2 text-center">
                <p className="text-[10px] text-primary">Total</p>
                <p className="text-xs font-semibold font-mono">{fmt$(project.totalCost || 0)}</p>
              </div>
            </div>

            {/* Dates */}
            {(project.startDate || project.endDate) && (
              <div className="text-xs text-muted-foreground">
                {project.startDate && <span>Start: {project.startDate}</span>}
                {project.startDate && project.endDate && <span className="mx-2">|</span>}
                {project.endDate && <span>End: {project.endDate}</span>}
              </div>
            )}

            {/* Notes */}
            {project.notes && (
              <div className="text-xs text-muted-foreground italic">
                {project.notes}
              </div>
            )}
          </div>
        )}
      </CardContent>
    </Card>
  );
}
