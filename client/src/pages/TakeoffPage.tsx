import { useState, useEffect } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useLocation, useSearch } from "wouter";
import { Trash2, Download, Calculator, ChevronDown, ChevronRight, FileText, Archive, ArchiveRestore, Eye, EyeOff, AlertTriangle, Image, GitCompare, Menu, X as XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import AppLayout from "@/components/AppLayout";
import UploadZone from "@/components/UploadZone";
import TakeoffBomTable from "@/components/TakeoffBomTable";
import SummaryCards from "@/components/SummaryCards";
import PivotSummary from "@/components/PivotSummary";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { exportMechanicalPdf, exportStructuralPdf, exportCivilPdf } from "@/lib/pdfExport";
import type { TakeoffProject } from "@shared/schema";

const DISCIPLINE_META = {
  mechanical: {
    label: "Mechanical Takeoff",
    description: "Piping isometric BOM extraction",
    badge: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  },
  structural: {
    label: "Structural Takeoff",
    description: "Steel, concrete & rebar schedule",
    badge: "bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300",
  },
  civil: {
    label: "Civil Takeoff",
    description: "Utilities, sitework & paving quantities",
    badge: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300",
  },
};

interface TakeoffPageProps {
  discipline: "mechanical" | "structural" | "civil";
}

export default function TakeoffPage({ discipline }: TakeoffPageProps) {
  const meta = DISCIPLINE_META[discipline];
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [showArchived, setShowArchived] = useState(false);
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [, navigate] = useLocation();
  const searchString = useSearch();
  const { toast } = useToast();

  // Read project ID from URL query parameter (e.g., ?project=abc)
  useEffect(() => {
    const params = new URLSearchParams(searchString);
    const projectId = params.get("project");
    if (projectId) {
      setSelectedId(projectId);
    }
  }, [searchString]);

  const { data: projects = [], isLoading } = useQuery<TakeoffProject[]>({
    queryKey: ["/api/takeoff/projects", discipline],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/takeoff/projects?discipline=${discipline}`);
      return res.json();
    },
  });

  const { data: selectedProject } = useQuery<TakeoffProject>({
    queryKey: ["/api/takeoff/projects", selectedId],
    queryFn: async () => {
      if (!selectedId) throw new Error("No project selected");
      const res = await apiRequest("GET", `/api/takeoff/projects/${selectedId}`);
      return res.json();
    },
    enabled: !!selectedId,
  });

  // Verification viewer state
  const [verifyOpen, setVerifyOpen] = useState(false);
  const [verifyPage, setVerifyPage] = useState<number | null>(null);

  // Page thumbnails available for this project
  const { data: availablePages = [] } = useQuery<number[]>({
    queryKey: ["/api/takeoff/projects", selectedId, "pages"],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/takeoff/projects/${selectedId}/pages`);
      return res.json();
    },
    enabled: !!selectedId,
  });

  // Scope gap detection
  const { data: scopeGaps = [] } = useQuery<{ type: string; message: string; severity: string }[]>({
    queryKey: ["/api/takeoff/projects", selectedId, "scope-gaps"],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/takeoff/projects/${selectedId}/scope-gaps`);
      const data = await res.json();
      return data.gaps || [];
    },
    enabled: !!selectedId && !!selectedProject && selectedProject.items.length > 0,
  });

  // Change order state
  const [changeOrderOpen, setChangeOrderOpen] = useState(false);
  const [changeOrderData, setChangeOrderData] = useState<any>(null);

  const changeOrderMutation = useMutation({
    mutationFn: async (projectId: string) => {
      const res = await apiRequest("POST", `/api/takeoff/projects/${projectId}/change-order`);
      return res.json();
    },
    onSuccess: (data) => {
      setChangeOrderData(data);
      setChangeOrderOpen(true);
    },
    onError: (err: any) => {
      toast({ title: "Change order failed", description: err.message, variant: "destructive" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/takeoff/projects/${id}`),
    onSuccess: (_, id) => {
      queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
      if (selectedId === id) setSelectedId(null);
      toast({ title: "Project deleted" });
    },
  });

  const archiveMutation = useMutation({
    mutationFn: ({ id, archived }: { id: string; archived: boolean }) =>
      apiRequest("PATCH", `/api/takeoff/projects/${id}/archive`, { archived }),
    onSuccess: (_, { id, archived }) => {
      queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
      if (archived && selectedId === id) setSelectedId(null);
      toast({ title: archived ? "Project archived" : "Project restored" });
    },
  });

  const importToEstimateMutation = useMutation({
    mutationFn: (takeoffProjectId: string) =>
      apiRequest("POST", "/api/takeoff/import-to-estimate", { takeoffProjectId }).then(r => r.json()),
    onSuccess: (estimate: any) => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates"] });
      toast({ title: "Estimate created!", description: "Navigating to estimating..." });
      setTimeout(() => navigate("/estimating"), 300);
    },
    onError: (err: any) => {
      toast({ title: "Failed to create estimate", description: err.message, variant: "destructive" });
    },
  });

  const handleExportPdf = () => {
    if (!selectedProject) return;
    if (discipline === "mechanical") exportMechanicalPdf(selectedProject);
    else if (discipline === "structural") exportStructuralPdf(selectedProject);
    else exportCivilPdf(selectedProject);
  };

  return (
    <AppLayout subtitle={meta.label}>
      <div className="flex h-full relative">
        {/* Mobile project list toggle */}
        <button
          className="md:hidden fixed bottom-4 right-4 z-50 w-12 h-12 rounded-full bg-primary text-primary-foreground shadow-lg flex items-center justify-center"
          onClick={() => setMobileSidebarOpen(!mobileSidebarOpen)}
        >
          {mobileSidebarOpen ? <XIcon size={20} /> : <FileText size={20} />}
        </button>

        {/* Mobile overlay */}
        {mobileSidebarOpen && (
          <div className="fixed inset-0 z-40 bg-black/40 md:hidden" onClick={() => setMobileSidebarOpen(false)} />
        )}

        {/* Left panel — project list */}
        <div className={`fixed inset-y-0 left-0 z-50 w-64 bg-background border-r border-border flex flex-col overflow-hidden transition-transform duration-200
          ${mobileSidebarOpen ? "translate-x-0" : "-translate-x-full"} md:relative md:translate-x-0 md:shrink-0`}>
          <div className="p-3 border-b border-border flex items-center justify-between">
            <h2 className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">Projects</h2>
            {projects.some(p => p.archived) && (
              <button
                className="text-muted-foreground hover:text-foreground transition-colors"
                onClick={() => setShowArchived(!showArchived)}
                title={showArchived ? "Hide archived" : "Show archived"}
              >
                {showArchived ? <EyeOff size={13} /> : <Eye size={13} />}
              </button>
            )}
          </div>
          <div className="flex-1 overflow-y-auto">
            {isLoading ? (
              <div className="p-3 space-y-2">
                {[1, 2, 3].map(i => <Skeleton key={i} className="h-16 rounded" />)}
              </div>
            ) : projects.filter(p => showArchived || !p.archived).length === 0 ? (
              <div className="p-4 text-xs text-muted-foreground text-center mt-4">
                No {discipline} projects yet.
                <br />Upload a PDF to start.
              </div>
            ) : (
              <div className="p-2 space-y-1">
                {projects.filter(p => showArchived || !p.archived).map(p => (
                  <div
                    key={p.id}
                    data-testid={`project-item-${p.id}`}
                    className={`group flex items-start gap-2 p-2.5 rounded-md cursor-pointer transition-colors
                      ${selectedId === p.id ? "bg-primary/10 border border-primary/20" : "hover:bg-accent"}
                      ${p.archived ? "opacity-50" : ""}`}
                    onClick={() => { setSelectedId(p.id); setMobileSidebarOpen(false); }}
                  >
                    <FileText size={14} className="text-muted-foreground mt-0.5 shrink-0" />
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-1">
                        <p className="text-xs font-medium truncate">{p.name}</p>
                        {p.archived && <Archive size={10} className="text-muted-foreground shrink-0" />}
                      </div>
                      <p className="text-[10px] text-muted-foreground">{p.items.length} items</p>
                      <p className="text-[10px] text-muted-foreground">{new Date(p.createdAt).toLocaleDateString()}</p>
                    </div>
                    <div className="flex flex-col gap-0.5 opacity-0 group-hover:opacity-100 shrink-0">
                      <button
                        className="text-muted-foreground hover:text-foreground p-0.5"
                        onClick={e => { e.stopPropagation(); archiveMutation.mutate({ id: p.id, archived: !p.archived }); }}
                        title={p.archived ? "Restore" : "Archive"}
                      >
                        {p.archived ? <ArchiveRestore size={13} /> : <Archive size={13} />}
                      </button>
                      <button
                        className="text-muted-foreground hover:text-destructive p-0.5"
                        onClick={e => { e.stopPropagation(); deleteMutation.mutate(p.id); }}
                        data-testid={`btn-delete-${p.id}`}
                        title="Delete permanently"
                      >
                        <Trash2 size={13} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Right panel */}
        <div className="flex-1 overflow-y-auto">
          <div className="p-5 space-y-5">
            {/* Upload zone */}
            <UploadZone
              discipline={discipline}
              onProjectCreated={(id) => {
                queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
                setSelectedId(id);
              }}
            />

            {/* Selected project view */}
            {selectedId && selectedProject ? (
              <div className="space-y-4">
                {/* Project header */}
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <h2 className="text-base font-semibold">{selectedProject.name}</h2>
                      <Badge variant="outline" className={`text-xs ${meta.badge}`}>{meta.label}</Badge>
                      {selectedProject.revision && (
                        <Badge variant="outline" className="text-xs">Rev {selectedProject.revision}</Badge>
                      )}
                    </div>
                    <div className="flex gap-4 mt-1 text-xs text-muted-foreground flex-wrap">
                      {selectedProject.lineNumber && <span>Line: {selectedProject.lineNumber}</span>}
                      {selectedProject.area && <span>Area: {selectedProject.area}</span>}
                      {selectedProject.drawingDate && <span>Date: {selectedProject.drawingDate}</span>}
                      <span>{selectedProject.items.length} items</span>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0 flex-wrap md:flex-nowrap">
                    {availablePages.length > 0 && (
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => { setVerifyOpen(!verifyOpen); if (!verifyOpen && availablePages.length) setVerifyPage(availablePages[0]); }}
                        data-testid="btn-verify-viewer"
                      >
                        <Image size={14} className="mr-1.5" />
                        {verifyOpen ? "Hide Pages" : "View Pages"}
                      </Button>
                    )}
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={handleExportPdf}
                      data-testid="btn-export-pdf"
                    >
                      <Download size={14} className="mr-1.5" />
                      Export PDF
                    </Button>
                    {selectedProject.items.some(i => i.revisionClouded) && (
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-orange-700 border-orange-300 hover:bg-orange-50 dark:text-orange-400 dark:border-orange-700 dark:hover:bg-orange-900/30"
                        onClick={() => changeOrderMutation.mutate(selectedProject.id)}
                        disabled={changeOrderMutation.isPending}
                        data-testid="btn-change-order"
                      >
                        <GitCompare size={14} className="mr-1.5" />
                        {changeOrderMutation.isPending ? "Generating..." : "Change Order"}
                      </Button>
                    )}
                    <Button
                      size="sm"
                      className="bg-green-600 hover:bg-green-700 text-white"
                      onClick={() => importToEstimateMutation.mutate(selectedProject.id)}
                      disabled={importToEstimateMutation.isPending}
                      data-testid="btn-run-estimate"
                    >
                      <Calculator size={14} className="mr-1.5" />
                      {importToEstimateMutation.isPending ? "Creating..." : "Run Estimate"}
                    </Button>
                  </div>
                </div>

                {/* Summary cards */}
                <SummaryCards items={selectedProject.items} discipline={discipline} />

                {/* Scope Gap Alerts */}
                {scopeGaps.length > 0 && (
                  <div className="space-y-2">
                    <h3 className="text-xs font-semibold text-amber-700 dark:text-amber-400 uppercase tracking-wider flex items-center gap-1.5">
                      <AlertTriangle size={13} /> Scope Gap Warnings
                    </h3>
                    <div className="grid gap-2 sm:grid-cols-2">
                      {scopeGaps.map((gap, i) => (
                        <div key={i} className="flex items-start gap-2 px-3 py-2 rounded-md bg-amber-50 border border-amber-200 dark:bg-amber-900/20 dark:border-amber-800">
                          <AlertTriangle size={13} className="text-amber-600 dark:text-amber-400 mt-0.5 shrink-0" />
                          <p className="text-xs text-amber-800 dark:text-amber-300">{gap.message}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Verification Viewer */}
                {verifyOpen && availablePages.length > 0 && (
                  <Card className="border-card-border">
                    <CardHeader className="p-3 pb-2">
                      <CardTitle className="text-xs flex items-center gap-1.5">
                        <Image size={12} /> Page Verification Viewer
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="p-3 pt-0">
                      <div className="flex gap-1 mb-2 flex-wrap">
                        {availablePages.map(pg => (
                          <Button
                            key={pg}
                            variant={verifyPage === pg ? "default" : "outline"}
                            size="sm"
                            className="h-7 text-xs px-2"
                            onClick={() => setVerifyPage(pg)}
                          >
                            Page {pg}
                          </Button>
                        ))}
                      </div>
                      {verifyPage != null && (
                        <div className="border rounded-md overflow-hidden bg-white dark:bg-black">
                          <img
                            src={`/api/takeoff/projects/${selectedId}/page/${verifyPage}`}
                            alt={`Page ${verifyPage}`}
                            className="w-full h-auto"
                            onError={(e) => { (e.target as HTMLImageElement).alt = "Thumbnail not available"; }}
                          />
                        </div>
                      )}
                    </CardContent>
                  </Card>
                )}

                {/* Change Order Dialog */}
                <Dialog open={changeOrderOpen} onOpenChange={setChangeOrderOpen}>
                  <DialogContent className="max-w-lg">
                    <DialogHeader>
                      <DialogTitle className="flex items-center gap-2">
                        <GitCompare size={16} /> Change Order Summary
                      </DialogTitle>
                    </DialogHeader>
                    {changeOrderData && (
                      <div className="space-y-3 text-sm">
                        <div className="grid grid-cols-2 gap-3">
                          <div className="p-3 rounded-md bg-orange-50 dark:bg-orange-900/20 border border-orange-200 dark:border-orange-800">
                            <p className="text-xs text-muted-foreground">Changed Items</p>
                            <p className="text-lg font-semibold">{changeOrderData.cloudedItems?.length || 0}</p>
                          </div>
                          <div className="p-3 rounded-md bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800">
                            <p className="text-xs text-muted-foreground">Unchanged Items</p>
                            <p className="text-lg font-semibold">{changeOrderData.unchangedItems?.length || 0}</p>
                          </div>
                        </div>
                        {changeOrderData.costImpact && (
                          <div className="p-3 rounded-md bg-card border border-border">
                            <p className="text-xs font-semibold mb-2">Cost Impact</p>
                            <div className="space-y-1 text-xs">
                              <div className="flex justify-between"><span>Material</span><span>${(changeOrderData.costImpact.materialCost || 0).toLocaleString("en-US", { minimumFractionDigits: 2 })}</span></div>
                              <div className="flex justify-between"><span>Labor</span><span>${(changeOrderData.costImpact.laborCost || 0).toLocaleString("en-US", { minimumFractionDigits: 2 })}</span></div>
                              <div className="flex justify-between font-semibold border-t pt-1 mt-1"><span>Total</span><span>${(changeOrderData.costImpact.totalCost || 0).toLocaleString("en-US", { minimumFractionDigits: 2 })}</span></div>
                            </div>
                          </div>
                        )}
                        {changeOrderData.cloudedItems?.length > 0 && (
                          <div>
                            <p className="text-xs font-semibold mb-1">Changed Items</p>
                            <div className="max-h-48 overflow-y-auto border rounded-md">
                              <table className="w-full text-xs">
                                <thead className="bg-muted sticky top-0"><tr><th className="p-1.5 text-left">Description</th><th className="p-1.5 text-left">Size</th><th className="p-1.5 text-right">Qty</th></tr></thead>
                                <tbody>
                                  {changeOrderData.cloudedItems.map((item: any, i: number) => (
                                    <tr key={i} className="border-t"><td className="p-1.5">{item.description}</td><td className="p-1.5">{item.size}</td><td className="p-1.5 text-right">{item.quantity}</td></tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </DialogContent>
                </Dialog>

                {/* Tabs */}
                <Tabs defaultValue="bom">
                  <TabsList className="mb-3">
                    <TabsTrigger value="bom" className="text-xs" data-testid="tab-bom">BOM Table</TabsTrigger>
                    <TabsTrigger value="pivot" className="text-xs" data-testid="tab-pivot">Pivot Summary</TabsTrigger>
                  </TabsList>
                  <TabsContent value="bom">
                    <TakeoffBomTable items={selectedProject.items} discipline={discipline} />
                  </TabsContent>
                  <TabsContent value="pivot">
                    <PivotSummary items={selectedProject.items} />
                  </TabsContent>
                </Tabs>
              </div>
            ) : selectedId ? (
              <div className="flex items-center justify-center h-32">
                <Skeleton className="h-full w-full rounded-lg" />
              </div>
            ) : null}
          </div>
        </div>
      </div>
    </AppLayout>
  );
}
