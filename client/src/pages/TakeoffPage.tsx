import { useState, useEffect } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useLocation, useSearch } from "wouter";
import { Trash2, Download, Calculator, ChevronDown, ChevronRight, FileText, Archive, ArchiveRestore, Eye, EyeOff, AlertTriangle, Image, GitCompare, Menu, X as XIcon, FolderOpen, FolderPlus, Folder, Plus, MoreVertical, FileSpreadsheet } from "lucide-react";
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
import ConnectionsSummary from "@/components/ConnectionsSummary";
import FabScopeSplitter from "@/components/FabScopeSplitter";
import { ReviewMode } from "@/components/ReviewMode";
import { useToast } from "@/hooks/use-toast";
import { Input } from "@/components/ui/input";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
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
  const [scopeGapsHidden, setScopeGapsHidden] = useState(false);
  const [showArchived, setShowArchived] = useState(false);
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [, navigate] = useLocation();
  const searchString = useSearch();
  const { toast } = useToast();

  // Folder state
  const [selectedFolderId, setSelectedFolderId] = useState<string | null>(null);
  const [showFolders, setShowFolders] = useState(true);
  const [newFolderName, setNewFolderName] = useState("");
  const [creatingFolder, setCreatingFolder] = useState(false);

  // Folder queries
  const { data: folders = [], refetch: refetchFolders } = useQuery<any[]>({
    queryKey: ["/api/folders"],
  });

  const { data: folderBom } = useQuery<any>({
    queryKey: ["/api/folders", selectedFolderId, "combined-bom"],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/folders/${selectedFolderId}/combined-bom`);
      return res.json();
    },
    enabled: !!selectedFolderId,
  });

  const createFolderMutation = useMutation({
    mutationFn: (name: string) => apiRequest("POST", "/api/folders", { name }).then(r => r.json()),
    onSuccess: () => {
      refetchFolders();
      setNewFolderName("");
      setCreatingFolder(false);
      toast({ title: "Folder created" });
    },
  });

  const deleteFolderMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/folders/${id}`),
    onSuccess: (_, id) => {
      refetchFolders();
      if (selectedFolderId === id) setSelectedFolderId(null);
      queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
      toast({ title: "Folder deleted" });
    },
  });

  const moveToFolderMutation = useMutation({
    mutationFn: ({ folderId, projectId }: { folderId: string; projectId: string }) =>
      apiRequest("POST", `/api/folders/${folderId}/projects`, { projectId }),
    onSuccess: () => {
      refetchFolders();
      queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
      toast({ title: "Project moved to folder" });
    },
  });

  const removeFromFolderMutation = useMutation({
    mutationFn: ({ folderId, projectId }: { folderId: string; projectId: string }) =>
      apiRequest("DELETE", `/api/folders/${folderId}/projects/${projectId}`),
    onSuccess: () => {
      refetchFolders();
      queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects"] });
      toast({ title: "Project removed from folder" });
    },
  });

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
      const data = await res.json();
      return data.pages || [];
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
          className="md:hidden fixed bottom-6 right-4 z-50 w-12 h-12 rounded-full bg-primary text-primary-foreground shadow-lg flex items-center justify-center"
          aria-label={mobileSidebarOpen ? "Close project list" : "Open project list"}
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
          {/* Folders section */}
          <div className="p-3 border-b border-border">
            <div className="flex items-center justify-between mb-2">
              <button
                className="flex items-center gap-1 text-xs font-semibold text-muted-foreground uppercase tracking-wider hover:text-foreground transition-colors"
                onClick={() => setShowFolders(!showFolders)}
              >
                {showFolders ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                Folders
              </button>
              <button
                className="text-muted-foreground hover:text-foreground transition-colors"
                onClick={() => setCreatingFolder(true)}
                title="Create folder"
              >
                <FolderPlus size={13} />
              </button>
            </div>
            {showFolders && (
              <div className="space-y-0.5">
                {creatingFolder && (
                  <form
                    className="flex gap-1 mb-1"
                    onSubmit={e => {
                      e.preventDefault();
                      if (newFolderName.trim()) createFolderMutation.mutate(newFolderName.trim());
                    }}
                  >
                    <Input
                      autoFocus
                      value={newFolderName}
                      onChange={e => setNewFolderName(e.target.value)}
                      placeholder="Folder name..."
                      className="h-7 text-xs"
                    />
                    <Button type="submit" size="sm" className="h-7 px-2 text-xs" disabled={!newFolderName.trim()}>
                      <Plus size={12} />
                    </Button>
                  </form>
                )}
                {folders.map((f: any) => (
                  <div
                    key={f.id}
                    className={`group flex items-center gap-2 px-2 py-1.5 rounded-md cursor-pointer transition-colors text-xs
                      ${selectedFolderId === f.id ? "bg-primary/10 border border-primary/20" : "hover:bg-accent"}`}
                    onClick={() => { setSelectedFolderId(f.id); setSelectedId(null); setMobileSidebarOpen(false); }}
                  >
                    <Folder size={13} className="text-teal-600 dark:text-teal-400 shrink-0" />
                    <span className="truncate flex-1 font-medium">{f.name}</span>
                    <span className="text-[10px] text-muted-foreground">{f.projectCount || 0}</span>
                    <DropdownMenu>
                      <DropdownMenuTrigger asChild onClick={e => e.stopPropagation()}>
                        <button className="opacity-0 group-hover:opacity-100 text-muted-foreground hover:text-foreground p-0.5">
                          <MoreVertical size={12} />
                        </button>
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="end" className="w-32">
                        <DropdownMenuItem
                          className="text-xs text-destructive"
                          onClick={e => {
                            e.stopPropagation();
                            if (window.confirm(`Delete folder "${f.name}"?`)) deleteFolderMutation.mutate(f.id);
                          }}
                        >
                          <Trash2 size={12} className="mr-1.5" /> Delete
                        </DropdownMenuItem>
                      </DropdownMenuContent>
                    </DropdownMenu>
                  </div>
                ))}
                {folders.length === 0 && !creatingFolder && (
                  <p className="text-[10px] text-muted-foreground px-2">No folders yet</p>
                )}
              </div>
            )}
          </div>

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
              <div className="p-4 flex flex-col items-center text-center mt-4">
                <div className="w-10 h-10 rounded-full bg-muted flex items-center justify-center mb-2">
                  <FolderOpen size={16} className="text-muted-foreground" />
                </div>
                <p className="text-xs font-medium text-foreground">No {discipline} projects yet</p>
                <p className="text-[10px] text-muted-foreground mt-0.5">Upload a PDF to start.</p>
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
                      <p className="text-[10px] text-muted-foreground">{(p as any).itemCount ?? p.items.length} items</p>
                      <p className="text-[10px] text-muted-foreground">{new Date(p.createdAt).toLocaleDateString()}</p>
                    </div>
                    <div className="flex flex-col gap-0.5 opacity-0 group-hover:opacity-100 shrink-0">
                      <button
                        className="text-muted-foreground hover:text-foreground p-0.5"
                        onClick={e => { e.stopPropagation(); archiveMutation.mutate({ id: p.id, archived: !p.archived }); }}
                        title={p.archived ? "Restore" : "Archive"}
                        aria-label={p.archived ? `Restore ${p.name}` : `Archive ${p.name}`}
                      >
                        {p.archived ? <ArchiveRestore size={13} /> : <Archive size={13} />}
                      </button>
                      <button
                        className="text-muted-foreground hover:text-destructive p-0.5"
                        onClick={e => { e.stopPropagation(); if (window.confirm(`Delete "${p.name}"? This cannot be undone.`)) deleteMutation.mutate(p.id); }}
                        data-testid={`btn-delete-${p.id}`}
                        title="Delete permanently"
                        aria-label={`Delete ${p.name}`}
                      >
                        <Trash2 size={13} />
                      </button>
                      {folders.length > 0 && (
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild onClick={e => e.stopPropagation()}>
                            <button className="text-muted-foreground hover:text-foreground p-0.5" title="Move to folder">
                              <Folder size={13} />
                            </button>
                          </DropdownMenuTrigger>
                          <DropdownMenuContent align="end" className="w-40">
                            {folders.map((f: any) => (
                              <DropdownMenuItem
                                key={f.id}
                                className="text-xs"
                                onClick={e => { e.stopPropagation(); moveToFolderMutation.mutate({ folderId: f.id, projectId: p.id }); }}
                              >
                                <Folder size={11} className="mr-1.5 text-teal-600 dark:text-teal-400" /> {f.name}
                              </DropdownMenuItem>
                            ))}
                          </DropdownMenuContent>
                        </DropdownMenu>
                      )}
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

            {/* Combined BOM view for selected folder */}
            {selectedFolderId && folderBom && (
              <div className="space-y-4">
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <Folder size={16} className="text-teal-600 dark:text-teal-400" />
                      <h2 className="text-base font-semibold">{folderBom.folder?.name}</h2>
                      <Badge variant="outline" className="text-xs">{folderBom.projects?.length || 0} projects</Badge>
                      <Badge variant="outline" className="text-xs">{folderBom.combinedItems?.length || 0} total items</Badge>
                    </div>
                    {folderBom.folder?.description && (
                      <p className="text-xs text-muted-foreground mt-1">{folderBom.folder.description}</p>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                      onClick={async () => {
                        try {
                          const res = await apiRequest("GET", `/api/folders/${selectedFolderId}/export-combined`);
                          const blob = await res.blob();
                          const a = document.createElement("a");
                          a.href = URL.createObjectURL(blob);
                          a.download = `${folderBom.folder?.name || "Combined"} - Combined BOM.xlsx`;
                          a.click();
                          URL.revokeObjectURL(a.href);
                          toast({ title: "Downloaded Combined BOM workbook" });
                        } catch { toast({ title: "Export failed", variant: "destructive" }); }
                      }}
                    >
                      <FileSpreadsheet size={14} className="mr-1.5" />
                      Export Combined
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => setSelectedFolderId(null)}
                    >
                      Close Folder
                    </Button>
                  </div>
                </div>

                {/* Folder projects */}
                {folderBom.projects?.length > 0 && (
                  <div className="flex gap-2 flex-wrap">
                    {folderBom.projects.map((fp: any) => (
                      <div key={fp.id} className="flex items-center gap-2 px-3 py-1.5 rounded-md bg-card border border-border text-xs">
                        <FileText size={12} className="text-muted-foreground" />
                        <span className="font-medium">{fp.name}</span>
                        <span className="text-muted-foreground">{fp.items?.length || 0} items</span>
                        <button
                          className="text-muted-foreground hover:text-destructive ml-1"
                          onClick={() => removeFromFolderMutation.mutate({ folderId: selectedFolderId!, projectId: fp.id })}
                          title="Remove from folder"
                        >
                          <XIcon size={12} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}

                {/* Combined BOM */}
                {folderBom.combinedItems?.length > 0 && (
                  <Tabs defaultValue="bom">
                    <TabsList className="mb-3">
                      <TabsTrigger value="bom" className="text-xs">Combined BOM</TabsTrigger>
                      <TabsTrigger value="connections" className="text-xs">Connections</TabsTrigger>
                      <TabsTrigger value="pivot" className="text-xs">Pivot Summary</TabsTrigger>
                    </TabsList>
                    <TabsContent value="bom">
                      <TakeoffBomTable items={folderBom.combinedItems} discipline={discipline} onItemUpdated={() => queryClient.invalidateQueries({ queryKey: ["/api/folders", selectedFolderId, "combined-bom"] })} />
                    </TabsContent>
                    <TabsContent value="connections">
                      <ConnectionsSummary items={folderBom.combinedItems} />
                    </TabsContent>
                    <TabsContent value="pivot">
                      <PivotSummary items={folderBom.combinedItems} />
                    </TabsContent>
                  </Tabs>
                )}
              </div>
            )}

            {/* Selected project view */}
            {selectedId && selectedProject && !selectedFolderId ? (
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
                    <Button
                      variant="outline"
                      size="sm"
                      className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                      onClick={async () => {
                        try {
                          const res = await apiRequest("GET", `/api/takeoff-projects/${selectedProject.id}/export-bom`);
                          const blob = await res.blob();
                          const a = document.createElement("a");
                          a.href = URL.createObjectURL(blob);
                          a.download = `${selectedProject.name} - BOM.xlsx`;
                          a.click();
                          URL.revokeObjectURL(a.href);
                          toast({ title: "Downloaded BOM workbook" });
                        } catch { toast({ title: "Export failed", variant: "destructive" }); }
                      }}
                    >
                      <FileSpreadsheet size={14} className="mr-1.5" />
                      Export BOM
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                      onClick={async () => {
                        try {
                          const res = await apiRequest("GET", `/api/takeoff-projects/${selectedProject.id}/export-connections`);
                          const blob = await res.blob();
                          const a = document.createElement("a");
                          a.href = URL.createObjectURL(blob);
                          a.download = `${selectedProject.name} - Connections.xlsx`;
                          a.click();
                          URL.revokeObjectURL(a.href);
                          toast({ title: "Downloaded Connections workbook" });
                        } catch { toast({ title: "Export failed", variant: "destructive" }); }
                      }}
                    >
                      <FileSpreadsheet size={14} className="mr-1.5" />
                      Export Connections
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      className="text-purple-700 border-purple-300 hover:bg-purple-50 dark:text-purple-400 dark:border-purple-700 dark:hover:bg-purple-900/30"
                      title="Re-run dedup with the latest rules (catches duplicate drawing-numbers like multi-revision and multi-binding)"
                      onClick={async () => {
                        try {
                          const res = await apiRequest("POST", `/api/takeoff-projects/${selectedProject.id}/redup`);
                          const data = await res.json();
                          toast({ title: data.message || "Dedup complete" });
                          // Refresh project data
                          queryClient.invalidateQueries({ queryKey: ["/api/takeoff-projects", selectedProject.id] });
                          queryClient.invalidateQueries({ queryKey: ["/api/takeoff-projects"] });
                        } catch { toast({ title: "Re-dedup failed", variant: "destructive" }); }
                      }}
                    >
                      Re-run Dedup
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

                {/* Scope Gap Alerts — collapsible */}
                {scopeGaps.length > 0 && !scopeGapsHidden && (
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <h3 className="text-xs font-semibold text-amber-700 dark:text-amber-400 uppercase tracking-wider flex items-center gap-1.5">
                        <AlertTriangle size={13} /> Scope Gap Warnings ({scopeGaps.length})
                      </h3>
                      <button onClick={() => setScopeGapsHidden(true)} className="text-[10px] text-muted-foreground hover:text-foreground">
                        Hide
                      </button>
                    </div>
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
                {scopeGaps.length > 0 && scopeGapsHidden && (
                  <button onClick={() => setScopeGapsHidden(false)} className="text-[10px] text-amber-600 dark:text-amber-400 hover:underline flex items-center gap-1">
                    <AlertTriangle size={10} /> Show {scopeGaps.length} scope gap warnings
                  </button>
                )}

                {/* Verification Viewer */}
                {verifyOpen && availablePages.length > 0 && (() => {
                  // Compute drawing-number label for the current page (shows up
                  // alongside the page number so the estimator can locate the ISO).
                  const pageDrawing: Record<number, string> = {};
                  for (const it of (selectedProject?.items || [])) {
                    if (it.sourcePage && (it as any).drawingNumber && !pageDrawing[it.sourcePage]) {
                      pageDrawing[it.sourcePage] = (it as any).drawingNumber;
                    }
                  }
                  const currentIdx = verifyPage != null ? availablePages.indexOf(verifyPage) : -1;
                  const goPrev = () => { if (currentIdx > 0) setVerifyPage(availablePages[currentIdx - 1]); };
                  const goNext = () => { if (currentIdx >= 0 && currentIdx < availablePages.length - 1) setVerifyPage(availablePages[currentIdx + 1]); };
                  return (
                  <Card className="border-card-border shadow-sm">
                    <CardHeader className="p-3 pb-2">
                      <CardTitle className="text-xs flex items-center justify-between gap-2">
                        <span className="flex items-center gap-1.5">
                          <Image size={12} /> ISO Page Viewer
                          {verifyPage != null && (
                            <span className="text-muted-foreground font-normal ml-2">
                              Page {verifyPage} of {availablePages.length}
                              {pageDrawing[verifyPage] && <span className="ml-2">· {pageDrawing[verifyPage]}</span>}
                            </span>
                          )}
                        </span>
                        <span className="flex items-center gap-1">
                          <Button variant="ghost" size="sm" className="h-6 w-6 p-0" onClick={goPrev} disabled={currentIdx <= 0} title="Previous page">←</Button>
                          <Button variant="ghost" size="sm" className="h-6 w-6 p-0" onClick={goNext} disabled={currentIdx < 0 || currentIdx >= availablePages.length - 1} title="Next page">→</Button>
                          {verifyPage != null && (
                            <Button variant="outline" size="sm" className="h-6 px-2 text-[10px]" onClick={() => window.open(`/api/takeoff/projects/${selectedId}/page/${verifyPage}`, "_blank")} title="Open full size in new tab">⛶ Open</Button>
                          )}
                        </span>
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="p-3 pt-0">
                      <div className="flex gap-1 mb-2 flex-wrap max-h-24 overflow-y-auto">
                        {availablePages.map(pg => (
                          <Button
                            key={pg}
                            variant={verifyPage === pg ? "default" : "outline"}
                            size="sm"
                            className="h-7 text-xs px-2"
                            onClick={() => setVerifyPage(pg)}
                            title={pageDrawing[pg] || `Page ${pg}`}
                          >
                            {pageDrawing[pg] ? <>P{pg} <span className="ml-1 text-[9px] text-muted-foreground">{pageDrawing[pg].slice(0, 16)}</span></> : <>Page {pg}</>}
                          </Button>
                        ))}
                      </div>
                      {verifyPage != null && (
                        <div className="border rounded-md overflow-hidden bg-white dark:bg-black cursor-pointer" onClick={() => window.open(`/api/takeoff/projects/${selectedId}/page/${verifyPage}`, "_blank")} title="Click to open full size">
                          <img
                            src={`/api/takeoff/projects/${selectedId}/page/${verifyPage}`}
                            alt={`Page ${verifyPage}`}
                            className="w-full h-auto max-h-[700px] object-contain"
                            onError={(e) => { (e.target as HTMLImageElement).alt = "Thumbnail not available"; }}
                          />
                        </div>
                      )}
                    </CardContent>
                  </Card>
                  );
                })()}

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
                    <TabsTrigger value="review" className="text-xs" data-testid="tab-review">
                      Review
                      {(() => {
                        const needsReview = selectedProject.items.filter(i => {
                          if ((i as any)._dedupCandidate) return false;
                          const rs = i.reviewStatus || "unreviewed";
                          return rs === "unreviewed" || rs === "rejected";
                        }).length;
                        return needsReview > 0 ? <Badge className="ml-1.5 h-4 px-1 text-[10px] bg-amber-500">{needsReview}</Badge> : null;
                      })()}
                    </TabsTrigger>
                    <TabsTrigger value="connections" className="text-xs" data-testid="tab-connections">Connections</TabsTrigger>
                    <TabsTrigger value="pivot" className="text-xs" data-testid="tab-pivot">Pivot Summary</TabsTrigger>
                    <TabsTrigger value="fab-scope" className="text-xs">Fab Scope</TabsTrigger>
                  </TabsList>
                  <TabsContent value="bom">
                    <TakeoffBomTable items={selectedProject.items} discipline={discipline} onItemUpdated={() => queryClient.invalidateQueries({ queryKey: ["/api/takeoff/projects", selectedId] })} />
                  </TabsContent>
                  <TabsContent value="review">
                    <ReviewMode project={selectedProject} />
                  </TabsContent>
                  <TabsContent value="connections">
                    <ConnectionsSummary items={selectedProject.items} />
                  </TabsContent>
                  <TabsContent value="pivot">
                    <PivotSummary items={selectedProject.items} />
                  </TabsContent>
                  <TabsContent value="fab-scope">
                    <FabScopeSplitter items={selectedProject.items} />
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
