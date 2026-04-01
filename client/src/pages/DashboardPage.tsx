import { useQuery } from "@tanstack/react-query";
import { Link } from "wouter";
import { Wrench, Building2, HardHat, Calculator, FileText, Download, TrendingUp, Package, Clock, BarChart3, FolderOpen } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import AppLayout from "@/components/AppLayout";
import { apiRequest } from "@/lib/queryClient";
import type { TakeoffProject, EstimateProject } from "@shared/schema";

const DISCIPLINE_CONFIG = {
  mechanical: { label: "Mechanical", color: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300", icon: Wrench, accent: "border-l-blue-500" },
  structural: { label: "Structural", color: "bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300", icon: Building2, accent: "border-l-purple-500" },
  civil: { label: "Civil", color: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300", icon: HardHat, accent: "border-l-green-500" },
};

function relativeDate(dateStr: string): string {
  const d = new Date(dateStr);
  const now = new Date();
  const diffMs = now.getTime() - d.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  if (diffMins < 1) return "Just now";
  if (diffMins < 60) return `${diffMins}m ago`;
  const diffHrs = Math.floor(diffMins / 60);
  if (diffHrs < 24) return `${diffHrs}h ago`;
  const diffDays = Math.floor(diffHrs / 24);
  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return `${diffDays}d ago`;
  return d.toLocaleDateString();
}

export default function DashboardPage() {
  const { data: takeoffs = [], isLoading: loadingTakeoffs } = useQuery<TakeoffProject[]>({
    queryKey: ["/api/takeoff/projects"],
  });
  const { data: estimates = [], isLoading: loadingEstimates } = useQuery<EstimateProject[]>({
    queryKey: ["/api/estimates"],
  });

  const activeTakeoffs = takeoffs.filter(t => !t.archived);
  const recentTakeoffs = activeTakeoffs.slice(0, 6);
  const recentEstimates = estimates.slice(0, 4);
  const totalItems = activeTakeoffs.reduce((s, t) => s + t.items.length, 0);
  const today = new Date().toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" });

  return (
    <AppLayout subtitle="Dashboard">
      <div className="p-6 max-w-7xl mx-auto space-y-8">
        {/* Welcome Banner */}
        <div className="relative overflow-hidden rounded-xl bg-gradient-to-r from-[#01696F] to-[#1E3448] p-6 md:p-8 text-white shadow-lg">
          <div className="absolute inset-0 bg-[url('data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDAiIGhlaWdodD0iNDAiIHZpZXdCb3g9IjAgMCA0MCA0MCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cGF0aCBkPSJNMCAwaDQwdjQwSDBWMHptMSAxdjM4aDM4VjFIMXoiIGZpbGw9InJnYmEoMjU1LDI1NSwyNTUsMC4wMykiLz48L3N2Zz4=')] opacity-50" />
          <div className="relative z-10">
            <h1 className="text-xl md:text-2xl font-bold tracking-tight">Picou Group Contractors</h1>
            <p className="text-white/70 text-sm mt-1">Takeoff & Estimating Dashboard</p>
            <p className="text-white/50 text-xs mt-3">{today}</p>
          </div>
          <div className="absolute right-6 top-6 hidden md:block">
            <Button
              variant="outline"
              size="sm"
              data-testid="btn-download-backup"
              className="bg-white/10 border-white/20 text-white hover:bg-white/20 hover:text-white"
              onClick={async () => {
                try {
                  const res = await apiRequest("GET", "/api/backup");
                  const blob = await res.blob();
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement("a");
                  a.href = url;
                  a.download = `pg-unified-backup-${new Date().toISOString().slice(0, 10)}.json`;
                  a.click();
                  URL.revokeObjectURL(url);
                } catch (err: any) {
                  console.error("Backup failed:", err);
                }
              }}
            >
              <Download size={14} className="mr-1.5" />
              Backup
            </Button>
          </div>
        </div>

        {/* Stats Row */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          {[
            { label: "Active Takeoffs", val: activeTakeoffs.length, icon: FileText, iconBg: "bg-blue-100 dark:bg-blue-900/40", iconColor: "text-blue-600 dark:text-blue-400", border: "border-l-blue-500" },
            { label: "Mechanical", val: activeTakeoffs.filter(t => t.discipline === "mechanical").length, icon: Wrench, iconBg: "bg-sky-100 dark:bg-sky-900/40", iconColor: "text-sky-600 dark:text-sky-400", border: "border-l-sky-500" },
            { label: "Structural", val: activeTakeoffs.filter(t => t.discipline === "structural").length, icon: Building2, iconBg: "bg-purple-100 dark:bg-purple-900/40", iconColor: "text-purple-600 dark:text-purple-400", border: "border-l-purple-500" },
            { label: "Estimates", val: estimates.length, icon: Calculator, iconBg: "bg-amber-100 dark:bg-amber-900/40", iconColor: "text-amber-600 dark:text-amber-400", border: "border-l-amber-500" },
          ].map(({ label, val, icon: Icon, iconBg, iconColor, border }) => (
            <Card key={label} className={`border-l-4 ${border} shadow-sm hover:shadow-md transition-shadow`}>
              <CardContent className="p-4 flex items-center gap-3">
                <div className={`w-10 h-10 rounded-lg ${iconBg} flex items-center justify-center shrink-0`}>
                  <Icon size={18} className={iconColor} />
                </div>
                <div>
                  <p className="text-2xl font-bold leading-none">{val}</p>
                  <p className="text-xs text-muted-foreground mt-1">{label}</p>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>

        {/* Quick Actions */}
        <div>
          <h2 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider mb-3">Quick Actions</h2>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <Link href="/mechanical">
              <a data-testid="btn-new-mechanical">
                <Card className="cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border">
                  <CardContent className="p-4 flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg bg-blue-100 dark:bg-blue-900/40 flex items-center justify-center shrink-0">
                      <Wrench size={16} className="text-blue-700 dark:text-blue-400" />
                    </div>
                    <div>
                      <p className="text-sm font-medium">Mechanical</p>
                      <p className="text-xs text-muted-foreground">Piping BOM</p>
                    </div>
                  </CardContent>
                </Card>
              </a>
            </Link>
            <Link href="/structural">
              <a data-testid="btn-new-structural">
                <Card className="cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border">
                  <CardContent className="p-4 flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg bg-purple-100 dark:bg-purple-900/40 flex items-center justify-center shrink-0">
                      <Building2 size={16} className="text-purple-700 dark:text-purple-400" />
                    </div>
                    <div>
                      <p className="text-sm font-medium">Structural</p>
                      <p className="text-xs text-muted-foreground">Steel/Concrete</p>
                    </div>
                  </CardContent>
                </Card>
              </a>
            </Link>
            <Link href="/civil">
              <a data-testid="btn-new-civil">
                <Card className="cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border">
                  <CardContent className="p-4 flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg bg-green-100 dark:bg-green-900/40 flex items-center justify-center shrink-0">
                      <HardHat size={16} className="text-green-700 dark:text-green-400" />
                    </div>
                    <div>
                      <p className="text-sm font-medium">Civil</p>
                      <p className="text-xs text-muted-foreground">Utilities/Sitework</p>
                    </div>
                  </CardContent>
                </Card>
              </a>
            </Link>
            <Link href="/estimating">
              <a data-testid="btn-new-estimate">
                <Card className="cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border">
                  <CardContent className="p-4 flex items-center gap-3">
                    <div className="w-9 h-9 rounded-lg bg-amber-100 dark:bg-amber-900/40 flex items-center justify-center shrink-0">
                      <Calculator size={16} className="text-amber-700 dark:text-amber-400" />
                    </div>
                    <div>
                      <p className="text-sm font-medium">Estimating</p>
                      <p className="text-xs text-muted-foreground">Cost Estimate</p>
                    </div>
                  </CardContent>
                </Card>
              </a>
            </Link>
          </div>
        </div>

        {/* Recent Takeoffs */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider">Recent Takeoffs</h2>
            {activeTakeoffs.length > 6 && (
              <Link href="/mechanical"><a className="text-xs text-primary hover:underline">View all</a></Link>
            )}
          </div>
          {loadingTakeoffs ? (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              {[1, 2, 3].map(i => <Skeleton key={i} className="h-24 rounded-lg" />)}
            </div>
          ) : recentTakeoffs.length === 0 ? (
            <Card className="shadow-sm border-card-border border-dashed">
              <CardContent className="p-10 flex flex-col items-center justify-center text-center">
                <div className="w-12 h-12 rounded-full bg-muted flex items-center justify-center mb-3">
                  <FolderOpen size={20} className="text-muted-foreground" />
                </div>
                <p className="text-sm font-medium text-foreground">No takeoffs yet</p>
                <p className="text-xs text-muted-foreground mt-1">Upload a PDF to get started with your first takeoff.</p>
              </CardContent>
            </Card>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              {recentTakeoffs.map(project => {
                const cfg = DISCIPLINE_CONFIG[project.discipline];
                const Icon = cfg.icon;
                return (
                  <Link key={project.id} href={`/${project.discipline}`}>
                    <a data-testid={`card-takeoff-${project.id}`}>
                      <Card className={`cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border border-l-4 ${cfg.accent} h-full`}>
                        <CardContent className="p-4">
                          <div className="flex items-start gap-3">
                            <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${cfg.color}`}>
                              <Icon size={14} />
                            </div>
                            <div className="min-w-0 flex-1">
                              <p className="text-sm font-medium truncate">{project.name}</p>
                              <div className="flex items-center gap-2 mt-1">
                                <Badge variant="outline" className={`text-[10px] px-1.5 py-0 ${cfg.color}`}>
                                  {cfg.label}
                                </Badge>
                                <span className="text-xs text-muted-foreground">{project.items.length} items</span>
                              </div>
                              <p className="text-[11px] text-muted-foreground mt-1.5 flex items-center gap-1">
                                <Clock size={10} className="shrink-0" />
                                {relativeDate(project.createdAt)}
                              </p>
                            </div>
                          </div>
                        </CardContent>
                      </Card>
                    </a>
                  </Link>
                );
              })}
            </div>
          )}
        </div>

        {/* Recent Estimates */}
        {(recentEstimates.length > 0 || !loadingEstimates) && (
          <div>
            <h2 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider mb-3">Recent Estimates</h2>
            {loadingEstimates ? (
              <Skeleton className="h-20 rounded-lg" />
            ) : recentEstimates.length === 0 ? (
              <Card className="shadow-sm border-card-border border-dashed">
                <CardContent className="p-10 flex flex-col items-center justify-center text-center">
                  <div className="w-12 h-12 rounded-full bg-muted flex items-center justify-center mb-3">
                    <Calculator size={20} className="text-muted-foreground" />
                  </div>
                  <p className="text-sm font-medium text-foreground">No estimates yet</p>
                  <p className="text-xs text-muted-foreground mt-1">Create an estimate from the Estimating page.</p>
                </CardContent>
              </Card>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {recentEstimates.map(est => {
                  const total = est.items.reduce((s, i) => s + (i.totalCost || 0), 0);
                  return (
                    <Link key={est.id} href="/estimating">
                      <a data-testid={`card-estimate-${est.id}`}>
                        <Card className="cursor-pointer shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 border-card-border border-l-4 border-l-amber-500">
                          <CardContent className="p-4 flex items-center justify-between">
                            <div>
                              <p className="text-sm font-medium">{est.name}</p>
                              <p className="text-xs text-muted-foreground mt-0.5 flex items-center gap-1">
                                {est.items.length} items
                                <span className="mx-1">&#183;</span>
                                <Clock size={10} className="shrink-0" />
                                {relativeDate(est.createdAt)}
                              </p>
                            </div>
                            <div className="text-right">
                              <p className="text-sm font-semibold text-primary">${total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                              <p className="text-xs text-muted-foreground">subtotal</p>
                            </div>
                          </CardContent>
                        </Card>
                      </a>
                    </Link>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* Quick Stats */}
        <div>
          <h2 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider mb-3">Quick Stats</h2>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <Card className="shadow-sm">
              <CardContent className="p-4 flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-teal-100 dark:bg-teal-900/40 flex items-center justify-center shrink-0">
                  <Package size={16} className="text-teal-600 dark:text-teal-400" />
                </div>
                <div>
                  <p className="text-lg font-bold leading-none">{totalItems.toLocaleString()}</p>
                  <p className="text-[11px] text-muted-foreground mt-1">Total Items</p>
                </div>
              </CardContent>
            </Card>
            <Card className="shadow-sm">
              <CardContent className="p-4 flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-green-100 dark:bg-green-900/40 flex items-center justify-center shrink-0">
                  <HardHat size={16} className="text-green-600 dark:text-green-400" />
                </div>
                <div>
                  <p className="text-lg font-bold leading-none">{activeTakeoffs.filter(t => t.discipline === "civil").length}</p>
                  <p className="text-[11px] text-muted-foreground mt-1">Civil Projects</p>
                </div>
              </CardContent>
            </Card>
            <Card className="shadow-sm">
              <CardContent className="p-4 flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-indigo-100 dark:bg-indigo-900/40 flex items-center justify-center shrink-0">
                  <BarChart3 size={16} className="text-indigo-600 dark:text-indigo-400" />
                </div>
                <div>
                  <p className="text-lg font-bold leading-none">{takeoffs.filter(t => t.archived).length}</p>
                  <p className="text-[11px] text-muted-foreground mt-1">Archived</p>
                </div>
              </CardContent>
            </Card>
            <Card className="shadow-sm">
              <CardContent className="p-4 flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-rose-100 dark:bg-rose-900/40 flex items-center justify-center shrink-0">
                  <TrendingUp size={16} className="text-rose-600 dark:text-rose-400" />
                </div>
                <div>
                  <p className="text-lg font-bold leading-none">
                    ${estimates.reduce((s, e) => s + e.items.reduce((t, i) => t + (i.totalCost || 0), 0), 0).toLocaleString(undefined, { maximumFractionDigits: 0 })}
                  </p>
                  <p className="text-[11px] text-muted-foreground mt-1">Est. Value</p>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        {/* Mobile backup button */}
        <div className="md:hidden flex justify-center pb-4">
          <Button
            variant="outline"
            size="sm"
            data-testid="btn-download-backup-mobile"
            onClick={async () => {
              try {
                const res = await apiRequest("GET", "/api/backup");
                const blob = await res.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `pg-unified-backup-${new Date().toISOString().slice(0, 10)}.json`;
                a.click();
                URL.revokeObjectURL(url);
              } catch (err: any) {
                console.error("Backup failed:", err);
              }
            }}
          >
            <Download size={14} className="mr-1.5" />
            Download Backup
          </Button>
        </div>
      </div>
    </AppLayout>
  );
}
