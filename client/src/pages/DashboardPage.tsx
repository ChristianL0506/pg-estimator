import { useQuery } from "@tanstack/react-query";
import { Link } from "wouter";
import { Wrench, Building2, HardHat, Calculator, FileText, Plus, Download } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import AppLayout from "@/components/AppLayout";
import { apiRequest } from "@/lib/queryClient";
import type { TakeoffProject, EstimateProject } from "@shared/schema";

const DISCIPLINE_CONFIG = {
  mechanical: { label: "Mechanical", color: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300", icon: Wrench },
  structural: { label: "Structural", color: "bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300", icon: Building2 },
  civil: { label: "Civil", color: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300", icon: HardHat },
};

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

  return (
    <AppLayout subtitle="Dashboard">
      <div className="p-6 max-w-7xl mx-auto">
        {/* Quick Actions */}
        <div className="mb-8">
          <h2 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider mb-3">Quick Actions</h2>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <Link href="/mechanical">
              <a data-testid="btn-new-mechanical">
                <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border">
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
                <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border">
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
                <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border">
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
                <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border">
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

        {/* Backup Button */}
        <div className="mb-6 flex justify-end">
          <Button
            variant="outline"
            size="sm"
            data-testid="btn-download-backup"
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

        {/* Stats Row */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-8">
          {[
            { label: "Total Takeoffs", val: activeTakeoffs.length, icon: FileText },
            { label: "Mechanical", val: activeTakeoffs.filter(t => t.discipline === "mechanical").length, icon: Wrench },
            { label: "Structural", val: activeTakeoffs.filter(t => t.discipline === "structural").length, icon: Building2 },
            { label: "Estimates", val: estimates.length, icon: Calculator },
          ].map(({ label, val, icon: Icon }) => (
            <Card key={label} className="border-card-border">
              <CardContent className="p-4">
                <p className="text-xs text-muted-foreground">{label}</p>
                <p className="text-2xl font-semibold mt-1">{val}</p>
              </CardContent>
            </Card>
          ))}
        </div>

        {/* Recent Takeoffs */}
        <div className="mb-8">
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
            <Card className="border-card-border border-dashed">
              <CardContent className="p-8 text-center text-muted-foreground text-sm">
                No takeoffs yet. Upload a PDF to get started.
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
                      <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border h-full">
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
                              <p className="text-xs text-muted-foreground mt-1">
                                {new Date(project.createdAt).toLocaleDateString()}
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
              <Card className="border-card-border border-dashed">
                <CardContent className="p-6 text-center text-muted-foreground text-sm">
                  No estimates yet.
                </CardContent>
              </Card>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {recentEstimates.map(est => {
                  const total = est.items.reduce((s, i) => s + (i.totalCost || 0), 0);
                  return (
                    <Link key={est.id} href="/estimating">
                      <a data-testid={`card-estimate-${est.id}`}>
                        <Card className="cursor-pointer hover:bg-accent transition-colors border-card-border">
                          <CardContent className="p-4 flex items-center justify-between">
                            <div>
                              <p className="text-sm font-medium">{est.name}</p>
                              <p className="text-xs text-muted-foreground mt-0.5">
                                {est.items.length} items · {new Date(est.createdAt).toLocaleDateString()}
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
      </div>
    </AppLayout>
  );
}
