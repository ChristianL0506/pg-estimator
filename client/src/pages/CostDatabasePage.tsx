import { useState, useMemo, useRef } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import {
  Search, Upload, Trash2, DollarSign, TrendingUp, Building2, Package,
  Filter, ChevronDown, ChevronRight, FileText, BarChart3, ArrowUpDown, X,
  AlertTriangle, Plus, ShoppingCart,
} from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import AppLayout from "@/components/AppLayout";
import { apiRequest, queryClient, apiUpload } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";
import type { PurchaseRecord } from "@shared/schema";

const CATEGORY_COLORS: Record<string, string> = {
  pipe: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  elbow: "bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300",
  tee: "bg-indigo-100 text-indigo-800 dark:bg-indigo-900/40 dark:text-indigo-300",
  reducer: "bg-cyan-100 text-cyan-800 dark:bg-cyan-900/40 dark:text-cyan-300",
  valve: "bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300",
  flange: "bg-orange-100 text-orange-800 dark:bg-orange-900/40 dark:text-orange-300",
  gasket: "bg-yellow-100 text-yellow-800 dark:bg-yellow-900/40 dark:text-yellow-300",
  bolt: "bg-gray-100 text-gray-800 dark:bg-gray-800/40 dark:text-gray-300",
  support: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300",
  other: "bg-slate-100 text-slate-800 dark:bg-slate-800/40 dark:text-slate-300",
};

function getCategoryColor(cat: string) {
  return CATEGORY_COLORS[cat.toLowerCase()] || CATEGORY_COLORS.other;
}

function formatCurrency(val: number) {
  return val.toLocaleString(undefined, { style: "currency", currency: "USD", minimumFractionDigits: 2 });
}

function parseDate(d: string | undefined): Date | null {
  if (!d) return null;
  const iso = new Date(d);
  if (!isNaN(iso.getTime())) return iso;
  const parts = d.split("/");
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
  }
  return null;
}

export default function CostDatabasePage() {
  const { toast } = useToast();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const vendorFileInputRef = useRef<HTMLInputElement>(null);
  const [search, setSearch] = useState("");
  const [supplierFilter, setSupplierFilter] = useState<string>("");
  const [categoryFilter, setCategoryFilter] = useState<string>("");
  const [sortField, setSortField] = useState<"date" | "cost" | "supplier" | "description">("date");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");
  const [activeTab, setActiveTab] = useState("history");

  // Vendor quote add form
  const [showVendorForm, setShowVendorForm] = useState(false);
  const [vfVendor, setVfVendor] = useState("");
  const [vfDesc, setVfDesc] = useState("");
  const [vfSize, setVfSize] = useState("");
  const [vfCategory, setVfCategory] = useState("other");
  const [vfUnit, setVfUnit] = useState("EA");
  const [vfUnitPrice, setVfUnitPrice] = useState(0);
  const [vfQty, setVfQty] = useState(1);
  const [vfQuoteNum, setVfQuoteNum] = useState("");
  const [vfNotes, setVfNotes] = useState("");

  const { data: records = [], isLoading } = useQuery<PurchaseRecord[]>({
    queryKey: ["/api/purchase-history"],
  });

  const { data: suppliers = [] } = useQuery<any[]>({
    queryKey: ["/api/purchase-history/suppliers"],
  });

  const { data: categories = [] } = useQuery<any[]>({
    queryKey: ["/api/purchase-history/categories"],
  });

  const { data: alerts = [] } = useQuery<any[]>({
    queryKey: ["/api/cost-database/alerts"],
  });

  const { data: vendorQuotes = [] } = useQuery<any[]>({
    queryKey: ["/api/vendor-quotes"],
  });

  const { data: vendorComparisons = [] } = useQuery<any[]>({
    queryKey: ["/api/vendor-quotes/compare"],
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/purchase-history/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history"] });
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history/suppliers"] });
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history/categories"] });
      queryClient.invalidateQueries({ queryKey: ["/api/cost-database/alerts"] });
    },
  });

  const importMutation = useMutation({
    mutationFn: async (file: File) => {
      const formData = new FormData();
      formData.append("file", file);
      const res = await apiUpload("/api/purchase-history/import", formData);
      return res.json();
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history"] });
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history/suppliers"] });
      queryClient.invalidateQueries({ queryKey: ["/api/purchase-history/categories"] });
      queryClient.invalidateQueries({ queryKey: ["/api/cost-database/alerts"] });
      toast({ title: `Imported ${data.imported} records` });
    },
    onError: (err: any) => {
      toast({ title: "Import failed", description: err.message, variant: "destructive" });
    },
  });

  const addVendorQuoteMutation = useMutation({
    mutationFn: async (data: any) => {
      const res = await apiRequest("POST", "/api/vendor-quotes", data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes"] });
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes/compare"] });
      toast({ title: "Vendor quote added" });
      setShowVendorForm(false);
      setVfVendor(""); setVfDesc(""); setVfSize(""); setVfCategory("other");
      setVfUnit("EA"); setVfUnitPrice(0); setVfQty(1); setVfQuoteNum(""); setVfNotes("");
    },
    onError: (err: any) => { toast({ title: "Error", description: err.message, variant: "destructive" }); },
  });

  const deleteVendorQuoteMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/vendor-quotes/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes"] });
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes/compare"] });
    },
  });

  const importVendorMutation = useMutation({
    mutationFn: async (file: File) => {
      const formData = new FormData();
      formData.append("file", file);
      const res = await apiUpload("/api/vendor-quotes/import", formData);
      return res.json();
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes"] });
      queryClient.invalidateQueries({ queryKey: ["/api/vendor-quotes/compare"] });
      toast({ title: `Imported ${data.imported} vendor quotes` });
    },
    onError: (err: any) => { toast({ title: "Import failed", description: err.message, variant: "destructive" }); },
  });

  // Filtering and sorting
  const filteredRecords = useMemo(() => {
    let filtered = records;
    if (search) {
      const q = search.toLowerCase();
      filtered = filtered.filter(r =>
        r.description.toLowerCase().includes(q) ||
        (r.supplier || "").toLowerCase().includes(q) ||
        (r.invoiceNumber || "").toLowerCase().includes(q) ||
        (r.size || "").toLowerCase().includes(q)
      );
    }
    if (supplierFilter) filtered = filtered.filter(r => r.supplier === supplierFilter);
    if (categoryFilter) filtered = filtered.filter(r => r.category === categoryFilter);

    filtered = [...filtered].sort((a, b) => {
      let cmp = 0;
      if (sortField === "date") {
        const da = parseDate(a.invoiceDate)?.getTime() || 0;
        const db = parseDate(b.invoiceDate)?.getTime() || 0;
        cmp = da - db;
      } else if (sortField === "cost") {
        cmp = a.unitCost - b.unitCost;
      } else if (sortField === "supplier") {
        cmp = (a.supplier || "").localeCompare(b.supplier || "");
      } else {
        cmp = a.description.localeCompare(b.description);
      }
      return sortDir === "desc" ? -cmp : cmp;
    });
    return filtered;
  }, [records, search, supplierFilter, categoryFilter, sortField, sortDir]);

  // Supplier summary for comparison tab
  const supplierSummary = useMemo(() => {
    const map: Record<string, { supplier: string; items: number; totalSpend: number; avgCost: number; latestDate: string }> = {};
    for (const r of records) {
      const s = r.supplier || "Unknown";
      if (!map[s]) map[s] = { supplier: s, items: 0, totalSpend: 0, avgCost: 0, latestDate: "" };
      map[s].items++;
      map[s].totalSpend += r.totalCost || 0;
      const d = r.invoiceDate || "";
      if (d > map[s].latestDate) map[s].latestDate = d;
    }
    for (const v of Object.values(map)) {
      v.avgCost = v.items > 0 ? v.totalSpend / v.items : 0;
    }
    return Object.values(map).sort((a, b) => b.totalSpend - a.totalSpend);
  }, [records]);

  // Price trends
  const categoryTrends = useMemo(() => {
    const byCategory: Record<string, { month: string; avgCost: number; count: number }[]> = {};
    const monthMap: Record<string, Record<string, { sum: number; count: number }>> = {};
    for (const r of records) {
      const d = parseDate(r.invoiceDate);
      if (!d) continue;
      const month = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const cat = r.category || "other";
      if (!monthMap[cat]) monthMap[cat] = {};
      if (!monthMap[cat][month]) monthMap[cat][month] = { sum: 0, count: 0 };
      monthMap[cat][month].sum += r.unitCost;
      monthMap[cat][month].count++;
    }
    for (const [cat, months] of Object.entries(monthMap)) {
      byCategory[cat] = Object.entries(months)
        .map(([month, data]) => ({ month, avgCost: data.sum / data.count, count: data.count }))
        .sort((a, b) => a.month.localeCompare(b.month));
    }
    return byCategory;
  }, [records]);

  const totalSpend = records.reduce((s, r) => s + (r.totalCost || 0), 0);
  const uniqueSuppliers = new Set(records.map(r => r.supplier)).size;
  const uniqueCategories = new Set(records.map(r => r.category)).size;
  const staleAlerts = alerts.filter((a: any) => a.alertType === "stale");
  const shiftAlerts = alerts.filter((a: any) => a.alertType === "price_shift");

  const handleSort = (field: typeof sortField) => {
    if (sortField === field) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortField(field); setSortDir("desc"); }
  };

  return (
    <AppLayout subtitle="Cost Database">
      <div className="p-5 max-w-7xl mx-auto space-y-5">
        {/* Stats row */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <Card className="border-card-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2 text-muted-foreground mb-1">
                <Package size={14} />
                <span className="text-xs">Total Records</span>
              </div>
              <p className="text-2xl font-semibold">{records.length.toLocaleString()}</p>
            </CardContent>
          </Card>
          <Card className="border-card-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2 text-muted-foreground mb-1">
                <DollarSign size={14} />
                <span className="text-xs">Total Spend</span>
              </div>
              <p className="text-2xl font-semibold">{formatCurrency(totalSpend)}</p>
            </CardContent>
          </Card>
          <Card className="border-card-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2 text-muted-foreground mb-1">
                <Building2 size={14} />
                <span className="text-xs">Suppliers</span>
              </div>
              <p className="text-2xl font-semibold">{uniqueSuppliers}</p>
            </CardContent>
          </Card>
          <Card className="border-card-border">
            <CardContent className="p-4">
              <div className="flex items-center gap-2 text-muted-foreground mb-1">
                {alerts.length > 0 ? <AlertTriangle size={14} className="text-amber-500" /> : <BarChart3 size={14} />}
                <span className="text-xs">{alerts.length > 0 ? "Alerts" : "Categories"}</span>
              </div>
              <p className="text-2xl font-semibold">{alerts.length > 0 ? alerts.length : uniqueCategories}</p>
            </CardContent>
          </Card>
        </div>

        {/* Toolbar */}
        <div className="flex items-center gap-3 flex-wrap">
          <div className="relative flex-1 min-w-[200px]">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
            <Input placeholder="Search by item, supplier, invoice #..." value={search} onChange={e => setSearch(e.target.value)} className="pl-9 h-9 text-sm" />
          </div>
          <select className="h-9 text-sm border rounded-md px-3 bg-background text-foreground border-border" value={supplierFilter} onChange={e => setSupplierFilter(e.target.value)}>
            <option value="">All Suppliers</option>
            {suppliers.map((s: any) => <option key={s.supplier} value={s.supplier}>{s.supplier} ({s.itemCount})</option>)}
          </select>
          <select className="h-9 text-sm border rounded-md px-3 bg-background text-foreground border-border" value={categoryFilter} onChange={e => setCategoryFilter(e.target.value)}>
            <option value="">All Categories</option>
            {categories.map((c: any) => <option key={c.category} value={c.category}>{c.category} ({c.itemCount})</option>)}
          </select>
          {(supplierFilter || categoryFilter || search) && (
            <Button variant="ghost" size="sm" onClick={() => { setSearch(""); setSupplierFilter(""); setCategoryFilter(""); }}>
              <X size={14} className="mr-1" /> Clear
            </Button>
          )}
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx" className="hidden" onChange={e => { if (e.target.files?.[0]) importMutation.mutate(e.target.files[0]); e.target.value = ""; }} />
          <Button variant="outline" size="sm" onClick={() => fileInputRef.current?.click()} disabled={importMutation.isPending}>
            <Upload size={14} className="mr-1.5" />
            {importMutation.isPending ? "Importing..." : "Import CSV"}
          </Button>
        </div>

        {/* Tabs */}
        <Tabs value={activeTab} onValueChange={setActiveTab}>
          <TabsList>
            <TabsTrigger value="history" className="text-xs">Purchase History</TabsTrigger>
            <TabsTrigger value="suppliers" className="text-xs">Supplier Comparison</TabsTrigger>
            <TabsTrigger value="trends" className="text-xs">Price Trends</TabsTrigger>
            <TabsTrigger value="alerts" className="text-xs">
              Alerts {alerts.length > 0 && <Badge variant="destructive" className="ml-1 text-[9px] px-1 py-0">{alerts.length}</Badge>}
            </TabsTrigger>
            <TabsTrigger value="vendor-quotes" className="text-xs">
              <ShoppingCart size={12} className="mr-1" /> Vendor Quotes
            </TabsTrigger>
          </TabsList>

          {/* Purchase History Table */}
          <TabsContent value="history">
            <Card className="border-card-border">
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-border bg-muted/30">
                      <th className="text-left p-3 font-medium cursor-pointer hover:text-foreground" onClick={() => handleSort("date")}><div className="flex items-center gap-1">Date <ArrowUpDown size={10} /></div></th>
                      <th className="text-left p-3 font-medium cursor-pointer hover:text-foreground" onClick={() => handleSort("description")}><div className="flex items-center gap-1">Description <ArrowUpDown size={10} /></div></th>
                      <th className="text-left p-3 font-medium">Size</th>
                      <th className="text-left p-3 font-medium">Category</th>
                      <th className="text-right p-3 font-medium">Qty</th>
                      <th className="text-right p-3 font-medium cursor-pointer hover:text-foreground" onClick={() => handleSort("cost")}><div className="flex items-center justify-end gap-1">Unit Cost <ArrowUpDown size={10} /></div></th>
                      <th className="text-right p-3 font-medium">Total</th>
                      <th className="text-left p-3 font-medium cursor-pointer hover:text-foreground" onClick={() => handleSort("supplier")}><div className="flex items-center gap-1">Supplier <ArrowUpDown size={10} /></div></th>
                      <th className="text-left p-3 font-medium">Invoice #</th>
                      <th className="text-center p-3 font-medium w-8"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {isLoading ? (
                      <tr><td colSpan={10} className="p-8 text-center text-muted-foreground">Loading...</td></tr>
                    ) : filteredRecords.length === 0 ? (
                      <tr><td colSpan={10} className="p-8 text-center text-muted-foreground">
                        {records.length === 0 ? "No purchase records yet. Import a CSV to get started." : "No records match your filters."}
                      </td></tr>
                    ) : (
                      filteredRecords.slice(0, 200).map(r => (
                        <tr key={r.id} className="border-b border-border/50 hover:bg-muted/20 transition-colors">
                          <td className="p-3 whitespace-nowrap text-muted-foreground">{r.invoiceDate || "—"}</td>
                          <td className="p-3 max-w-[300px]"><p className="truncate font-medium">{r.description}</p></td>
                          <td className="p-3 whitespace-nowrap">{r.size || "—"}</td>
                          <td className="p-3"><Badge variant="outline" className={`text-[10px] px-1.5 py-0 ${getCategoryColor(r.category)}`}>{r.category}</Badge></td>
                          <td className="p-3 text-right tabular-nums">{r.quantity}</td>
                          <td className="p-3 text-right tabular-nums font-medium">{formatCurrency(r.unitCost)}</td>
                          <td className="p-3 text-right tabular-nums">{formatCurrency(r.totalCost || 0)}</td>
                          <td className="p-3 whitespace-nowrap">{r.supplier}</td>
                          <td className="p-3 whitespace-nowrap text-muted-foreground">{r.invoiceNumber || "—"}</td>
                          <td className="p-3 text-center">
                            <button className="text-muted-foreground hover:text-destructive p-0.5" onClick={() => deleteMutation.mutate(r.id)}><Trash2 size={12} /></button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
              {filteredRecords.length > 200 && (
                <div className="p-3 text-center text-xs text-muted-foreground border-t border-border">
                  Showing 200 of {filteredRecords.length} records. Use filters to narrow down.
                </div>
              )}
            </Card>
          </TabsContent>

          {/* Supplier Comparison */}
          <TabsContent value="suppliers">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Card className="border-card-border">
                <CardHeader className="pb-2"><CardTitle className="text-sm">Supplier Spend Ranking</CardTitle></CardHeader>
                <CardContent>
                  <div className="space-y-2">
                    {supplierSummary.map((s, i) => {
                      const maxSpend = supplierSummary[0]?.totalSpend || 1;
                      return (
                        <div key={s.supplier}>
                          <div className="flex items-center justify-between text-xs mb-1">
                            <button className="font-medium hover:text-primary truncate max-w-[200px]" onClick={() => { setSupplierFilter(s.supplier); setActiveTab("history"); }}>
                              {i + 1}. {s.supplier}
                            </button>
                            <span className="tabular-nums font-medium">{formatCurrency(s.totalSpend)}</span>
                          </div>
                          <div className="h-1.5 bg-muted rounded-full overflow-hidden">
                            <div className="h-full bg-primary rounded-full transition-all" style={{ width: `${(s.totalSpend / maxSpend) * 100}%` }} />
                          </div>
                          <div className="flex justify-between text-[10px] text-muted-foreground mt-0.5">
                            <span>{s.items} items</span><span>Last: {s.latestDate || "N/A"}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </CardContent>
              </Card>
              <Card className="border-card-border">
                <CardHeader className="pb-2"><CardTitle className="text-sm">Price Comparison (Same Items)</CardTitle></CardHeader>
                <CardContent><PriceComparison records={records} /></CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Price Trends */}
          <TabsContent value="trends">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {Object.entries(categoryTrends).map(([cat, months]) => {
                if (months.length < 1) return null;
                const maxCost = Math.max(...months.map(m => m.avgCost));
                return (
                  <Card key={cat} className="border-card-border">
                    <CardHeader className="pb-2">
                      <CardTitle className="text-sm flex items-center gap-2">
                        <Badge variant="outline" className={`text-[10px] ${getCategoryColor(cat)}`}>{cat}</Badge>
                        Average Unit Cost Over Time
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="flex items-end gap-1 h-24">
                        {months.map(m => (
                          <div key={m.month} className="flex-1 flex flex-col items-center gap-1">
                            <div className="relative w-full flex justify-center">
                              <div className="w-full max-w-8 bg-primary/80 rounded-t" style={{ height: `${(m.avgCost / maxCost) * 80}px` }} title={`${m.month}: ${formatCurrency(m.avgCost)} (${m.count} items)`} />
                            </div>
                          </div>
                        ))}
                      </div>
                      <div className="flex gap-1 mt-1">
                        {months.map(m => (<div key={m.month} className="flex-1 text-center text-[9px] text-muted-foreground truncate">{m.month.slice(5)}</div>))}
                      </div>
                      <div className="mt-2 text-[10px] text-muted-foreground">
                        Range: {formatCurrency(Math.min(...months.map(m => m.avgCost)))} — {formatCurrency(Math.max(...months.map(m => m.avgCost)))}
                      </div>
                    </CardContent>
                  </Card>
                );
              })}
              {Object.keys(categoryTrends).length === 0 && (
                <Card className="border-card-border border-dashed col-span-2">
                  <CardContent className="p-8 text-center text-muted-foreground text-sm">No data with dates available for trend analysis.</CardContent>
                </Card>
              )}
            </div>
          </TabsContent>

          {/* Material Escalation Alerts */}
          <TabsContent value="alerts">
            <div className="space-y-4">
              {alerts.length === 0 ? (
                <Card className="border-dashed border-2 border-muted-foreground/20">
                  <CardContent className="p-8 text-center text-muted-foreground text-sm">
                    No material alerts. Alerts appear when prices are stale ({">"}90 days) or show significant shifts ({">"}15%).
                  </CardContent>
                </Card>
              ) : (
                <>
                  {staleAlerts.length > 0 && (
                    <div>
                      <h3 className="text-xs font-semibold text-amber-700 dark:text-amber-400 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                        <AlertTriangle size={13} /> Stale Prices ({staleAlerts.length})
                      </h3>
                      <div className="grid gap-2 sm:grid-cols-2">
                        {staleAlerts.map((alert: any, i: number) => (
                          <Card key={`stale-${i}`} className="border-amber-200 dark:border-amber-800 bg-amber-50/50 dark:bg-amber-900/10">
                            <CardContent className="p-3">
                              <div className="flex items-start justify-between gap-2">
                                <div>
                                  <p className="text-xs font-medium">{alert.category} — {alert.size || "all sizes"}</p>
                                  <p className="text-[10px] text-muted-foreground mt-0.5">{alert.description?.slice(0, 60) || "Price changed"}</p>
                                </div>
                                <Badge variant="outline" className="text-[9px] bg-amber-100 text-amber-800 dark:bg-amber-900/40 dark:text-amber-300 shrink-0">
                                  {alert.daysSinceLastPurchase}d old
                                </Badge>
                              </div>
                              <div className="mt-1 text-[10px] text-muted-foreground">
                                Last price: {formatCurrency(alert.latestPrice || 0)} | Since: {alert.daysSinceLastPurchase}d ago
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    </div>
                  )}
                  {shiftAlerts.length > 0 && (
                    <div>
                      <h3 className="text-xs font-semibold text-red-700 dark:text-red-400 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                        <TrendingUp size={13} /> Price Shifts ({shiftAlerts.length})
                      </h3>
                      <div className="grid gap-2 sm:grid-cols-2">
                        {shiftAlerts.map((alert: any, i: number) => (
                          <Card key={`shift-${i}`} className="border-red-200 dark:border-red-800 bg-red-50/50 dark:bg-red-900/10">
                            <CardContent className="p-3">
                              <div className="flex items-start justify-between gap-2">
                                <div>
                                  <p className="text-xs font-medium">{alert.category} — {alert.size || "all sizes"}</p>
                                  <p className="text-[10px] text-muted-foreground mt-0.5">{alert.description?.slice(0, 60) || "Price changed"}</p>
                                </div>
                                <Badge variant="outline" className="text-[9px] bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300 shrink-0">
                                  {(alert.pctChange || 0) > 0 ? "+" : ""}{(alert.pctChange || 0).toFixed(0)}%
                                </Badge>
                              </div>
                              <div className="mt-1 text-[10px] text-muted-foreground">
                                Avg: {formatCurrency(alert.averagePrice || 0)} | Last: {formatCurrency(alert.latestPrice || 0)}
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          </TabsContent>

          {/* Vendor Quotes */}
          <TabsContent value="vendor-quotes">
            <div className="space-y-4">
              {/* Toolbar */}
              <div className="flex items-center gap-2 flex-wrap">
                <Button size="sm" onClick={() => setShowVendorForm(!showVendorForm)}>
                  <Plus size={14} className="mr-1" /> Add Quote
                </Button>
                <input ref={vendorFileInputRef} type="file" accept=".csv,.xlsx" className="hidden"
                  onChange={e => { if (e.target.files?.[0]) importVendorMutation.mutate(e.target.files[0]); e.target.value = ""; }} />
                <Button variant="outline" size="sm" onClick={() => vendorFileInputRef.current?.click()} disabled={importVendorMutation.isPending}>
                  <Upload size={14} className="mr-1" /> {importVendorMutation.isPending ? "Importing..." : "Import CSV"}
                </Button>
                <span className="text-xs text-muted-foreground ml-auto">{vendorQuotes.length} quotes from {new Set(vendorQuotes.map((q: any) => q.vendorName)).size} vendors</span>
              </div>

              {/* Add form */}
              {showVendorForm && (
                <Card className="border-primary/30">
                  <CardContent className="p-4 space-y-3">
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      <div><Label className="text-xs">Vendor *</Label><Input className="h-8 text-xs mt-1" value={vfVendor} onChange={e => setVfVendor(e.target.value)} placeholder="Vendor name" /></div>
                      <div><Label className="text-xs">Description *</Label><Input className="h-8 text-xs mt-1" value={vfDesc} onChange={e => setVfDesc(e.target.value)} placeholder="Item description" /></div>
                      <div><Label className="text-xs">Size</Label><Input className="h-8 text-xs mt-1" value={vfSize} onChange={e => setVfSize(e.target.value)} placeholder='e.g. 2"' /></div>
                      <div><Label className="text-xs">Quote #</Label><Input className="h-8 text-xs mt-1" value={vfQuoteNum} onChange={e => setVfQuoteNum(e.target.value)} placeholder="Q-12345" /></div>
                    </div>
                    <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                      <div>
                        <Label className="text-xs">Category</Label>
                        <select className="h-8 text-xs mt-1 w-full border rounded px-2 bg-background border-border" value={vfCategory} onChange={e => setVfCategory(e.target.value)}>
                          {Object.keys(CATEGORY_COLORS).map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                      <div><Label className="text-xs">Unit</Label><Input className="h-8 text-xs mt-1" value={vfUnit} onChange={e => setVfUnit(e.target.value)} /></div>
                      <div><Label className="text-xs">Unit Price</Label><Input className="h-8 text-xs mt-1" type="number" value={vfUnitPrice || ""} onChange={e => setVfUnitPrice(parseFloat(e.target.value) || 0)} /></div>
                      <div><Label className="text-xs">Qty</Label><Input className="h-8 text-xs mt-1" type="number" value={vfQty || ""} onChange={e => setVfQty(parseFloat(e.target.value) || 1)} /></div>
                      <div><Label className="text-xs">Notes</Label><Input className="h-8 text-xs mt-1" value={vfNotes} onChange={e => setVfNotes(e.target.value)} /></div>
                    </div>
                    <div className="flex gap-2">
                      <Button size="sm" disabled={addVendorQuoteMutation.isPending || !vfVendor || !vfDesc}
                        onClick={() => addVendorQuoteMutation.mutate({
                          vendorName: vfVendor, description: vfDesc, size: vfSize || undefined,
                          category: vfCategory, unit: vfUnit, unitPrice: vfUnitPrice, quantity: vfQty,
                          quoteNumber: vfQuoteNum || undefined, notes: vfNotes || undefined,
                        })}>
                        {addVendorQuoteMutation.isPending ? "Saving..." : "Save Quote"}
                      </Button>
                      <Button size="sm" variant="outline" onClick={() => setShowVendorForm(false)}>Cancel</Button>
                    </div>
                  </CardContent>
                </Card>
              )}

              {/* Comparison section */}
              {vendorComparisons.length > 0 && (
                <Card className="border-card-border">
                  <CardHeader className="pb-2"><CardTitle className="text-sm">Vendor Price Comparison</CardTitle></CardHeader>
                  <CardContent>
                    <div className="space-y-3">
                      {vendorComparisons.slice(0, 10).map((comp: any, ci: number) => (
                        <div key={ci} className="border border-border/50 rounded-md p-2">
                          <div className="flex items-center justify-between mb-1.5">
                            <p className="text-xs font-medium truncate">{comp.description} {comp.size && <span className="text-muted-foreground">({comp.size})</span>}</p>
                            <Badge variant="outline" className="text-[9px] bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300">
                              Save {formatCurrency(comp.potentialSavings)}
                            </Badge>
                          </div>
                          {comp.quotes.map((q: any, qi: number) => (
                            <div key={qi} className="flex items-center justify-between text-[11px] py-0.5">
                              <span className="truncate max-w-[180px]">{q.vendorName}</span>
                              <div className="flex items-center gap-2">
                                <span className="tabular-nums font-medium">{formatCurrency(q.unitPrice)}</span>
                                {q.isBest ? (
                                  <Badge className="bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300 text-[9px] px-1 py-0">Best</Badge>
                                ) : (
                                  <span className="text-[9px] text-red-500">+{q.savings.toFixed(0)}%</span>
                                )}
                              </div>
                            </div>
                          ))}
                        </div>
                      ))}
                    </div>
                  </CardContent>
                </Card>
              )}

              {/* Quote list */}
              <Card className="border-card-border">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-border bg-muted/30">
                        <th className="text-left p-3 font-medium">Vendor</th>
                        <th className="text-left p-3 font-medium">Description</th>
                        <th className="text-left p-3 font-medium">Size</th>
                        <th className="text-left p-3 font-medium">Category</th>
                        <th className="text-right p-3 font-medium">Unit Price</th>
                        <th className="text-right p-3 font-medium">Qty</th>
                        <th className="text-right p-3 font-medium">Total</th>
                        <th className="text-left p-3 font-medium">Quote #</th>
                        <th className="w-8"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {vendorQuotes.length === 0 ? (
                        <tr><td colSpan={9} className="p-8 text-center text-muted-foreground">No vendor quotes yet. Add manually or import a CSV.</td></tr>
                      ) : (
                        vendorQuotes.slice(0, 200).map((q: any) => (
                          <tr key={q.id} className="border-b border-border/50 hover:bg-muted/20">
                            <td className="p-3 whitespace-nowrap font-medium">{q.vendorName}</td>
                            <td className="p-3 max-w-[250px]"><p className="truncate">{q.description}</p></td>
                            <td className="p-3">{q.size || "—"}</td>
                            <td className="p-3"><Badge variant="outline" className={`text-[10px] px-1.5 py-0 ${getCategoryColor(q.category || "other")}`}>{q.category || "other"}</Badge></td>
                            <td className="p-3 text-right tabular-nums font-medium">{formatCurrency(q.unitPrice)}</td>
                            <td className="p-3 text-right tabular-nums">{q.quantity}</td>
                            <td className="p-3 text-right tabular-nums">{formatCurrency(q.totalPrice || q.unitPrice * q.quantity)}</td>
                            <td className="p-3 text-muted-foreground">{q.quoteNumber || "—"}</td>
                            <td className="p-3 text-center">
                              <button className="text-muted-foreground hover:text-destructive p-0.5" onClick={() => deleteVendorQuoteMutation.mutate(q.id)}><Trash2 size={12} /></button>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
            </div>
          </TabsContent>
        </Tabs>
      </div>
    </AppLayout>
  );
}

// Sub-component: find items bought from multiple suppliers and compare prices
function PriceComparison({ records }: { records: PurchaseRecord[] }) {
  const groups: Record<string, { supplier: string; unitCost: number; date: string }[]> = {};
  for (const r of records) {
    const key = `${r.description.toLowerCase().trim().slice(0, 60)}|${(r.size || "").toLowerCase().trim()}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push({ supplier: r.supplier, unitCost: r.unitCost, date: r.invoiceDate || "" });
  }

  const comparisons = Object.entries(groups)
    .filter(([, items]) => new Set(items.map(i => i.supplier)).size > 1)
    .slice(0, 10);

  if (comparisons.length === 0) {
    return <p className="text-xs text-muted-foreground text-center py-4">Need items purchased from multiple suppliers to compare.</p>;
  }

  return (
    <div className="space-y-3">
      {comparisons.map(([key, items]) => {
        const [desc] = key.split("|");
        const sorted = [...items].sort((a, b) => a.unitCost - b.unitCost);
        const lowest = sorted[0].unitCost;
        return (
          <div key={key} className="border border-border/50 rounded-md p-2">
            <p className="text-xs font-medium truncate mb-1.5">{desc}</p>
            <div className="space-y-1">
              {sorted.map((item, i) => {
                const pctMore = lowest > 0 ? ((item.unitCost - lowest) / lowest * 100) : 0;
                return (
                  <div key={i} className="flex items-center justify-between text-[11px]">
                    <span className="truncate max-w-[140px]">{item.supplier}</span>
                    <div className="flex items-center gap-2">
                      <span className="tabular-nums font-medium">{formatCurrency(item.unitCost)}</span>
                      {i === 0 ? (
                        <Badge className="bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300 text-[9px] px-1 py-0">Best</Badge>
                      ) : (
                        <span className="text-[9px] text-red-500">+{pctMore.toFixed(0)}%</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}
