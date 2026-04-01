import { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { Target, Plus, Trash2, Edit2, Check, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import AppLayout from "@/components/AppLayout";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";

const STATUS_COLORS: Record<string, string> = {
  draft: "bg-gray-100 text-gray-800 dark:bg-gray-800 dark:text-gray-300",
  submitted: "bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300",
  won: "bg-green-100 text-green-800 dark:bg-green-900/40 dark:text-green-300",
  lost: "bg-red-100 text-red-800 dark:bg-red-900/40 dark:text-red-300",
  no_bid: "bg-slate-100 text-slate-800 dark:bg-slate-800 dark:text-slate-300",
};

function fmtDollar(n: number) {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

export default function BidDashboardPage() {
  const { toast } = useToast();
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [form, setForm] = useState({
    projectName: "", client: "", bidDate: "", dueDate: "", bidAmount: "",
    status: "draft", awardAmount: "", competitor: "", notes: "",
  });

  const { data: bids = [] } = useQuery<any[]>({
    queryKey: ["/api/bids"],
    queryFn: async () => (await apiRequest("GET", "/api/bids")).json(),
  });

  const { data: stats = [] } = useQuery<any[]>({
    queryKey: ["/api/bids/stats"],
    queryFn: async () => (await apiRequest("GET", "/api/bids/stats")).json(),
  });

  const createMutation = useMutation({
    mutationFn: (data: any) => apiRequest("POST", "/api/bids", data).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/bids"] });
      setShowForm(false);
      resetForm();
      toast({ title: "Bid created" });
    },
  });

  const updateMutation = useMutation({
    mutationFn: ({ id, data }: { id: string; data: any }) =>
      apiRequest("PATCH", `/api/bids/${id}`, data).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/bids"] });
      setEditId(null);
      resetForm();
      toast({ title: "Bid updated" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/bids/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/bids"] });
      toast({ title: "Bid deleted" });
    },
  });

  function resetForm() {
    setForm({ projectName: "", client: "", bidDate: "", dueDate: "", bidAmount: "", status: "draft", awardAmount: "", competitor: "", notes: "" });
  }

  function startEdit(bid: any) {
    setEditId(bid.id);
    setForm({
      projectName: bid.projectName || "", client: bid.client || "",
      bidDate: bid.bidDate || "", dueDate: bid.dueDate || "",
      bidAmount: String(bid.bidAmount || ""), status: bid.status || "draft",
      awardAmount: String(bid.awardAmount || ""), competitor: bid.competitor || "", notes: bid.notes || "",
    });
  }

  function handleSubmit() {
    const data = {
      ...form,
      bidAmount: parseFloat(form.bidAmount) || 0,
      awardAmount: form.awardAmount ? parseFloat(form.awardAmount) : undefined,
    };
    if (editId) {
      updateMutation.mutate({ id: editId, data });
    } else {
      createMutation.mutate(data);
    }
  }

  // Compute stats
  const totalBids = bids.length;
  const wonBids = bids.filter((b: any) => b.status === "won");
  const lostBids = bids.filter((b: any) => b.status === "lost");
  const winRate = (wonBids.length + lostBids.length) > 0
    ? Math.round((wonBids.length / (wonBids.length + lostBids.length)) * 100)
    : 0;
  const totalBidValue = bids.reduce((sum: number, b: any) => sum + (b.bidAmount || 0), 0);
  const avgBid = totalBids > 0 ? totalBidValue / totalBids : 0;

  return (
    <AppLayout subtitle="Bid Tracker">
      <div className="p-5 space-y-5">
        {/* Stats cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <Card>
            <CardContent className="pt-4 pb-3 px-4">
              <p className="text-xs text-muted-foreground">Total Bids</p>
              <p className="text-2xl font-bold">{totalBids}</p>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-4 pb-3 px-4">
              <p className="text-xs text-muted-foreground">Win Rate</p>
              <p className="text-2xl font-bold text-green-600">{winRate}%</p>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-4 pb-3 px-4">
              <p className="text-xs text-muted-foreground">Total Bid Value</p>
              <p className="text-2xl font-bold">{fmtDollar(totalBidValue)}</p>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-4 pb-3 px-4">
              <p className="text-xs text-muted-foreground">Avg Bid Size</p>
              <p className="text-2xl font-bold">{fmtDollar(avgBid)}</p>
            </CardContent>
          </Card>
        </div>

        {/* Add button */}
        <div className="flex justify-between items-center">
          <h2 className="text-sm font-semibold">All Bids</h2>
          <Button size="sm" onClick={() => { setShowForm(!showForm); setEditId(null); resetForm(); }}>
            <Plus size={14} className="mr-1.5" />
            Add Bid
          </Button>
        </div>

        {/* Add/Edit Form */}
        {(showForm || editId) && (
          <Card>
            <CardContent className="pt-4 space-y-3">
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <div>
                  <Label className="text-xs">Project Name *</Label>
                  <Input value={form.projectName} onChange={e => setForm({ ...form, projectName: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Client</Label>
                  <Input value={form.client} onChange={e => setForm({ ...form, client: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Bid Date</Label>
                  <Input type="date" value={form.bidDate} onChange={e => setForm({ ...form, bidDate: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Due Date</Label>
                  <Input type="date" value={form.dueDate} onChange={e => setForm({ ...form, dueDate: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Bid Amount ($)</Label>
                  <Input type="number" value={form.bidAmount} onChange={e => setForm({ ...form, bidAmount: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Status</Label>
                  <Select value={form.status} onValueChange={v => setForm({ ...form, status: v })}>
                    <SelectTrigger className="h-8 text-xs"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="draft">Draft</SelectItem>
                      <SelectItem value="submitted">Submitted</SelectItem>
                      <SelectItem value="won">Won</SelectItem>
                      <SelectItem value="lost">Lost</SelectItem>
                      <SelectItem value="no_bid">No Bid</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label className="text-xs">Award Amount ($)</Label>
                  <Input type="number" value={form.awardAmount} onChange={e => setForm({ ...form, awardAmount: e.target.value })} className="h-8 text-xs" />
                </div>
                <div>
                  <Label className="text-xs">Competitor</Label>
                  <Input value={form.competitor} onChange={e => setForm({ ...form, competitor: e.target.value })} className="h-8 text-xs" />
                </div>
              </div>
              <div>
                <Label className="text-xs">Notes</Label>
                <Input value={form.notes} onChange={e => setForm({ ...form, notes: e.target.value })} className="h-8 text-xs" />
              </div>
              <div className="flex gap-2">
                <Button size="sm" onClick={handleSubmit} disabled={!form.projectName || createMutation.isPending || updateMutation.isPending}>
                  <Check size={14} className="mr-1" />
                  {editId ? "Update" : "Create"}
                </Button>
                <Button size="sm" variant="outline" onClick={() => { setShowForm(false); setEditId(null); resetForm(); }}>
                  <X size={14} className="mr-1" />
                  Cancel
                </Button>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Bids table */}
        <Card>
          <CardContent className="p-0">
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="border-b bg-muted/40">
                    <th className="text-left px-3 py-2 font-medium">Project</th>
                    <th className="text-left px-3 py-2 font-medium">Client</th>
                    <th className="text-left px-3 py-2 font-medium">Bid Date</th>
                    <th className="text-left px-3 py-2 font-medium">Due</th>
                    <th className="text-right px-3 py-2 font-medium">Bid Amount</th>
                    <th className="text-center px-3 py-2 font-medium">Status</th>
                    <th className="text-right px-3 py-2 font-medium">Award</th>
                    <th className="text-left px-3 py-2 font-medium">Competitor</th>
                    <th className="text-center px-3 py-2 font-medium">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {bids.length === 0 ? (
                    <tr><td colSpan={9} className="text-center py-10 text-muted-foreground">
                      <p className="text-sm font-medium text-foreground">No bids yet</p>
                      <p className="text-xs text-muted-foreground mt-1">Add your first bid using the form above.</p>
                    </td></tr>
                  ) : bids.map((bid: any) => (
                    <tr key={bid.id} className="border-b hover:bg-muted/20">
                      <td className="px-3 py-2 font-medium">{bid.projectName}</td>
                      <td className="px-3 py-2">{bid.client || "—"}</td>
                      <td className="px-3 py-2">{bid.bidDate || "—"}</td>
                      <td className="px-3 py-2">{bid.dueDate || "—"}</td>
                      <td className="px-3 py-2 text-right font-mono">{bid.bidAmount ? fmtDollar(bid.bidAmount) : "—"}</td>
                      <td className="px-3 py-2 text-center">
                        <Badge variant="outline" className={`text-[10px] ${STATUS_COLORS[bid.status] || ""}`}>
                          {bid.status?.replace("_", " ").toUpperCase()}
                        </Badge>
                      </td>
                      <td className="px-3 py-2 text-right font-mono">{bid.awardAmount ? fmtDollar(bid.awardAmount) : "—"}</td>
                      <td className="px-3 py-2">{bid.competitor || "—"}</td>
                      <td className="px-3 py-2 text-center">
                        <div className="flex justify-center gap-1">
                          <button className="p-1 hover:text-primary" onClick={() => startEdit(bid)} title="Edit">
                            <Edit2 size={13} />
                          </button>
                          <button className="p-1 hover:text-destructive" onClick={() => deleteMutation.mutate(bid.id)} title="Delete">
                            <Trash2 size={13} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
      </div>
    </AppLayout>
  );
}
