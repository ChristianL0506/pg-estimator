import { useState, useMemo, useEffect } from "react";
import { useMutation } from "@tanstack/react-query";
import { Copy, Mail, ExternalLink, Check, Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { useToast } from "@/hooks/use-toast";
import { apiRequest } from "@/lib/queryClient";
import type { EstimateProject, EstimateItem } from "@shared/schema";

interface RfqModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  project: EstimateProject;
}

export default function RfqModal({ open, onOpenChange, project }: RfqModalProps) {
  const { toast } = useToast();
  const [selectedIds, setSelectedIds] = useState<Set<string>>(() => new Set(project.items.map(i => i.id)));
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    setSelectedIds(new Set(project.items.map(i => i.id)));
  }, [project.id]);

  const allSelected = selectedIds.size === project.items.length;
  const noneSelected = selectedIds.size === 0;

  const toggleItem = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleAll = () => {
    if (allSelected) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(project.items.map(i => i.id)));
    }
  };

  const rfqMutation = useMutation({
    mutationFn: async () => {
      const res = await apiRequest("POST", `/api/estimates/${project.id}/generate-rfq`, {
        selectedItemIds: Array.from(selectedIds),
      });
      return res.json();
    },
  });

  const rfqData = rfqMutation.data;

  const handleGenerate = () => {
    if (noneSelected) {
      toast({ title: "Select at least one item", variant: "destructive" });
      return;
    }
    rfqMutation.mutate();
  };

  const handleCopy = async () => {
    if (!rfqData?.emailText) return;
    try {
      await navigator.clipboard.writeText(rfqData.emailText);
      setCopied(true);
      toast({ title: "Copied to clipboard" });
      setTimeout(() => setCopied(false), 2000);
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  const handleMailto = () => {
    if (!rfqData?.emailText) return;
    const subject = encodeURIComponent(`Request for Quotation - ${project.name}`);
    const body = encodeURIComponent(rfqData.emailText);
    window.open(`mailto:?subject=${subject}&body=${body}`, "_blank");
  };

  const handleDownloadTxt = () => {
    if (!rfqData?.emailText) return;
    const blob = new Blob([rfqData.emailText], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `RFQ-${project.name.replace(/[^a-zA-Z0-9]/g, "_")}.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-4xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-sm">Request Material Quotes</DialogTitle>
        </DialogHeader>

        <div className="space-y-4">
          {/* Material selection */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-semibold text-muted-foreground uppercase">Select Materials ({selectedIds.size}/{project.items.length})</span>
              <button className="text-xs text-primary hover:underline" onClick={toggleAll}>
                {allSelected ? "Deselect All" : "Select All"}
              </button>
            </div>
            <div className="max-h-48 overflow-y-auto rounded-md border border-border">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-muted/50 border-b border-border sticky top-0">
                    <th className="px-2 py-1.5 w-8"></th>
                    <th className="px-2 py-1.5 text-left">Description</th>
                    <th className="px-2 py-1.5 text-left">Size</th>
                    <th className="px-2 py-1.5 text-right">Qty</th>
                    <th className="px-2 py-1.5 text-left">Unit</th>
                    <th className="px-2 py-1.5 text-left">Category</th>
                  </tr>
                </thead>
                <tbody>
                  {project.items.map(item => (
                    <tr
                      key={item.id}
                      className={`border-b border-border cursor-pointer hover:bg-muted/20 ${selectedIds.has(item.id) ? "bg-primary/5" : ""}`}
                      onClick={() => toggleItem(item.id)}
                    >
                      <td className="px-2 py-1">
                        <input
                          type="checkbox"
                          checked={selectedIds.has(item.id)}
                          readOnly
                          className="rounded pointer-events-none"
                        />
                      </td>
                      <td className="px-2 py-1 max-w-[200px] truncate">{item.description}</td>
                      <td className="px-2 py-1 font-mono">{item.size}</td>
                      <td className="px-2 py-1 text-right font-mono">{item.quantity}</td>
                      <td className="px-2 py-1 text-muted-foreground">{item.unit}</td>
                      <td className="px-2 py-1">
                        <Badge variant="outline" className="text-[9px] px-1 py-0">{item.category}</Badge>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Generate button */}
          <Button size="sm" onClick={handleGenerate} disabled={rfqMutation.isPending || noneSelected}>
            {rfqMutation.isPending ? "Generating..." : "Generate RFQ Email"}
          </Button>

          {/* Generated email */}
          {rfqData?.emailText && (
            <>
              <Separator />
              <div>
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-semibold text-muted-foreground uppercase">Generated RFQ Email</span>
                  <div className="flex gap-2">
                    <Button size="sm" className="h-7 text-xs" onClick={handleCopy}>
                      {copied ? <Check size={12} className="mr-1" /> : <Copy size={12} className="mr-1" />}
                      {copied ? "Copied" : "Copy to Clipboard"}
                    </Button>
                    <Button variant="outline" size="sm" className="h-7 text-xs" onClick={handleDownloadTxt}>
                      <Download size={12} className="mr-1" />
                      Download .txt
                    </Button>
                    <Button variant="ghost" size="sm" className="h-7 text-xs text-muted-foreground" onClick={handleMailto} title="Best for small quotes — may truncate long material lists">
                      <Mail size={12} className="mr-1" />
                      Email
                    </Button>
                  </div>
                </div>
                <pre className="text-xs bg-muted/30 border border-border rounded-md p-3 max-h-64 overflow-y-auto whitespace-pre-wrap font-mono">
                  {rfqData.emailText}
                </pre>
              </div>

              {/* Supplier suggestions */}
              {rfqData.supplierSuggestions && rfqData.supplierSuggestions.length > 0 && (
                <>
                  <Separator />
                  <div>
                    <span className="text-xs font-semibold text-muted-foreground uppercase block mb-2">Suggested Suppliers</span>
                    <div className="space-y-2">
                      {rfqData.supplierSuggestions.filter((s: any) => s.type === "database").length > 0 && (
                        <div>
                          <p className="text-[10px] font-semibold text-muted-foreground mb-1">From Cost Database</p>
                          {rfqData.supplierSuggestions.filter((s: any) => s.type === "database").map((s: any, idx: number) => (
                            <p key={idx} className="text-xs text-muted-foreground ml-2">{s.suggestion}</p>
                          ))}
                        </div>
                      )}
                      {rfqData.supplierSuggestions.filter((s: any) => s.type === "location").length > 0 && (
                        <div>
                          <p className="text-[10px] font-semibold text-muted-foreground mb-1">Location-Based</p>
                          {rfqData.supplierSuggestions.filter((s: any) => s.type === "location").map((s: any, idx: number) => (
                            <div key={idx} className="flex items-center gap-2 ml-2">
                              <p className="text-xs">{s.suggestion}</p>
                              {s.link && (
                                <a href={s.link} target="_blank" rel="noopener noreferrer" className="text-primary hover:underline inline-flex items-center gap-0.5 text-xs">
                                  <ExternalLink size={10} /> Map
                                </a>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                      {rfqData.supplierSuggestions.filter((s: any) => s.type === "category").length > 0 && (
                        <div>
                          <p className="text-[10px] font-semibold text-muted-foreground mb-1">By Material Category</p>
                          <ul className="space-y-0.5 ml-2">
                            {rfqData.supplierSuggestions.filter((s: any) => s.type === "category").map((s: any, idx: number) => (
                              <li key={idx} className="text-xs text-muted-foreground">• {s.suggestion}</li>
                            ))}
                          </ul>
                        </div>
                      )}
                    </div>
                  </div>
                </>
              )}
            </>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
