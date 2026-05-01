// Command Palette — keyboard-first navigation.
// Triggered by ⌘K / Ctrl+K. Lets the user jump to any page or recent project
// without clicking through the sidebar. Includes:
//   - Page navigation (Dashboard, Mechanical, Estimating, etc.)
//   - Recent projects (jump to a takeoff)
//   - Quick actions (toggle dark mode, open Help)

import { useEffect, useState, useMemo } from "react";
import { useLocation } from "wouter";
import { Dialog, DialogContent } from "@/components/ui/dialog";
import { Search, ArrowRight, Wrench, Building2, HardHat, LayoutDashboard, Calculator, DollarSign, Clock, Target, Settings, BookOpen, FileText, Moon, Sun, Sparkles } from "lucide-react";
import type { TakeoffProject } from "@shared/schema";

type Action = {
  id: string;
  label: string;
  hint?: string;
  icon: any;
  group: string;
  run: () => void;
};

interface Props {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  recentProjects: TakeoffProject[];
}

export default function CommandPalette({ open, onOpenChange, recentProjects }: Props) {
  const [, setLocation] = useLocation();
  const [query, setQuery] = useState("");
  const [activeIdx, setActiveIdx] = useState(0);

  // Reset state when opened/closed
  useEffect(() => {
    if (open) {
      setQuery("");
      setActiveIdx(0);
    }
  }, [open]);

  const disciplinePath: Record<string, string> = {
    mechanical: "/mechanical",
    structural: "/structural",
    civil: "/civil",
  };

  const allActions = useMemo<Action[]>(() => {
    const close = () => onOpenChange(false);
    const navTo = (path: string) => () => { setLocation(path); close(); };
    const items: Action[] = [
      // Pages
      { id: "page-dashboard", label: "Dashboard", group: "Pages", icon: LayoutDashboard, run: navTo("/") },
      { id: "page-mechanical", label: "Mechanical Takeoff", group: "Pages", icon: Wrench, run: navTo("/mechanical") },
      { id: "page-structural", label: "Structural Takeoff", group: "Pages", icon: Building2, run: navTo("/structural") },
      { id: "page-civil", label: "Civil Takeoff", group: "Pages", icon: HardHat, run: navTo("/civil") },
      { id: "page-estimating", label: "Estimating", group: "Pages", icon: Calculator, run: navTo("/estimating") },
      { id: "page-cost-database", label: "Cost Database", group: "Pages", icon: DollarSign, run: navTo("/cost-database") },
      { id: "page-history", label: "Project History", group: "Pages", icon: Clock, run: navTo("/project-history") },
      { id: "page-bids", label: "Bid Tracker", group: "Pages", icon: Target, run: navTo("/bids") },
      { id: "page-help", label: "Help & How-To", group: "Pages", icon: BookOpen, run: navTo("/help") },
      { id: "page-settings", label: "Settings", group: "Pages", icon: Settings, run: navTo("/settings") },
      // Quick actions
      {
        id: "action-toggle-dark",
        label: "Toggle dark mode",
        group: "Actions",
        icon: document.documentElement.classList.contains("dark") ? Sun : Moon,
        run: () => {
          const next = !document.documentElement.classList.contains("dark");
          document.documentElement.classList.toggle("dark", next);
          localStorage.setItem("pg-dark-mode", String(next));
          close();
        },
      },
    ];

    // Recent projects
    for (const p of recentProjects) {
      items.push({
        id: `proj-${p.id}`,
        label: p.name,
        hint: `${p.discipline} · ${p.items.length} items`,
        group: "Recent Projects",
        icon: FileText,
        run: () => {
          const basePath = disciplinePath[p.discipline] || "/mechanical";
          setLocation(`${basePath}?project=${p.id}`);
          close();
        },
      });
    }

    return items;
  }, [recentProjects, setLocation, onOpenChange]);

  // Filter
  const filtered = useMemo(() => {
    if (!query.trim()) return allActions;
    const q = query.toLowerCase();
    return allActions.filter(a => a.label.toLowerCase().includes(q) || (a.hint || "").toLowerCase().includes(q));
  }, [allActions, query]);

  // Group for display
  const grouped = useMemo(() => {
    const map: Record<string, Action[]> = {};
    for (const a of filtered) {
      if (!map[a.group]) map[a.group] = [];
      map[a.group].push(a);
    }
    return map;
  }, [filtered]);

  // Keep activeIdx within bounds when filtered changes
  useEffect(() => {
    if (activeIdx >= filtered.length) setActiveIdx(Math.max(0, filtered.length - 1));
  }, [filtered, activeIdx]);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActiveIdx(i => Math.min(filtered.length - 1, i + 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveIdx(i => Math.max(0, i - 1));
    } else if (e.key === "Enter") {
      e.preventDefault();
      const action = filtered[activeIdx];
      if (action) action.run();
    } else if (e.key === "Escape") {
      onOpenChange(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-xl p-0 gap-0 overflow-hidden [&>button.absolute]:hidden">
        <div className="flex items-center gap-2 px-4 py-3 border-b border-border">
          <Search size={16} className="text-muted-foreground shrink-0" />
          <input
            autoFocus
            value={query}
            onChange={e => { setQuery(e.target.value); setActiveIdx(0); }}
            onKeyDown={onKeyDown}
            placeholder="Search pages, projects, actions..."
            className="flex-1 bg-transparent outline-none text-sm placeholder:text-muted-foreground"
          />
          <kbd className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-muted text-muted-foreground border border-border">ESC</kbd>
        </div>
        <div className="max-h-[50vh] overflow-y-auto py-1">
          {filtered.length === 0 ? (
            <div className="px-4 py-8 text-center text-sm text-muted-foreground">
              <Sparkles size={20} className="mx-auto mb-2 opacity-50" />
              No matches for "{query}"
            </div>
          ) : (
            Object.entries(grouped).map(([group, items]) => (
              <div key={group} className="py-1">
                <p className="text-[10px] uppercase tracking-wider text-muted-foreground font-semibold px-4 py-1.5">{group}</p>
                {items.map(action => {
                  const Icon = action.icon;
                  const idx = filtered.indexOf(action);
                  const active = idx === activeIdx;
                  return (
                    <button
                      key={action.id}
                      onClick={action.run}
                      onMouseEnter={() => setActiveIdx(idx)}
                      className={`w-full flex items-center gap-3 px-4 py-2 text-left text-sm transition-colors ${active ? "bg-accent text-accent-foreground" : "hover:bg-accent/50"}`}
                    >
                      <Icon size={14} className="text-muted-foreground shrink-0" />
                      <span className="flex-1 truncate">{action.label}</span>
                      {action.hint && <span className="text-[10px] text-muted-foreground truncate max-w-[180px]">{action.hint}</span>}
                      {active && <ArrowRight size={12} className="text-muted-foreground" />}
                    </button>
                  );
                })}
              </div>
            ))
          )}
        </div>
        <div className="px-4 py-2 border-t border-border bg-muted/30 flex items-center gap-3 text-[10px] text-muted-foreground">
          <span className="flex items-center gap-1"><kbd className="font-mono px-1 rounded bg-background border border-border">↑</kbd> <kbd className="font-mono px-1 rounded bg-background border border-border">↓</kbd> navigate</span>
          <span className="flex items-center gap-1"><kbd className="font-mono px-1 rounded bg-background border border-border">↵</kbd> select</span>
          <span className="flex items-center gap-1"><kbd className="font-mono px-1 rounded bg-background border border-border">esc</kbd> close</span>
        </div>
      </DialogContent>
    </Dialog>
  );
}
