import { Link, useLocation } from "wouter";
import { useState, useEffect, useMemo } from "react";
import {
  LayoutDashboard, Wrench, Building2, HardHat, Calculator, Moon, Sun, Menu, FileText, Settings, DollarSign, Clock, Target, Sparkles, BookOpen, Search, Sliders, TrendingUp
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { useQuery } from "@tanstack/react-query";
import PerplexityAttribution from "./PerplexityAttribution";
import ChatAssistant from "./ChatAssistant";
import CommandPalette from "./CommandPalette";
import logoPic from "@assets/logo-pic.jpg";
import type { TakeoffProject } from "@shared/schema";

// Sidebar nav items grouped for easier scanning. The previous layout had 9 items
// at uniform visual weight which made it hard to find anything; now they're
// clustered by workflow.
const NAV_GROUPS: { label: string; items: { path: string; label: string; icon: any }[] }[] = [
  {
    label: "Project",
    items: [
      { path: "/", label: "Dashboard", icon: LayoutDashboard },
      { path: "/project-history", label: "Project History", icon: Clock },
    ],
  },
  {
    label: "Takeoff",
    items: [
      { path: "/mechanical", label: "Mechanical", icon: Wrench },
      { path: "/structural", label: "Structural", icon: Building2 },
      { path: "/civil", label: "Civil", icon: HardHat },
    ],
  },
  {
    label: "Estimating",
    items: [
      { path: "/estimating", label: "Estimating", icon: Calculator },
      { path: "/methods", label: "Methods", icon: Sliders },
      { path: "/cost-database", label: "Cost Database", icon: DollarSign },
      { path: "/bids", label: "Bid Tracker", icon: Target },
      { path: "/performance", label: "Performance", icon: TrendingUp },
    ],
  },
  {
    label: "Account",
    items: [
      { path: "/help", label: "Help & How-To", icon: BookOpen },
      { path: "/settings", label: "Settings", icon: Settings },
    ],
  },
];

interface AppLayoutProps {
  children: React.ReactNode;
  subtitle?: string;
}

export default function AppLayout({ children, subtitle }: AppLayoutProps) {
  const [location, setLocation] = useLocation();
  const [dark, setDark] = useState(() => {
    const saved = localStorage.getItem("pg-dark-mode");
    if (saved !== null) return saved === "true";
    return window.matchMedia("(prefers-color-scheme: dark)").matches;
  });
  const [mobileOpen, setMobileOpen] = useState(false);
  const [paletteOpen, setPaletteOpen] = useState(false);

  const toggleDark = () => {
    const next = !dark;
    setDark(next);
    document.documentElement.classList.toggle("dark", next);
    localStorage.setItem("pg-dark-mode", String(next));
  };

  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
  }, [dark]);

  // Keyboard shortcut: ⌘K / Ctrl+K opens the command palette
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setPaletteOpen(true);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Fetch recent takeoff projects for sidebar — only poll on relevant pages
  const shouldPoll = useMemo(
    () => ["/", "/mechanical", "/structural", "/civil"].some(p => p === "/" ? location === "/" : location.startsWith(p)),
    [location]
  );
  const { data: recentProjects = [] } = useQuery<TakeoffProject[]>({
    queryKey: ["/api/takeoff/projects"],
    refetchInterval: shouldPoll ? 10000 : false,
  });

  const sidebarProjects = useMemo(() => recentProjects.filter(p => !p.archived).slice(0, 5), [recentProjects]);

  const disciplinePath: Record<string, string> = {
    mechanical: "/mechanical",
    structural: "/structural",
    civil: "/civil",
  };

  return (
    <div className="flex h-screen bg-background text-foreground overflow-hidden">
      {/* Sidebar */}
      <aside
        className={`fixed inset-y-0 left-0 z-50 w-60 flex flex-col bg-sidebar border-r border-sidebar-border transition-transform duration-200
          ${mobileOpen ? "translate-x-0" : "-translate-x-full"} md:relative md:translate-x-0`}
      >
        {/* Logo / Branding */}
        <div className="flex items-center gap-3 px-4 py-4 border-b border-sidebar-border bg-gradient-to-br from-sidebar to-transparent">
          <img src={logoPic} alt="Picou Group" className="h-9 w-9 rounded-lg object-contain shadow-sm ring-1 ring-black/5 dark:ring-white/10" />
          <div className="min-w-0 flex-1">
            <p className="text-sm font-bold leading-tight text-sidebar-foreground truncate">Picou Group</p>
            <p className="text-[10px] text-muted-foreground truncate uppercase tracking-wider">Takeoff & Estimating</p>
          </div>
        </div>

        {/* Quick search trigger */}
        <button
          onClick={() => setPaletteOpen(true)}
          className="mx-3 mt-3 mb-1 flex items-center gap-2 px-3 py-2 rounded-md border border-sidebar-border bg-background/60 hover:bg-sidebar-accent transition-colors text-xs text-muted-foreground"
          title="Search anywhere (⌘K)"
        >
          <Search size={12} />
          <span className="flex-1 text-left">Search...</span>
          <kbd className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-muted text-muted-foreground border border-border">⌘K</kbd>
        </button>

        {/* Nav Groups */}
        <nav className="flex-1 overflow-y-auto py-2 px-2 space-y-3">
          {NAV_GROUPS.map(group => (
            <div key={group.label}>
              <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider px-3 mb-1">{group.label}</p>
              <div className="space-y-0.5">
                {group.items.map(({ path, label, icon: Icon }) => {
                  const active = path === "/" ? location === "/" : location.startsWith(path);
                  return (
                    <Link key={path} href={path}>
                      <a
                        onClick={() => setMobileOpen(false)}
                        data-testid={`nav-${label.toLowerCase().replace(/\s+/g, "-")}`}
                        className={`group relative flex items-center gap-3 px-3 py-2 rounded-md text-sm font-medium transition-all duration-150 cursor-pointer
                          ${active
                            ? "bg-sidebar-primary text-sidebar-primary-foreground shadow-sm"
                            : "text-sidebar-foreground hover:bg-sidebar-accent hover:text-sidebar-accent-foreground"
                          }`}
                      >
                        {active && (
                          <span className="absolute left-0 top-1/2 -translate-y-1/2 w-[3px] h-5 rounded-r-full bg-[#01696F] dark:bg-teal-400" />
                        )}
                        <Icon size={15} className={`shrink-0 transition-colors ${active ? "text-sidebar-primary-foreground" : "text-muted-foreground group-hover:text-sidebar-accent-foreground"}`} />
                        {label}
                      </a>
                    </Link>
                  );
                })}
              </div>
            </div>
          ))}

          {/* Recent Projects section */}
          {sidebarProjects.length > 0 && (
            <div>
              <p className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wider px-3 mb-1">Recent</p>
              <div className="space-y-0.5">
                {sidebarProjects.map(p => (
                  <button
                    key={p.id}
                    onClick={() => {
                      setMobileOpen(false);
                      const basePath = disciplinePath[p.discipline] || "/mechanical";
                      setLocation(`${basePath}?project=${p.id}`);
                    }}
                    className="group w-full flex items-center gap-2 px-3 py-1.5 rounded-md text-xs text-sidebar-foreground hover:bg-sidebar-accent hover:text-sidebar-accent-foreground cursor-pointer transition-all duration-150"
                    title={p.name}
                  >
                    <FileText size={11} className="shrink-0 text-muted-foreground group-hover:text-sidebar-accent-foreground transition-colors" />
                    <span className="truncate flex-1 text-left">{p.name}</span>
                    <span className="text-[9px] text-muted-foreground shrink-0 font-mono">{p.items.length}</span>
                  </button>
                ))}
              </div>
            </div>
          )}
        </nav>

        {/* Footer */}
        <div className="px-4 pb-3 pt-2 border-t border-sidebar-border space-y-2">
          <div className="flex items-center gap-1.5 px-1">
            <Sparkles size={10} className="text-teal-500 dark:text-teal-400" />
            <p className="text-[10px] text-muted-foreground">Powered by AI</p>
          </div>
          <PerplexityAttribution />
        </div>
      </aside>

      {/* Mobile overlay */}
      {mobileOpen && (
        <div
          className="fixed inset-0 z-40 bg-black/40 md:hidden"
          onClick={() => setMobileOpen(false)}
        />
      )}

      {/* Main content */}
      <div className="flex flex-col flex-1 min-w-0 overflow-hidden">
        {/* Top header */}
        <header className="flex items-center justify-between px-4 py-3 border-b border-border bg-card shrink-0">
          <div className="flex items-center gap-3">
            <Button
              variant="ghost"
              size="icon"
              className="md:hidden"
              onClick={() => setMobileOpen(true)}
              data-testid="btn-mobile-menu"
            >
              <Menu size={18} />
            </Button>
            <div>
              <h1 className="text-sm font-semibold text-foreground leading-tight">Picou Group Contractors</h1>
              {subtitle && <p className="text-xs text-muted-foreground">{subtitle}</p>}
            </div>
          </div>
          <div className="flex items-center gap-1">
            <Button
              variant="ghost"
              size="icon"
              onClick={() => setPaletteOpen(true)}
              className="shrink-0"
              title="Search (⌘K)"
            >
              <Search size={16} />
            </Button>
            <Button
              variant="ghost"
              size="icon"
              onClick={toggleDark}
              data-testid="btn-dark-mode"
              className="shrink-0"
              title={dark ? "Switch to light mode" : "Switch to dark mode"}
            >
              {dark ? <Sun size={16} /> : <Moon size={16} />}
            </Button>
          </div>
        </header>

        {/* Page content */}
        <main className="flex-1 overflow-auto">
          {children}
        </main>
      </div>

      {/* Chat Assistant — floating on all pages */}
      <ChatAssistant
        pageContext={
          subtitle ||
          NAV_GROUPS.flatMap(g => g.items).find(n => n.path === "/" ? location === "/" : location.startsWith(n.path))?.label ||
          "Unknown"
        }
      />

      {/* Command palette (⌘K) */}
      <CommandPalette
        open={paletteOpen}
        onOpenChange={setPaletteOpen}
        recentProjects={recentProjects.filter(p => !p.archived).slice(0, 8)}
      />
    </div>
  );
}
