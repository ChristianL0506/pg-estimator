import { useState, useRef, useEffect, useCallback } from "react";
import { MessageCircle, X, Send, Sparkles, HelpCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";

// ============================================================
// Knowledge Base (Help Mode — no API calls)
// ============================================================

const KNOWLEDGE_BASE: { keywords: string[]; answer: string }[] = [
  {
    keywords: ["mechanical", "takeoff", "piping", "run mechanical", "start mechanical"],
    answer: "To run a mechanical takeoff: Go to the Mechanical page, click Upload Drawing, select your PDF (piping isometric drawings), choose revision detection if needed, and click Start Extraction. The AI reads BOM tables from each ISO and extracts pipe, fittings, valves, flanges, bolts, and gaskets with sizes, quantities, and specs.",
  },
  {
    keywords: ["structural", "takeoff", "steel", "run structural"],
    answer: "For structural takeoff: Go to the Structural page and upload structural drawing PDFs. The AI extracts wide flange beams, HSS tubes, angles, channels, plates, base plates, footings, grade beams, rebar, and more. Items are categorized into steel, concrete, rebar, and earthwork.",
  },
  {
    keywords: ["civil", "takeoff", "site", "run civil"],
    answer: "For civil takeoff: Go to the Civil page and upload civil/site drawing PDFs. The AI extracts storm pipe, sewer pipe, water pipe, manholes, catch basins, fire hydrants, earthwork, paving, curb & gutter, and other site items.",
  },
  {
    keywords: ["upload", "drawing", "pdf", "file", "how to upload"],
    answer: "Click the upload area on any discipline page (Mechanical, Structural, or Civil). Only PDF files are accepted. After selecting your PDF, you'll be asked if the drawing has revisions (clouded changes). You can also enable dual-model extraction (Claude + Gemini) for higher accuracy. Then click Start Extraction.",
  },
  {
    keywords: ["revision", "cloud", "clouded", "changes", "revision detection"],
    answer: "When uploading, toggle 'Has Revisions' if your drawing has revision clouds marking changes. The AI will detect clouded items and flag them with a revision indicator in the BOM table, so you can quickly identify what changed between revisions.",
  },
  {
    keywords: ["bom", "table", "bill of materials", "items", "extracted items"],
    answer: "The BOM table shows all extracted items with columns for Line #, Category, Description, Size, Qty, Unit, Spec, Material, Schedule, Rating, and Notes. You can filter by category, size, sheet number, confidence level, and search text. Items from SHOP and FIELD sections are tracked separately.",
  },
  {
    keywords: ["connections", "weld", "bolt", "connection tab", "weld tab"],
    answer: "The Connections tab shows a summary of all welds and bolt-ups grouped by size. It includes butt welds, socket welds, and threaded connections with their quantities. This helps you quickly see the total weld count for manhour estimation.",
  },
  {
    keywords: ["pivot", "summary", "pivot summary", "category summary"],
    answer: "The Pivot Summary provides a bird's-eye view of your takeoff grouped by category. It shows total quantities, sizes, and line items for each category (pipe, elbows, tees, valves, flanges, etc.) with color-coded bars. Great for quick validation of extraction accuracy.",
  },
  {
    keywords: ["filter", "search", "sort", "find items"],
    answer: "Use the filter bar above the BOM table to filter by: Category (pipe, valve, elbow, etc.), Size, Sheet Number, Confidence Level (high/medium/low), and free-text Search. Filters combine — e.g., show only 4\" valves from Sheet 3. Click column headers to sort.",
  },
  {
    keywords: ["estimate", "estimating", "run estimate", "create estimate", "labor hours"],
    answer: "To create an estimate: Go to the Estimating page. You can either 'Estimate from Takeoff' (select an existing takeoff project) or import items manually. The estimator calculates material costs from the cost database and labor hours using either Bill's EI method or Justin's IPMH method. It generates line-by-line manhour calculations.",
  },
  {
    keywords: ["labor rate", "rates", "overtime", "double time", "per diem", "st rate", "ot rate", "dt rate"],
    answer: "Default labor rates: Straight Time (ST) = $56/hr, Overtime (OT) = $79/hr (1.5x base), Double Time (DT) = $100/hr (2x base). Per Diem = $75/day per crew member. These are fully burdened rates. You can adjust OT/DT hours per week on the Estimating page.",
  },
  {
    keywords: ["rack", "factor", "rack factor", "pipe rack", "elevated"],
    answer: "The rack factor (default 1.3x) is a multiplier applied to labor hours for pipe installed on elevated racks vs. ground level. Rack pipe is harder to access, requires scaffolding, and takes longer. You can adjust this factor on the Estimating page.",
  },
  {
    keywords: ["bill", "method", "ei", "engineering index", "bill's method"],
    answer: "Bill's Engineering Index (EI) method uses manhour-per-equivalent-inch rates for each item type. It factors in pipe location (sleeper rack, open rack, underground), elevation, alloy group, and schedule. Rates come from the estimator data JSON file. Best for detailed bottom-up estimates.",
  },
  {
    keywords: ["justin", "ipmh", "justin's method", "factor method"],
    answer: "Justin's IPMH (Installed Piping Manhour) method uses factor tables for each item type. It matches items by NPS (Nominal Pipe Size) and applies size-specific manhour factors for pipe, welds, bolts, valves, fittings, supports, etc. Calibrated against Stolthaven Phase 6 data (IPMH 0.437, within 3.8% of actual).",
  },
  {
    keywords: ["cost database", "material cost", "unit cost", "pricing"],
    answer: "The Cost Database has 232+ records with material and labor unit costs by category, size, and description. It supports vendor quote comparison — upload quotes from multiple vendors and the system highlights the best price. Use the Cost Database page to view, search, and manage pricing records.",
  },
  {
    keywords: ["vendor", "quote", "rfq", "request for quote"],
    answer: "The RFQ Builder lets you generate Request for Quote documents from your takeoff items. Select items, choose a vendor, and generate an RFQ. You can also import vendor quotes via CSV and compare pricing across vendors to find the best deal for each item.",
  },
  {
    keywords: ["project history", "completed project", "historical"],
    answer: "Project History stores data from completed projects — actual manhours, crew composition, duration, and costs. This data feeds calibration and helps validate future estimates. Add completed projects with their actual field data to improve estimation accuracy over time.",
  },
  {
    keywords: ["bid", "tracker", "bid tracker", "proposal"],
    answer: "The Bid Tracker helps manage your active bids and proposals. Track bid status, due dates, estimated values, and outcomes. It connects to your estimates and takeoffs so you can see the full pipeline from takeoff to bid submission.",
  },
  {
    keywords: ["crew", "planner", "crew planner", "crew size", "duration"],
    answer: "The Crew Planner calculates project duration and cost based on crew composition. It uses role-specific rates from Stolthaven Phase 6: Welders $35/hr, Fitters $35/hr, Helpers $25/hr, Firewatches $20/hr (no per diem), Foreman $35/hr, Superintendent $45/hr. Set your pipe size mix and number of areas for optimal crew sizing.",
  },
  {
    keywords: ["archive", "unarchive", "hide", "delete project"],
    answer: "To archive a takeoff project, open it and click the Archive button. Archived projects are hidden from the sidebar and main lists but not deleted. To unarchive, go to the discipline page, find the archived project in the list, and click Unarchive. You can also permanently delete projects.",
  },
  {
    keywords: ["confidence", "scoring", "high", "medium", "low", "accuracy"],
    answer: "Confidence scoring rates each extracted item as High (green), Medium (yellow), or Low (red) based on extraction quality. High confidence means clear, unambiguous data. Medium may have minor parsing issues. Low means the AI was uncertain — review these items manually. Filter by confidence to focus your QC effort.",
  },
  {
    keywords: ["weld inference", "infer welds", "auto weld", "fitting welds"],
    answer: "Weld Inference automatically generates weld items from fittings: Elbows = 2 butt welds each, Tees = 3 butt welds, Reducers = 2 butt welds (at larger size), Caps = 1 butt weld, Couplings = 2 socket welds, Flanges = 1 slip-on weld + 1 bolt-up. Small-bore items (<=1.5\") are skipped — their MH is rolled into weld factors.",
  },
  {
    keywords: ["api key", "settings", "configure", "anthropic", "claude key"],
    answer: "Go to Settings to configure your Anthropic API key for AI extraction. Enter your key (starts with 'sk-ant-') and click Save. The key is stored locally and never shared. Without a key, extraction won't work. The platform may also provide a key automatically if configured.",
  },
  {
    keywords: ["export", "pdf", "excel", "download", "report"],
    answer: "You can export takeoff and estimate data to PDF or Excel. On the takeoff page, look for the Export button in the toolbar. Excel exports include all item details, pivot summaries, and connection counts. PDF exports create a formatted report suitable for printing or sharing with clients.",
  },
  {
    keywords: ["dual model", "gemini", "two models", "dual extraction"],
    answer: "Dual-model extraction runs both Claude (Anthropic) and Gemini (Google) on each page, then merges the results. This catches items that one model might miss and improves accuracy by cross-referencing. It's slower and costs more API calls but produces more reliable results.",
  },
  {
    keywords: ["scope gap", "missing items", "gap detection"],
    answer: "Scope Gap Detection compares your takeoff against expected items for the project type. It identifies potentially missing categories — for example, if you have pipe and fittings but no gaskets or bolts, it flags that as a gap. Helps ensure nothing was missed in extraction.",
  },
  {
    keywords: ["change order", "variation", "extra work"],
    answer: "Change Order generation creates documentation for extra work beyond the original scope. It compares the original takeoff against revised drawings and generates a detailed change order with added/removed items, cost impact, and manhour changes.",
  },
  {
    keywords: ["login", "password", "credentials", "access", "sign in"],
    answer: "Default login credentials: Username: admin, Password: picougroup. You can change these in Settings after logging in. The app uses session-based authentication with sliding expiry.",
  },
  {
    keywords: ["url", "website", "address", "link", "access online"],
    answer: "The app is hosted at: pg-estimator.onrender.com. It's a web application — just open the URL in any modern browser (Chrome, Firefox, Edge, Safari). No installation needed.",
  },
  {
    keywords: ["calibration", "stolthaven", "phase 6", "benchmark"],
    answer: "The app is calibrated against Stolthaven Phase 6 data — a completed project with known actual manhours. Calibrated IPMH: 0.437, target: 0.45, variance: 3.8%, total MH: 56,412. SS weld factors: 3\" = 4.68 MH/weld, 4\" = 5.56 MH/weld. This calibration ensures estimates are grounded in real field performance.",
  },
  {
    keywords: ["small bore", "1.5", "nipple", "plug", "coupling", "rollup"],
    answer: "Small-bore items (1.5\" and under) like plugs, nipples, couplings, sockolets, and weldolets have their manhours rolled into weld factors rather than being counted individually. This matches industry practice where small-bore work is bundled with nearby large-bore welds.",
  },
];

function findBestAnswer(query: string): string {
  const q = query.toLowerCase().trim();
  if (q.length < 2) return "Please type a question and I'll help you out!";

  const words = q.split(/\s+/).filter(w => w.length > 2);
  let bestScore = 0;
  let bestAnswer = "";

  for (const entry of KNOWLEDGE_BASE) {
    let score = 0;
    for (const kw of entry.keywords) {
      const kwLower = kw.toLowerCase();
      if (q.includes(kwLower)) {
        score += kwLower.length * 3;
      } else {
        for (const word of words) {
          if (kwLower.includes(word)) score += word.length;
          if (word.includes(kwLower)) score += kwLower.length;
        }
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestAnswer = entry.answer;
    }
  }

  return bestScore >= 4
    ? bestAnswer
    : "I'm not sure about that. Try switching to AI mode for a more detailed answer, or rephrase your question.";
}

// ============================================================
// Types
// ============================================================

interface ChatMessage {
  id: number;
  role: "user" | "assistant";
  text: string;
}

type ChatMode = "help" | "ai";

// ============================================================
// Component
// ============================================================

export default function ChatAssistant({ pageContext }: { pageContext?: string }) {
  const [open, setOpen] = useState(false);
  const [mode, setMode] = useState<ChatMode>("help");
  const [messages, setMessages] = useState<ChatMessage[]>([
    { id: 0, role: "assistant", text: "Hi! I'm PG Assistant. Ask me anything about the Picou Group Estimator. Switch to AI mode for deeper questions." },
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const msgEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const idRef = useRef(1);

  useEffect(() => {
    msgEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  useEffect(() => {
    if (open) inputRef.current?.focus();
  }, [open]);

  const handleSend = useCallback(async () => {
    const text = input.trim();
    if (!text || loading) return;
    setInput("");

    const userMsg: ChatMessage = { id: idRef.current++, role: "user", text };
    setMessages(prev => [...prev, userMsg]);

    if (mode === "help") {
      const answer = findBestAnswer(text);
      setMessages(prev => [...prev, { id: idRef.current++, role: "assistant", text: answer }]);
      return;
    }

    // AI mode
    setLoading(true);
    try {
      const resp = await fetch("/api/chat-assistant", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message: text, context: pageContext }),
      });
      const data = await resp.json();
      setMessages(prev => [...prev, {
        id: idRef.current++,
        role: "assistant",
        text: data.response || "Sorry, I couldn't process that.",
      }]);
    } catch {
      setMessages(prev => [...prev, {
        id: idRef.current++,
        role: "assistant",
        text: "Connection error. Please check your network and try again.",
      }]);
    } finally {
      setLoading(false);
    }
  }, [input, loading, mode, pageContext]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleClose = () => {
    setOpen(false);
  };

  const handleOpen = () => {
    setOpen(true);
  };

  return (
    <>
      {/* ===== Floating Button ===== */}
      {!open && (
        <button
          onClick={handleOpen}
          className="fixed bottom-5 right-5 z-50 h-14 w-14 rounded-full bg-primary text-primary-foreground shadow-lg hover:shadow-xl hover:scale-105 transition-all duration-200 flex items-center justify-center"
          aria-label="Open chat assistant"
        >
          <MessageCircle size={24} />
        </button>
      )}

      {/* ===== Chat Panel ===== */}
      {open && (
        <div
          className="fixed bottom-5 right-5 z-50 flex flex-col bg-card border border-border rounded-xl shadow-2xl overflow-hidden
            w-[calc(100vw-2.5rem)] sm:w-[380px] h-[min(500px,calc(100vh-3rem))]
            animate-in slide-in-from-bottom-4 fade-in duration-200"
        >
          {/* Header */}
          <div className="flex items-center justify-between px-4 py-3 border-b border-border bg-muted/50 shrink-0">
            <div className="flex items-center gap-2">
              <Sparkles size={16} className="text-primary" />
              <span className="text-sm font-semibold text-foreground">PG Assistant</span>
            </div>
            <div className="flex items-center gap-2">
              {/* Mode toggle */}
              <div className="flex items-center bg-muted rounded-md p-0.5 gap-0.5">
                <button
                  onClick={() => setMode("help")}
                  className={`flex items-center gap-1 px-2 py-1 rounded text-[10px] font-medium transition-colors ${
                    mode === "help"
                      ? "bg-background text-foreground shadow-sm"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  <HelpCircle size={10} />
                  Help
                  {mode === "help" && (
                    <Badge variant="secondary" className="text-[8px] px-1 py-0 h-3.5 bg-green-100 text-green-700 dark:bg-green-900/40 dark:text-green-400">
                      Free
                    </Badge>
                  )}
                </button>
                <button
                  onClick={() => setMode("ai")}
                  className={`flex items-center gap-1 px-2 py-1 rounded text-[10px] font-medium transition-colors ${
                    mode === "ai"
                      ? "bg-background text-foreground shadow-sm"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  <Sparkles size={10} />
                  AI
                  {mode === "ai" && (
                    <Badge variant="secondary" className="text-[8px] px-1 py-0 h-3.5 bg-blue-100 text-blue-700 dark:bg-blue-900/40 dark:text-blue-400">
                      AI
                    </Badge>
                  )}
                </button>
              </div>
              <button
                onClick={handleClose}
                className="text-muted-foreground hover:text-foreground transition-colors p-0.5 rounded"
                aria-label="Close chat"
              >
                <X size={16} />
              </button>
            </div>
          </div>

          {/* Messages */}
          <div className="flex-1 overflow-y-auto px-3 py-3 space-y-3 min-h-0">
            {messages.map(msg => (
              <div
                key={msg.id}
                className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
              >
                <div
                  className={`max-w-[85%] px-3 py-2 rounded-xl text-xs leading-relaxed whitespace-pre-wrap ${
                    msg.role === "user"
                      ? "bg-primary text-primary-foreground rounded-br-sm"
                      : "bg-muted text-foreground rounded-bl-sm"
                  }`}
                >
                  {msg.text}
                </div>
              </div>
            ))}

            {/* Typing indicator */}
            {loading && (
              <div className="flex justify-start">
                <div className="bg-muted rounded-xl rounded-bl-sm px-4 py-2.5 flex items-center gap-1">
                  <span className="w-1.5 h-1.5 bg-muted-foreground/60 rounded-full animate-bounce [animation-delay:0ms]" />
                  <span className="w-1.5 h-1.5 bg-muted-foreground/60 rounded-full animate-bounce [animation-delay:150ms]" />
                  <span className="w-1.5 h-1.5 bg-muted-foreground/60 rounded-full animate-bounce [animation-delay:300ms]" />
                </div>
              </div>
            )}

            <div ref={msgEndRef} />
          </div>

          {/* Input */}
          <div className="border-t border-border px-3 py-2.5 shrink-0 bg-card">
            <div className="flex items-center gap-2">
              <Input
                ref={inputRef}
                className="h-9 text-xs flex-1"
                placeholder={mode === "help" ? "Ask about the app..." : "Ask anything (uses AI)..."}
                value={input}
                onChange={e => setInput(e.target.value)}
                onKeyDown={handleKeyDown}
                disabled={loading}
              />
              <Button
                size="icon"
                variant="default"
                className="h-9 w-9 shrink-0"
                onClick={handleSend}
                disabled={!input.trim() || loading}
                aria-label="Send message"
              >
                <Send size={14} />
              </Button>
            </div>
            <p className="text-[9px] text-muted-foreground mt-1.5 text-center">
              {mode === "help" ? "Answers from built-in knowledge base" : "Powered by Claude \u2014 uses your API key"}
            </p>
          </div>
        </div>
      )}
    </>
  );
}
