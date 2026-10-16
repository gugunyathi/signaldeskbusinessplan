import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import {
  ArrowDownToLine,
  ArrowLeft,
  ArrowRight,
  FileText,
  ChartNoAxesCombined,
  Layers,
  Check,
  ChevronRight,
  Radio,
  FileSliders,
  Sparkles,
  TrendingUp,
  Cpu,
  Users,
  Building2,
  DollarSign,
  ShieldCheck,
  Award,
  BarChart3,
  Compass,
} from "lucide-react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
} from "recharts";
import { Button } from "@/components/ui/button";
import plan from "@/assets/plan-manifest.json";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Signal Desk | Business Plan & Operational Model" },
      {
        name: "description",
        content:
          "Strategic framework for a cloud-based public infrastructure triage & asset-liquidation platform. 5-year financials and downloadable PowerPoint.",
      },
      { property: "og:title", content: "Signal Desk | Business Plan & Operational Model" },
      {
        property: "og:description",
        content:
          "5-year financial trajectory, AI lab dataset licensing model, and downloadable PowerPoint.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

const pptxDownloadPath = "/downloads/Signal_Desk_Business_Plan.pptx";
const pdfDownloadPath = "/downloads/Signal_Desk_Business_Plan.pdf";

async function triggerDirectDownload(url: string, filename: string) {
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP error ${res.status}`);
    const blob = await res.blob();
    const blobUrl = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = filename;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => window.URL.revokeObjectURL(blobUrl), 3000);
  } catch {
    // Direct link fallback
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}

function DownloadButtons({ compact = false }: { compact?: boolean }) {
  const [downloading, setDownloading] = useState<string | null>(null);

  const handleDownload = async (url: string, filename: string) => {
    setDownloading(filename);
    try {
      await triggerDirectDownload(url, filename);
    } finally {
      setTimeout(() => setDownloading(null), 1200);
    }
  };

  return (
    <div className="flex flex-wrap items-center gap-2">
      <Button
        type="button"
        size={compact ? "sm" : "default"}
        onClick={() => handleDownload(pdfDownloadPath, "Signal_Desk_Business_Plan.pdf")}
        className="cursor-pointer font-medium bg-primary text-primary-foreground hover:bg-primary/90 shadow-xs"
      >
        <FileText className="size-4" />
        {downloading === "Signal_Desk_Business_Plan.pdf"
          ? "Downloading..."
          : compact
            ? "Download PDF"
            : "Direct Download PDF"}
      </Button>
      <Button
        type="button"
        size={compact ? "sm" : "default"}
        variant="outline"
        onClick={() => handleDownload(pptxDownloadPath, "Signal_Desk_Business_Plan.pptx")}
        className="cursor-pointer font-medium"
      >
        <FileSliders className="size-4" />
        {downloading === "Signal_Desk_Business_Plan.pptx"
          ? "Downloading..."
          : compact
            ? ".pptx"
            : "Download PowerPoint (.pptx)"}
      </Button>
    </div>
  );
}

// Chart Colors matching the design system
const COLORS = {
  navy: "#162B3A",
  primary: "#008A68",
  blue: "#0068A0",
  cyan: "#00A896",
  amber: "#E3B341",
  coral: "#E76F51",
  muted: "#62717D",
  slate: "#8E9AAF",
};

const PIE_COLORS_Y1 = [COLORS.blue, COLORS.primary, COLORS.cyan];
const PIE_COLORS_Y5 = [
  COLORS.primary,
  COLORS.blue,
  COLORS.cyan,
  COLORS.amber,
  COLORS.coral,
  COLORS.muted,
  COLORS.slate,
];

function SlideRenderer({ slideIndex }: { slideIndex: number }) {
  const currentSlide = plan.slides[slideIndex];
  if (!currentSlide) return null;

  // Custom rich preview for each slide
  switch (slideIndex) {
    case 0: // Cover
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-[#0E1A24] p-6 text-white sm:p-10">
          <div>
            <div className="flex items-center gap-2 text-xs font-bold tracking-widest text-emerald-400 uppercase sm:text-sm">
              <Radio className="size-4" /> SIGNAL DESK / STRATEGIC BUSINESS PLAN
            </div>
            <h1 className="mt-4 text-2xl font-bold tracking-tight text-white sm:text-4xl md:text-5xl">
              SIGNAL DESK
            </h1>
            <p className="mt-2 text-base font-semibold text-emerald-300 sm:text-xl">
              Signal Desk Business Plan
            </p>
            <p className="mt-2 max-w-3xl text-xs font-medium text-slate-200 sm:text-sm">
              Strategic Framework for an Autonomous Orchestration Layer for Municipal Infrastructure
              Platform.
            </p>
            <p className="mt-1.5 max-w-3xl text-xs text-emerald-400/90 sm:text-xs">
              AI-powered data routing and smart workflows bridging citizen intelligence with
              municipal action.
            </p>
          </div>
          <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
            <div className="rounded border border-emerald-500/30 bg-[#162736] p-3">
              <span className="text-[10px] font-bold text-teal-400 uppercase">Seed Runway</span>
              <p className="mt-1 text-base font-bold text-white sm:text-lg">$100k (18–24 mo)</p>
            </div>
            <div className="rounded border border-emerald-500/30 bg-[#162736] p-3">
              <span className="text-[10px] font-bold text-teal-400 uppercase">Year 5 Revenue</span>
              <p className="mt-1 text-base font-bold text-white sm:text-lg">$17.425M</p>
            </div>
            <div className="rounded border border-emerald-500/30 bg-[#162736] p-3">
              <span className="text-[10px] font-bold text-teal-400 uppercase">Year 5 Profit</span>
              <p className="mt-1 text-base font-bold text-white sm:text-lg">+$15.260M</p>
            </div>
            <div className="rounded border border-emerald-500/30 bg-[#162736] p-3">
              <span className="text-[10px] font-bold text-teal-400 uppercase">Gross Margin</span>
              <p className="mt-1 text-base font-bold text-white sm:text-lg">89.9%</p>
            </div>
          </div>
        </div>
      );

    case 1: // Executive Summary
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-6 sm:p-8">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / Executive Summary
            </div>
            <h2 className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
              High-Margin Infrastructure Data Ecosystem
            </h2>
            <p className="mt-2 text-xs text-muted-foreground sm:text-sm">
              By capturing real-time municipal triage logs, citizen routing, and multi-team field
              dispatches, Signal Desk creates a proprietary dataset licensed directly to frontier AI
              labs.
            </p>
          </div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div className="rounded-md border border-border bg-muted/40 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-primary uppercase">Source Layer</span>
              <h3 className="mt-1 text-sm font-bold text-foreground">Proprietary Ingestion</h3>
              <p className="mt-2 text-xs text-muted-foreground">
                Triage logs and routing patterns captured from active ward communities.
              </p>
            </div>
            <div className="rounded-md border border-border bg-muted/40 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-primary uppercase">
                High-Margin Scale
              </span>
              <h3 className="mt-1 text-sm font-bold text-foreground">Frontier AI Labs</h3>
              <p className="mt-2 text-xs text-muted-foreground">
                Licensed to Anthropic, Google, OpenAI, X/Cursor, and Meta for RL agent training.
              </p>
            </div>
            <div className="rounded-md border border-border bg-muted/40 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-primary uppercase">Runway Security</span>
              <h3 className="mt-1 text-sm font-bold text-foreground">18–24 Month Runtime</h3>
              <p className="mt-2 text-xs text-muted-foreground">
                $100k seed capital funds Year 1 baseline outflows ($75k) to achieve data volume.
              </p>
            </div>
          </div>
        </div>
      );

    case 2: // Runway Analysis
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-6 sm:p-8">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / Runway Analysis
            </div>
            <h2 className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
              Financial Feasibility & Seed Runway Allocation ($100k)
            </h2>
          </div>
          <div className="space-y-2.5">
            <div className="flex items-center justify-between rounded border border-border bg-muted/30 p-3">
              <div>
                <strong className="text-xs font-semibold sm:text-sm">
                  Sunk Assets (Historical R&D)
                </strong>
                <p className="text-[11px] text-muted-foreground">
                  Foundational engine engineering capitalized as sunk setup asset
                </p>
              </div>
              <span className="text-sm font-bold text-muted-foreground sm:text-base">$30,000</span>
            </div>
            <div className="flex items-center justify-between rounded border border-emerald-500/30 bg-emerald-500/10 p-3">
              <div>
                <strong className="text-xs font-semibold text-emerald-800 dark:text-emerald-300 sm:text-sm">
                  Net Cash Position (Working Capital)
                </strong>
                <p className="text-[11px] text-muted-foreground">
                  Initial liquid runway deployed for operational validation
                </p>
              </div>
              <span className="text-sm font-bold text-emerald-600 sm:text-base">$70,000</span>
            </div>
            <div className="flex items-center justify-between rounded border border-border bg-muted/30 p-3">
              <div>
                <strong className="text-xs font-semibold sm:text-sm">
                  Fixed Core Infrastructure
                </strong>
                <p className="text-[11px] text-muted-foreground">
                  Cloud server ($500/mo) and Claude API workflow layers ($1,000/mo)
                </p>
              </div>
              <span className="text-sm font-bold text-primary sm:text-base">$18,000 / yr</span>
            </div>
            <div className="flex items-center justify-between rounded border border-border bg-muted/30 p-3">
              <div>
                <strong className="text-xs font-semibold sm:text-sm">
                  Year 1 Baseline Outflows
                </strong>
                <p className="text-[11px] text-muted-foreground">
                  $57,000 lean opex ($45k payroll + $12k other) + $18,000 fixed infrastructure
                </p>
              </div>
              <span className="text-sm font-bold text-foreground sm:text-base">$75,000 / yr</span>
            </div>
          </div>
        </div>
      );

    case 3: // GTM Strategy
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-6 sm:p-8">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / GTM Strategy
            </div>
            <h2 className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
              Triple-Pronged Wedge Strategy
            </h2>
            <p className="mt-1 text-xs text-muted-foreground">
              Circumventing traditional municipal procurement friction through bottom-up adoption.
            </p>
          </div>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
            <div className="rounded border border-border bg-muted/30 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-blue-600 uppercase">
                Stage 1 • Low Friction
              </span>
              <h3 className="mt-1 text-sm font-bold">Councillor Dashboard</h3>
              <span className="mt-1 inline-block text-xs font-semibold text-primary">
                $10 / month
              </span>
              <p className="mt-2 text-xs text-muted-foreground">
                Direct outreach loops introducing ward councillors to dedicated tracking tools.
              </p>
            </div>
            <div className="rounded border border-border bg-muted/30 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-blue-600 uppercase">
                Stage 2 • Workflow Link
              </span>
              <h3 className="mt-1 text-sm font-bold">B2B Service Providers</h3>
              <span className="mt-1 inline-block text-xs font-semibold text-primary">
                $50 / month
              </span>
              <p className="mt-2 text-xs text-muted-foreground">
                Contracted external vendors use the interface for structured dispatches.
              </p>
            </div>
            <div className="rounded border border-emerald-500/40 bg-emerald-500/10 p-3 sm:p-4">
              <span className="text-[10px] font-bold text-emerald-600 uppercase">
                Stage 3 • Institutional
              </span>
              <h3 className="mt-1 text-sm font-bold">AI Data Asset Liquidation</h3>
              <span className="mt-1 inline-block text-xs font-semibold text-emerald-600">
                $100k → $1.0M / package
              </span>
              <p className="mt-2 text-xs text-muted-foreground">
                Anonymized municipal triage datasets packaged and licensed to frontier AI labs.
              </p>
            </div>
          </div>
        </div>
      );

    case 4: // CivicRewards
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-6 sm:p-8">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / Product Architecture
            </div>
            <h2 className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
              CivicRewards Loyalty Framework
            </h2>
            <p className="mt-1 text-xs text-muted-foreground">
              Maintaining reporting volume & data verification without increasing cash burn.
            </p>
          </div>
          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
            <div className="rounded border border-border bg-muted/30 p-2.5 sm:p-3">
              <div className="text-xs font-bold text-primary">01</div>
              <strong className="text-xs">Citizen Report</strong>
              <p className="mt-1 text-[11px] text-muted-foreground">
                Citizen submits verified utility gap via Telegram/web.
              </p>
            </div>
            <div className="rounded border border-border bg-muted/30 p-2.5 sm:p-3">
              <div className="text-xs font-bold text-primary">02</div>
              <strong className="text-xs">Geolocation & De-dupe</strong>
              <p className="mt-1 text-[11px] text-muted-foreground">
                Automated validation of GPS, duplicates, and official IDs.
              </p>
            </div>
            <div className="rounded border border-border bg-muted/30 p-2.5 sm:p-3">
              <div className="text-xs font-bold text-primary">03</div>
              <strong className="text-xs">Token Allocation</strong>
              <p className="mt-1 text-[11px] text-muted-foreground">
                CivicPoints awarded upon community verified resolution.
              </p>
            </div>
            <div className="rounded border border-border bg-muted/30 p-2.5 sm:p-3">
              <div className="text-xs font-bold text-primary">04</div>
              <strong className="text-xs">Partner Redemption</strong>
              <p className="mt-1 text-[11px] text-muted-foreground">
                Funded by B2B Advertisers ($1k/mo) with zero cash burn.
              </p>
            </div>
          </div>
          <div className="rounded border border-border bg-card p-3 text-xs text-muted-foreground">
            <strong>Loop Monetization:</strong> Citizens redeem points for local micro-discounts.
            These incentives are funded entirely by premium local advertisers ($1,000/mo), creating
            a sustainable engagement loop.
          </div>
        </div>
      );

    case 7: // Table A P&L
      return (
        <div className="flex h-full flex-col justify-between overflow-x-auto rounded-lg bg-card p-4 sm:p-6">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / Financial Model
            </div>
            <h2 className="text-base font-bold text-foreground sm:text-xl">
              Table A: 5-Year Consolidated Financial Projection (USD)
            </h2>
          </div>
          <table className="w-full text-left text-[11px] sm:text-xs">
            <thead>
              <tr className="border-b border-border bg-muted/50 font-bold">
                <th className="py-1.5 pr-2">Metric</th>
                <th className="py-1.5 text-right">Year 1</th>
                <th className="py-1.5 text-right">Year 2</th>
                <th className="py-1.5 text-right">Year 3</th>
                <th className="py-1.5 text-right">Year 4</th>
                <th className="py-1.5 text-right">Year 5</th>
              </tr>
            </thead>
            <tbody>
              <tr className="border-b border-border/50 font-semibold text-primary">
                <td className="py-1">Gross Annual Revenue</td>
                <td className="py-1 text-right">$30,000</td>
                <td className="py-1 text-right">$315,000</td>
                <td className="py-1 text-right">$1,523,000</td>
                <td className="py-1 text-right">$13,261,000</td>
                <td className="py-1 text-right">$17,425,000</td>
              </tr>
              <tr className="border-b border-border/50">
                <td className="py-1 text-muted-foreground">Cost of Revenue (COR)</td>
                <td className="py-1 text-right">$22,500</td>
                <td className="py-1 text-right">$62,100</td>
                <td className="py-1 text-right">$186,990</td>
                <td className="py-1 text-right">$1,361,520</td>
                <td className="py-1 text-right">$1,760,500</td>
              </tr>
              <tr className="border-b border-border/50 font-medium">
                <td className="py-1">Gross Profit</td>
                <td className="py-1 text-right">$7,500</td>
                <td className="py-1 text-right">$252,900</td>
                <td className="py-1 text-right">$1,336,010</td>
                <td className="py-1 text-right">$11,899,480</td>
                <td className="py-1 text-right">$15,664,500</td>
              </tr>
              <tr className="border-b border-border/50 text-emerald-600">
                <td className="py-1">Gross Margin %</td>
                <td className="py-1 text-right">25.0%</td>
                <td className="py-1 text-right">80.3%</td>
                <td className="py-1 text-right">87.7%</td>
                <td className="py-1 text-right">89.7%</td>
                <td className="py-1 text-right">89.9%</td>
              </tr>
              <tr className="border-b border-border/50">
                <td className="py-1 text-muted-foreground">Total Operating Expenses</td>
                <td className="py-1 text-right">$57,000</td>
                <td className="py-1 text-right">$90,000</td>
                <td className="py-1 text-right">$155,000</td>
                <td className="py-1 text-right">$260,000</td>
                <td className="py-1 text-right">$405,000</td>
              </tr>
              <tr className="bg-emerald-500/10 font-bold">
                <td className="py-1.5">Net Operating Profit / (Loss)</td>
                <td className="py-1.5 text-right text-rose-600">(-$49,500)</td>
                <td className="py-1.5 text-right text-emerald-600">+$162,900</td>
                <td className="py-1.5 text-right text-emerald-600">+$1,181,010</td>
                <td className="py-1.5 text-right text-emerald-600">+$11,639,480</td>
                <td className="py-1.5 text-right text-emerald-600">+$15,259,500</td>
              </tr>
            </tbody>
          </table>
        </div>
      );

    case 9: // Chart 1 Trajectory
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-4 sm:p-6">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              Signal Desk / Visual Projections
            </div>
            <h2 className="text-base font-bold text-foreground sm:text-xl">
              5-Year Financial Trajectory & AI Lab Data Scaling
            </h2>
          </div>
          <div className="h-56 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart
                data={plan.financials.years.map((y, i) => ({
                  year: y,
                  Revenue: plan.financials.revenue[i]! / 1000,
                  Expenses: plan.financials.totalExpenses[i]! / 1000,
                  NetProfit: plan.financials.netProfitLoss[i]! / 1000,
                }))}
              >
                <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                <XAxis dataKey="year" fontSize={11} />
                <YAxis fontSize={11} unit="k" />
                <Tooltip formatter={(val: number) => `$${val.toLocaleString()}k`} />
                <Legend />
                <Bar dataKey="Revenue" fill={COLORS.blue} radius={[4, 4, 0, 0]} />
                <Bar dataKey="Expenses" fill={COLORS.muted} radius={[4, 4, 0, 0]} />
                <Bar dataKey="NetProfit" fill={COLORS.primary} radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      );

    default: // Generic slide representation
      return (
        <div className="flex h-full flex-col justify-between rounded-lg bg-card p-6 sm:p-8">
          <div>
            <div className="text-xs font-bold text-primary uppercase">
              SIGNAL DESK / {currentSlide.section.toUpperCase()}
            </div>
            <h2 className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
              {currentSlide.title}
            </h2>
            <p className="mt-2 text-sm text-muted-foreground">{currentSlide.summary}</p>
          </div>
          <div className="rounded-md border border-border bg-muted/20 p-4">
            <span className="text-xs font-semibold text-primary">Slide Content Highlight</span>
            <p className="mt-1 text-xs text-muted-foreground">
              Editable native PowerPoint component included in the downloadable .pptx file. Fully
              reconciled with the October 2026 Comprehensive Business Plan & Operational Model.
            </p>
          </div>
          <div className="flex items-center justify-between text-xs text-muted-foreground">
            <span>Signal Desk • Comprehensive Business Plan</span>
            <span>
              Slide {slideIndex + 1} of {plan.slideCount}
            </span>
          </div>
        </div>
      );
  }
}

function Index() {
  const [view, setView] = useState<"overview" | "financials" | "strategy" | "full">("overview");
  const [slide, setSlide] = useState(0);
  const selected = plan.slides[slide];

  // Recharts Trajectory Data
  const trajectoryData = plan.financials.years.map((y, i) => ({
    name: y,
    Revenue: plan.financials.revenue[i],
    Expenses: plan.financials.totalExpenses[i],
    NetProfit: plan.financials.netProfitLoss[i],
    GrossProfit: plan.financials.grossProfit[i],
  }));

  return (
    <div className="workspace">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-symbol">
            <Radio size={20} />
          </span>
          Signal Desk
        </div>
        <div className="sidebar-label">BUSINESS PLAN VIEWER</div>
        <nav aria-label="Business plan views">
          {[
            { id: "overview", label: "Executive overview", icon: FileText },
            { id: "financials", label: "Financial model (P&L)", icon: ChartNoAxesCombined },
            { id: "strategy", label: "GTM & AI data refinery", icon: Compass },
            { id: "full", label: "All 15 slides", icon: Layers },
          ].map((item) => (
            <Button
              key={item.id}
              className="nav-item cursor-pointer"
              variant={view === item.id ? "secondary" : "ghost"}
              onClick={() => setView(item.id as typeof view)}
            >
              <item.icon className="size-4" />
              {item.label}
            </Button>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="flex items-center gap-2 text-foreground">
            <span className="brand-symbol">SD</span>
            <strong>Signal Desk</strong>
          </div>
          <p className="mt-2 text-xs text-muted-foreground">
            Public Infrastructure Triage & Asset-Liquidation Platform
            <br />
            Johannesburg, South Africa • Oct 2026
          </p>
        </div>
      </aside>

      <div className="min-w-0">
        <header className="topbar">
          <div className="flex items-center gap-2 text-muted-foreground">
            Documents
            <ChevronRight size={12} />
            <span className="text-foreground">Comprehensive Business Plan</span>
          </div>
          <div className="flex items-center gap-3">
            <span className="hidden items-center gap-1.5 text-xs text-muted-foreground sm:flex">
              <span className="status-dot" />
              Updated October 2026 Model
            </span>
            <DownloadButtons compact />
          </div>
        </header>

        <main className="content">
          <div className="page-heading">
            <div>
              <div className="eyebrow flex items-center gap-1.5">
                <Sparkles size={12} /> Strategic Framework & Operational Model
              </div>
              <h1 className="text-2xl font-bold sm:text-3xl">
                {view === "financials"
                  ? "5-Year Pro-Forma Income Statement & Unit Economics"
                  : view === "strategy"
                    ? "GTM Strategy & Frontier AI Data Partnerships"
                    : view === "full"
                      ? "Complete 15-Slide Presentation Deck"
                      : "Signal Desk Business Plan"}
              </h1>
              <p className="mt-1 text-sm font-medium text-foreground">
                Strategic Framework for an Autonomous Orchestration Layer for Municipal Infrastructure Platform.
              </p>
              <p className="mt-1 text-xs text-muted-foreground">
                AI-powered data routing and smart workflows bridging citizen intelligence with municipal action.
              </p>
            </div>
            <DownloadButtons />
          </div>

          {/* OVERVIEW VIEW */}
          {view === "overview" && (
            <>
              {/* Highlight Metrics */}
              <div className="mb-6 grid grid-cols-2 gap-3 sm:grid-cols-4">
                <div className="rounded-lg border border-border bg-card p-4 shadow-xs">
                  <span className="text-xs font-semibold text-muted-foreground">Seed Runway</span>
                  <div className="mt-1 text-xl font-bold text-foreground sm:text-2xl">
                    18–24 Months
                  </div>
                  <span className="text-[11px] text-muted-foreground">
                    $100k seed funds $75k Y1 burn
                  </span>
                </div>
                <div className="rounded-lg border border-border bg-card p-4 shadow-xs">
                  <span className="text-xs font-semibold text-muted-foreground">
                    Year 5 Gross Revenue
                  </span>
                  <div className="mt-1 text-xl font-bold text-emerald-600 sm:text-2xl">$17.43M</div>
                  <span className="text-[11px] text-muted-foreground">
                    86.1% AI Data Broker sales
                  </span>
                </div>
                <div className="rounded-lg border border-border bg-card p-4 shadow-xs">
                  <span className="text-xs font-semibold text-muted-foreground">
                    Year 5 Net Profit
                  </span>
                  <div className="mt-1 text-xl font-bold text-emerald-600 sm:text-2xl">
                    +$15.26M
                  </div>
                  <span className="text-[11px] text-muted-foreground">89.9% Gross Margin</span>
                </div>
                <div className="rounded-lg border border-border bg-card p-4 shadow-xs">
                  <span className="text-xs font-semibold text-muted-foreground">
                    Single-Ward CAC Payback
                  </span>
                  <div className="mt-1 text-xl font-bold text-primary sm:text-2xl">2.25 Months</div>
                  <span className="text-[11px] text-muted-foreground">
                    $45 CAC vs $20/mo blended yield
                  </span>
                </div>
              </div>

              {/* Interactive Slide Viewer & Details Panel */}
              <div className="document-layout">
                <div>
                  <div className="preview-shell">
                    <div className="preview-label">
                      <span className="flex items-center gap-2 font-medium">
                        <FileSliders size={15} />
                        PRESENTATION PREVIEW / SLIDE {String(slide + 1).padStart(2, "0")} OF{" "}
                        {plan.slideCount}
                      </span>
                      <span className="text-xs text-muted-foreground">{selected?.section}</span>
                    </div>

                    {/* 16:9 Interactive Slide Canvas */}
                    <div className="relative aspect-video w-full bg-slate-900/5 p-2 sm:p-4">
                      <SlideRenderer slideIndex={slide} />
                    </div>

                    <div className="slide-controls">
                      <span className="text-xs text-muted-foreground">
                        Slide {slide + 1} of {plan.slideCount} • {selected?.title}
                      </span>
                      <div className="flex items-center gap-1">
                        <Button
                          aria-label="Previous slide"
                          variant="ghost"
                          size="sm"
                          disabled={slide === 0}
                          onClick={() => setSlide((s) => s - 1)}
                        >
                          <ArrowLeft className="size-4" /> Prev
                        </Button>
                        <Button
                          aria-label="Next slide"
                          variant="ghost"
                          size="sm"
                          disabled={slide === plan.slideCount - 1}
                          onClick={() => setSlide((s) => s + 1)}
                        >
                          Next <ArrowRight className="size-4" />
                        </Button>
                      </div>
                    </div>
                  </div>
                  <h2 className="slide-title">{selected?.title}</h2>
                  <p className="mt-1 text-xs text-muted-foreground">{selected?.summary}</p>
                </div>

                {/* Document Details Panel */}
                <aside className="detail-panel">
                  <h2>Operational document</h2>
                  <div className="detail-row">
                    <span>Format</span>
                    <strong>PDF (.pdf) & PowerPoint (.pptx)</strong>
                  </div>
                  <div className="detail-row">
                    <span>Deck Length</span>
                    <strong>{plan.slideCount} editable slides</strong>
                  </div>
                  <div className="detail-row">
                    <span>Planning Horizon</span>
                    <strong>5 years (USD)</strong>
                  </div>
                  <div className="detail-row">
                    <span>Initial Seed Raise</span>
                    <strong>$100,000 (18–24 mo)</strong>
                  </div>
                  <div className="detail-row">
                    <span>Frontier AI Buyers</span>
                    <strong>Anthropic, Google, OpenAI</strong>
                  </div>
                  <div className="detail-divider" />
                  <div className="flex items-center gap-3">
                    <FileSliders size={32} className="format-icon" />
                    <div className="text-xs">
                      <strong>Executive Document & Slides</strong>
                      <p className="mt-0.5 text-muted-foreground">
                        Vector charts • Direct PDF & PowerPoint download
                      </p>
                    </div>
                  </div>
                  <div className="detail-divider" />
                  <div className="flex items-center gap-2 text-xs text-primary">
                    <Check size={14} />
                    Fully Reconciled with Oct 2026 Model
                  </div>
                  <p className="note mt-2">
                    Verified seed runway, Table A P&L, triple-pronged GTM, CivicRewards loop, and
                    single-ward CTS decay curve.
                  </p>
                  <div className="mt-4">
                    <DownloadButtons compact />
                  </div>
                </aside>
              </div>

              {/* 8 Plan Modules Overview */}
              <h2 className="section-heading">Operational Architecture & Modules</h2>
              <div className="section-grid">
                {[
                  {
                    num: "01",
                    title: "Executive Summary & Runway Analysis",
                    body: "B2G/B2B data ecosystem. $100k seed funds 18–24 mo runtime ($30k sunk asset, $70k net cash, $75k Y1 burn).",
                  },
                  {
                    num: "02",
                    title: "Triple-Pronged Wedge (GTM)",
                    body: "Stage 1: Councillor dashboard ($10/mo) → Stage 2: B2B service providers ($50/mo) → Stage 3: High-yield AI data liquidation.",
                  },
                  {
                    num: "03",
                    title: "CivicRewards Loyalty Framework",
                    body: "GPS-verified citizen incident reporting earning CivicPoints, redeemed through B2B advertisers paying $1k/mo without cash burn.",
                  },
                  {
                    num: "04",
                    title: "AI Broker Monetization Timeline",
                    body: "Y2 micro1 ($100k) → Y3 2 brokers ($500k) → Y4/Y5 3 municipal datasets to Anthropic, Google, OpenAI, X/Cursor, Meta ($1M each).",
                  },
                  {
                    num: "05",
                    title: "5-Year Scaling Milestones",
                    body: "Scaling from 100 councillors & 20 providers in Year 1 to 3,200 councillors, 500 providers, 90 municipalities in Year 5.",
                  },
                  {
                    num: "06",
                    title: "Pro-Forma P&L (Table A)",
                    body: "Turnaround from (-$49.5k) in Year 1 to +$162.9k in Year 2, reaching $17.425M revenue and +$15.260M profit in Year 5.",
                  },
                  {
                    num: "07",
                    title: "Visual Financial Trajectory",
                    body: "Exponential top-line growth with 89.9% Year-5 gross margins while holding operational headcount lean (13 staff).",
                  },
                  {
                    num: "08",
                    title: "Single-Ward Micro Unit Economics",
                    body: "$20/mo blended revenue yield vs $3/mo variable token cost; $18/mo Y1 CTS declining to $3.25/mo by Y5; $45 CAC.",
                  },
                ].map((item) => (
                  <article className="section-card" key={item.num}>
                    <span className="section-number">{item.num}</span>
                    <h3>{item.title}</h3>
                    <p>{item.body}</p>
                  </article>
                ))}
              </div>
            </>
          )}

          {/* FINANCIALS VIEW */}
          {view === "financials" && (
            <div className="space-y-8">
              {/* Top Key Metrics */}
              <div className="metrics">
                <div className="metric">
                  <span>Year 5 Gross Revenue</span>
                  <strong className="text-emerald-600">$17.425M</strong>
                  <span>86.1% AI Data Asset sales</span>
                </div>
                <div className="metric">
                  <span>Year 5 Net Operating Profit</span>
                  <strong className="text-emerald-600">+$15.260M</strong>
                  <span>89.9% Gross Margin</span>
                </div>
                <div className="metric">
                  <span>Seed Round Runway</span>
                  <strong className="text-primary">18–24 Months</strong>
                  <span>$100k capital / $75k Y1 burn</span>
                </div>
              </div>

              {/* Table A: 5-Year Consolidated Financial Projection */}
              <div>
                <div className="flex flex-col justify-between sm:flex-row sm:items-center">
                  <div>
                    <h2 className="text-lg font-bold text-foreground sm:text-xl">
                      Table A: 5-Year Consolidated Financial Projection
                    </h2>
                    <p className="text-xs text-muted-foreground">
                      All figures in USD. Mid-year active averages; lean South African engineering
                      cost structures.
                    </p>
                  </div>
                </div>

                <div className="table-scroll mt-4 rounded-lg border border-border bg-card">
                  <table className="finance-table">
                    <thead>
                      <tr>
                        <th>Metric (USD)</th>
                        <th>Year 1</th>
                        <th>Year 2</th>
                        <th>Year 3</th>
                        <th>Year 4</th>
                        <th>Year 5</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="font-semibold text-primary">
                        <td>Gross Annual Revenue</td>
                        {plan.financials.revenue.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="text-muted-foreground">Cost of Revenue (COR)</td>
                        {plan.financials.cor.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr className="font-semibold">
                        <td>Gross Profit</td>
                        {plan.financials.grossProfit.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr className="font-medium text-emerald-600">
                        <td>Gross Margin %</td>
                        {plan.financials.grossMarginPct.map((v, i) => (
                          <td key={i}>{v.toFixed(1)}%</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="text-muted-foreground">Staffing Headcount</td>
                        {plan.financials.headcount.map((v, i) => (
                          <td key={i}>{v} Staff</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="text-muted-foreground">Baseline Core Payroll</td>
                        {plan.financials.payroll.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="text-muted-foreground">Other Operational Expenses</td>
                        {plan.financials.otherOpex.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr className="font-semibold">
                        <td>Total Operating Expenses (Opex)</td>
                        {plan.financials.totalOpex.map((v, i) => (
                          <td key={i}>${v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr className="bg-emerald-500/10 font-bold text-base">
                        <td>Net Operating Profit / (Loss)</td>
                        {plan.financials.netProfitLoss.map((v, i) => (
                          <td
                            key={i}
                            className={
                              v < 0 ? "text-rose-600" : "text-emerald-700 dark:text-emerald-400"
                            }
                          >
                            {v < 0
                              ? `(-$${Math.abs(v).toLocaleString()})`
                              : `+$${v.toLocaleString()}`}
                          </td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Chart 1: 5-Year Trajectory */}
              <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                <div className="mb-4">
                  <h3 className="text-base font-bold text-foreground sm:text-lg">
                    Chart 1: 5-Year Financial Trajectory & AI Lab Data Scaling
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Comparison of Annual Revenue, Total Operational Costs (COR + Opex), and Net
                    Operating Profit.
                  </p>
                </div>
                <div className="h-72 w-full sm:h-80">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={trajectoryData}>
                      <CartesianGrid strokeDasharray="3 3" opacity={0.25} />
                      <XAxis dataKey="name" fontSize={12} />
                      <YAxis fontSize={12} tickFormatter={(v) => `$${(v / 1000000).toFixed(1)}M`} />
                      <Tooltip
                        formatter={(val: number) => `$${val.toLocaleString()}`}
                        labelStyle={{ fontWeight: "bold" }}
                      />
                      <Legend />
                      <Bar dataKey="Revenue" fill={COLORS.blue} radius={[4, 4, 0, 0]} />
                      <Bar dataKey="Expenses" fill={COLORS.muted} radius={[4, 4, 0, 0]} />
                      <Bar dataKey="NetProfit" fill={COLORS.primary} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Charts 2 & 3: Revenue Stream Evolution Pie Charts */}
              <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
                <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                  <h3 className="text-base font-bold text-foreground">
                    Chart 2: Year 1 Revenue Stream Distribution ($30k)
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Early-stage SaaS footprint: 40% Councillors, 40% Providers, 20% Analytics.
                  </p>
                  <div className="mt-4 h-64 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={plan.financials.y1RevenueDistribution}
                          dataKey="value"
                          nameKey="name"
                          cx="50%"
                          cy="50%"
                          outerRadius={80}
                          label={({ name, percent }) =>
                            `${name.split(" ")[0]} ${(percent * 100).toFixed(0)}%`
                          }
                          fontSize={11}
                        >
                          {plan.financials.y1RevenueDistribution.map((_, index) => (
                            <Cell
                              key={`cell-${index}`}
                              fill={PIE_COLORS_Y1[index % PIE_COLORS_Y1.length]}
                            />
                          ))}
                        </Pie>
                        <Tooltip formatter={(val: number) => `${val}%`} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                  <h3 className="text-base font-bold text-foreground">
                    Chart 3: Year 5 Revenue Stream Distribution ($17.425M)
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    AI Data Refinery scale: 86.1% AI Data Broker sales ($15M).
                  </p>
                  <div className="mt-4 h-64 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={plan.financials.y5RevenueDistribution}
                          dataKey="value"
                          nameKey="name"
                          cx="50%"
                          cy="50%"
                          outerRadius={80}
                          label={({ name, percent }) =>
                            `${name.split(" ")[0]} ${(percent * 100).toFixed(1)}%`
                          }
                          fontSize={11}
                        >
                          {plan.financials.y5RevenueDistribution.map((_, index) => (
                            <Cell
                              key={`cell-${index}`}
                              fill={PIE_COLORS_Y5[index % PIE_COLORS_Y5.length]}
                            />
                          ))}
                        </Pie>
                        <Tooltip formatter={(val: number) => `${val}%`} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>

              {/* Chart 4 & Unit Economics Micro Analysis */}
              <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
                <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                  <h3 className="text-base font-bold text-foreground">
                    Chart 4: Unit Economics — Cost to Serve (CTS) Decay Curve
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Monthly ward-level cost drops from $18.00 to $3.25 as platform amortizes.
                  </p>
                  <div className="mt-4 h-64 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={plan.financials.unitEconomics.monthlyCostToServe}>
                        <CartesianGrid strokeDasharray="3 3" opacity={0.25} />
                        <XAxis dataKey="year" fontSize={11} />
                        <YAxis unit="$" domain={[0, 20]} fontSize={11} />
                        <Tooltip formatter={(val: number) => `$${val.toFixed(2)} / mo`} />
                        <Line
                          type="monotone"
                          dataKey="cost"
                          name="Cost to Serve (CTS)"
                          stroke={COLORS.blue}
                          strokeWidth={3}
                          dot={{ r: 5 }}
                        />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                  <h3 className="text-base font-bold text-foreground">
                    Single-Ward Microeconomics Analysis
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Evaluating profitability and unit economics at the individual paying ward
                    footprint.
                  </p>
                  <div className="mt-4 space-y-2.5 text-xs">
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">Target Ward Composition</span>
                      <strong className="text-foreground">1 Councillor + 0.2 Provider</strong>
                    </div>
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">Blended Ward Monthly Revenue</span>
                      <strong className="text-primary font-bold">$20.00 / mo ($240 / yr)</strong>
                    </div>
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">Variable AI Token Processing</span>
                      <strong className="text-foreground">$3.00 / mo (150 tix @ $0.02)</strong>
                    </div>
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">Amortized Server Overhead (Y1)</span>
                      <strong className="text-foreground">$15.00 / mo (100 wards)</strong>
                    </div>
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">Total Ward CTS (Y1 → Y5)</span>
                      <strong className="text-foreground">$18.00 → $3.25 / month</strong>
                    </div>
                    <div className="flex items-center justify-between border-b border-border/60 pb-2">
                      <span className="text-muted-foreground">
                        Ward Contribution Margin (Y1 → Y5)
                      </span>
                      <strong className="text-emerald-600 font-bold">10% → 83.8%</strong>
                    </div>
                    <div className="flex items-center justify-between pt-1">
                      <span className="text-muted-foreground">Customer Acquisition Cost (CAC)</span>
                      <strong className="text-primary font-bold">$45.00 (2.25 mo payback)</strong>
                    </div>
                  </div>
                </div>
              </div>

              {/* 5-Year Target Accounts Matrix */}
              <div className="rounded-lg border border-border bg-card p-4 sm:p-6">
                <h3 className="text-base font-bold text-foreground">
                  5-Year Accounts & Operational Scaling Matrix
                </h3>
                <div className="table-scroll mt-4">
                  <table className="finance-table text-xs">
                    <thead>
                      <tr>
                        <th>Account Tier</th>
                        <th>Year 1</th>
                        <th>Year 2</th>
                        <th>Year 3</th>
                        <th>Year 4</th>
                        <th>Year 5</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="font-semibold">Ward Councillors ($10/mo)</td>
                        {plan.financials.accounts.councillors.map((v, i) => (
                          <td key={i}>{v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="font-semibold">B2B Service Providers ($50/mo)</td>
                        {plan.financials.accounts.serviceProviders.map((v, i) => (
                          <td key={i}>{v.toLocaleString()}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="font-semibold">Partner Municipalities</td>
                        {plan.financials.accounts.municipalities.map((v, i) => (
                          <td key={i}>{v}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="font-semibold">Premium Advertisers ($1k/mo)</td>
                        {plan.financials.accounts.advertisers.map((v, i) => (
                          <td key={i}>{v}</td>
                        ))}
                      </tr>
                      <tr>
                        <td className="font-semibold">Data Analytics Clients</td>
                        {plan.financials.accounts.analyticsClients.map((v, i) => (
                          <td key={i}>{v}</td>
                        ))}
                      </tr>
                      <tr className="bg-emerald-500/10 font-bold text-primary">
                        <td>AI Data Brokers (Frontier Labs)</td>
                        {plan.financials.accounts.dataBrokers.map((v, i) => (
                          <td key={i}>{v === 1 ? "1 (micro1)" : v}</td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* STRATEGY & MODEL VIEW */}
          {view === "strategy" && (
            <div className="space-y-8">
              {/* Wedge Strategy Detail */}
              <div className="rounded-lg border border-border bg-card p-6 shadow-xs">
                <div className="max-w-2xl">
                  <span className="text-xs font-bold text-primary uppercase">
                    Market Positioning
                  </span>
                  <h2 className="mt-1 text-xl font-bold sm:text-2xl">
                    Triple-Pronged Wedge: Organic Ingestion to High-Yield Data
                  </h2>
                  <p className="mt-2 text-sm text-muted-foreground">
                    Public sector software usually stumbles on slow procurement cycles. Signal Desk
                    bypasses bureaucracy through bottom-up ward adoption, converting community
                    resolution into valuable AI datasets.
                  </p>
                </div>

                <div className="mt-6 grid grid-cols-1 gap-4 md:grid-cols-3">
                  <div className="rounded-lg border border-border bg-muted/30 p-5">
                    <span className="text-xs font-bold text-blue-600 uppercase">
                      Stage 1 • Low Friction Wedge
                    </span>
                    <h3 className="mt-2 text-base font-bold text-foreground">
                      Ward Councillor Dashboard
                    </h3>
                    <div className="mt-2 text-lg font-bold text-primary">$10 / month</div>
                    <p className="mt-2 text-xs text-muted-foreground leading-relaxed">
                      Councillors monitor local infrastructure issues in real-time. Fast personal
                      budget sign-off eliminates tender roadblocks and locks in ward-level
                      footprint.
                    </p>
                  </div>

                  <div className="rounded-lg border border-border bg-muted/30 p-5">
                    <span className="text-xs font-bold text-blue-600 uppercase">
                      Stage 2 • Workflow Optimization
                    </span>
                    <h3 className="mt-2 text-base font-bold text-foreground">
                      B2B Service Providers
                    </h3>
                    <div className="mt-2 text-lg font-bold text-primary">$50 / month</div>
                    <p className="mt-2 text-xs text-muted-foreground leading-relaxed">
                      Contracted plumbing, electrical, and roads teams use dispatch feeds to reduce
                      manual inspection overhead and track job resolution verified by residents.
                    </p>
                  </div>

                  <div className="rounded-lg border border-emerald-500/40 bg-emerald-500/10 p-5">
                    <span className="text-xs font-bold text-emerald-600 uppercase">
                      Stage 3 • Institutional Liquidation
                    </span>
                    <h3 className="mt-2 text-base font-bold text-foreground">
                      Frontier AI Lab Licensing
                    </h3>
                    <div className="mt-2 text-lg font-bold text-emerald-600">
                      $100k → $1.0M / package
                    </div>
                    <p className="mt-2 text-xs text-muted-foreground leading-relaxed">
                      End-to-end municipal triage, routing, and resolution graph datasets licensed
                      to Anthropic, Google, OpenAI, X/Cursor, and Meta for RL agentic workflows.
                    </p>
                  </div>
                </div>
              </div>

              {/* CivicRewards Framework */}
              <div className="rounded-lg border border-border bg-card p-6 shadow-xs">
                <span className="text-xs font-bold text-primary uppercase">
                  Community Incentive Engine
                </span>
                <h2 className="mt-1 text-xl font-bold sm:text-2xl">
                  CivicRewards Loyalty Framework
                </h2>
                <p className="mt-2 text-sm text-muted-foreground">
                  How Signal Desk generates high-quality verified data without increasing cash burn.
                </p>

                <div className="mt-6 grid grid-cols-1 gap-3 sm:grid-cols-4">
                  <div className="rounded-md border border-border bg-muted/20 p-4">
                    <div className="text-sm font-bold text-primary">01</div>
                    <h4 className="mt-1 text-sm font-bold">Verified Citizen Report</h4>
                    <p className="mt-2 text-xs text-muted-foreground">
                      Submitted via Telegram/web with photo and automated metadata capture.
                    </p>
                  </div>
                  <div className="rounded-md border border-border bg-muted/20 p-4">
                    <div className="text-sm font-bold text-primary">02</div>
                    <h4 className="mt-1 text-sm font-bold">Geolocation & Deduplication</h4>
                    <p className="mt-2 text-xs text-muted-foreground">
                      AI routing links incoming reports to existing canonical incidents.
                    </p>
                  </div>
                  <div className="rounded-md border border-border bg-muted/20 p-4">
                    <div className="text-sm font-bold text-primary">03</div>
                    <h4 className="mt-1 text-sm font-bold">CivicPoints Generation</h4>
                    <p className="mt-2 text-xs text-muted-foreground">
                      Tokens credited once field repairs are confirmed by community verification.
                    </p>
                  </div>
                  <div className="rounded-md border border-border bg-muted/20 p-4">
                    <div className="text-sm font-bold text-primary">04</div>
                    <h4 className="mt-1 text-sm font-bold">B2B Advertiser Redemption</h4>
                    <p className="mt-2 text-xs text-muted-foreground">
                      Redeemed at local merchants funded by advertisers paying $1,000/month.
                    </p>
                  </div>
                </div>
              </div>

              {/* AI Broker Monetization Roadmap */}
              <div className="rounded-lg border border-border bg-card p-6 shadow-xs">
                <span className="text-xs font-bold text-primary uppercase">
                  Data Asset Value Curve
                </span>
                <h2 className="mt-1 text-xl font-bold sm:text-2xl">
                  AI Broker & Frontier Lab Monetization Timeline
                </h2>
                <div className="mt-6 space-y-4">
                  <div className="flex flex-col gap-3 rounded-lg border border-border bg-muted/20 p-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <span className="text-xs font-bold text-blue-600">
                        Year 2 • Pilot Partnerships
                      </span>
                      <h4 className="text-base font-bold text-foreground">
                        Basic Operational Logs & Routing Histories
                      </h4>
                      <p className="text-xs text-muted-foreground">
                        1 Qualified broker (micro1) licensing initial triage datasets for model
                        fine-tuning.
                      </p>
                    </div>
                    <div className="text-left sm:text-right">
                      <strong className="text-lg font-bold text-primary">$100,000</strong>
                      <div className="text-xs text-muted-foreground">1 Broker package</div>
                    </div>
                  </div>

                  <div className="flex flex-col gap-3 rounded-lg border border-border bg-muted/20 p-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <span className="text-xs font-bold text-blue-600">
                        Year 3 • Deep Transactional Collections
                      </span>
                      <h4 className="text-base font-bold text-foreground">
                        Multi-Team Field Work Histories
                      </h4>
                      <p className="text-xs text-muted-foreground">
                        Deep incident lifecycle graphs across 20 partner municipalities delivered to
                        2 active brokers.
                      </p>
                    </div>
                    <div className="text-left sm:text-right">
                      <strong className="text-lg font-bold text-primary">$1,000,000</strong>
                      <div className="text-xs text-muted-foreground">2 Brokers @ $500k each</div>
                    </div>
                  </div>

                  <div className="flex flex-col gap-3 rounded-lg border border-emerald-500/40 bg-emerald-500/10 p-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <span className="text-xs font-bold text-emerald-600">
                        Years 4 & 5 • Proprietary Decision Workflows
                      </span>
                      <h4 className="text-base font-bold text-foreground">
                        Enterprise RL Training Datasets
                      </h4>
                      <p className="text-xs text-muted-foreground">
                        Distinctive operational workflows sold to Anthropic, Google, OpenAI,
                        X/Cursor, and Meta.
                      </p>
                    </div>
                    <div className="text-left sm:text-right">
                      <strong className="text-lg font-bold text-emerald-600">$15,000,000</strong>
                      <div className="text-xs text-muted-foreground">
                        Year 5: 3 Datasets × 5 Brokers
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ALL 15 SLIDES VIEW */}
          {view === "full" && (
            <div>
              <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <p className="text-sm text-muted-foreground">
                  Browse all 15 native slides formatted in the downloadable PDF and PowerPoint.
                </p>
                <DownloadButtons compact />
              </div>
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
                {plan.slides.map((s, i) => (
                  <article
                    key={i}
                    onClick={() => {
                      setSlide(i);
                      setView("overview");
                    }}
                    className="group cursor-pointer rounded-lg border border-border bg-card p-4 transition-all hover:border-primary hover:shadow-md"
                  >
                    <div className="relative aspect-video w-full overflow-hidden rounded border border-border/80 bg-slate-900/5">
                      <SlideRenderer slideIndex={i} />
                    </div>
                    <div className="mt-3">
                      <div className="flex items-center justify-between">
                        <span className="text-[11px] font-bold text-primary">
                          SLIDE {String(i + 1).padStart(2, "0")}
                        </span>
                        <span className="text-[10px] text-muted-foreground uppercase">
                          {s.section}
                        </span>
                      </div>
                      <h3 className="mt-1 text-xs font-bold text-foreground line-clamp-1 group-hover:text-primary">
                        {s.title}
                      </h3>
                      <p className="mt-1 text-[11px] text-muted-foreground line-clamp-2">
                        {s.summary}
                      </p>
                    </div>
                  </article>
                ))}
              </div>
            </div>
          )}

          {/* Footer Note */}
          <footer className="footer-note">
            <span>CONFIDENTIAL • SIGNAL DESK • OCTOBER 2026</span>
            <span>
              Management planning case • Cloud-Based Public Infrastructure Triage &
              Asset-Liquidation Platform
            </span>
          </footer>
        </main>
      </div>
    </div>
  );
}
