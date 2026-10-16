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
} from "lucide-react";
import { Button } from "@/components/ui/button";
import plan from "@/assets/plan-manifest.json";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Signal Desk — Strategic Business Plan & Financials" },
      {
        name: "description",
        content:
          "Review Signal Desk’s five-year business plan, financial projections and funding analysis. Download the editable Arial PowerPoint.",
      },
      { property: "og:title", content: "Signal Desk — Strategic Business Plan & Financials" },
      {
        property: "og:description",
        content:
          "Source-reconciled strategy, five-year financial statements and an editable PowerPoint download.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

const downloadPath = "/downloads/Signal_Desk_Business_Plan.pptx";
function Download({ compact = false }: { compact?: boolean }) {
  return (
    <Button asChild size={compact ? "default" : "lg"}>
      <a href={downloadPath} download="Signal_Desk_Business_Plan.pptx">
        <ArrowDownToLine />
        {compact ? "Download .pptx" : "Download PowerPoint"}
      </a>
    </Button>
  );
}

function Index() {
  const [view, setView] = useState("overview");
  const [slide, setSlide] = useState(0);
  const selected = plan.slides[slide];
  const imagePath = (i: number) => `/plan-preview/slide-${String(i + 1).padStart(2, "0")}.jpg`;
  return (
    <div className="workspace">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-symbol">
            <Radio size={20} />
          </span>
          Signal Desk
        </div>
        <div className="sidebar-label">INVESTOR DOCUMENTS</div>
        <nav aria-label="Business plan views">
          {[
            { id: "overview", label: "Business plan", icon: FileText },
            { id: "financials", label: "Financial projections", icon: ChartNoAxesCombined },
            { id: "full", label: "All slides", icon: Layers },
          ].map((item) => (
            <Button
              key={item.id}
              className="nav-item"
              variant={view === item.id ? "secondary" : "ghost"}
              onClick={() => setView(item.id)}
            >
              <item.icon />
              {item.label}
            </Button>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="flex items-center gap-2 text-foreground">
            <span className="brand-symbol">SD</span>
            <strong>Signal Desk</strong>
          </div>
          <p className="mt-3">
            Johannesburg, South Africa
            <br />
            October 2026
          </p>
        </div>
      </aside>
      <div className="min-w-0">
        <header className="topbar">
          <div className="flex items-center gap-2 text-muted-foreground">
            Documents
            <ChevronRight size={12} />
            <span className="text-foreground">Business plan</span>
          </div>
          <span>
            <span className="status-dot" />
            Prepared October 2026
          </span>
        </header>
        <main className="content">
          <div className="page-heading">
            <div>
              <div className="eyebrow">Strategy & capital investment</div>
              <h1>
                {view === "financials"
                  ? "Financial projections"
                  : view === "full"
                    ? "The complete business plan"
                    : "Signal Desk Business Plan"}
              </h1>
              <p>Five-year strategy, financial outlook and the path to regional scale.</p>
            </div>
            <Download />
          </div>
          {view === "overview" && (
            <>
              <div className="document-layout">
                <div>
                  <div className="preview-shell">
                    <div className="preview-label">
                      <span className="flex items-center gap-2">
                        <FileSliders size={15} />
                        BUSINESS PLAN / OCT 2026
                      </span>
                      <span className="text-muted-foreground">{plan.slideCount} slides</span>
                    </div>
                    <img
                      className="slide-image"
                      src={imagePath(slide)}
                      alt={`Slide ${slide + 1}: ${selected?.title}`}
                    />
                    <div className="slide-controls">
                      <span className="text-muted-foreground">
                        {String(slide + 1).padStart(2, "0")} / {plan.slideCount}
                      </span>
                      <span className="flex gap-2">
                        <Button
                          aria-label="Previous slide"
                          variant="ghost"
                          size="icon"
                          disabled={slide === 0}
                          onClick={() => setSlide((s) => s - 1)}
                        >
                          <ArrowLeft />
                        </Button>
                        <Button
                          aria-label="Next slide"
                          variant="ghost"
                          size="icon"
                          disabled={slide === plan.slideCount - 1}
                          onClick={() => setSlide((s) => s + 1)}
                        >
                          <ArrowRight />
                        </Button>
                      </span>
                    </div>
                  </div>
                  <h2 className="slide-title">{selected?.title}</h2>
                </div>
                <aside className="detail-panel">
                  <h2>Document details</h2>
                  <div className="detail-row">
                    <span>Format</span>
                    <strong>PowerPoint (.pptx)</strong>
                  </div>
                  <div className="detail-row">
                    <span>Length</span>
                    <strong>{plan.slideCount} slides</strong>
                  </div>
                  <div className="detail-row">
                    <span>Planning period</span>
                    <strong>5 years</strong>
                  </div>
                  <div className="detail-row">
                    <span>Currency</span>
                    <strong>US dollars</strong>
                  </div>
                  <div className="detail-divider" />
                  <div className="flex items-center gap-3">
                    <FileSliders size={39} className="format-icon" />
                    <div className="text-xs">
                      <strong>Editable PowerPoint</strong>
                      <p className="mt-1 text-muted-foreground">Arial · editable tables & chart</p>
                    </div>
                  </div>
                  <div className="detail-divider" />
                  <div className="flex items-center gap-2 text-xs text-primary">
                    <Check size={14} />
                    All three sources included
                  </div>
                  <p className="note mt-3">
                    Original prospectus retained in full. Source differences and financial
                    assumptions are disclosed.
                  </p>
                  <div className="mt-5">
                    <Download compact />
                  </div>
                </aside>
              </div>
              <h2 className="section-heading">Inside the business plan</h2>
              <div className="section-grid">
                {[
                  [
                    "01",
                    "Strategy & opportunity",
                    "Executive case, market, product, go-to-market and continental expansion.",
                  ],
                  [
                    "02",
                    "Financial projections",
                    "Five-year revenue, costs, profit & loss, cash flow, balance sheet and sensitivities.",
                  ],
                  [
                    "03",
                    "Capital & execution",
                    "Funding requirements, milestones, governance, risks and original-text appendix.",
                  ],
                ].map(([n, title, body]) => (
                  <article className="section-card" key={n}>
                    <span className="section-number">{n}</span>
                    <h3>{title}</h3>
                    <p>{body}</p>
                  </article>
                ))}
              </div>
            </>
          )}
          {view === "financials" && (
            <>
              <div className="metrics">
                <div className="metric">
                  <span>Year 5 source revenue</span>
                  <strong>$581k</strong>
                  <span>Paid-seat reference case</span>
                </div>
                <div className="metric">
                  <span>Peak loss funding</span>
                  <strong>$981k</strong>
                  <span>Before working capital</span>
                </div>
                <div className="metric">
                  <span>$100k initial runway</span>
                  <strong>4.8 months</strong>
                  <span>Even-burn reference</span>
                </div>
              </div>
              <div className="table-scroll">
                <table className="finance-table">
                  <thead>
                    <tr>
                      <th>US dollars</th>
                      {[1, 2, 3, 4, 5].map((y) => (
                        <th key={y}>Year {y}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      ["Source revenue", plan.financials.revenue],
                      ["Source EBITDA", plan.financials.ebitda],
                      ["Free-core revenue*", plan.financials.freeCoreRevenue],
                      ["Free-core EBITDA*", plan.financials.freeCoreEbitda],
                    ].map(([label, values]) => (
                      <tr key={String(label)}>
                        <td>{String(label)}</td>
                        {(values as number[]).map((v, i) => (
                          <td key={i}>
                            {v < 0
                              ? `(${Math.abs(v).toLocaleString("en-US")})`
                              : v.toLocaleString("en-US")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="note mt-5">
                *Councillor revenue removed; no sponsor replacement assumed. Other unvalidated
                source assumptions retained.
              </p>
              <p className="note mt-2">
                Illustrative cash funding: ${plan.financials.cashFunding.toLocaleString("en-US")}{" "}
                with 30-day receivables. Forecasts are not audited actuals or committed capital.
              </p>
              <h2 className="section-heading">Financial outlook</h2>
              <img
                src={imagePath(22)}
                alt="Editable financial trajectory chart preview"
                className="slide-image border border-border"
              />
            </>
          )}
          {view === "full" && (
            <div className="plan-list">
              {plan.slides.map((s, i) => (
                <article key={i}>
                  <img
                    src={imagePath(i)}
                    loading="lazy"
                    className="plan-thumb"
                    alt={`Slide ${i + 1}: ${s.title}`}
                  />
                  <p className="thumb-caption">
                    <span className="text-primary mr-2">{String(i + 1).padStart(2, "0")}</span>
                    {s.title}
                  </p>
                </article>
              ))}
            </div>
          )}
          <footer className="footer-note">
            <span>CONFIDENTIAL · SIGNAL DESK · OCTOBER 2026</span>
            <span>Management planning · Not audited · No KPMG or McKinsey affiliation</span>
          </footer>
        </main>
      </div>
    </div>
  );
}
