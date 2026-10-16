import pptxgen from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const DOWNLOAD_PATH = path.join(ROOT, "public", "downloads", "Signal_Desk_Business_Plan.pptx");
const MANIFEST_PATH = path.join(ROOT, "src", "assets", "plan-manifest.json");

const P = new pptxgen();
P.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 inches

const C = {
  navy: "162B3A",
  primary: "008A68",
  blue: "0068A0",
  cyan: "00A896",
  white: "FFFFFF",
  muted: "62717D",
  light: "F4F7F9",
  card: "FFFFFF",
  line: "D6E0E5",
  danger: "B74444",
  darkBg: "0E1A24",
};

// Data models from the business plan document
const financialModel = {
  years: ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"],
  grossRevenue: [30000, 315000, 1523000, 13261000, 17425000],
  costOfRevenue: [22500, 62100, 186990, 1361520, 1760500],
  grossProfit: [7500, 252900, 1336010, 11899480, 15664500],
  grossMarginPct: [25.0, 80.3, 87.7, 89.7, 89.9],
  headcount: [3, 4, 6, 9, 13],
  payroll: [45000, 70000, 120000, 200000, 310000],
  otherOpex: [12000, 20000, 35000, 60000, 95000],
  totalOpex: [57000, 90000, 155000, 260000, 405000],
  totalExpenses: [79500, 152100, 341990, 1621520, 2165500], // COR + Opex
  netProfitLoss: [-49500, 162900, 1181010, 11639480, 15259500],
  accounts: {
    councillors: [100, 300, 800, 1800, 3200],
    serviceProviders: [20, 50, 120, 250, 500],
    municipalities: [3, 8, 20, 45, 90],
    advertisers: [3, 8, 15, 30, 50],
    analyticsClients: [1, 3, 7, 15, 30],
    dataBrokers: [0, 1, 2, 4, 5],
  },
  y1RevenueDistribution: [
    { name: "Ward Councillors Subscriptions", value: 40.0, amount: 12000 },
    { name: "B2B Service Providers", value: 40.0, amount: 12000 },
    { name: "Data Analytics Clients", value: 20.0, amount: 6000 },
  ],
  y5RevenueDistribution: [
    { name: "AI Data Broker Asset Sales", value: 86.1, amount: 15000000 },
    { name: "Municipal Enterprise Tiers", value: 6.2, amount: 1080000 },
    { name: "Data Analytics Clients", value: 2.1, amount: 365000 },
    { name: "Premium Advertisers Network", value: 1.7, amount: 300000 },
    { name: "B2B Service Providers", value: 1.4, amount: 245000 },
    { name: "Ward Councillors Subscriptions", value: 0.2, amount: 35000 },
    { name: "Others / Ancillary", value: 2.3, amount: 400000 },
  ],
  unitEconomics: {
    monthlyCostToServe: [
      { year: "Year 1", cost: 18.0 },
      { year: "Year 2", cost: 6.2 },
      { year: "Year 3", cost: 4.8 },
      { year: "Year 4", cost: 3.8 },
      { year: "Year 5", cost: 3.25 },
    ],
    blendedWardRevenue: 20.0,
    annualWardRevenue: 240.0,
    variableTicketCost: 0.02,
    monthlyTickets: 150,
    variableTokenCost: 3.0,
    serverOverheadY1: 15.0,
    totalCtsY1: 18.0,
    totalCtsY5: 3.25,
    cac: 45.0,
  },
  runway: {
    seedCapital: 100000,
    sunkAssets: 30000,
    netCashWorkingCapital: 70000,
    annualCloudBaseline: 6000, // $500/mo
    annualClaudeApiBaseline: 12000, // $1,000/mo
    fixedInfrastructureAnnual: 18000,
    y1Opex: 57000,
    totalY1BaselineOutflows: 75000,
    monthsRunwayMin: 18,
    monthsRunwayMax: 24,
  },
};

const slideRegistry: Array<{ title: string; section: string; summary: string }> = [];

function addSlideFrame(title: string, section: string, sub = "", dark = false): pptxgen.Slide {
  const slide = P.addSlide();
  slide.background = { color: dark ? C.darkBg : C.white };

  // Category tag
  slide.addText(`SIGNAL DESK  /  ${section.toUpperCase()}`, {
    x: 0.6,
    y: 0.4,
    w: 12.0,
    h: 0.3,
    fontSize: 11,
    bold: true,
    color: dark ? C.primary : C.blue,
    fontFace: "Arial",
  });

  // Main Slide Title
  slide.addText(title, {
    x: 0.6,
    y: 0.75,
    w: 12.1,
    h: 0.8,
    fontSize: 26,
    bold: true,
    color: dark ? C.white : C.navy,
    fontFace: "Arial",
  });

  // Subtitle / context
  if (sub) {
    slide.addText(sub, {
      x: 0.6,
      y: 1.55,
      w: 12.1,
      h: 0.5,
      fontSize: 14,
      color: dark ? C.light : C.muted,
      fontFace: "Arial",
    });
  }

  // Footer
  slide.addText(
    "BUSINESS PLAN • OCTOBER 2026  |  CONFIDENTIAL • MANAGEMENT PLANNING, NOT AUDITED",
    {
      x: 0.6,
      y: 7.0,
      w: 10.5,
      h: 0.3,
      fontSize: 9,
      color: dark ? C.muted : C.muted,
      fontFace: "Arial",
    },
  );

  slide.addText(`${String(slideRegistry.length + 1).padStart(2, "0")}`, {
    x: 12.2,
    y: 7.0,
    w: 0.6,
    h: 0.3,
    fontSize: 10,
    align: "right",
    bold: true,
    color: dark ? C.white : C.navy,
    fontFace: "Arial",
  });

  slideRegistry.push({ title, section, summary: sub });
  return slide;
}

// -------------------------------------------------------------
// SLIDE 1: Cover Slide
// -------------------------------------------------------------
const s1 = addSlideFrame(
  "SIGNAL DESK",
  "Strategic Business Plan",
  "Signal Desk Business Plan",
  true,
);
s1.addText(
  "Strategic Framework for an Autonomous Orchestration Layer for Municipal Infrastructure Platform.",
  {
    x: 0.6,
    y: 2.2,
    w: 12.0,
    h: 1.1,
    fontSize: 26,
    bold: true,
    color: C.white,
    fontFace: "Arial",
  },
);
s1.addText(
  "AI-powered data routing and smart workflows bridging citizen intelligence with municipal action.",
  {
    x: 0.6,
    y: 3.4,
    w: 11.5,
    h: 0.8,
    fontSize: 16,
    color: C.cyan,
    fontFace: "Arial",
  },
);
s1.addText(
  "A high-margin B2G and B2B public infrastructure data ecosystem capturing real-time municipal triage logs, citizen routing, and multi-team field dispatch dynamics to license proprietary operational datasets to frontier AI engineering labs.",
  {
    x: 0.6,
    y: 4.3,
    w: 10.5,
    h: 0.9,
    fontSize: 14,
    color: C.light,
    fontFace: "Arial",
  },
);

// Key badges on cover
const badges = [
  { label: "SEED RUNWAY", val: "$100k (18–24 mo)" },
  { label: "YEAR 5 REVENUE", val: "$17.425M" },
  { label: "YEAR 5 PROFIT", val: "+$15.260M" },
  { label: "GROSS MARGIN", val: "89.9%" },
];
badges.forEach((b, idx) => {
  const x = 0.6 + idx * 3.0;
  s1.addShape(P.ShapeType.rect, {
    x,
    y: 5.6,
    w: 2.8,
    h: 0.9,
    fill: { color: "1A2E3D" },
    line: { color: C.primary, width: 1 },
  });
  s1.addText(b.label, {
    x: x + 0.15,
    y: 5.7,
    w: 2.5,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: C.cyan,
    fontFace: "Arial",
  });
  s1.addText(b.val, {
    x: x + 0.15,
    y: 5.95,
    w: 2.5,
    h: 0.45,
    fontSize: 18,
    bold: true,
    color: C.white,
    fontFace: "Arial",
  });
});

// -------------------------------------------------------------
// SLIDE 2: Executive Summary & Data Asset Thesis
// -------------------------------------------------------------
const s2 = addSlideFrame(
  "Executive Summary: High-Margin Infrastructure Data Ecosystem",
  "Executive Summary",
  "Commercial deployment of a cloud-based public infrastructure triage & asset-liquidation platform.",
);

const pillars = [
  {
    title: "Proprietary Data Generation",
    desc: "By capturing real-time municipal triage logs, citizen utility routing patterns, and multi-team field dispatch dynamics, the platform creates a highly proprietary operational dataset unavailable in open public repositories.",
    tag: "Source Layer",
  },
  {
    title: "Frontier AI Lab Licensing",
    desc: "Beginning in Year 2, processed structural datasets are licensed directly to frontier AI labs (Anthropic, Google, OpenAI, X/Cursor, Meta) for specialized neural networks, agentic workflows, and Reinforcement Learning (RL) training.",
    tag: "High-Margin Scale",
  },
  {
    title: "Capital Efficient Runway",
    desc: "A $100,000 initial seed round funds 18 to 24 months of core runtime with a lean $75,000 Year 1 total baseline outflow, enabling proof of unit economics and data refinery scale before raising institutional capital.",
    tag: "Runway Security",
  },
];

pillars.forEach((p, idx) => {
  const x = 0.6 + idx * 4.05;
  s2.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 3.9,
    h: 4.3,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s2.addText(p.tag, {
    x: x + 0.25,
    y: 2.55,
    w: 3.4,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: C.blue,
    fontFace: "Arial",
  });
  s2.addText(p.title, {
    x: x + 0.25,
    y: 2.9,
    w: 3.4,
    h: 0.7,
    fontSize: 18,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s2.addText(p.desc, {
    x: x + 0.25,
    y: 3.75,
    w: 3.4,
    h: 2.6,
    fontSize: 13,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 18,
  });
});

// -------------------------------------------------------------
// SLIDE 3: Financial Feasibility & Seed Runway Allocation
// -------------------------------------------------------------
const s3 = addSlideFrame(
  "Financial Feasibility & Seed Runway Allocation ($100k)",
  "Runway Analysis",
  "$100,000 seed round funds roughly 18 to 24 months of core runtime.",
);

const runwayItems = [
  {
    label: "Sunk Assets (Historical R&D)",
    amount: "$30,000",
    detail:
      "Prior software design iterations and foundational engine engineering capitalised as a sunk setup asset.",
    iconColor: C.muted,
  },
  {
    label: "Net Working Cash Position",
    amount: "$70,000",
    detail:
      "Initial working capital deployed directly into active operational runtime, cloud hosting, and outreach.",
    iconColor: C.primary,
  },
  {
    label: "Fixed Core Infrastructure",
    amount: "$18,000 / yr",
    detail:
      "Baseline cloud server deployments ($500/mo) and Claude API workflow layers ($1,000/mo).",
    iconColor: C.blue,
  },
  {
    label: "Total Year 1 Baseline Outflows",
    amount: "$75,000 / yr",
    detail:
      "Combined with a lean Year 1 Operating Expense framework of $57,000 ($45k payroll + $12k other opex).",
    iconColor: C.navy,
  },
];

runwayItems.forEach((item, idx) => {
  const y = 2.2 + idx * 1.15;
  s3.addShape(P.ShapeType.rect, {
    x: 0.6,
    y,
    w: 12.1,
    h: 1.0,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s3.addText(item.label, {
    x: 0.9,
    y: y + 0.15,
    w: 4.5,
    h: 0.35,
    fontSize: 14,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s3.addText(item.amount, {
    x: 5.5,
    y: y + 0.15,
    w: 2.2,
    h: 0.35,
    fontSize: 16,
    bold: true,
    color: item.iconColor,
    fontFace: "Arial",
  });
  s3.addText(item.detail, {
    x: 7.8,
    y: y + 0.15,
    w: 4.6,
    h: 0.7,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
  });
});

s3.addShape(P.ShapeType.rect, {
  x: 0.6,
  y: 6.0,
  w: 12.1,
  h: 0.7,
  fill: { color: "E6F4EA" },
  line: { color: C.primary, width: 1 },
});
s3.addText(
  "Key Takeaway: The $100,000 seed round provides 18–24 months of runway to build proprietary data assets and secure sustainable commercial contracts before follow-on expansion capital.",
  {
    x: 0.9,
    y: 6.15,
    w: 11.5,
    h: 0.4,
    fontSize: 12,
    bold: true,
    color: C.primary,
    fontFace: "Arial",
  },
);

// -------------------------------------------------------------
// SLIDE 4: Market Positioning & Triple-Pronged GTM Strategy
// -------------------------------------------------------------
const s4 = addSlideFrame(
  "Market Positioning & Go-To-Market (GTM) Strategy",
  "GTM Strategy",
  "Circumventing traditional public sector procurement friction via a bottom-up adoption wedge.",
);

const stages = [
  {
    stage: "Stage 1",
    name: "Bottom-Up Councillor Wedge",
    price: "$10 / month",
    friction: "Low friction organic footprint",
    desc: "Direct outreach loops introduce individual ward councillors to dedicated utility tracking dashboards to monitor local service requests without heavy municipal procurement cycles.",
  },
  {
    stage: "Stage 2",
    name: "B2B Service Provider Onboarding",
    price: "$50 / month",
    friction: "Workflow optimization & operational link",
    desc: "Contracted external municipal vendors (plumbing, road maintenance, electrical repair crews) use the interface for structured job dispatches, lowering verification & inspection overhead.",
  },
  {
    stage: "Stage 3",
    name: "Data Broker Asset Liquidation",
    price: "$100k → $1.0M per package",
    friction: "Institutional dataset licensing to AI Labs",
    desc: "As data deepens, the platform transitions into a high-tier data refinery. Anonymized municipal incident, triage, and resolution datasets are packaged and sold to frontier AI labs.",
  },
];

stages.forEach((st, idx) => {
  const x = 0.6 + idx * 4.05;
  s4.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 3.9,
    h: 4.4,
    fill: { color: C.light },
    line: { color: idx === 2 ? C.primary : C.line, width: idx === 2 ? 2 : 1 },
  });
  s4.addText(`${st.stage.toUpperCase()} • ${st.friction}`, {
    x: x + 0.25,
    y: 2.5,
    w: 3.4,
    h: 0.3,
    fontSize: 10,
    bold: true,
    color: idx === 2 ? C.primary : C.blue,
    fontFace: "Arial",
  });
  s4.addText(st.name, {
    x: x + 0.25,
    y: 2.85,
    w: 3.4,
    h: 0.7,
    fontSize: 17,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s4.addText(`Pricing: ${st.price}`, {
    x: x + 0.25,
    y: 3.6,
    w: 3.4,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s4.addText(st.desc, {
    x: x + 0.25,
    y: 4.0,
    w: 3.4,
    h: 2.4,
    fontSize: 12,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 18,
  });
});

// -------------------------------------------------------------
// SLIDE 5: CivicRewards Loyalty Framework
// -------------------------------------------------------------
const s5 = addSlideFrame(
  "CivicRewards Loyalty Framework",
  "Product Architecture",
  "Maintaining reporting volume and verified field data without increasing operational service costs.",
);

const loopSteps = [
  {
    step: "01",
    title: "Citizen Verified Report",
    desc: "Citizen submits real-time infrastructure gap via Telegram / web.",
  },
  {
    step: "02",
    title: "Geolocation & De-dupe",
    desc: "System validates GPS coordinates, cross-references duplicates, links to cases.",
  },
  {
    step: "03",
    title: "Token Allocation (CivicPoints)",
    desc: "CivicPoints awarded upon independent community resolution validation.",
  },
  {
    step: "04",
    title: "Partner B2B Redemption",
    desc: "Points redeemed for discounts funded by Premium B2B Advertisers ($1k/mo).",
  },
];

loopSteps.forEach((st, idx) => {
  const x = 0.6 + idx * 3.05;
  s5.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 2.9,
    h: 2.6,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s5.addText(st.step, {
    x: x + 0.2,
    y: 2.5,
    w: 2.5,
    h: 0.35,
    fontSize: 20,
    bold: true,
    color: C.primary,
    fontFace: "Arial",
  });
  s5.addText(st.title, {
    x: x + 0.2,
    y: 2.9,
    w: 2.5,
    h: 0.55,
    fontSize: 14,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s5.addText(st.desc, {
    x: x + 0.2,
    y: 3.5,
    w: 2.5,
    h: 1.2,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 16,
  });
});

// Dual column callouts on bottom
s5.addShape(P.ShapeType.rect, {
  x: 0.6,
  y: 5.15,
  w: 5.9,
  h: 1.5,
  fill: { color: C.white },
  line: { color: C.line, width: 1 },
});
s5.addText("Validation Engine", {
  x: 0.85,
  y: 5.3,
  w: 5.4,
  h: 0.3,
  fontSize: 13,
  bold: true,
  color: C.navy,
  fontFace: "Arial",
});
s5.addText(
  "Citizens earn loyalty tokens ('CivicPoints') when they log verified public service gaps that are subsequently validated by community resolution engines, ensuring high signal-to-noise ratio.",
  {
    x: 0.85,
    y: 5.65,
    w: 5.4,
    h: 0.8,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
  },
);

s5.addShape(P.ShapeType.rect, {
  x: 6.8,
  y: 5.15,
  w: 5.9,
  h: 1.5,
  fill: { color: C.white },
  line: { color: C.line, width: 1 },
});
s5.addText("Loop Monetization (Zero Cash Burn)", {
  x: 7.05,
  y: 5.3,
  w: 5.4,
  h: 0.3,
  fontSize: 13,
  bold: true,
  color: C.navy,
  fontFace: "Arial",
});
s5.addText(
  "Points are redeemable for localized value-added discounts. These incentives are funded by Premium B2B Advertisers paying $1,000/month, keeping customer engagement high without creating direct company cash burn.",
  {
    x: 7.05,
    y: 5.65,
    w: 5.4,
    h: 0.8,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
  },
);

// -------------------------------------------------------------
// SLIDE 6: Data Refinement & AI Broker Monetization Timeline
// -------------------------------------------------------------
const s6 = addSlideFrame(
  "Data Refinement & AI Broker Monetization Timeline",
  "Data Asset Refinery",
  "Data assets increase in valuation as transactional complexity and running history accumulate.",
);

const timeline = [
  {
    period: "Year 2",
    tier: "Qualified Partnerships",
    buyer: "1 Broker (micro1)",
    pricing: "$100,000 / package",
    desc: "The platform licenses basic operational logs, citizen routing patterns, and triage histories under pilot data partnership programs.",
  },
  {
    period: "Year 3",
    tier: "Large Data Collections",
    buyer: "2 Brokers",
    pricing: "$500,000 / package",
    desc: "Capturing ongoing, multi-team, deep transactional logs and work histories across 20 municipalities. Scales to $1.0M total revenue.",
  },
  {
    period: "Years 4 & 5",
    tier: "Highly Unique / Proprietary",
    buyer: "4 Brokers (Y4) → 5 Brokers (Y5)",
    pricing: "$1,000,000 / package",
    desc: "Distinctive operational workflows involving complex human decisions, triage, and escalations. Each of 3 core partner municipalities produces 1 distinct proprietary dataset sold to Anthropic, Google, OpenAI, X/Cursor, Meta ($15M Y5 revenue).",
  },
];

timeline.forEach((tl, idx) => {
  const x = 0.6 + idx * 4.05;
  s6.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 3.9,
    h: 4.4,
    fill: { color: C.light },
    line: { color: idx === 2 ? C.primary : C.line, width: idx === 2 ? 2 : 1 },
  });
  s6.addText(`${tl.period} • ${tl.tier}`, {
    x: x + 0.25,
    y: 2.5,
    w: 3.4,
    h: 0.3,
    fontSize: 11,
    bold: true,
    color: idx === 2 ? C.primary : C.blue,
    fontFace: "Arial",
  });
  s6.addText(tl.pricing, {
    x: x + 0.25,
    y: 2.85,
    w: 3.4,
    h: 0.45,
    fontSize: 20,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s6.addText(`Target Buyers: ${tl.buyer}`, {
    x: x + 0.25,
    y: 3.4,
    w: 3.4,
    h: 0.35,
    fontSize: 12,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s6.addText(tl.desc, {
    x: x + 0.25,
    y: 3.85,
    w: 3.4,
    h: 2.6,
    fontSize: 12,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 18,
  });
});

// -------------------------------------------------------------
// SLIDE 7: 5-Year Operational & Accounts Scaling Targets
// -------------------------------------------------------------
const s7 = addSlideFrame(
  "5-Year Operational & Accounts Scaling Targets",
  "Scaling Milestones",
  "Account growth across all tiers driving software and data flywheel.",
);

const accountTableData: pptxgen.TableCell[][] = [
  [
    { text: "Account / Segment", options: { bold: true, fill: { color: C.navy }, color: C.white } },
    {
      text: "Year 1",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 2",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 3",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 4",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 5",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
  ],
  [
    { text: "Ward Councillors ($10/mo)", options: { bold: true } },
    { text: "100", options: { align: "right" } },
    { text: "300", options: { align: "right" } },
    { text: "800", options: { align: "right" } },
    { text: "1,800", options: { align: "right" } },
    { text: "3,200", options: { align: "right", bold: true } },
  ],
  [
    { text: "B2B Service Providers ($50/mo)", options: { bold: true } },
    { text: "20", options: { align: "right" } },
    { text: "50", options: { align: "right" } },
    { text: "120", options: { align: "right" } },
    { text: "250", options: { align: "right" } },
    { text: "500", options: { align: "right", bold: true } },
  ],
  [
    { text: "Partner Municipalities", options: { bold: true } },
    { text: "3", options: { align: "right" } },
    { text: "8", options: { align: "right" } },
    { text: "20", options: { align: "right" } },
    { text: "45", options: { align: "right" } },
    { text: "90", options: { align: "right", bold: true } },
  ],
  [
    { text: "Premium Advertisers ($1k/mo)", options: { bold: true } },
    { text: "3", options: { align: "right" } },
    { text: "8", options: { align: "right" } },
    { text: "15", options: { align: "right" } },
    { text: "30", options: { align: "right" } },
    { text: "50", options: { align: "right", bold: true } },
  ],
  [
    { text: "Data Analytics Clients", options: { bold: true } },
    { text: "1", options: { align: "right" } },
    { text: "3", options: { align: "right" } },
    { text: "7", options: { align: "right" } },
    { text: "15", options: { align: "right" } },
    { text: "30", options: { align: "right", bold: true } },
  ],
  [
    { text: "AI Data Brokers (Frontier Labs)", options: { bold: true, color: C.primary } },
    { text: "0", options: { align: "right" } },
    { text: "1 (micro1)", options: { align: "right" } },
    { text: "2", options: { align: "right" } },
    { text: "4", options: { align: "right" } },
    { text: "5 (Tier 1)", options: { align: "right", bold: true, color: C.primary } },
  ],
];

s7.addTable(accountTableData, {
  x: 0.6,
  y: 2.3,
  w: 12.1,
  h: 3.5,
  colW: [4.1, 1.6, 1.6, 1.6, 1.6, 1.6],
  fontSize: 12,
  fontFace: "Arial",
  border: { color: C.line, pt: 1 },
  autoPage: false,
});

s7.addText(
  "Notes: Year 4 & 5 data brokers represent enterprise licenses to Anthropic, Google, OpenAI, X/Cursor, and Meta under structured data partnership programs.",
  {
    x: 0.6,
    y: 6.1,
    w: 12.1,
    h: 0.5,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
  },
);

// -------------------------------------------------------------
// SLIDE 8: Table A - Pro-Forma Income Statement (P&L Financial Model)
// -------------------------------------------------------------
const s8 = addSlideFrame(
  "Table A: 5-Year Consolidated Financial Projection (P&L)",
  "Financial Model",
  "All figures modeled in USD. Revenue metrics represent mid-year active customer averages.",
);

const plTableData: pptxgen.TableCell[][] = [
  [
    { text: "Metric (USD)", options: { bold: true, fill: { color: C.navy }, color: C.white } },
    {
      text: "Year 1",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 2",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 3",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 4",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
    {
      text: "Year 5",
      options: { bold: true, fill: { color: C.navy }, color: C.white, align: "right" },
    },
  ],
  [
    { text: "Gross Annual Revenue", options: { bold: true, color: C.blue } },
    { text: "$30,000", options: { align: "right", bold: true } },
    { text: "$315,000", options: { align: "right", bold: true } },
    { text: "$1,523,000", options: { align: "right", bold: true } },
    { text: "$13,261,000", options: { align: "right", bold: true } },
    { text: "$17,425,000", options: { align: "right", bold: true, color: C.blue } },
  ],
  [
    { text: "Cost of Revenue (COR)", options: { bold: false } },
    { text: "$22,500", options: { align: "right" } },
    { text: "$62,100", options: { align: "right" } },
    { text: "$186,990", options: { align: "right" } },
    { text: "$1,361,520", options: { align: "right" } },
    { text: "$1,760,500", options: { align: "right" } },
  ],
  [
    { text: "Gross Profit", options: { bold: true } },
    { text: "$7,500", options: { align: "right", bold: true } },
    { text: "$252,900", options: { align: "right", bold: true } },
    { text: "$1,336,010", options: { align: "right", bold: true } },
    { text: "$11,899,480", options: { align: "right", bold: true } },
    { text: "$15,664,500", options: { align: "right", bold: true } },
  ],
  [
    { text: "Gross Margin %", options: { bold: true, color: C.primary } },
    { text: "25.0%", options: { align: "right", bold: true } },
    { text: "80.3%", options: { align: "right", bold: true } },
    { text: "87.7%", options: { align: "right", bold: true } },
    { text: "89.7%", options: { align: "right", bold: true } },
    { text: "89.9%", options: { align: "right", bold: true, color: C.primary } },
  ],
  [
    { text: "Staffing Headcount", options: { bold: false } },
    { text: "3 Staff", options: { align: "right" } },
    { text: "4 Staff", options: { align: "right" } },
    { text: "6 Staff", options: { align: "right" } },
    { text: "9 Staff", options: { align: "right" } },
    { text: "13 Staff", options: { align: "right" } },
  ],
  [
    { text: "Baseline Core Payroll", options: { bold: false } },
    { text: "$45,000", options: { align: "right" } },
    { text: "$70,000", options: { align: "right" } },
    { text: "$120,000", options: { align: "right" } },
    { text: "$200,000", options: { align: "right" } },
    { text: "$310,000", options: { align: "right" } },
  ],
  [
    { text: "Other Operational Expenses", options: { bold: false } },
    { text: "$12,000", options: { align: "right" } },
    { text: "$20,000", options: { align: "right" } },
    { text: "$35,000", options: { align: "right" } },
    { text: "$60,000", options: { align: "right" } },
    { text: "$95,000", options: { align: "right" } },
  ],
  [
    { text: "Total Operating Expenses (Opex)", options: { bold: true } },
    { text: "$57,000", options: { align: "right", bold: true } },
    { text: "$90,000", options: { align: "right", bold: true } },
    { text: "$155,000", options: { align: "right", bold: true } },
    { text: "$260,000", options: { align: "right", bold: true } },
    { text: "$405,000", options: { align: "right", bold: true } },
  ],
  [
    {
      text: "Net Operating Profit / (Loss)",
      options: { bold: true, color: C.navy, fill: { color: "EAF4FB" } },
    },
    {
      text: "($49,500)",
      options: { align: "right", bold: true, color: C.danger, fill: { color: "EAF4FB" } },
    },
    {
      text: "+$162,900",
      options: { align: "right", bold: true, color: C.primary, fill: { color: "EAF4FB" } },
    },
    {
      text: "+$1,181,010",
      options: { align: "right", bold: true, color: C.primary, fill: { color: "EAF4FB" } },
    },
    {
      text: "+$11,639,480",
      options: { align: "right", bold: true, color: C.primary, fill: { color: "EAF4FB" } },
    },
    {
      text: "+$15,259,500",
      options: { align: "right", bold: true, color: C.primary, fill: { color: "EAF4FB" } },
    },
  ],
];

s8.addTable(plTableData, {
  x: 0.6,
  y: 2.2,
  w: 12.1,
  h: 4.5,
  colW: [3.8, 1.66, 1.66, 1.66, 1.66, 1.66],
  fontSize: 11,
  fontFace: "Arial",
  border: { color: C.line, pt: 1 },
  autoPage: false,
});

// -------------------------------------------------------------
// SLIDE 9: Operational Cost Architecture
// -------------------------------------------------------------
const s9 = addSlideFrame(
  "Operational Cost Architecture & Lean Delivery Framework",
  "Cost Discipline",
  "Structural efficiency ensures operational expenses remain below 2.5% of revenue by Year 5.",
);

s9.addShape(P.ShapeType.rect, {
  x: 0.6,
  y: 2.3,
  w: 5.9,
  h: 4.4,
  fill: { color: C.light },
  line: { color: C.line, width: 1 },
});
s9.addText("Cost of Revenue (COR) Architecture", {
  x: 0.9,
  y: 2.55,
  w: 5.3,
  h: 0.35,
  fontSize: 16,
  bold: true,
  color: C.navy,
  fontFace: "Arial",
});
s9.addText(
  "• Baseline Infrastructure: $18,000 annual fixed cost comprising cloud hosting ($500/mo) and Claude API connectivity ($1,000/mo).\n\n• Variable Platform Usage: Drops systematically from 15% in Year 1 down to 10% in Year 5 as server efficiency, ticket batching, and cache hit rates scale.\n\n• High Operational Leverage: Data asset extraction requires near-zero marginal reproduction cost once raw municipal dispatches are normalized.",
  {
    x: 0.9,
    y: 3.1,
    w: 5.3,
    h: 3.3,
    fontSize: 13,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 20,
  },
);

s9.addShape(P.ShapeType.rect, {
  x: 6.8,
  y: 2.3,
  w: 5.9,
  h: 4.4,
  fill: { color: C.light },
  line: { color: C.line, width: 1 },
});
s9.addText("Operating Expense (Opex) Discipline", {
  x: 7.1,
  y: 2.55,
  w: 5.3,
  h: 0.35,
  fontSize: 16,
  bold: true,
  color: C.navy,
  fontFace: "Arial",
});
s9.addText(
  "• Lean Engineering Payroll: Capitalizes on South African technical talent hubs ($45k Year 1 -> $310k Year 5 across 13 engineering and operations staff).\n\n• Below-Market Founder Compensation: Early founder salaries held strictly below market benchmarks, aligning cash preservation with equity value creation.\n\n• Strict Travel & Procurement Cap: Zero speculative municipal lobbying or travel; spend strictly bounded inside the 'Other Opex' budget envelope ($12k -> $95k).",
  {
    x: 7.1,
    y: 3.1,
    w: 5.3,
    h: 3.3,
    fontSize: 13,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 20,
  },
);

// -------------------------------------------------------------
// SLIDE 10: Chart 1 - 5-Year Financial Trajectory & AI Lab Data Scaling
// -------------------------------------------------------------
const s10 = addSlideFrame(
  "Chart 1: 5-Year Financial Trajectory & AI Lab Data Scaling",
  "Financial Projections",
  "Exponential top-line revenue growth scaling profits while keeping operational headcount lean.",
);

// Chart: Bar chart of Revenue, Total Expenses, Net Operating Profit
const trajectoryChartData = [
  {
    name: "Annual Revenue",
    labels: financialModel.years,
    values: financialModel.grossRevenue.map((v) => v / 1000), // in $k
  },
  {
    name: "Total Expenses (COR + Opex)",
    labels: financialModel.years,
    values: financialModel.totalExpenses.map((v) => v / 1000),
  },
  {
    name: "Net Operating Profit / (Loss)",
    labels: financialModel.years,
    values: financialModel.netProfitLoss.map((v) => v / 1000),
  },
];

s10.addChart(P.ChartType.bar, trajectoryChartData, {
  x: 0.6,
  y: 2.2,
  w: 12.1,
  h: 4.5,
  barGrouping: "clustered",
  barDir: "col",
  chartColors: [C.blue, C.muted, C.primary],
  showValue: false,
  showLegend: true,
  legendPos: "t",
  valAxisTitle: "Amount ($ Thousands USD)",
  catAxisTitle: "Timeline Horizon",
});

// -------------------------------------------------------------
// SLIDE 11: Chart 2 & 3 - Revenue Stream Distribution Shift
// -------------------------------------------------------------
const s11 = addSlideFrame(
  "Charts 2 & 3: Revenue Stream Evolution (Year 1 vs. Year 5)",
  "Revenue Mix",
  "Transition from early B2G/B2B software subscriptions to global AI data asset refinery.",
);

// Pie 1: Year 1
const y1PieData = [
  {
    name: "Year 1 Revenue ($30k)",
    labels: financialModel.y1RevenueDistribution.map((d) => d.name),
    values: financialModel.y1RevenueDistribution.map((d) => d.value),
  },
];
s11.addChart(P.ChartType.pie, y1PieData, {
  x: 0.6,
  y: 2.3,
  w: 5.8,
  h: 3.8,
  chartColors: [C.blue, C.primary, C.cyan],
  showPercent: true,
  showLegend: true,
  legendPos: "b",
});
s11.addText(
  "Year 1: Software Seats ($30k Total)\n40% Councillors | 40% Providers | 20% Analytics",
  {
    x: 0.6,
    y: 6.2,
    w: 5.8,
    h: 0.6,
    fontSize: 11,
    bold: true,
    align: "center",
    color: C.navy,
    fontFace: "Arial",
  },
);

// Pie 2: Year 5
const y5PieData = [
  {
    name: "Year 5 Revenue ($17.425M)",
    labels: financialModel.y5RevenueDistribution.map((d) => d.name),
    values: financialModel.y5RevenueDistribution.map((d) => d.value),
  },
];
s11.addChart(P.ChartType.pie, y5PieData, {
  x: 6.9,
  y: 2.3,
  w: 5.8,
  h: 3.8,
  chartColors: [C.primary, C.blue, C.cyan, "E3B341", "E76F51", C.muted, "8E9AAF"],
  showPercent: true,
  showLegend: true,
  legendPos: "b",
});
s11.addText("Year 5: Data Refinery ($17.425M Total)\n86.1% AI Data Broker Asset Sales ($15.0M)", {
  x: 6.9,
  y: 6.2,
  w: 5.8,
  h: 0.6,
  fontSize: 11,
  bold: true,
  align: "center",
  color: C.primary,
  fontFace: "Arial",
});

// -------------------------------------------------------------
// SLIDE 12: Chart 4 - Unit Economics & Cost to Serve (CTS) Decay Curve
// -------------------------------------------------------------
const s12 = addSlideFrame(
  "Chart 4: Unit Economics — Cost to Serve (CTS) Decay Curve",
  "Unit Economics",
  "Single-ward server costs decrease as system throughput and data amortization scale.",
);

const ctsLineData = [
  {
    name: "Monthly Cost to Serve (CTS) per Ward",
    labels: financialModel.unitEconomics.monthlyCostToServe.map((d) => d.year),
    values: financialModel.unitEconomics.monthlyCostToServe.map((d) => d.cost),
  },
];

s12.addChart(P.ChartType.line, ctsLineData, {
  x: 0.6,
  y: 2.2,
  w: 7.2,
  h: 4.5,
  chartColors: [C.blue],
  lineSize: 3,
  showValue: true,
  valAxisTitle: "Monthly Cost to Serve ($ USD)",
  catAxisTitle: "Timeline Horizon",
});

// Summary cards on the right
const ctsCards = [
  {
    title: "Year 1 CTS",
    val: "$18.00 / mo",
    desc: "$3 token parsing + $15 amortized server overhead across first 100 wards.",
  },
  {
    title: "Year 2 CTS",
    val: "$6.20 / mo",
    desc: "Significant infrastructure amortization as ward counts scale to 300.",
  },
  {
    title: "Year 5 CTS",
    val: "$3.25 / mo",
    desc: "Fixed infrastructure fully amortized by high-yield data broker asset contracts.",
  },
];

ctsCards.forEach((c, idx) => {
  const y = 2.2 + idx * 1.5;
  s12.addShape(P.ShapeType.rect, {
    x: 8.2,
    y,
    w: 4.5,
    h: 1.35,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s12.addText(c.title, {
    x: 8.4,
    y: y + 0.15,
    w: 2.3,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s12.addText(c.val, {
    x: 10.7,
    y: y + 0.15,
    w: 1.8,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: C.primary,
    fontFace: "Arial",
    align: "right",
  });
  s12.addText(c.desc, {
    x: 8.4,
    y: y + 0.5,
    w: 4.1,
    h: 0.7,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
  });
});

// -------------------------------------------------------------
// SLIDE 13: Core Unit Economics & Single-Ward Micro Analysis
// -------------------------------------------------------------
const s13 = addSlideFrame(
  "Core Unit Economics & Single-Ward Micro Analysis",
  "Microeconomics",
  "Evaluating the model at the individual Paying Ward level illustrates underlying network efficiency.",
);

const microMetrics = [
  {
    label: "Target Ward Composition",
    val: "1 Councillor + 0.2 Provider",
    detail:
      "1 local Councillor seat ($10/mo) paired with ~0.2 integrated Service Provider access configs ($10/mo pro-rata).",
  },
  {
    label: "Blended Monthly Revenue",
    val: "$20.00 / month",
    detail:
      "$240.00 annualized recurring subscription yield per active ward footprint before data monetization.",
  },
  {
    label: "Variable Cost to Serve (CTS)",
    val: "$3.00 / month",
    detail:
      "150 incoming tickets/month @ $0.02 variable AI token cost per validated incident dispatch.",
  },
  {
    label: "Server Overhead Amortization",
    val: "$15.00 / month (Y1)",
    detail:
      "$18k/yr hosting + dedicated Claude API allocated across initial cohort of 100 active wards.",
  },
  {
    label: "Net Contribution (Y1 → Y5)",
    val: "$2.00 → $16.75 / mo",
    detail:
      "Ward contribution margin expands from 10% in Year 1 to 83.8% in Year 5 as server costs decline to $3.25.",
  },
  {
    label: "Customer Acquisition Cost",
    val: "$45.00 / ward",
    detail:
      "Direct outreach to municipal councillor cohorts yields capital payback within 2.25 months (< 1 quarter).",
  },
];

microMetrics.forEach((m, idx) => {
  const col = idx % 2;
  const row = Math.floor(idx / 2);
  const x = 0.6 + col * 6.2;
  const y = 2.2 + row * 1.5;
  s13.addShape(P.ShapeType.rect, {
    x,
    y,
    w: 5.9,
    h: 1.35,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s13.addText(m.label, {
    x: x + 0.2,
    y: y + 0.15,
    w: 3.5,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s13.addText(m.val, {
    x: x + 3.5,
    y: y + 0.15,
    w: 2.2,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: C.primary,
    fontFace: "Arial",
    align: "right",
  });
  s13.addText(m.detail, {
    x: x + 0.2,
    y: y + 0.5,
    w: 5.5,
    h: 0.75,
    fontSize: 11,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 16,
  });
});

// -------------------------------------------------------------
// SLIDE 14: Data Governance & Ethical AI Pipeline
// -------------------------------------------------------------
const s14 = addSlideFrame(
  "Data Governance, Privacy (POPIA) & Frontier AI Training",
  "Governance & Compliance",
  "Transforming raw civic messages into clean, legally compliant structured training sets.",
);

const govPillars = [
  {
    title: "Privacy & De-identification",
    desc: "Complete stripping of personal identifiers (PII), telephone numbers, and resident identities in compliance with South Africa's Protection of Personal Information Act (POPIA) and global data privacy standards.",
  },
  {
    title: "Canonical Triage Graph",
    desc: "Structuring real-time incident lifecycle: raw report -> automated categorization -> departmental assignment -> contractor dispatch -> physical repair -> community verification.",
  },
  {
    title: "RL & Agentic Training Utility",
    desc: "Frontier AI labs specifically acquire this data to train autonomous multi-agent operational decision-making systems capable of coordinating complex physical logistics and emergency municipal triage.",
  },
];

govPillars.forEach((gp, idx) => {
  const x = 0.6 + idx * 4.05;
  s14.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 3.9,
    h: 4.4,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  s14.addText(`PILLAR ${idx + 1}`, {
    x: x + 0.25,
    y: 2.5,
    w: 3.4,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: C.blue,
    fontFace: "Arial",
  });
  s14.addText(gp.title, {
    x: x + 0.25,
    y: 2.85,
    w: 3.4,
    h: 0.7,
    fontSize: 17,
    bold: true,
    color: C.navy,
    fontFace: "Arial",
  });
  s14.addText(gp.desc, {
    x: x + 0.25,
    y: 3.65,
    w: 3.4,
    h: 2.7,
    fontSize: 12,
    color: C.muted,
    fontFace: "Arial",
    lineSpacing: 18,
  });
});

// -------------------------------------------------------------
// SLIDE 15: Conclusion & Investment Milestones
// -------------------------------------------------------------
const s15 = addSlideFrame(
  "Summary & Investment Milestones",
  "Strategic Roadmap",
  "Clear execution path from initial seed validation to institutional data scale.",
  true,
);

const roadmapPhases = [
  {
    phase: "Phase 1: Seed Validation",
    time: "Months 1–12",
    goals:
      "• Deploy $100k seed runway\n• Onboard 100 Ward Councillors\n• Onboard 20 Service Providers\n• Establish $30k Year 1 baseline\n• Validate 150 tickets/mo/ward",
  },
  {
    phase: "Phase 2: Data Commercialization",
    time: "Months 13–24",
    goals:
      "• Scale to 300 wards & 8 municipalities\n• License pilot dataset to micro1 ($100k)\n• Reach profitability ($162.9k profit)\n• Launch CivicRewards advertising loop\n• 18–24 mo self-funding runway",
  },
  {
    phase: "Phase 3: Global AI Refinery",
    time: "Years 3–5",
    goals:
      "• Expand to 3,200 wards & 90 municipalities\n• Deliver proprietary datasets to 5 AI labs\n• Achieve $17.425M revenue\n• 89.9% gross margin & $15.26M profit\n• Global standard for municipal triage AI",
  },
];

roadmapPhases.forEach((rp, idx) => {
  const x = 0.6 + idx * 4.05;
  s15.addShape(P.ShapeType.rect, {
    x,
    y: 2.3,
    w: 3.9,
    h: 4.4,
    fill: { color: "132433" },
    line: { color: C.primary, width: 1 },
  });
  s15.addText(rp.time.toUpperCase(), {
    x: x + 0.25,
    y: 2.5,
    w: 3.4,
    h: 0.25,
    fontSize: 11,
    bold: true,
    color: C.cyan,
    fontFace: "Arial",
  });
  s15.addText(rp.phase, {
    x: x + 0.25,
    y: 2.8,
    w: 3.4,
    h: 0.6,
    fontSize: 17,
    bold: true,
    color: C.white,
    fontFace: "Arial",
  });
  s15.addText(rp.goals, {
    x: x + 0.25,
    y: 3.5,
    w: 3.4,
    h: 3.0,
    fontSize: 12,
    color: C.light,
    fontFace: "Arial",
    lineSpacing: 20,
  });
});

async function main() {
  P.title = "Signal Desk Business Plan";
  P.subject =
    "Strategic Framework for an Autonomous Orchestration Layer for Municipal Infrastructure Platform";
  P.author = "Signal Desk";

  fs.mkdirSync(path.dirname(DOWNLOAD_PATH), { recursive: true });
  await P.writeFile({ fileName: DOWNLOAD_PATH });
  console.log(`PowerPoint generated successfully at ${DOWNLOAD_PATH}`);

  const manifest = {
    title: "Signal Desk Business Plan",
    subtitle:
      "Strategic Framework for an Autonomous Orchestration Layer for Municipal Infrastructure Platform.",
    slideCount: slideRegistry.length,
    slides: slideRegistry,
    financials: {
      years: financialModel.years,
      revenue: financialModel.grossRevenue,
      cor: financialModel.costOfRevenue,
      grossProfit: financialModel.grossProfit,
      grossMarginPct: financialModel.grossMarginPct,
      headcount: financialModel.headcount,
      payroll: financialModel.payroll,
      otherOpex: financialModel.otherOpex,
      totalOpex: financialModel.totalOpex,
      totalExpenses: financialModel.totalExpenses,
      netProfitLoss: financialModel.netProfitLoss,
      accounts: financialModel.accounts,
      y1RevenueDistribution: financialModel.y1RevenueDistribution,
      y5RevenueDistribution: financialModel.y5RevenueDistribution,
      unitEconomics: financialModel.unitEconomics,
      runway: financialModel.runway,
    },
  };

  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2), "utf-8");
  console.log(`Plan manifest updated at ${MANIFEST_PATH}`);
}

main().catch((err) => {
  console.error("Error generating presentation:", err);
  process.exit(1);
});
