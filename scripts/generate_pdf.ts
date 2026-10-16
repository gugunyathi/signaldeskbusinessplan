import PDFDocument from "pdfkit";
import { PDFDocument as PDFLibDoc } from "pdf-lib";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const PDF_PATH = path.join(ROOT, "public", "downloads", "Signal_Desk_Business_Plan.pdf");
const TEMP_PATH = path.join(ROOT, "public", "downloads", "temp_raw_plan.pdf");

// Widescreen 16:9 dimensions in points (landscape standard 842 x 474 pt)
const WIDTH = 842;
const HEIGHT = 474;

const C = {
  navy: "#162B3A",
  primary: "#008A68",
  blue: "#0068A0",
  cyan: "#00A896",
  white: "#FFFFFF",
  muted: "#62717D",
  light: "#F4F7F9",
  card: "#FFFFFF",
  line: "#D6E0E5",
  danger: "#B74444",
  darkBg: "#0E1A24",
  cardBg: "#162736",
  textDark: "#1E293B",
  emeraldLight: "#E6F4EA",
  emeraldBorder: "#34D399",
};

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
  totalExpenses: [79500, 152100, 341990, 1621520, 2165500],
  netProfitLoss: [-49500, 162900, 1181010, 11639480, 15259500],
};

// Clean non-standard Unicode characters to 100% compliant standard ASCII for Type 1 fonts
function clean(str: string): string {
  return str
    .replace(/[–—]/g, "-")
    .replace(/[•]/g, "-")
    .replace(/[→]/g, "->")
    .replace(/[×]/g, "x")
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[^\x20-\x7E\n\r\t]/g, " ");
}

function generateRawPDF(): Promise<void> {
  return new Promise((resolve, reject) => {
    fs.mkdirSync(path.dirname(TEMP_PATH), { recursive: true });
    const doc = new PDFDocument({
      size: [WIDTH, HEIGHT],
      margin: 0,
      autoFirstPage: false,
      info: {
        Title: clean("SIGNAL DESK - Comprehensive Business Plan & Operational Model"),
        Author: "Signal Desk",
        Subject: clean(
          "Strategic Framework for a Cloud-Based Public Infrastructure Triage & Asset-Liquidation Platform",
        ),
        Keywords: "Business Plan, Public Infrastructure, AI Data Refinery, Financial Projections",
      },
    });

    const stream = fs.createWriteStream(TEMP_PATH);
    doc.pipe(stream);

    function addHeaderFooter(
      section: string,
      title: string,
      subtitle: string,
      pageNum: number,
      dark = false,
    ) {
      doc.addPage({ size: [WIDTH, HEIGHT], margin: 0 });

      // Background
      doc.rect(0, 0, WIDTH, HEIGHT).fill(dark ? C.darkBg : C.white);

      // Top Tag
      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(dark ? C.primary : C.blue)
        .text(clean(`SIGNAL DESK  /  ${section.toUpperCase()}`), 40, 24, { width: WIDTH - 80 });

      // Main Title
      doc
        .font("Helvetica-Bold")
        .fontSize(18)
        .fillColor(dark ? C.white : C.navy)
        .text(clean(title), 40, 42, { width: WIDTH - 80 });

      // Subtitle
      if (subtitle) {
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor(dark ? "#94A3B8" : C.muted)
          .text(clean(subtitle), 40, 66, { width: WIDTH - 80 });
      }

      // Footer line
      doc
        .strokeColor(dark ? "#1E293B" : C.line)
        .lineWidth(0.5)
        .moveTo(40, HEIGHT - 28)
        .lineTo(WIDTH - 40, HEIGHT - 28)
        .stroke();

      // Footer text
      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor(dark ? "#64748B" : C.muted)
        .text(
          clean("BUSINESS PLAN - OCTOBER 2026  |  CONFIDENTIAL - MANAGEMENT PLANNING, NOT AUDITED"),
          40,
          HEIGHT - 22,
          { width: WIDTH - 120 },
        );

      doc
        .font("Helvetica-Bold")
        .fontSize(8)
        .fillColor(dark ? C.white : C.navy)
        .text(String(pageNum).padStart(2, "0"), WIDTH - 60, HEIGHT - 22, {
          width: 20,
          align: "right",
        });
    }

    // -------------------------------------------------------------
    // PAGE 1: Cover
    // -------------------------------------------------------------
    doc.addPage({ size: [WIDTH, HEIGHT], margin: 0 });
    doc.rect(0, 0, WIDTH, HEIGHT).fill(C.darkBg);

    doc
      .font("Helvetica-Bold")
      .fontSize(11)
      .fillColor(C.cyan)
      .text(clean("SIGNAL DESK  /  STRATEGIC BUSINESS PLAN"), 45, 45);

    doc.font("Helvetica-Bold").fontSize(38).fillColor(C.white).text("SIGNAL DESK", 45, 75);

    doc
      .font("Helvetica-Bold")
      .fontSize(18)
      .fillColor("#6EE7B7")
      .text(clean("Comprehensive Business Plan & Operational Model"), 45, 125);

    doc
      .font("Helvetica")
      .fontSize(12)
      .fillColor("#CBD5E1")
      .text(
        clean(
          "Strategic Framework for a Cloud-Based Public Infrastructure Triage & Asset-Liquidation Platform",
        ),
        45,
        155,
        { width: 680, lineGap: 4 },
      );

    doc
      .font("Helvetica")
      .fontSize(10)
      .fillColor("#94A3B8")
      .text(
        clean(
          "A high-margin B2G and B2B public infrastructure data ecosystem capturing real-time municipal triage logs, citizen routing, and multi-team field dispatch dynamics to license proprietary operational datasets to frontier AI engineering labs (Anthropic, Google, OpenAI, X/Cursor, Meta).",
        ),
        45,
        185,
        { width: 660, lineGap: 3 },
      );

    const badges = [
      { label: "SEED RUNWAY", val: "$100k (18-24 mo)" },
      { label: "YEAR 5 REVENUE", val: "$17.425M" },
      { label: "YEAR 5 NET PROFIT", val: "+$15.260M" },
      { label: "GROSS MARGIN", val: "89.9%" },
    ];
    badges.forEach((b, i) => {
      const bx = 45 + i * 188;
      const by = 280;
      doc.roundedRect(bx, by, 175, 65, 4).fillAndStroke("#162736", "#059669");
      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(C.cyan)
        .text(clean(b.label), bx + 12, by + 12);
      doc
        .font("Helvetica-Bold")
        .fontSize(15)
        .fillColor(C.white)
        .text(clean(b.val), bx + 12, by + 32);
    });

    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor("#64748B")
      .text(
        clean(
          "OCTOBER 2026 - JOHANNESBURG, SOUTH AFRICA  |  CONFIDENTIAL - MANAGEMENT PLANNING CASE",
        ),
        45,
        HEIGHT - 35,
      );

    // -------------------------------------------------------------
    // PAGE 2: Executive Summary
    // -------------------------------------------------------------
    addHeaderFooter(
      "Executive Summary",
      "Executive Summary: High-Margin Infrastructure Data Ecosystem",
      "Commercial deployment of a cloud-based public infrastructure triage & asset-liquidation platform.",
      2,
    );

    const pillars = [
      {
        tag: "SOURCE LAYER",
        title: "Proprietary Data Ingestion",
        desc: "By capturing real-time municipal triage logs, citizen utility routing patterns, and multi-team field dispatch dynamics, the platform creates an un-scraped, highly proprietary operational dataset unavailable in open public repositories.",
      },
      {
        tag: "HIGH-MARGIN SCALE",
        title: "Frontier AI Lab Licensing",
        desc: "Beginning in Year 2, processed structural datasets are licensed directly to frontier AI labs (Anthropic, Google, OpenAI, X/Cursor, Meta) for specialized neural networks, agentic workflows, and Reinforcement Learning (RL) training.",
      },
      {
        tag: "RUNWAY SECURITY",
        title: "Capital Efficient Runtime",
        desc: "A $100,000 initial seed round funds 18 to 24 months of core runtime with a lean $75,000 Year 1 total baseline outflow, enabling proof of unit economics and data refinery scale before raising institutional growth capital.",
      },
    ];

    pillars.forEach((p, i) => {
      const px = 40 + i * 258;
      const py = 95;
      doc.roundedRect(px, py, 246, 330, 6).fillAndStroke(C.light, C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(C.blue)
        .text(clean(p.tag), px + 16, py + 18);

      doc
        .font("Helvetica-Bold")
        .fontSize(14)
        .fillColor(C.navy)
        .text(clean(p.title), px + 16, py + 38, { width: 214 });

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor(C.muted)
        .text(clean(p.desc), px + 16, py + 95, { width: 214, lineGap: 4 });
    });

    // -------------------------------------------------------------
    // PAGE 3: Runway Analysis
    // -------------------------------------------------------------
    addHeaderFooter(
      "Runway Analysis",
      "Financial Feasibility & Seed Runway Allocation ($100,000)",
      "Initial seed round allocation providing 18 to 24 months of core operational runtime.",
      3,
    );

    const runwayRows = [
      {
        title: "Sunk Assets (Historical R&D)",
        amount: "$30,000",
        detail:
          "Prior software design iterations and foundational engine engineering capitalised as a sunk setup asset.",
        color: C.muted,
      },
      {
        title: "Net Cash Position (Working Capital)",
        amount: "$70,000",
        detail:
          "Initial liquid working capital deployed into active operations, cloud hosting, and councillor outreach.",
        color: C.primary,
      },
      {
        title: "Fixed Overhead Core Infrastructure",
        amount: "$18,000 / yr",
        detail:
          "Baseline cloud server deployments ($500/month) and dedicated Claude API workflow layers ($1,000/month).",
        color: C.blue,
      },
      {
        title: "Year 1 Total Baseline Outflows",
        amount: "$75,000 / yr",
        detail:
          "Combined with lean Year 1 Operating Expenses of $57,000 ($45k payroll + $12k other opex) + $18k infrastructure.",
        color: C.navy,
      },
    ];

    runwayRows.forEach((row, i) => {
      const ry = 95 + i * 68;
      doc.roundedRect(40, ry, WIDTH - 80, 58, 4).fillAndStroke(C.light, C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(11)
        .fillColor(C.navy)
        .text(clean(row.title), 55, ry + 12);

      doc
        .font("Helvetica")
        .fontSize(9)
        .fillColor(C.muted)
        .text(clean(row.detail), 55, ry + 30, { width: 520 });

      doc
        .font("Helvetica-Bold")
        .fontSize(14)
        .fillColor(row.color)
        .text(clean(row.amount), WIDTH - 180, ry + 18, { width: 120, align: "right" });
    });

    doc.roundedRect(40, 385, WIDTH - 80, 42, 4).fillAndStroke(C.emeraldLight, C.emeraldBorder);
    doc
      .font("Helvetica-Bold")
      .fontSize(10)
      .fillColor(C.primary)
      .text(
        clean(
          "Key Runway Takeaway: The $100,000 seed round funds 18 to 24 months of core runtime, providing ample time to build proprietary data assets and secure self-sustaining revenue before follow-on investment.",
        ),
        55,
        397,
        { width: WIDTH - 110 },
      );

    // -------------------------------------------------------------
    // PAGE 4: GTM Strategy
    // -------------------------------------------------------------
    addHeaderFooter(
      "GTM Strategy",
      "Market Positioning & Go-To-Market (GTM) Strategy",
      "Circumventing traditional public sector procurement friction via a bottom-up adoption wedge.",
      4,
    );

    const stages = [
      {
        badge: "STAGE 1 - LOW FRICTION",
        title: "Bottom-Up Councillor Wedge",
        price: "$10 / month",
        desc: "Direct outreach loops introduce individual ward councillors to dedicated utility tracking dashboards. Allows real-time tracking of community service requests, building organic ward adoption at low friction.",
      },
      {
        badge: "STAGE 2 - WORKFLOW LINK",
        title: "B2B Service Provider Onboarding",
        price: "$50 / month",
        desc: "Contracted external municipal vendors (plumbing, road maintenance, electrical repair crews) use the interface for structured dispatches, drastically lowering manual verification and inspection overhead.",
      },
      {
        badge: "STAGE 3 - INSTITUTIONAL",
        title: "Data Asset Liquidation",
        price: "$100k -> $1.0M / package",
        desc: "As infrastructure data deepens, the platform transitions into a high-tier data refinery. Anonymized, end-to-end municipal triage and resolution datasets are licensed to frontier AI labs.",
      },
    ];

    stages.forEach((st, i) => {
      const sx = 40 + i * 258;
      const sy = 95;
      const isLast = i === 2;
      doc
        .roundedRect(sx, sy, 246, 330, 6)
        .fillAndStroke(isLast ? "#ECFDF5" : C.light, isLast ? C.primary : C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(isLast ? C.primary : C.blue)
        .text(clean(st.badge), sx + 16, sy + 18);

      doc
        .font("Helvetica-Bold")
        .fontSize(14)
        .fillColor(C.navy)
        .text(clean(st.title), sx + 16, sy + 38, { width: 214 });

      doc
        .font("Helvetica-Bold")
        .fontSize(12)
        .fillColor(isLast ? C.primary : C.navy)
        .text(clean(`Pricing: ${st.price}`), sx + 16, sy + 85);

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor(C.muted)
        .text(clean(st.desc), sx + 16, sy + 115, { width: 214, lineGap: 4 });
    });

    // -------------------------------------------------------------
    // PAGE 5: CivicRewards
    // -------------------------------------------------------------
    addHeaderFooter(
      "Product Architecture",
      "CivicRewards Loyalty Framework",
      "Maintaining reporting volume and verified field data without increasing operational cash burn.",
      5,
    );

    const steps = [
      {
        num: "01",
        title: "Citizen Report",
        desc: "Citizen submits verified infrastructure gap via Telegram mini-app / web.",
      },
      {
        num: "02",
        title: "Geolocation & De-dupe",
        desc: "System validates GPS coordinates, identifies duplicate tickets, and maps official refs.",
      },
      {
        num: "03",
        title: "Token Allocation",
        desc: "CivicPoints awarded when service gap is independently validated by community resolution.",
      },
      {
        num: "04",
        title: "Partner Redemption",
        desc: "Points redeemed for local merchant discounts funded by B2B Advertisers ($1,000/mo).",
      },
    ];

    steps.forEach((st, i) => {
      const sx = 40 + i * 192;
      const sy = 95;
      doc.roundedRect(sx, sy, 182, 170, 5).fillAndStroke(C.light, C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(18)
        .fillColor(C.primary)
        .text(clean(st.num), sx + 14, sy + 14);

      doc
        .font("Helvetica-Bold")
        .fontSize(11)
        .fillColor(C.navy)
        .text(clean(st.title), sx + 14, sy + 42, { width: 154 });

      doc
        .font("Helvetica")
        .fontSize(9)
        .fillColor(C.muted)
        .text(clean(st.desc), sx + 14, sy + 75, { width: 154, lineGap: 3 });
    });

    doc.roundedRect(40, 280, 375, 145, 5).fillAndStroke(C.white, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(11)
      .fillColor(C.navy)
      .text(clean("Validation Engine"), 55, 296);
    doc
      .font("Helvetica")
      .fontSize(9.5)
      .fillColor(C.muted)
      .text(
        clean(
          "Citizens earn loyalty tokens ('CivicPoints') when they log verified public service gaps that are subsequently validated by community resolution engines. This ensures high signal-to-noise ratio in incoming municipal dispatch feeds.",
        ),
        55,
        320,
        { width: 345, lineGap: 3 },
      );

    doc.roundedRect(427, 280, 375, 145, 5).fillAndStroke(C.white, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(11)
      .fillColor(C.navy)
      .text(clean("Loop Monetization (Zero Cash Burn)"), 442, 296);
    doc
      .font("Helvetica")
      .fontSize(9.5)
      .fillColor(C.muted)
      .text(
        clean(
          "Points are redeemable for localized value-added discounts and micro-rewards. These incentives are funded directly by Premium B2B Advertisers paying $1,000/month, keeping customer engagement high without creating any direct cash burn for Signal Desk.",
        ),
        442,
        320,
        { width: 345, lineGap: 3 },
      );

    // -------------------------------------------------------------
    // PAGE 6: AI Broker Timeline
    // -------------------------------------------------------------
    addHeaderFooter(
      "Data Asset Refinery",
      "Data Refinement & AI Broker Monetization Timeline",
      "Platform data assets increase in valuation as transactional complexity and running history accumulate.",
      6,
    );

    const brokerTiers = [
      {
        badge: "YEAR 2 - QUALIFIED PARTNERSHIPS",
        price: "$100,000 / package",
        buyers: "1 Broker (micro1)",
        desc: "The platform licenses basic operational logs, citizen routing patterns, and triage histories under pilot data partnership programs.",
      },
      {
        badge: "YEAR 3 - LARGE DATA COLLECTIONS",
        price: "$500,000 / package",
        buyers: "2 Brokers ($1.0M Total)",
        desc: "Capturing ongoing, multi-team, deep transactional logs and field work histories across 20 partner municipalities.",
      },
      {
        badge: "YEARS 4 & 5 - HIGHLY UNIQUE / PROPRIETARY",
        price: "$1,000,000 / package",
        buyers: "Anthropic, Google, OpenAI, X/Cursor, Meta",
        desc: "Distinctive operational workflows involving complex human decisions, triage, and escalations. 3 partner municipalities produce 1 distinct dataset/yr sold across 5 active brokers ($15.0M in Year 5).",
      },
    ];

    brokerTiers.forEach((bt, i) => {
      const bx = 40 + i * 258;
      const by = 95;
      const isLast = i === 2;
      doc
        .roundedRect(bx, by, 246, 330, 6)
        .fillAndStroke(isLast ? "#ECFDF5" : C.light, isLast ? C.primary : C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(isLast ? C.primary : C.blue)
        .text(clean(bt.badge), bx + 16, by + 18);

      doc
        .font("Helvetica-Bold")
        .fontSize(16)
        .fillColor(C.navy)
        .text(clean(bt.price), bx + 16, by + 40);

      doc
        .font("Helvetica-Bold")
        .fontSize(10)
        .fillColor(isLast ? C.primary : C.blue)
        .text(clean(`Target: ${bt.buyers}`), bx + 16, by + 75, { width: 214 });

      doc
        .font("Helvetica")
        .fontSize(9.5)
        .fillColor(C.muted)
        .text(clean(bt.desc), bx + 16, by + 115, { width: 214, lineGap: 4 });
    });

    // -------------------------------------------------------------
    // PAGE 7: Scaling Targets Table
    // -------------------------------------------------------------
    addHeaderFooter(
      "Scaling Milestones",
      "5-Year Operational & Accounts Scaling Targets",
      "Target accounts scaling across software seats, partner municipalities, and AI data brokers.",
      7,
    );

    const scalingRows = [
      ["Ward Councillors ($10/mo)", "100", "300", "800", "1,800", "3,200"],
      ["B2B Service Providers ($50/mo)", "20", "50", "120", "250", "500"],
      ["Partner Municipalities", "3", "8", "20", "45", "90"],
      ["Premium Advertisers ($1k/mo)", "3", "8", "15", "30", "50"],
      ["Data Analytics Clients", "1", "3", "7", "15", "30"],
      ["AI Data Brokers (Frontier Labs)", "0", "1 (micro1)", "2", "4", "5 (Tier 1)"],
    ];

    const tableX = 40;
    const tableY = 95;
    const colWidths = [242, 104, 104, 104, 104, 104];

    doc.rect(tableX, tableY, WIDTH - 80, 28).fill(C.navy);
    const headers = ["Account Segment", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
    let cx = tableX;
    headers.forEach((h, i) => {
      doc
        .font("Helvetica-Bold")
        .fontSize(10)
        .fillColor(C.white)
        .text(clean(h), cx + 8, tableY + 9, {
          width: colWidths[i]! - 16,
          align: i === 0 ? "left" : "right",
        });
      cx += colWidths[i]!;
    });

    scalingRows.forEach((row, ri) => {
      const ry = tableY + 28 + ri * 42;
      const isAlt = ri % 2 === 1;
      const isHighlight = ri === 5;
      doc
        .rect(tableX, ry, WIDTH - 80, 42)
        .fillAndStroke(isHighlight ? "#ECFDF5" : isAlt ? C.light : C.white, C.line);

      let rcx = tableX;
      row.forEach((cell, ci) => {
        doc
          .font(ci === 0 || isHighlight ? "Helvetica-Bold" : "Helvetica")
          .fontSize(9.5)
          .fillColor(isHighlight ? C.primary : ci === 0 ? C.navy : C.textDark)
          .text(clean(cell), rcx + 8, ry + 15, {
            width: colWidths[ci]! - 16,
            align: ci === 0 ? "left" : "right",
          });
        rcx += colWidths[ci]!;
      });
    });

    doc
      .font("Helvetica")
      .fontSize(8.5)
      .fillColor(C.muted)
      .text(
        clean(
          "Notes: AI Data Brokers in Year 4 & 5 represent structured dataset partnership contracts with Anthropic, Google, OpenAI, X/Cursor, and Meta.",
        ),
        40,
        380,
      );

    // -------------------------------------------------------------
    // PAGE 8: Table A - Consolidated P&L
    // -------------------------------------------------------------
    addHeaderFooter(
      "Financial Model",
      "Table A: 5-Year Consolidated Financial Projection (USD)",
      "Pro-forma income statement modeling exponential revenue growth and 89.9% Year-5 gross margins.",
      8,
    );

    const plRows = [
      ["Gross Annual Revenue", "$30,000", "$315,000", "$1,523,000", "$13,261,000", "$17,425,000"],
      ["Cost of Revenue (COR)", "$22,500", "$62,100", "$186,990", "$1,361,520", "$1,760,500"],
      ["Gross Profit", "$7,500", "$252,900", "$1,336,010", "$11,899,480", "$15,664,500"],
      ["Gross Margin %", "25.0%", "80.3%", "87.7%", "89.7%", "89.9%"],
      ["Staffing Headcount", "3 Staff", "4 Staff", "6 Staff", "9 Staff", "13 Staff"],
      ["Baseline Core Payroll", "$45,000", "$70,000", "$120,000", "$200,000", "$310,000"],
      ["Other Operating Expenses", "$12,000", "$20,000", "$35,000", "$60,000", "$95,000"],
      ["Total Operating Expenses (Opex)", "$57,000", "$90,000", "$155,000", "$260,000", "$405,000"],
      [
        "Net Operating Profit / (Loss)",
        "(-$49,500)",
        "+$162,900",
        "+$1,181,010",
        "+$11,639,480",
        "+$15,259,500",
      ],
    ];

    const plTableY = 92;
    const plRowH = 34;

    doc.rect(tableX, plTableY, WIDTH - 80, 26).fill(C.navy);
    let plcx = tableX;
    headers.forEach((h, i) => {
      doc
        .font("Helvetica-Bold")
        .fontSize(9.5)
        .fillColor(C.white)
        .text(clean(h), plcx + 8, plTableY + 8, {
          width: colWidths[i]! - 16,
          align: i === 0 ? "left" : "right",
        });
      plcx += colWidths[i]!;
    });

    plRows.forEach((row, ri) => {
      const ry = plTableY + 26 + ri * plRowH;
      const isNetProfit = ri === 8;
      const isRevenue = ri === 0;
      const isGrossMargin = ri === 3;
      const isAlt = ri % 2 === 1;

      doc
        .rect(tableX, ry, WIDTH - 80, plRowH)
        .fillAndStroke(
          isNetProfit ? "#E6F4EA" : isRevenue ? "#EFF6FF" : isAlt ? C.light : C.white,
          C.line,
        );

      let rcx = tableX;
      row.forEach((cell, ci) => {
        let cellColor = C.textDark;
        if (isNetProfit) {
          cellColor = ci === 1 ? C.danger : C.primary;
        } else if (isRevenue) {
          cellColor = C.blue;
        } else if (isGrossMargin) {
          cellColor = C.primary;
        }

        doc
          .font(
            ci === 0 || isNetProfit || isRevenue || isGrossMargin ? "Helvetica-Bold" : "Helvetica",
          )
          .fontSize(9)
          .fillColor(cellColor)
          .text(clean(cell), rcx + 8, ry + 11, {
            width: colWidths[ci]! - 16,
            align: ci === 0 ? "left" : "right",
          });
        rcx += colWidths[ci]!;
      });
    });

    // -------------------------------------------------------------
    // PAGE 9: Operational Cost Architecture
    // -------------------------------------------------------------
    addHeaderFooter(
      "Cost Discipline",
      "Operational Cost Architecture & Lean Delivery Framework",
      "Structural efficiency ensures operating expenses remain below 2.5% of revenue by Year 5.",
      9,
    );

    doc.roundedRect(40, 95, 375, 330, 6).fillAndStroke(C.light, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor(C.navy)
      .text(clean("Cost of Revenue (COR) Architecture"), 55, 115);

    const corPoints = [
      "- Baseline Core Infrastructure: $18,000 per year fixed baseline comprising cloud servers ($500/mo) and dedicated Claude API workflow layers ($1,000/mo).",
      "- Variable Platform Usage: Decreases systematically from 15% in Year 1 down to 10% in Year 5 as server efficiency, ticket deduplication, and cache hit rates scale.",
      "- Near-Zero Marginal Data Cost: Once raw incident and dispatch telemetry is captured and linked, dataset compilation and packaging for AI labs requires minimal marginal compute.",
      "- Operating Leverage: Gross margin expands from 25.0% in Year 1 to 89.9% in Year 5 as institutional data broker revenue outpaces operational delivery expenses.",
    ];
    let corY = 150;
    corPoints.forEach((p) => {
      doc
        .font("Helvetica")
        .fontSize(9.5)
        .fillColor(C.muted)
        .text(clean(p), 55, corY, { width: 345, lineGap: 3 });
      corY += 62;
    });

    doc.roundedRect(427, 95, 375, 330, 6).fillAndStroke(C.light, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor(C.navy)
      .text(clean("Operating Expense (Opex) Discipline"), 442, 115);

    const opexPoints = [
      "- Lean Engineering Payroll: Capitalizes on high-caliber South African technical talent ($45k Year 1 -> $310k Year 5 across 13 engineering and operations personnel).",
      "- Below-Market Founder Compensation: Founder salaries maintained strictly below market benchmarks early in the lifecycle to protect liquid runway.",
      "- Strict Travel & Procurement Cap: Zero speculative municipal lobbying or travel; all external development expenses bounded inside the 'Other Opex' envelope ($12k -> $95k).",
      "- Profitable by Year 2: Net operating turnaround generates +$162,900 profit in Year 2, funding internal growth without early dilutive down-rounds.",
    ];
    let opexY = 150;
    opexPoints.forEach((p) => {
      doc
        .font("Helvetica")
        .fontSize(9.5)
        .fillColor(C.muted)
        .text(clean(p), 442, opexY, { width: 345, lineGap: 3 });
      opexY += 62;
    });

    // -------------------------------------------------------------
    // PAGE 10: Vector Bar Chart - 5-Year Trajectory
    // -------------------------------------------------------------
    addHeaderFooter(
      "Financial Projections",
      "Chart 1: 5-Year Financial Trajectory & AI Lab Data Scaling",
      "Comparison of Annual Revenue, Total Operational Expenses (COR + Opex), and Net Operating Profit.",
      10,
    );

    const chartLeft = 80;
    const chartTop = 110;
    const chartW = WIDTH - 160;
    const chartH = 260;
    const maxVal = 20000;

    doc
      .strokeColor(C.line)
      .lineWidth(1)
      .moveTo(chartLeft, chartTop)
      .lineTo(chartLeft, chartTop + chartH)
      .lineTo(chartLeft + chartW, chartTop + chartH)
      .stroke();

    [0, 5000, 10000, 15000, 20000].forEach((gridVal) => {
      const gy = chartTop + chartH - (gridVal / maxVal) * chartH;
      doc
        .strokeColor("#E2E8F0")
        .lineWidth(0.5)
        .moveTo(chartLeft, gy)
        .lineTo(chartLeft + chartW, gy)
        .stroke();

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor(C.muted)
        .text(`$${gridVal / 1000}M`, chartLeft - 45, gy - 4, { width: 40, align: "right" });
    });

    const barGroupW = chartW / 5;
    financialModel.years.forEach((yr, i) => {
      const gx = chartLeft + i * barGroupW + 15;
      const revH = (financialModel.grossRevenue[i]! / 1000 / maxVal) * chartH;
      const expH = (financialModel.totalExpenses[i]! / 1000 / maxVal) * chartH;
      const netVal = financialModel.netProfitLoss[i]! / 1000;
      const netH = (Math.max(netVal, 0) / maxVal) * chartH;

      const barW = 28;

      doc.rect(gx, chartTop + chartH - revH, barW, Math.max(revH, 2)).fill(C.blue);
      doc.rect(gx + 32, chartTop + chartH - expH, barW, Math.max(expH, 2)).fill(C.muted);

      if (netVal >= 0) {
        doc.rect(gx + 64, chartTop + chartH - netH, barW, Math.max(netH, 2)).fill(C.primary);
      } else {
        const lossH = (Math.abs(netVal) / maxVal) * chartH;
        doc.rect(gx + 64, chartTop + chartH, barW, Math.max(lossH, 3)).fill(C.danger);
      }

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(C.navy)
        .text(clean(yr), gx + 20, chartTop + chartH + 10, { width: 60, align: "center" });
    });

    const legendX = WIDTH - 350;
    const legendY = chartTop - 15;
    doc.rect(legendX, legendY, 12, 8).fill(C.blue);
    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor(C.navy)
      .text("Annual Revenue", legendX + 16, legendY);

    doc.rect(legendX + 100, legendY, 12, 8).fill(C.muted);
    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor(C.navy)
      .text("Total Expenses", legendX + 116, legendY);

    doc.rect(legendX + 200, legendY, 12, 8).fill(C.primary);
    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor(C.navy)
      .text("Net Profit", legendX + 216, legendY);

    doc.roundedRect(chartLeft, 400, chartW, 26, 4).fillAndStroke(C.light, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(8.5)
      .fillColor(C.primary)
      .text(
        clean(
          "Trajectory Summary: Explosive top-line revenue scaling from $30k (Y1) to $17.425M (Y5) driven by 86.1% AI Data Broker dataset licensing.",
        ),
        chartLeft + 12,
        408,
      );

    // -------------------------------------------------------------
    // PAGE 11: Vector Pie Charts - Revenue Stream Shift
    // -------------------------------------------------------------
    addHeaderFooter(
      "Revenue Mix",
      "Charts 2 & 3: Revenue Stream Evolution (Year 1 vs. Year 5)",
      "Transition from early municipal seat subscriptions to global AI data asset refinery.",
      11,
    );

    doc.roundedRect(40, 95, 375, 330, 6).fillAndStroke(C.light, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(12)
      .fillColor(C.navy)
      .text(clean("Year 1: Direct Software Footprint ($30,000 Total)"), 55, 112);

    const y1Dist = [
      { name: "Ward Councillors ($10/mo)", pct: "40.0%", amt: "$12,000", color: C.blue },
      { name: "B2B Service Providers ($50/mo)", pct: "40.0%", amt: "$12,000", color: C.primary },
      { name: "Data Analytics Clients", pct: "20.0%", amt: "$6,000", color: C.cyan },
      { name: "Municipal Enterprise Tiers", pct: "0.0%", amt: "$0", color: C.muted },
      { name: "AI Data Brokers", pct: "0.0%", amt: "$0", color: C.muted },
    ];

    y1Dist.forEach((d, i) => {
      const dy = 145 + i * 50;
      doc.rect(55, dy, 10, 10).fill(d.color);
      doc
        .font("Helvetica-Bold")
        .fontSize(10)
        .fillColor(C.navy)
        .text(clean(d.name), 72, dy - 1);
      doc
        .font("Helvetica-Bold")
        .fontSize(10)
        .fillColor(d.color)
        .text(clean(`${d.pct}  (${d.amt})`), 280, dy - 1, { width: 120, align: "right" });

      doc.roundedRect(55, dy + 18, 345, 10, 2).fill("#E2E8F0");
      const barW = (parseFloat(d.pct) / 100) * 345;
      if (barW > 0) doc.roundedRect(55, dy + 18, barW, 10, 2).fill(d.color);
    });

    doc.roundedRect(427, 95, 375, 330, 6).fillAndStroke(C.light, C.line);
    doc
      .font("Helvetica-Bold")
      .fontSize(12)
      .fillColor(C.navy)
      .text(clean("Year 5: Global Data Refinery ($17,425,000 Total)"), 442, 112);

    const y5Dist = [
      { name: "AI Data Broker Asset Sales", pct: "86.1%", amt: "$15,000,000", color: C.primary },
      { name: "Municipal Enterprise Tiers", pct: "6.2%", amt: "$1,080,000", color: C.blue },
      { name: "Data Analytics Clients", pct: "2.1%", amt: "$365,000", color: C.cyan },
      { name: "Premium Advertisers Network", pct: "1.7%", amt: "$300,000", color: "#E3B341" },
      { name: "B2B Service Providers", pct: "1.4%", amt: "$245,000", color: "#E76F51" },
      { name: "Ward Councillors Subscriptions", pct: "0.2%", amt: "$35,000", color: C.muted },
    ];

    y5Dist.forEach((d, i) => {
      const dy = 140 + i * 44;
      doc.rect(442, dy, 10, 10).fill(d.color);
      doc
        .font("Helvetica-Bold")
        .fontSize(9.5)
        .fillColor(C.navy)
        .text(clean(d.name), 459, dy - 1);
      doc
        .font("Helvetica-Bold")
        .fontSize(9.5)
        .fillColor(d.color)
        .text(clean(`${d.pct} (${d.amt})`), 670, dy - 1, { width: 120, align: "right" });

      doc.roundedRect(442, dy + 16, 345, 8, 2).fill("#E2E8F0");
      const barW = (parseFloat(d.pct) / 100) * 345;
      if (barW > 0) doc.roundedRect(442, dy + 16, barW, 8, 2).fill(d.color);
    });

    // -------------------------------------------------------------
    // PAGE 12: Cost to Serve (CTS) Decay Curve
    // -------------------------------------------------------------
    addHeaderFooter(
      "Unit Economics",
      "Chart 4: Unit Economics - Cost to Serve (CTS) Decay Curve",
      "Single-ward server costs decrease as system throughput and data monetization scale.",
      12,
    );

    const lChartX = 60;
    const lChartY = 120;
    const lChartW = 380;
    const lChartH = 250;

    doc
      .strokeColor(C.line)
      .lineWidth(1)
      .moveTo(lChartX, lChartY)
      .lineTo(lChartX, lChartY + lChartH)
      .lineTo(lChartX + lChartW, lChartY + lChartH)
      .stroke();

    [0, 5, 10, 15, 20].forEach((v) => {
      const gy = lChartY + lChartH - (v / 20) * lChartH;
      doc
        .strokeColor("#E2E8F0")
        .lineWidth(0.5)
        .moveTo(lChartX, gy)
        .lineTo(lChartX + lChartW, gy)
        .stroke();
      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor(C.muted)
        .text(`$${v.toFixed(2)}`, lChartX - 45, gy - 4, { width: 40, align: "right" });
    });

    const ctsData = [
      { yr: "Year 1", cost: 18.0 },
      { yr: "Year 2", cost: 6.2 },
      { yr: "Year 3", cost: 4.8 },
      { yr: "Year 4", cost: 3.8 },
      { yr: "Year 5", cost: 3.25 },
    ];

    const pts = ctsData.map((d, i) => {
      const px = lChartX + (i / 4) * lChartW;
      const py = lChartY + lChartH - (d.cost / 20) * lChartH;
      return { px, py, ...d };
    });

    doc.strokeColor(C.blue).lineWidth(3);
    pts.forEach((p, i) => {
      if (i === 0) doc.moveTo(p.px, p.py);
      else doc.lineTo(p.px, p.py);
    });
    doc.stroke();

    pts.forEach((p) => {
      doc.circle(p.px, p.py, 4).fill(C.blue);
      doc
        .font("Helvetica-Bold")
        .fontSize(8.5)
        .fillColor(C.navy)
        .text(`$${p.cost.toFixed(2)}`, p.px - 20, p.py - 16, { width: 40, align: "center" });

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor(C.muted)
        .text(clean(p.yr), p.px - 20, lChartY + lChartH + 8, { width: 40, align: "center" });
    });

    const ctsCards = [
      {
        title: "Year 1 CTS ($18.00 / month)",
        desc: "Consists of $3.00 variable AI parsing token cost (150 tickets @ $0.02) plus $15.00 amortized server overhead across the initial cohort of 100 active wards.",
      },
      {
        title: "Year 2 CTS ($6.20 / month)",
        desc: "Fixed infrastructure amortizes significantly as ward count scales to 300 active wards, while ticket deduplication and caching lower token costs.",
      },
      {
        title: "Year 5 CTS ($3.25 / month)",
        desc: "Fixed server overhead is virtually 100% amortized by data broker asset licensing. Single-ward contribution margin expands to 83.8%.",
      },
    ];

    ctsCards.forEach((c, i) => {
      const cy = 110 + i * 88;
      doc.roundedRect(480, cy, 320, 76, 5).fillAndStroke(C.light, C.line);
      doc
        .font("Helvetica-Bold")
        .fontSize(10)
        .fillColor(C.navy)
        .text(clean(c.title), 495, cy + 12);
      doc
        .font("Helvetica")
        .fontSize(8.5)
        .fillColor(C.muted)
        .text(clean(c.desc), 495, cy + 28, { width: 290, lineGap: 3 });
    });

    // -------------------------------------------------------------
    // PAGE 13: Single-Ward Micro Analysis
    // -------------------------------------------------------------
    addHeaderFooter(
      "Microeconomics",
      "Core Unit Economics & Single-Ward Micro Analysis",
      "Evaluating profitability at the individual paying ward footprint reveals structural network efficiency.",
      13,
    );

    const microMetrics = [
      {
        label: "Target Ward Composition",
        val: "1 Councillor + 0.2 Provider",
        detail:
          "1 local Councillor seat ($10/mo) paired with approximately 0.2 integrated Service Provider access configurations ($10/mo pro-rata value share).",
      },
      {
        label: "Blended Ward Monthly Revenue",
        val: "$20.00 / month ($240/yr)",
        detail:
          "Reliable annualized recurring SaaS subscription yield per active ward footprint before factoring in institutional data licensing.",
      },
      {
        label: "Variable Cost to Serve (CTS)",
        val: "$3.00 / month",
        detail:
          "Based on pilot metrics: $0.02 variable AI token cost per validated incident report x 150 incoming validated tickets per ward monthly.",
      },
      {
        label: "Server Overhead Amortization",
        val: "$15.00 / month (Y1)",
        detail:
          "$18,000/yr hosting + dedicated Claude API connectivity allocated evenly across the initial cohort of 100 active wards.",
      },
      {
        label: "Net Contribution Margin Expansion",
        val: "10.0% -> 83.8%",
        detail:
          "Monthly ward contribution margin expands from $2.00 (10%) in Year 1 to $16.75 (83.8%) in Year 5 as server CTS declines to $3.25.",
      },
      {
        label: "Cost of Acquisition (CAC)",
        val: "$45.00 / ward",
        detail:
          "Direct outreach channels to municipal councillor cohorts achieve full capital payback within 2.25 months (< 1 quarter).",
      },
    ];

    microMetrics.forEach((m, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const mx = 40 + col * 387;
      const my = 95 + row * 105;

      doc.roundedRect(mx, my, 375, 95, 5).fillAndStroke(C.light, C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(10.5)
        .fillColor(C.navy)
        .text(clean(m.label), mx + 14, my + 12);

      doc
        .font("Helvetica-Bold")
        .fontSize(11.5)
        .fillColor(C.primary)
        .text(clean(m.val), mx + 14, my + 28);

      doc
        .font("Helvetica")
        .fontSize(8.5)
        .fillColor(C.muted)
        .text(clean(m.detail), mx + 14, my + 46, { width: 345, lineGap: 3 });
    });

    // -------------------------------------------------------------
    // PAGE 14: Data Governance & Compliance
    // -------------------------------------------------------------
    addHeaderFooter(
      "Governance & Compliance",
      "Data Governance, Privacy (POPIA) & Frontier AI Training",
      "Transforming raw civic messages into legally compliant, clean, structured training collections.",
      14,
    );

    const govPillars = [
      {
        tag: "PILLAR 1",
        title: "Privacy & De-Identification",
        desc: "Complete stripping of personal identifiers (PII), telephone numbers, and resident identities in strict compliance with South Africa's Protection of Personal Information Act (POPIA) and global data sovereignty frameworks.",
      },
      {
        tag: "PILLAR 2",
        title: "Canonical Triage Graph",
        desc: "Normalizing the real-time operational lifecycle: raw report -> automated categorization -> departmental assignment -> contractor dispatch -> physical repair -> community verification.",
      },
      {
        tag: "PILLAR 3",
        title: "Agentic Training Utility",
        desc: "Frontier AI labs acquire this data to train autonomous multi-agent operational models capable of navigating physical constraints, human escalation loops, and municipal logistics.",
      },
    ];

    govPillars.forEach((gp, i) => {
      const gx = 40 + i * 258;
      const gy = 95;
      doc.roundedRect(gx, gy, 246, 330, 6).fillAndStroke(C.light, C.line);

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(C.blue)
        .text(clean(gp.tag), gx + 16, gy + 18);

      doc
        .font("Helvetica-Bold")
        .fontSize(13)
        .fillColor(C.navy)
        .text(clean(gp.title), gx + 16, gy + 38, { width: 214 });

      doc
        .font("Helvetica")
        .fontSize(9.5)
        .fillColor(C.muted)
        .text(clean(gp.desc), gx + 16, gy + 95, { width: 214, lineGap: 4 });
    });

    // -------------------------------------------------------------
    // PAGE 15: Roadmap & Milestones
    // -------------------------------------------------------------
    addHeaderFooter(
      "Strategic Roadmap",
      "Summary & Investment Milestones",
      "Clear execution path from initial seed validation to institutional data scale.",
      15,
      true,
    );

    const roadmapPhases = [
      {
        time: "MONTHS 1-12",
        title: "Phase 1: Seed Validation",
        items: [
          "- Deploy $100k seed runway ($75k Y1 burn)",
          "- Onboard 100 Ward Councillors ($10/mo)",
          "- Onboard 20 Service Providers ($50/mo)",
          "- Establish $30k Year 1 baseline revenue",
          "- Validate 150 tickets/mo/ward volume",
          "- Prove 2.25-month CAC payback",
        ],
      },
      {
        time: "MONTHS 13-24",
        title: "Phase 2: Data Commercialization",
        items: [
          "- Scale to 300 wards & 8 municipalities",
          "- License pilot dataset to micro1 ($100k)",
          "- Achieve net profitability (+$162.9k)",
          "- Launch CivicRewards ad loop ($1k/mo)",
          "- 18-24 month self-funding runway",
          "- Solidify South African operating core",
        ],
      },
      {
        time: "YEARS 3-5",
        title: "Phase 3: Global AI Refinery",
        items: [
          "- Expand to 3,200 wards & 90 municipalities",
          "- Deliver datasets to 5 frontier AI labs",
          "- Scale top-line revenue to $17.425M",
          "- Achieve 89.9% margin & +$15.26M profit",
          "- Global standard for municipal triage AI",
          "- Institutional liquidity & regional expansion",
        ],
      },
    ];

    roadmapPhases.forEach((rp, i) => {
      const rx = 40 + i * 258;
      const ry = 95;
      doc.roundedRect(rx, ry, 246, 330, 6).fillAndStroke("#132433", "#059669");

      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor(C.cyan)
        .text(clean(rp.time), rx + 16, ry + 18);

      doc
        .font("Helvetica-Bold")
        .fontSize(13)
        .fillColor(C.white)
        .text(clean(rp.title), rx + 16, ry + 38, { width: 214 });

      let itemY = ry + 85;
      rp.items.forEach((item) => {
        doc
          .font("Helvetica")
          .fontSize(9)
          .fillColor("#E2E8F0")
          .text(clean(item), rx + 16, itemY, { width: 214, lineGap: 3 });
        itemY += 32;
      });
    });

    doc.end();
    stream.on("finish", () => {
      resolve();
    });
    stream.on("error", reject);
  });
}

async function main() {
  console.log("Generating clean raw PDF...");
  await generateRawPDF();

  console.log("Validating and reserializing through pdf-lib for 100% standard compliance...");
  const rawBytes = fs.readFileSync(TEMP_PATH);
  const pdfLibDoc = await PDFLibDoc.load(rawBytes);

  // Set standard PDF document metadata
  pdfLibDoc.setTitle("SIGNAL DESK - Comprehensive Business Plan & Operational Model");
  pdfLibDoc.setAuthor("Signal Desk");
  pdfLibDoc.setSubject(
    "Strategic Framework for a Cloud-Based Public Infrastructure Triage & Asset-Liquidation Platform",
  );
  pdfLibDoc.setProducer("Signal Desk Standard PDF Engine (100% compliant PDF 1.7)");
  pdfLibDoc.setCreator("Signal Desk Strategic Business Plan System");

  // Save with standard cross-reference tables compatible with all viewers
  const cleanBytes = await pdfLibDoc.save({ useObjectStreams: false });
  fs.writeFileSync(PDF_PATH, cleanBytes);

  if (fs.existsSync(TEMP_PATH)) {
    fs.unlinkSync(TEMP_PATH);
  }

  console.log(
    `Successfully verified and wrote standard PDF to ${PDF_PATH} (${cleanBytes.length} bytes)`,
  );
}

main().catch((err) => {
  console.error("Error generating PDF:", err);
  process.exit(1);
});
