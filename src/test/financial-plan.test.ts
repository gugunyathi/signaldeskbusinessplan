import { describe, it, expect } from "vitest";
import plan from "@/assets/plan-manifest.json";

describe("Updated Business Plan Financial Projections", () => {
  it("preserves the 5-year gross revenue projection", () => {
    expect(plan.financials.revenue).toEqual([30000, 315000, 1523000, 13261000, 17425000]);
  });

  it("validates Cost of Revenue (COR) and Gross Profit progression", () => {
    expect(plan.financials.cor).toEqual([22500, 62100, 186990, 1361520, 1760500]);
    expect(plan.financials.grossProfit).toEqual([7500, 252900, 1336010, 11899480, 15664500]);
    expect(plan.financials.grossMarginPct).toEqual([25.0, 80.3, 87.7, 89.7, 89.9]);
  });

  it("verifies operating expenses and net profit/loss turnaround", () => {
    expect(plan.financials.totalOpex).toEqual([57000, 90000, 155000, 260000, 405000]);
    expect(plan.financials.netProfitLoss).toEqual([-49500, 162900, 1181010, 11639480, 15259500]);
  });

  it("validates the $100,000 seed runway allocation", () => {
    const { runway } = plan.financials;
    expect(runway.seedCapital).toBe(100000);
    expect(runway.sunkAssets).toBe(30000);
    expect(runway.netCashWorkingCapital).toBe(70000);
    expect(runway.totalY1BaselineOutflows).toBe(75000);
    expect(runway.monthsRunwayMin).toBe(18);
    expect(runway.monthsRunwayMax).toBe(24);
  });

  it("validates single-ward microeconomics and cost decay curve", () => {
    const { unitEconomics } = plan.financials;
    expect(unitEconomics.blendedWardRevenue).toBe(20.0);
    expect(unitEconomics.totalCtsY1).toBe(18.0);
    expect(unitEconomics.totalCtsY5).toBe(3.25);
    expect(unitEconomics.cac).toBe(45.0);
  });
});
