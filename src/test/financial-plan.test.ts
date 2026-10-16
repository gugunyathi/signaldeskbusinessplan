import { describe, it, expect } from "vitest";
import plan from "@/assets/plan-manifest.json";

describe("Source financial rules", () => {
  it("preserves the supplied five-year total revenues", () => {
    expect(plan.financials.revenue).toEqual([7700, 41000, 117500, 287000, 581000]);
  });
  it("preserves the supplied EBITDA forecast", () => {
    expect(plan.financials.ebitda).toEqual([-248400, -268500, -254600, -209500, 5000]);
  });
  it("removes $150,000 Year-5 councillor revenue for free core access", () => {
    const revenue = plan.financials.revenue[4];
    const freeRevenue = plan.financials.freeCoreRevenue[4];
    if (revenue === undefined || freeRevenue === undefined)
      throw new Error("Missing Year-5 projection");
    expect(revenue - freeRevenue).toBe(150000);
    expect(plan.financials.freeCoreEbitda[4]).toBe(-145000);
  });
});
