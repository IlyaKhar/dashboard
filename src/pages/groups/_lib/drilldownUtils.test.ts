import { describe, it, expect } from "vitest";
import {
  drillToAttendanceRows,
  getAttendanceBadge,
  getRowBgClass,
} from "./drilldownUtils";

describe("drillToAttendanceRows", () => {
  it("возвращает 6 строк: 4 с данными + 2 с NaN", () => {
    const rows = drillToAttendanceRows(100, 10);
    expect(rows).toHaveLength(6);
    expect(rows[0]).toEqual({ max: 100, total: 90 });
    expect(rows[4].total).toBe(Number.NaN);
    expect(rows[5].total).toBe(Number.NaN);
  });

  it("present = total - absent", () => {
    const rows = drillToAttendanceRows(50, 5);
    expect(rows[0]).toEqual({ max: 50, total: 45 });
  });
});

describe("getAttendanceBadge", () => {
  it("возвращает secondary для total=0", () => {
    expect(getAttendanceBadge(0, 0)).toEqual({ variant: "secondary", text: "—" });
  });

  it("возвращает default (>= 95%)", () => {
    const result = getAttendanceBadge(100, 3);
    expect(result.variant).toBe("default");
    expect(result.text).toBe("97%");
  });

  it("возвращает outline (90-94%)", () => {
    const result = getAttendanceBadge(100, 8);
    expect(result.variant).toBe("outline");
    expect(result.text).toBe("92%");
  });

  it("возвращает destructive (< 90%)", () => {
    const result = getAttendanceBadge(100, 15);
    expect(result.variant).toBe("destructive");
    expect(result.text).toBe("85%");
  });
});

describe("getRowBgClass", () => {
  it("пустая строка для total=0", () => {
    expect(getRowBgClass(0, 0)).toBe("");
  });

  it("зелёный >= 95%", () => {
    expect(getRowBgClass(100, 3)).toContain("green");
  });

  it("жёлтый 90-94%", () => {
    expect(getRowBgClass(100, 8)).toContain("yellow");
  });

  it("красный < 90%", () => {
    expect(getRowBgClass(100, 15)).toContain("red");
  });
});
