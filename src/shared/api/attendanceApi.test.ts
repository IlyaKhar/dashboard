import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  loginRequest,
  fetchDashboardStats,
  fetchDrillDepartments,
} from "./attendanceApi";

// Re-export buildQueryString for testing if not exported - we need to test it
// Actually buildQueryString is not exported. Let me test the public API.

vi.mock("./apiClient", () => ({
  apiRequest: vi.fn(),
}));

const { apiRequest } = await import("./apiClient");

describe("loginRequest", () => {
  beforeEach(() => {
    vi.mocked(apiRequest).mockReset();
  });

  it("вызывает apiRequest с правильными параметрами", async () => {
    vi.mocked(apiRequest).mockResolvedValueOnce({
      token: "jwt-token",
      role: "admin",
    });

    const result = await loginRequest({
      username: "admin",
      password: "secret",
    });

    expect(apiRequest).toHaveBeenCalledWith("/login", {
      method: "POST",
      body: JSON.stringify({ username: "admin", password: "secret" }),
    });
    expect(result).toEqual({ token: "jwt-token", role: "admin" });
  });
});

describe("fetchDashboardStats", () => {
  beforeEach(() => {
    vi.mocked(apiRequest).mockReset();
  });

  it("вызывает /dashboard/stats", async () => {
    vi.mocked(apiRequest).mockResolvedValueOnce({
      totalStudents: 100,
      presentNow: 80,
      absentNow: 20,
      attendancePercent: 80,
    });

    await fetchDashboardStats();

    expect(apiRequest).toHaveBeenCalledWith("/dashboard/stats", { signal: undefined });
  });

  it("передаёт signal при наличии", async () => {
    const signal = new AbortController().signal;
    vi.mocked(apiRequest).mockResolvedValueOnce({});

    await fetchDashboardStats(signal);

    expect(apiRequest).toHaveBeenCalledWith("/dashboard/stats", { signal });
  });
});

describe("fetchDrillDepartments", () => {
  beforeEach(() => {
    vi.mocked(apiRequest).mockReset();
  });

  it("вызывает /attendance/drill/departments", async () => {
    vi.mocked(apiRequest).mockResolvedValueOnce([]);

    await fetchDrillDepartments();

    expect(apiRequest).toHaveBeenCalledWith("/attendance/drill/departments", {
      signal: undefined,
    });
  });
});
