import { describe, it, expect, vi, beforeEach } from "vitest";

vi.mock("@/store/auth", () => ({
  useAuthStore: {
    getState: () => ({ token: "test-token-123", logout: vi.fn() }),
  },
}));

const mockFetch = vi.fn();
global.fetch = mockFetch;

const { apiRequest } = await import("./apiClient");

describe("apiRequest", () => {
  beforeEach(() => {
    mockFetch.mockReset();
  });

  it("отправляет запрос с токеном авторизации", async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve({ data: "test" }),
    });

    await apiRequest("/test");

    expect(mockFetch).toHaveBeenCalledWith(
      "/api/test",
      expect.objectContaining({
        headers: expect.objectContaining({
          Authorization: "Bearer test-token-123",
          "Content-Type": "application/json",
        }),
      })
    );
  });

  it("возвращает JSON из ответа", async () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve({ id: 1, name: "test" }),
    });

    const result = await apiRequest("/test");
    expect(result).toEqual({ id: 1, name: "test" });
  });

  it("бросает ошибку при не-ok ответе", async () => {
    mockFetch.mockResolvedValueOnce({
      ok: false,
      status: 500,
      statusText: "Internal Server Error",
      json: () => Promise.resolve({ error: "Server error" }),
    });

    await expect(apiRequest("/test")).rejects.toThrow("Server error");
  });
});
