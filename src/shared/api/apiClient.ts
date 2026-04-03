import { useAuthStore } from "@/store/auth";
import { notifyApiError } from "./apiErrorHandler";

const API_BASE_URL = import.meta.env.VITE_API_URL || "/api";
const API_TIMEOUT_MS = 15_000;

export interface ApiRequestOptions extends Omit<RequestInit, "signal"> {
  signal?: AbortSignal;
}

/**
 * Выполняет запрос к API с timeout и поддержкой AbortSignal.
 * При 401 — сбрасывает авторизацию.
 */
export async function apiRequest<T>(
  endpoint: string,
  options: ApiRequestOptions = {}
): Promise<T> {
  const { signal: userSignal, ...fetchOptions } = options;

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), API_TIMEOUT_MS);

  if (userSignal) {
    userSignal.addEventListener("abort", () => controller.abort());
  }

  try {
    const token = useAuthStore.getState().token;
    const headers: HeadersInit = {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...fetchOptions.headers,
    };

    const response = await fetch(`${API_BASE_URL}${endpoint}`, {
      ...fetchOptions,
      headers,
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      if (response.status === 401) {
        useAuthStore.getState().logout();
      }

      let errorMessage = `Ошибка ${response.status}: ${response.statusText}`;
      try {
        const body = await response.json();
        if (body?.error) errorMessage = body.error;
      } catch {
        // ignore
      }

      const err = new Error(errorMessage);
      notifyApiError(err, endpoint);
      throw err;
    }

    return response.json();
  } catch (err) {
    clearTimeout(timeoutId);
    const error = err instanceof Error ? err : new Error(String(err));
    if (error.name === "AbortError") {
      const timeoutErr = new Error("Превышено время ожидания ответа");
      notifyApiError(timeoutErr, endpoint);
      throw timeoutErr;
    }
    notifyApiError(error, endpoint);
    throw error;
  }
}
