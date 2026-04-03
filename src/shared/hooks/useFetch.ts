import { useState, useEffect, useCallback, useRef } from "react";

interface UseFetchState<T> {
  data: T | null;
  isLoading: boolean;
  error: string | null;
}

interface UseFetchOptions {
  pollingInterval?: number;
}

/** Фетчер может принимать AbortSignal для отмены при unmount */
export type Fetcher<T> = (signal?: AbortSignal) => Promise<T>;

export function useFetch<T>(
  fetcher: Fetcher<T>,
  deps: unknown[] = [],
  options: UseFetchOptions = {}
): UseFetchState<T> & { refetch: () => void } {
  const { pollingInterval = 0 } = options;

  const [state, setState] = useState<UseFetchState<T>>({
    data: null,
    isLoading: true,
    error: null,
  });

  const lastDataRef = useRef<T | null>(null);
  const controllerRef = useRef<AbortController | null>(null);
  const fetcherRef = useRef(fetcher);
  fetcherRef.current = fetcher;

  const load = useCallback(
    async (signal?: AbortSignal) => {
      const controller = new AbortController();
      controllerRef.current = controller;
      const s = signal ?? controller.signal;

      setState((prev) => ({
        data: prev.data ?? lastDataRef.current,
        isLoading: true,
        error: null,
      }));

      try {
        const result = await fetcherRef.current(s);
        if (controller.signal.aborted) return;

        lastDataRef.current = result;
        setState({ data: result, isLoading: false, error: null });
      } catch (err) {
        if (controller.signal.aborted) return;

        const message =
          err instanceof Error ? err.message : "Произошла ошибка при загрузке";
        setState({
          data: lastDataRef.current,
          isLoading: false,
          error: message,
        });
      } finally {
        controllerRef.current = null;
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps -- deps передаётся вызывающим кодом
    deps
  );

  useEffect(() => {
    const controller = new AbortController();
    load(controller.signal);
    return () => {
      controller.abort();
      controllerRef.current?.abort();
    };
  }, [load]);

  useEffect(() => {
    if (!pollingInterval || pollingInterval <= 0) return;

    const interval = setInterval(() => {
      load();
    }, pollingInterval);

    return () => clearInterval(interval);
  }, [load, pollingInterval]);

  const refetch = useCallback(() => load(), [load]);

  return { ...state, refetch };
}
