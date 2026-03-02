/**
 * Централизованная обработка ошибок API.
 * Используется apiClient для уведомления о любых ошибках запросов.
 */

export type ApiErrorHandlerFn = (error: Error, endpoint: string) => void;

let globalHandler: ApiErrorHandlerFn | null = null;

/** Устанавливает глобальный обработчик ошибок API */
export function setApiErrorHandler(h: ApiErrorHandlerFn | null): void {
  globalHandler = h;
}

/** Вызывается из apiClient при ошибке запроса */
export function notifyApiError(error: Error, endpoint: string): void {
  globalHandler?.(error, endpoint);
}
