function formatDateForInput(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** Дата с данными для демо: должна совпадать с периодом в schedule.json (сейчас 23.02–01.03.2026) */
export const OPERATIONAL_FALLBACK_DATE = "2026-02-23";

const DEMO_RANGE_START = "2026-02-23";
const DEMO_RANGE_END = "2026-03-01";

/**
 * Возвращает дату для оперативного режима: «сегодня», если он в периоде демо-данных,
 * иначе OPERATIONAL_FALLBACK_DATE, чтобы таблицы по парам не были пустыми.
 */
export function getOperationalDate(): string {
  const today = formatDateForInput(new Date());
  if (today >= DEMO_RANGE_START && today <= DEMO_RANGE_END) return today;
  return OPERATIONAL_FALLBACK_DATE;
}
