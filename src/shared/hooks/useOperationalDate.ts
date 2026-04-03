function formatDateForInput(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function getOperationalDate(): string {
  const today = formatDateForInput(new Date());
  // В боевом режиме всегда работаем с сегодняшней датой.
  return today;
}
