export interface PluralizeRuFormats {
  one: string;
  few: string;
  many: string;
}

export const pluralizeRu = (value: number, formats: PluralizeRuFormats): string => {
  const abs = Math.abs(value) % 100;
  const last = abs % 10;

  if (abs > 10 && abs < 20) return formats.many;
  if (last === 1) return formats.one;
  if (last >= 2 && last <= 4) return formats.few;
  return formats.many;
}