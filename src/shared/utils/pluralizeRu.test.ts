import { describe, it, expect } from "vitest";
import { pluralizeRu } from "./pluralizeRu";

const forms = { one: "студент", few: "студента", many: "студентов" };

describe("pluralizeRu", () => {
  it("возвращает one для 1", () => {
    expect(pluralizeRu(1, forms)).toBe("студент");
  });

  it("возвращает few для 2-4", () => {
    expect(pluralizeRu(2, forms)).toBe("студента");
    expect(pluralizeRu(3, forms)).toBe("студента");
    expect(pluralizeRu(4, forms)).toBe("студента");
  });

  it("возвращает many для 5-20", () => {
    expect(pluralizeRu(5, forms)).toBe("студентов");
    expect(pluralizeRu(11, forms)).toBe("студентов");
    expect(pluralizeRu(19, forms)).toBe("студентов");
  });

  it("корректно обрабатывает 21, 22, 25", () => {
    expect(pluralizeRu(21, forms)).toBe("студент");
    expect(pluralizeRu(22, forms)).toBe("студента");
    expect(pluralizeRu(25, forms)).toBe("студентов");
  });

  it("корректно обрабатывает 111, 112", () => {
    expect(pluralizeRu(111, forms)).toBe("студентов");
    expect(pluralizeRu(112, forms)).toBe("студентов");
  });

  it("корректно обрабатывает 0", () => {
    expect(pluralizeRu(0, forms)).toBe("студентов");
  });
});
