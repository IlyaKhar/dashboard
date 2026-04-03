import React from "react";
import { describe, it, expect } from "vitest";
import { render, screen } from "@testing-library/react";
import { AttendanceTable } from "./AttendanceTable";

describe("AttendanceTable", () => {
  const mockData = [
    { max: 20, total: 19 },
    { max: 20, total: 18 },
    { max: 20, total: 15 },
    { max: 20, total: NaN },
  ];

  it("рендерит заголовок таблицы", () => {
    render(<AttendanceTable attendance={mockData} header="Тест" />);
    expect(screen.getByText("Тест")).toBeInTheDocument();
  });

  it("рендерит номера пар", () => {
    render(<AttendanceTable attendance={mockData} header="Тест" />);
    expect(screen.getByText("1")).toBeInTheDocument();
    expect(screen.getByText("2")).toBeInTheDocument();
    expect(screen.getByText("3")).toBeInTheDocument();
    expect(screen.getByText("4")).toBeInTheDocument();
  });

  it("показывает --- для NaN", () => {
    render(<AttendanceTable attendance={mockData} header="Тест" />);
    expect(screen.getByText("---")).toBeInTheDocument();
  });

  it("показывает формат 'X из Y'", () => {
    render(<AttendanceTable attendance={mockData} header="Тест" />);
    expect(screen.getByText("19 из 20")).toBeInTheDocument();
    expect(screen.getByText("18 из 20")).toBeInTheDocument();
    expect(screen.getByText("15 из 20")).toBeInTheDocument();
  });
});
