import { twMerge } from "tailwind-merge";
import type { TableProps } from "./types";

export default function Table({ data, className }: TableProps) {
  return (
    <div className={twMerge("overflow-auto rounded-md", className)}>
      <table
        className={twMerge("w-full border-collapse text-sm", data.className)}
      >
        {data.header && (
          <thead className={data.header.className}>
            {data.header.rows.map((row, rowIndex) => (
              <tr key={rowIndex} className={row.className}>
                {row.cells.map((cell, cellIndex) => (
                  <cell.type
                    key={cellIndex}
                    className={twMerge(
                      "px-3 py-2 text-center font-medium",
                      "text-black border-gray-800!",
                      cellIndex + 1 < row.cells.length ? "border-r-2" : "",
                      (data.header && rowIndex + 1 < data.header.rows.length) ||
                        data.body
                        ? "border-b-2"
                        : "",
                      cell.className
                    )}
                  >
                    {cell.text}
                  </cell.type>
                ))}
              </tr>
            ))}
          </thead>
        )}

        {data.body && (
          <tbody className={data.body.className}>
            {data.body.rows.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                className={row.onClick ? "cursor-pointer " + (row.className ?? "") : row.className}
                onClick={row.onClick}
                role={row.onClick ? "button" : undefined}
              >
                {row.cells.map((cell, cellIndex) => (
                  <cell.type
                    key={cellIndex}
                    className={twMerge(
                      "px-3 py-2 text-center text-black border-gray-800!",
                      cellIndex + 1 < row.cells.length ? "border-r-2" : "",
                      data.body && rowIndex + 1 < data.body.rows.length
                        ? "border-b-2"
                        : "",
                      cell.className
                    )}
                  >
                    {cell.text}
                  </cell.type>
                ))}
              </tr>
            ))}
          </tbody>
        )}
      </table>
    </div>
  );
}

