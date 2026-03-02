import { twMerge } from "tailwind-merge";
import type { PieDiagramProps } from "../types";

export const PieLegend = ({ data, valueLabel, className }: PieDiagramProps) => {
  return (
    <div
      className={twMerge("flex flex-col justify-center space-y-3", className)}
    >
      {data.map((item, index) => (
        <div key={index} className="flex items-center gap-3">
          <div
            className="w-4 h-4 rounded"
            style={{
              backgroundColor: item.color,
            }}
          />
          <div className="flex-1">
            <div className="text-[18px] font-medium text-black">{item.name}</div>
            <div className="text-[16px] text-gray-600">
              {item.value} {valueLabel}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};
