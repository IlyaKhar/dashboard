import { ResponsiveContainer } from "recharts";
import type { PieDiagramProps } from "./types";
import { PieChart, PieLegend } from "./_components";
import { twMerge } from "tailwind-merge";

export const PieDiagram = ({
  data,
  valueLabel,
  className,
}: PieDiagramProps) => {
  return (
    <div
      className={twMerge("grid grid-cols-1 md:grid-cols-2 gap-6", className)}
    >
      <ResponsiveContainer width="100%" height={300}>
        <PieChart data={data} />
      </ResponsiveContainer>
      <PieLegend data={data} valueLabel={valueLabel} />
    </div>
  );
};
