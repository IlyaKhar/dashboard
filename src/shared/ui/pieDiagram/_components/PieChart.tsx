import { useEffect, useState } from "react";
import { Pie, Tooltip, PieChart as PieRechart } from "recharts";
import type { PieDiagramData } from "../types";

export const PieChart = ({ data }: { data: PieDiagramData[] }) => {
  const [animation, setAnimation] = useState(true);

  useEffect(() => {
    const t = setTimeout(() => setAnimation(false), 2200);

    return () => clearTimeout(t);
  }, []);

  return (
    <PieRechart>
      <Pie
        data={data.map((item) => ({
          ...item,
          fill: item.color,
        }))}
        cx="50%"
        cy="50%"
        labelLine={false}
        label={({ percent = 0 }) =>
          percent === 0
            ? "0%"
            : Math.round(percent * 100) === 0
              ? "<1%"
              : `${Math.round(percent * 100)}%`
        }
        outerRadius={100}
        animationDuration={1100}
        isAnimationActive={animation}
      />
      <Tooltip />
    </PieRechart>
  );
};
