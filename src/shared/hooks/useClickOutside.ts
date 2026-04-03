import { useEffect, type RefObject } from "react";

export function useClickOutside(
  detectionRef: RefObject<HTMLElement | null>,
  outsideAction: () => void
) {
  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      if (
        detectionRef.current &&
        !detectionRef.current.contains(e.target as Node)
      ) {
        outsideAction();
      }
    };

    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, [detectionRef, outsideAction]);
}
