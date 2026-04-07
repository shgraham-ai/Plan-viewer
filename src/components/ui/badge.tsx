import * as React from "react";
import { cn } from "@/lib/utils";

type BadgeVariant = "default" | "secondary" | "destructive";

const variants: Record<BadgeVariant, string> = {
  default: "bg-amber-500 text-slate-950",
  secondary: "border border-slate-600 bg-slate-800 text-slate-300",
  destructive: "bg-red-900/70 text-red-100",
};

export interface BadgeProps extends React.HTMLAttributes<HTMLDivElement> {
  variant?: BadgeVariant;
}

export function Badge({ className, variant = "default", ...props }: BadgeProps) {
  return (
    <div
      className={cn(
        "inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium",
        variants[variant],
        className
      )}
      {...props}
    />
  );
}
