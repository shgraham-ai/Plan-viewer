import type { Config } from "tailwindcss";

export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      boxShadow: {
        panel: "0 18px 45px rgba(15, 23, 42, 0.12)",
      },
      colors: {
        blueprint: {
          50: "#eff7ff",
          100: "#dceeff",
          200: "#b3ddff",
          300: "#7ec2ff",
          400: "#44a4ff",
          500: "#1b84ff",
          600: "#0a68db",
          700: "#0d54b2",
          800: "#124690",
          900: "#153d76",
        },
      },
    },
  },
  plugins: [],
} satisfies Config;

