const { tokens } = require("@fluentui/react-components");

/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./public/index.html", "./src/**/*.{js,jsx,ts,tsx}"],
  theme: {
    extend: {
      backgroundColor: {
        "teams-bg-1": tokens.colorNeutralBackground1,
        "teams-bg-3": tokens.colorNeutralBackground3,
      },
    },
  },
  plugins: [],
};
