import React from "react";
import { createRoot } from "react-dom/client";
import App from "./taskpane";

function render() {
  const rootEl = document.getElementById("root");
  if (!rootEl) throw new Error("No #root element found in HTML");
  const root = createRoot(rootEl);
  root.render(<App />);
}

if (window.Office) {
  // Office.js is present → wait for initialization
  Office.onReady(() => {
    render();
  });
} else {
  // Fallback for local dev in browser
  render();
}
