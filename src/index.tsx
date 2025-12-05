import React from "react";
import { createRoot } from "react-dom/client";
import App from "./taskpane/app";

function render() {
  const rootEl = document.getElementById("root");
  if (!rootEl) throw new Error("No #root element found in HTML");
  const root = createRoot(rootEl);
  root.render(<App />);
}

if (window.Office) {
  Office.onReady(() => {
    render();
  });
} else {
  render();
}