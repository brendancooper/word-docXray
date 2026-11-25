import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";

const mount = () => createRoot(document.getElementById("root")!).render(<App />);

if (window.Office) {
  Office.onReady(() => mount());
} else {
  mount(); // Enables browser-only UI development; Word actions explain their requirement.
}
