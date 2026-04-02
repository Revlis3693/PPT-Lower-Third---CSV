import React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { ErrorBoundary } from "./components/ErrorBoundary";
import "./styles/taskpane.css";

// Mount immediately. Do not wait for Office.onReady — the UI only needs Office when you run
// PowerPoint actions; waiting on onReady breaks plain-browser smoke tests and can delay paint.
const el = document.getElementById("root");
if (el) {
  createRoot(el).render(
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}

