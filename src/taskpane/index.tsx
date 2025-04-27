/* global Office */
import * as React from "react";
import * as ReactDOM from "react-dom/client";
import { App } from "./App";

Office.onReady(() => {
  const root = ReactDOM.createRoot(document.getElementById("app")!);
  root.render(<App />);
});
