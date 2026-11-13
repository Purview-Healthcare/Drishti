import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
// In the Microsoft sign-in pop-up, ms-guard.js hands the answer back and Drishti does not start.
window.__DRISHTI_MS_CALLBACK || createRoot(document.getElementById("root")).render(<App />);
