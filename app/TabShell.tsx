"use client";

import { useState, useCallback, lazy, Suspense } from "react";

const RunCheck = lazy(() => import("./page-run-check"));
const DataManager = lazy(() => import("./data/page"));
const MultiBom = lazy(() => import("./multi-bom/page"));
const MpsImport = lazy(() => import("./mps/page"));
const Sourcing = lazy(() => import("./sourcing/page"));

const TABS = [
  { key: "run-check", label: "Run Check" },
  { key: "data", label: "Data Manager" },
  { key: "multi-bom", label: "Multi-BOM Analysis" },
  { key: "mps", label: "MPS Import" },
  { key: "sourcing", label: "Sourcing" },
] as const;

type TabKey = (typeof TABS)[number]["key"];

const fallback = <div className="max-w-5xl mx-auto p-6 text-sm text-gray-500">Loading…</div>;

export default function TabShell() {
  const [active, setActive] = useState<TabKey>("run-check");
  // Track which tabs have been visited so we lazy-mount on first visit
  const [mounted, setMounted] = useState<Set<TabKey>>(new Set(["run-check"]));

  const switchTab = useCallback((tab: TabKey) => {
    setActive(tab);
    setMounted((prev) => {
      if (prev.has(tab)) return prev;
      const next = new Set(prev);
      next.add(tab);
      return next;
    });
  }, []);

  return (
    <>
      <header className="border-b">
        <div className="max-w-5xl mx-auto p-4 flex items-center gap-4 text-sm">
          {TABS.map((t) => (
            <button
              key={t.key}
              onClick={() => switchTab(t.key)}
              className={`px-2 py-1 rounded transition-colors ${
                active === t.key
                  ? "bg-black text-white font-semibold"
                  : "hover:bg-gray-100 underline"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </header>
      <div className="flex-1">
        {mounted.has("run-check") && (
          <div style={{ display: active === "run-check" ? "block" : "none" }}>
            <Suspense fallback={fallback}><RunCheck /></Suspense>
          </div>
        )}
        {mounted.has("data") && (
          <div style={{ display: active === "data" ? "block" : "none" }}>
            <Suspense fallback={fallback}><DataManager /></Suspense>
          </div>
        )}
        {mounted.has("multi-bom") && (
          <div style={{ display: active === "multi-bom" ? "block" : "none" }}>
            <Suspense fallback={fallback}><MultiBom /></Suspense>
          </div>
        )}
        {mounted.has("mps") && (
          <div style={{ display: active === "mps" ? "block" : "none" }}>
            <Suspense fallback={fallback}><MpsImport /></Suspense>
          </div>
        )}
        {mounted.has("sourcing") && (
          <div style={{ display: active === "sourcing" ? "block" : "none" }}>
            <Suspense fallback={fallback}><Sourcing /></Suspense>
          </div>
        )}
      </div>
    </>
  );
}
