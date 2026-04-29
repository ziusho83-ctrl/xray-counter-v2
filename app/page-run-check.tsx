"use client";

import { useState } from "react";

type CheckResult = {
  canRun: boolean;
  maxBuildable: number;
  source?: string;
  checkId?: string | null;
  shortages: Array<{ part: string; required: number; available: number; shortage: number }>;
};

export default function Home() {
  const [assembly, setAssembly] = useState("8E-03918-92");
  const [revision, setRevision] = useState("A");
  const [qty, setQty] = useState(1);
  const [result, setResult] = useState<CheckResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function runCheck() {
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const res = await fetch("/api/check", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ assembly, revision, qty }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.error || "Check failed");
      setResult(data);
    } catch (e: any) {
      setError(e.message || "Unknown error");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="max-w-4xl mx-auto p-6 space-y-6 w-full">
      <h1 className="text-2xl font-bold">XRAY Counter — Run Check</h1>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
        <input className="border rounded p-2" value={assembly} onChange={(e) => setAssembly(e.target.value)} placeholder="Assembly PN" />
        <input className="border rounded p-2" value={revision} onChange={(e) => setRevision(e.target.value)} placeholder="Revision" />
        <input className="border rounded p-2" type="number" min={1} value={qty} onChange={(e) => setQty(Number(e.target.value || 1))} placeholder="Build Qty" />
        <button onClick={runCheck} className="rounded bg-black text-white px-4 py-2">
          {loading ? "Checking..." : "Run Check"}
        </button>
      </div>

      {error && <div className="text-red-600">{error}</div>}

      {result && (
        <section className="space-y-3">
          <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
            <div className="border rounded p-3">Can Run: <b>{result.canRun ? "YES" : "NO"}</b></div>
            <div className="border rounded p-3">Max Buildable: <b>{result.maxBuildable}</b></div>
            <div className="border rounded p-3">Shortages: <b>{result.shortages.length}</b></div>
            <div className="border rounded p-3">Source: <b>{result.source || "n/a"}</b></div>
          </div>

          <div className="border rounded p-3">
            <h2 className="font-semibold mb-2">Shortage Details</h2>
            {result.shortages.length === 0 ? (
              <p>None</p>
            ) : (
              <ul className="space-y-1 text-sm">
                {result.shortages.map((s) => (
                  <li key={s.part}>- {s.part}: required={s.required}, available={s.available}, shortage={s.shortage}</li>
                ))}
              </ul>
            )}
          </div>
        </section>
      )}
    </main>
  );
}
