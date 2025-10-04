"use client";

import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";

type Result = {
  id: string;
  filename: string;
  ok: boolean;
  error?: string;
  dataUrl?: string;
  status?: "queued" | "processing" | "done" | "error";
  progress?: number;
  sizeBytes?: number;
};

type ApiResult = { filename: string; ok: boolean; error?: string; dataUrl?: string };

export default function Home() {
  const [results, setResults] = useState<Result[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const hasDownloads = results.some((r) => r.ok && r.dataUrl);
  const timersRef = useRef<Record<string, number>>({});
  const activeBatchesRef = useRef(0);

  function startProgress(id: string, sizeBytes?: number) {
    // Estimate duration based on file size (MB): 1.5s base + 1.5s per MB, clamped 1.2s–20s
    const sizeMB = (sizeBytes ?? 0) / (1024 * 1024);
    const minMs = 1200;
    const maxMs = 20000;
    const expectedMs = Math.min(maxMs, Math.max(minMs, 1500 + sizeMB * 1500));
    const tickMs = 100;
    const steps = Math.max(1, Math.round(expectedMs / tickMs));
    const stepPerTick = Math.max(0.5, 90 / steps);

    const timer = window.setInterval(() => {
      setResults((prev) =>
        prev.map((r) =>
          r.id === id
            ? { ...r, progress: Math.min(typeof r.progress === "number" ? r.progress + stepPerTick : stepPerTick, 90) }
            : r
        )
      );
    }, tickMs);
    timersRef.current[id] = timer;
  }

  function stopProgress(id: string) {
    const t = timersRef.current[id];
    if (t) {
      clearInterval(t);
      delete timersRef.current[id];
    }
  }

  function stopAllProgress() {
    Object.values(timersRef.current).forEach((t) => clearInterval(t));
    timersRef.current = {};
  }

  useEffect(() => {
    return () => stopAllProgress();
  }, []);

  async function onDrop(e: React.DragEvent<HTMLDivElement>) {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files);
    await upload(files);
  }

  async function onChange(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files ?? []);
    await upload(files);
  }

  async function upload(files: File[]) {
    activeBatchesRef.current += 1;
    setIsUploading(true);
    setErrors([]);
    // Add placeholders immediately (append to any existing results)
    const placeholders: Result[] = files.map((f, idx) => ({
      id: `${Date.now()}-${idx}`,
      filename: f.name,
      ok: false,
      status: "processing",
      progress: 5,
      sizeBytes: f.size,
    }));
    setResults((prev) => [...prev, ...placeholders]);

    // Process each file concurrently, updating as results arrive
    await Promise.all(
      files.map(async (file, i) => {
        const id = placeholders[i].id;
        startProgress(id, files[i]?.size);
        try {
          const form = new FormData();
          form.append("files[]", file);
          const res = await fetch("/api/unlock", { method: "POST", body: form });
          const json: { results?: ApiResult[]; error?: string } = await res.json();
          if (!res.ok) {
            setResults((prev) =>
              prev.map((r) =>
                r.id === id ? { ...r, ok: false, status: "error", error: json.error || "Upload failed", progress: 100 } : r
              )
            );
            stopProgress(id);
            return;
          }
          const single = (json.results ?? [])[0];
          const ok = !!single?.ok;
          const dataUrl: string | undefined = single?.dataUrl;
          const error: string | undefined = single?.error;
          setResults((prev) =>
            prev.map((r) =>
              r.id === id
                ? {
                    ...r,
                    ok,
                    dataUrl,
                    error,
                    status: ok ? "done" : "error",
                    progress: 100,
                  }
                : r
            )
          );
          stopProgress(id);
        } catch {
          setResults((prev) =>
            prev.map((r) => (r.id === id ? { ...r, ok: false, status: "error", error: "Unexpected error", progress: 100 } : r))
          );
          stopProgress(id);
        }
      })
    );
    activeBatchesRef.current = Math.max(0, activeBatchesRef.current - 1);
    setIsUploading(activeBatchesRef.current > 0);
  }

  function downloadSingle(filename: string, dataUrl: string) {
    const link = document.createElement("a");
    const href = dataUrl.startsWith("data:") ? dataUrl : `data:${dataUrl}`;
    link.href = href;
    link.download = filename.replace(/\.(xlsx|xlsm)$/i, "") + "_unlocked.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  async function downloadAllZip() {
    const zip = new JSZip();
    for (const r of results) {
      if (!r.ok || !r.dataUrl) continue;
      const href = r.dataUrl.startsWith("data:") ? r.dataUrl : `data:${r.dataUrl}`;
      const base64 = href.split(",")[1] ?? "";
      const fileName = r.filename.replace(/\.(xlsx|xlsm)$/i, "") + "_unlocked.xlsx";
      zip.file(fileName, base64, { base64: true });
    }
    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "unlocked_excels.zip";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  return (
    <div className="min-h-screen bg-white">
      <header className="border-b bg-white">
        <div className="max-w-6xl mx-auto px-4 py-6">
          <h1 className="text-3xl font-semibold">Microsoft Excel Protection Remover</h1>
          <p className="text-gray-600 mt-1">Upload locked .xlsx/.xlsm files and instantly download unlocked copies. No files are stored.</p>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-8">
        <section className="grid md:grid-cols-3 gap-6">
          <div className="md:col-span-2 rounded-lg p-8 text-center border bg-gray-50 hover:bg-gray-100 transition"
               onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
            <p className="mb-2 font-medium">Drag and drop Excel files here</p>
            <p className="text-sm text-gray-500 mb-4">or</p>
            <label className="inline-block cursor-pointer text-white bg-black hover:bg-gray-800 px-4 py-2 rounded transition-colors transition-transform active:scale-95 focus:outline-none focus:ring-2 focus:ring-black/20">
              <input type="file" multiple accept=".xlsx,.xlsm" className="hidden" onChange={onChange} />
              Choose files
            </label>
            <div className="text-xs text-gray-500 mt-3">Supported: .xlsx, .xlsm • Max of 10MB</div>
            {isUploading && <p className="mt-4 text-sm text-gray-700">Processing...</p>}
            {errors.length > 0 && (
              <div className="mt-4 text-red-600 text-sm">
                {errors.map((e, i) => (
                  <div key={i}>{e}</div>
                ))}
              </div>
            )}
          </div>

          <div className="rounded-lg border p-4 bg-white">
            <h2 className="font-semibold mb-2">How it works</h2>
            <ol className="list-decimal ml-5 space-y-1 text-sm text-gray-700">
              <li>Upload your protected workbook(s)</li>
              <li>We remove sheet/workbook protection in-memory</li>
              <li>Download unlocked copies instantly</li>
            </ol>
            <div className="text-xs text-gray-500 mt-3">
              We never store your files. Processing happens in temporary memory only.
            </div>
          </div>
        </section>

        <section>
          <div className="flex items-center justify-between mb-3">
            <h2 className="font-semibold">Results</h2>
            <div className="flex gap-2">
              <button
                onClick={() => {
                  stopAllProgress();
                  setResults([]);
                }}
                className="text-sm border border-gray-300 rounded px-3 py-1 text-gray-700 hover:bg-gray-100 transition-colors transition-transform active:scale-95"
              >
                Clear
              </button>
              <button
                onClick={downloadAllZip}
                disabled={!hasDownloads}
                className={`text-sm text-white px-3 py-1 rounded transition-colors transition-transform active:scale-95 ${hasDownloads ? "bg-black hover:bg-gray-800" : "bg-gray-300 cursor-not-allowed"}`}
              >
                Download all as ZIP
              </button>
            </div>
          </div>

          <div className="space-y-3">
            {results.length === 0 && (
              <div className="border rounded p-4 text-gray-500 text-sm">Processed files will appear here.</div>
            )}
            {results.map((r) => (
              <div key={r.id} className="flex items-center justify-between bg-white border rounded p-3">
                <div>
                  <div className="font-medium text-gray-900">{r.filename}</div>
                  {r.status === "processing" && (
                    <div className="mt-2 h-2 w-48 bg-gray-200 rounded overflow-hidden">
                      <div
                        className="h-full bg-gray-600 transition-[width] duration-200"
                        style={{ width: `${r.progress ?? 0}%` }}
                      />
                    </div>
                  )}
                  {r.status === "error" && <div className="text-sm text-red-600">{r.error || "Error processing file"}</div>}
                </div>
                {r.ok && r.dataUrl && (
                  <button
                    onClick={() => downloadSingle(r.filename, r.dataUrl!)}
                    className="text-sm text-white bg-black hover:bg-gray-800 px-3 py-1 rounded transition-colors transition-transform active:scale-95"
                  >
                    Download
                  </button>
                )}
              </div>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}
