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
  startedAt?: number;
  endedAt?: number;
  selected?: boolean;
};

type ApiResult = { filename: string; ok: boolean; error?: string; dataUrl?: string };

export default function Home() {
  const [results, setResults] = useState<Result[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const hasDownloads = results.some((r) => r.ok && r.dataUrl);
  const timersRef = useRef<Record<string, number>>({});
  const activeBatchesRef = useRef(0);
  const controllersRef = useRef<Record<string, AbortController>>({});
  const selectedCount = results.filter((r) => r.selected ?? true).length;
  const selectedDownloadable = results.filter((r) => (r.selected ?? true) && r.ok && r.dataUrl).length;

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
    const MAX_BYTES = 10 * 1024 * 1024; // 10MB
    const existingNames = new Set(results.map((r) => r.filename.toLowerCase()));
    const placeholders: Result[] = files.map((f, idx) => ({
      id: `${Date.now()}-${idx}`,
      filename: f.name,
      ok: false,
      status: "processing",
      progress: 5,
      sizeBytes: f.size,
      startedAt: Date.now(),
    }));
    // Validate size and duplicates; mark errors up-front
    const validated: Result[] = placeholders.map((p) => {
      if (p.sizeBytes && p.sizeBytes > MAX_BYTES) {
        return { ...p, status: "error", ok: false, error: "File exceeds 10MB limit", progress: 100 } as Result;
      }
      const lower = p.filename.toLowerCase();
      if (existingNames.has(lower)) {
        return { ...p, status: "error", ok: false, error: "Duplicate filename", progress: 100 } as Result;
      }
      existingNames.add(lower);
      return p;
    });
    setResults((prev) => [...prev, ...validated]);

    // Process each file concurrently, updating as results arrive
    await Promise.all(
      files.map(async (file, i) => {
        const id = placeholders[i].id;
        const placeholder = validated[i];
        if (placeholder.status === "error") return; // skip upload for invalid files
        startProgress(id, files[i]?.size);
        try {
          const form = new FormData();
          form.append("files[]", file);
          const controller = new AbortController();
          controllersRef.current[id] = controller;
          const res = await fetch("/api/unlock", { method: "POST", body: form, signal: controller.signal });
          const json: { results?: ApiResult[]; error?: string } = await res.json();
          if (!res.ok) {
            setResults((prev) =>
              prev.map((r) =>
                r.id === id
                  ? { ...r, ok: false, status: "error", error: json.error || "Upload failed", progress: 100, endedAt: Date.now() }
                  : r
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
                    endedAt: Date.now(),
                  }
                : r
            )
          );
          stopProgress(id);
        } catch {
          setResults((prev) =>
            prev.map((r) =>
              r.id === id ? { ...r, ok: false, status: "error", error: "Unexpected error", progress: 100, endedAt: Date.now() } : r
            )
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
    for (const r of results.filter((x) => x.ok && x.dataUrl && (x.selected ?? true))) {
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

  function removeFile(id: string) {
    // Abort if still processing, then remove entry
    const controller = controllersRef.current[id];
    if (controller) controller.abort();
    stopProgress(id);
    setResults((prev) => prev.filter((x) => x.id !== id));
  }

  function removeSelected() {
    // Abort any in-flight selected and remove
    const idsToRemove = results.filter((r) => r.selected ?? true).map((r) => r.id);
    for (const id of idsToRemove) {
      const c = controllersRef.current[id];
      if (c) c.abort();
      stopProgress(id);
    }
    setResults((prev) => prev.filter((x) => !(x.selected ?? true)));
  }

  return (
    <div className="min-h-screen bg-white">
      <header className="border-b bg-white">
        <div className="max-w-6xl mx-auto px-4 py-6">
          <h1 className="text-3xl font-semibold">Microsoft Excel Protection Remover</h1>
          <p className="text-gray-600 mt-1">Upload locked .xlsx/.xlsm files and instantly download unlocked copies. No files are stored.</p>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-4 md:py-8 space-y-6 md:space-y-8">
        <section className="grid md:grid-cols-3 gap-4 md:gap-6">
          <div className="md:col-span-2 rounded-lg p-4 md:p-8 text-center border bg-gray-50 hover:bg-gray-100 transition"
               onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
            {/* Desktop drag & drop */}
            <div className="hidden md:block">
              <p className="mb-2 font-medium">Drag and drop Excel files here</p>
              <p className="text-sm text-gray-500 mb-4">or</p>
            </div>
            {/* Mobile only button */}
            <div className="md:hidden mb-2">
              <p className="font-medium text-gray-900">Upload Excel files</p>
            </div>
            <label className="inline-block cursor-pointer text-white bg-black hover:bg-gray-800 px-3 py-2 md:px-4 rounded transition-colors transition-transform active:scale-95 focus:outline-none focus:ring-2 focus:ring-black/20 text-sm md:text-base">
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

          <div className="rounded-lg border p-3 md:p-4 bg-white">
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
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-3">
            <h2 className="font-semibold">Results</h2>
            <div className="flex gap-2 w-full sm:w-auto">
              <button
                onClick={() => {
                  stopAllProgress();
                  setResults([]);
                }}
                className="flex-1 sm:flex-none text-sm border border-gray-300 rounded px-3 py-1 text-gray-700 hover:bg-gray-100 transition-colors transition-transform active:scale-95"
              >
                Clear
              </button>
              <button
                onClick={downloadAllZip}
                disabled={!hasDownloads}
                className={`flex-1 sm:flex-none text-sm text-white px-3 py-1 rounded transition-colors transition-transform active:scale-95 ${hasDownloads ? "bg-black hover:bg-gray-800" : "bg-gray-300 cursor-not-allowed"}`}
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
              <div key={r.id} className="bg-white border rounded p-3">
                <div className="flex items-start gap-3 mb-3">
                  <input
                    type="checkbox"
                    checked={r.selected ?? true}
                    onChange={(e) => setResults((prev) => prev.map((x) => (x.id === r.id ? { ...x, selected: e.target.checked } : x)))}
                    className="accent-black mt-1"
                    title="Select for Download All"
                  />
                  <div className="flex-1 min-w-0">
                    <div className="font-medium text-gray-900">
                      <span className="break-words">{r.filename}</span>
                      {typeof r.sizeBytes === "number" && (
                        <span className="block sm:inline sm:ml-2 text-xs text-gray-500">({(r.sizeBytes / (1024 * 1024)).toFixed(2)} MB)</span>
                      )}
                    </div>
                  {r.status === "processing" && (
                    <div className="mt-2 h-2 w-full sm:w-48 bg-gray-200 rounded overflow-hidden">
                      <div
                        className="h-full bg-gray-600 transition-[width] duration-200"
                        style={{ width: `${r.progress ?? 0}%` }}
                      />
                    </div>
                  )}
                  {r.status === "error" && <div className="text-sm text-red-600">{r.error || "Error processing file"}</div>}
                  {r.status === "done" && r.startedAt && r.endedAt && (
                    <div className="text-xs text-gray-500 mt-1">Processed in {((r.endedAt - r.startedAt) / 1000).toFixed(2)}s</div>
                  )}
                  </div>
                </div>
                <div className="flex items-center justify-end gap-2">
                  {r.status === "processing" ? (
                    <button
                      onClick={() => {
                        const c = controllersRef.current[r.id];
                        if (c) c.abort();
                        stopProgress(r.id);
                        setResults((prev) => prev.map((x) => (x.id === r.id ? { ...x, status: "error", error: "Canceled", progress: 100 } : x)));
                      }}
                      className="text-sm text-gray-700 border border-gray-300 px-3 py-1 rounded hover:bg-gray-100 transition-colors flex-shrink-0"
                    >
                      Cancel
                    </button>
                  ) : (
                    <>
                      {r.ok && r.dataUrl && (
                        <button
                          onClick={() => downloadSingle(r.filename, r.dataUrl!)}
                          className="text-sm text-white bg-black hover:bg-gray-800 px-3 py-1 rounded transition-colors transition-transform active:scale-95 flex-shrink-0"
                        >
                          Download
                        </button>
                      )}
                      <button
                        onClick={() => removeFile(r.id)}
                        className="text-sm text-red-600 border border-red-200 px-3 py-1 rounded hover:bg-red-50 transition-colors flex-shrink-0"
                      >
                        Remove
                      </button>
                    </>
                  )}
                </div>
              </div>
            ))}
        </div>
        </section>
      </main>

      {/* Mobile sticky actions */}
      {results.length > 0 && (
        <div className="md:hidden fixed bottom-0 left-0 right-0 border-t bg-white/95 backdrop-blur supports-[backdrop-filter]:bg-white/80">
          <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
            <div className="text-xs text-gray-600">Selected: {selectedCount} • Ready: {selectedDownloadable}</div>
            <div className="flex gap-2">
              <button
                onClick={removeSelected}
                disabled={selectedCount === 0}
                className={`text-xs px-3 py-1 rounded border transition-colors ${selectedCount === 0 ? "text-gray-400 border-gray-200" : "text-gray-700 border-gray-300 hover:bg-gray-100"}`}
              >
                Remove selected
              </button>
              <button
                onClick={downloadAllZip}
                disabled={selectedDownloadable === 0}
                className={`text-xs text-white px-3 py-1 rounded transition-colors ${selectedDownloadable === 0 ? "bg-gray-300" : "bg-black hover:bg-gray-800"}`}
              >
                Download selected
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
