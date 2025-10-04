"use client";

import { useState } from "react";

type Result = { filename: string; ok: boolean; error?: string; dataUrl?: string };

export default function Home() {
  const [results, setResults] = useState<Result[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);

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
    setIsUploading(true);
    setErrors([]);
    setResults([]);
    try {
      const form = new FormData();
      for (const f of files) form.append("files[]", f);
      const res = await fetch("/api/unlock", { method: "POST", body: form });
      const json = await res.json();
      if (!res.ok) {
        setErrors([json.error ?? "Upload failed"]);
        return;
      }
      setResults(json.results as Result[]);
    } catch (err) {
      setErrors(["Unexpected error"]);
    } finally {
      setIsUploading(false);
    }
  }

  return (
    <div className="min-h-screen p-6 sm:p-10 bg-gray-50">
      <div className="max-w-3xl mx-auto">
        <h1 className="text-2xl font-semibold mb-4">Microsoft Excel Protection Remover</h1>
        <p className="text-gray-600 mb-6">Remove worksheet/workbook protection from .xlsx/.xlsm files in your browser session.</p>

        <div
          onDragOver={(e) => e.preventDefault()}
          onDrop={onDrop}
          className="border-2 border-dashed rounded-lg p-8 text-center bg-white"
        >
          <p className="mb-3">Drag and drop files here</p>
          <p className="text-sm text-gray-500 mb-4">or</p>
          <label className="inline-block cursor-pointer text-white bg-black px-4 py-2 rounded">
            <input type="file" multiple accept=".xlsx,.xlsm" className="hidden" onChange={onChange} />
            Choose files
          </label>
        </div>

        {isUploading && <p className="mt-4 text-sm">Processing...</p>}

        {errors.length > 0 && (
          <div className="mt-4 text-red-600 text-sm">
            {errors.map((e, i) => (
              <div key={i}>{e}</div>
            ))}
          </div>
        )}

        {results.length > 0 && (
          <div className="mt-6 space-y-3">
            {results.map((r, idx) => (
              <div key={idx} className="flex items-center justify-between bg-white border rounded p-3">
                <div>
                  <div className="font-medium">{r.filename}</div>
                  {!r.ok && <div className="text-sm text-red-600">{r.error}</div>}
                </div>
                {r.ok && r.dataUrl && (
                  <a
                    href={`data:${r.dataUrl}`}
                    download={`${r.filename.replace(/\.(xlsx|xlsm)$/i, "")}_unlocked.xlsx`}
                    className="text-sm text-white bg-black px-3 py-1 rounded"
                  >
                    Download
                  </a>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
