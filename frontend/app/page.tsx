"use client";

import { useCallback, useRef, useState } from "react";

type Status = "idle" | "converting" | "done" | "error";

const API_URL = process.env.NEXT_PUBLIC_API_URL;
// const API_URL = "http://localhost:8000";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [downloadName, setDownloadName] = useState<string>("");
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  // ── file selection ──────────────────────────────────────────────────────
  const handleFile = (f: File) => {
    if (!f.name.toLowerCase().endsWith(".idml")) {
      setErrorMsg(".idml 파일만 업로드할 수 있습니다.");
      setStatus("error");
      return;
    }
    setFile(f);
    setStatus("idle");
    setErrorMsg("");
    setDownloadUrl("");
  };

  const handleDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }, []);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) handleFile(f);
  };

  // ── conversion ──────────────────────────────────────────────────────────
  const handleConvert = async () => {
    if (!file) return;
    setStatus("converting");
    setErrorMsg("");

    if (!API_URL) {
      setErrorMsg("NEXT_PUBLIC_API_URL is not configured.");
      setStatus("error");
      return;
    }

    const form = new FormData();
    form.append("file", file);

    try {
      const res = await fetch(`${API_URL}/convert`, {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const detail = await res
          .json()
          .then((j) => j.detail)
          .catch(() => res.statusText);
        throw new Error(detail);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const name = file.name.replace(/\.idml$/i, ".pptx");

      setDownloadUrl(url);
      setDownloadName(name);
      setStatus("done");
    } catch (err: unknown) {
      setErrorMsg(
        err instanceof Error ? err.message : "알 수 없는 오류가 발생했습니다.",
      );
      setStatus("error");
    }
  };

  // ── derived UI state ─────────────────────────────────────────────────────
  const dropzoneClass = [
    "dropzone",
    isDragging ? "active" : "",
    file && status !== "error" ? "has-file" : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <main>
      <div className="card">
        <h1>📄 교재 변환기</h1>
        <p className="subtitle">
          IDML 파일을 드래그하거나 클릭해서 업로드하세요
        </p>

        {/* Drop zone */}
        <div
          className={dropzoneClass}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
        >
          <div className="drop-icon">
            {file && status !== "error" ? "✅" : "📂"}
          </div>
          <p className="drop-label">
            {file && status !== "error"
              ? "파일이 선택되었습니다"
              : "여기에 파일을 놓아주세요"}
          </p>
          <p className="drop-hint">
            {file && status !== "error"
              ? ""
              : "또는 클릭해서 파일 선택 (.idml)"}
          </p>
          {file && status !== "error" && (
            <p className="filename">{file.name}</p>
          )}
        </div>

        <input
          ref={inputRef}
          type="file"
          accept=".idml"
          style={{ display: "none" }}
          onChange={handleInputChange}
        />

        {/* Convert button */}
        <button
          className="btn btn-primary"
          disabled={!file || status === "converting"}
          onClick={handleConvert}
        >
          {status === "converting" ? "변환 중…" : "PPT로 변환하기"}
        </button>

        {/* Download button */}
        {status === "done" && (
          <a className="btn-success" href={downloadUrl} download={downloadName}>
            ⬇️ &nbsp;{downloadName} 다운로드
          </a>
        )}

        {/* Status messages */}
        {status === "converting" && (
          <div className="status-bar status-converting">
            <span className="spinner" />
            변환 중입니다. 잠시만 기다려주세요…
          </div>
        )}
        {status === "error" && (
          <div className="status-bar status-error">⚠️ &nbsp;{errorMsg}</div>
        )}
      </div>
    </main>
  );
}
