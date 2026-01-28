"use client";

import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState("");

  async function handleGenerate() {
    if (!file) return;
    setLoading(true);
    setMsg("");

    try {
      const form = new FormData();
      form.append("file", file);

      const res = await fetch("/api/generate", { method: "POST", body: form });
      if (!res.ok) throw new Error(await res.text());

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "BLING_IMPORT.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      setMsg("Gerado com sucesso ✅");
    } catch {
      setMsg("Erro ao gerar planilha");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main style={{ maxWidth: 760, margin: "40px auto", fontFamily: "system-ui" }}>
      <h1>Gerador de Planilha Bling</h1>
      <input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
      <div style={{ height: 12 }} />
      <button onClick={handleGenerate} disabled={!file || loading}>
        {loading ? "Gerando..." : "Gerar planilha Bling"}
      </button>
      <div style={{ marginTop: 12 }}>{msg}</div>
    </main>
  );
}
