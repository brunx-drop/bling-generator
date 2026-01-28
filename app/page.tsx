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

      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || "Erro ao gerar planilha");
      }

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
    } catch (e: any) {
      setMsg(e?.message || "Erro ao gerar planilha");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main style={{ maxWidth: 860, margin: "40px auto", fontFamily: "system-ui" }}>
      <h1 style={{ marginBottom: 8 }}>Gerador de Planilha Bling</h1>

      <p style={{ marginTop: 0, opacity: 0.85 }}>
        Envie o Excel no padrão de entrada (Código Pai, Marca, Peça, Estampa, Cores, Tamanhos).
      </p>

      <input
        type="file"
        accept=".xlsx"
        onChange={(e) => setFile(e.target.files?.[0] ?? null)}
      />

      <div style={{ height: 12 }} />

      <button onClick={handleGenerate} disabled={!file || loading}>
        {loading ? "Gerando..." : "Gerar planilha Bling"}
      </button>

      {msg && (
        <div style={{ marginTop: 14, whiteSpace: "pre-wrap" }}>
          {msg}
        </div>
      )}
    </main>
  );
}
