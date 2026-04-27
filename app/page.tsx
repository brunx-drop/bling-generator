"use client";

import { useState } from "react";

/**
 * v1.1
 *
 * Alteração principal:
 * - Quando o backend retorna uma planilha única, baixa BLING_IMPORT.xlsx.
 * - Quando o backend retorna múltiplas partes, baixa cada .xlsx em sequência:
 *   BLING_IMPORT_parte_01.xlsx
 *   BLING_IMPORT_parte_02.xlsx
 *   BLING_IMPORT_parte_03.xlsx
 * - Não usa ZIP.
 */

type GeneratedFile = {
  fileName: string;
  contentType: string;
  base64: string;
};

type MultipleFilesResponse = {
  multiple: true;
  totalParts: number;
  files: GeneratedFile[];
};

function getFileNameFromContentDisposition(contentDisposition: string | null) {
  if (!contentDisposition) return null;

  const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);

  if (utf8Match?.[1]) {
    return decodeURIComponent(utf8Match[1].replace(/"/g, ""));
  }

  const regularMatch = contentDisposition.match(/filename="?([^"]+)"?/i);

  if (regularMatch?.[1]) {
    return regularMatch[1].replace(/"/g, "");
  }

  return null;
}

function base64ToBlob(base64: string, contentType: string) {
  const byteCharacters = window.atob(base64);
  const byteArrays: Uint8Array[] = [];
  const sliceSize = 512;

  for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    const slice = byteCharacters.slice(offset, offset + sliceSize);
    const byteNumbers = new Array(slice.length);

    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }

    byteArrays.push(new Uint8Array(byteNumbers));
  }

  return new Blob(byteArrays, { type: contentType });
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");

  a.href = url;
  a.download = fileName;

  document.body.appendChild(a);
  a.click();
  a.remove();

  window.URL.revokeObjectURL(url);
}

async function downloadFilesInSequence(files: GeneratedFile[]) {
  for (const file of files) {
    const blob = base64ToBlob(file.base64, file.contentType);

    downloadBlob(blob, file.fileName);

    await new Promise((resolve) => setTimeout(resolve, 350));
  }
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState("");

  async function handleGenerate() {
    if (!file) {
      setMsg("Selecione uma planilha antes de gerar.");
      return;
    }

    setLoading(true);
    setMsg("");

    try {
      const form = new FormData();

      form.append("file", file);

      const res = await fetch("/api/generate", {
        method: "POST",
        body: form,
      });

      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || "Erro ao gerar planilha");
      }

      const contentType = res.headers.get("Content-Type") || "";

      if (contentType.includes("application/json")) {
        const data = (await res.json()) as MultipleFilesResponse;

        if (!data.files?.length) {
          throw new Error("Nenhum arquivo foi gerado.");
        }

        await downloadFilesInSequence(data.files);

        setMsg(`Gerado com sucesso ✅ ${data.totalParts} arquivos baixados em partes.`);
        return;
      }

      const blob = await res.blob();
      const contentDisposition = res.headers.get("Content-Disposition");
      const fileName = getFileNameFromContentDisposition(contentDisposition) || "BLING_IMPORT.xlsx";

      downloadBlob(blob, fileName);

      setMsg("Gerado com sucesso ✅");
    } catch (e: any) {
      setMsg(e?.message || "Erro ao gerar planilha");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        padding: 32,
        fontFamily: "Arial, sans-serif",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "#f5f5f5",
      }}
    >
      <section
        style={{
          width: "100%",
          maxWidth: 560,
          padding: 32,
          borderRadius: 16,
          background: "#ffffff",
          boxShadow: "0 10px 30px rgba(0,0,0,0.08)",
        }}
      >
        <h1 style={{ marginTop: 0, marginBottom: 12 }}>Gerador de Planilha Bling</h1>

        <p style={{ marginTop: 0, marginBottom: 24, color: "#555", lineHeight: 1.5 }}>
          Envie o Excel no padrão de entrada com as colunas Código Pai, Marca, Peça,
          Estampa, Cores e Tamanhos. Se a planilha final passar de 1.000 linhas,
          o sistema divide automaticamente em arquivos separados sem quebrar a
          estrutura de produto pai e variações.
        </p>

        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile(e.target.files?.[0] ?? null)}
          style={{
            display: "block",
            width: "100%",
            marginBottom: 20,
          }}
        />

        <button
          type="button"
          onClick={handleGenerate}
          disabled={!file || loading}
          style={{
            width: "100%",
            padding: "12px 16px",
            border: 0,
            borderRadius: 10,
            background: !file || loading ? "#999" : "#111",
            color: "#fff",
            fontWeight: 700,
            cursor: !file || loading ? "not-allowed" : "pointer",
          }}
        >
          {loading ? "Gerando..." : "Gerar planilha Bling"}
        </button>

        {msg && (
          <p
            style={{
              marginTop: 20,
              marginBottom: 0,
              color: msg.includes("sucesso") ? "#0a7a2f" : "#b00020",
              fontWeight: 600,
            }}
          >
            {msg}
          </p>
        )}

        <p style={{ marginTop: 18, marginBottom: 0, color: "#777", fontSize: 13, lineHeight: 1.4 }}>
          Observação: se o navegador pedir permissão para baixar vários arquivos,
          clique em permitir.
        </p>
      </section>
    </main>
  );
}
