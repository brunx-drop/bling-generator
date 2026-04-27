"use client";

import { useState } from "react";

/**
 * v1.2
 *
 * Alterações:
 * - Retorna o visual simples/original da página, sem CSS extra.
 * - Mantém suporte para baixar múltiplos arquivos em sequência quando a geração
 *   passar de 1.000 linhas.
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

  return new Blob(byteArrays as BlobPart[], { type: contentType });
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
    if (!file) return;

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

        setMsg(`Gerado com sucesso ✅ ${data.totalParts} arquivos baixados.`);
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
    <main>
      <h1>Gerador de Planilha Bling</h1>

      <p>Envie o Excel no padrão de entrada (Código Pai, Marca, Peça, Estampa, Cores, Tamanhos).</p>

      <input
        type="file"
        accept=".xlsx,.xls"
        onChange={(e) => setFile(e.target.files?.[0] ?? null)}
      />

      <br />
      <br />

      <button type="button" onClick={handleGenerate} disabled={!file || loading}>
        {loading ? "Gerando..." : "Gerar planilha Bling"}
      </button>

      {msg && <p>{msg}</p>}
    </main>
  );
}
