import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type Color = { name: string; code: string };

function norm(s: string) {
  return s
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseColors(raw: string, colors: Color[]) {
  const map = new Map(colors.map((c) => [norm(c.name), c]));
  const parts = String(raw || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  const out: Color[] = [];
  const unknown: string[] = [];

  for (const p of parts) {
    const found = map.get(norm(p));
    if (!found) unknown.push(p);
    else out.push(found);
  }

  return { out, unknown };
}

function parseSizes(raw: string) {
  const t = String(raw || "").trim();
  if (!t) return ["Único"];

  // Único
  if (norm(t) === "unico" || norm(t) === "único") return ["Único"];

  // formato "PP ao G3" ou "36 ao 46"
  const m = t.match(/^(.+?)\s+ao\s+(.+)$/i);
  if (m) {
    const a = m[1].trim();
    const b = m[2].trim();

    // numérico (assumindo passo 2)
    if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
      const start = Number(a);
      const end = Number(b);
      const arr: string[] = [];
      for (let x = start; x <= end; x += 2) arr.push(String(x));
      return arr;
    }

    // adulto
    const adult = ["PP","P","M","G","GG","XG","G2","G3","G4","G5","G6","G7"];
    const A = a.toUpperCase();
    const B = b.toUpperCase();
    const ia = adult.indexOf(A);
    const ib = adult.indexOf(B);
    if (ia !== -1 && ib !== -1 && ia <= ib) return adult.slice(ia, ib + 1);
  }

  // lista separada por vírgula (ex: "P, M, G")
  if (t.includes(",")) {
    return t.split(",").map((x) => x.trim()).filter(Boolean);
  }

  return [t];
}

async function readUploadedXlsx(file: File) {
  const ab = await file.arrayBuffer();
  const buf = Buffer.from(ab);
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: "" });
}

// Substitui tokens em qualquer célula de TEXTO
function applyTokens(values: any[], tokens: Record<string, string>) {
  return values.map((v) => {
    if (typeof v !== "string") return v;
    let out = v;
    for (const [k, val] of Object.entries(tokens)) {
      out = out.split(k).join(val);
    }
    return out;
  });
}

// Mapeia header -> índice da coluna (1-based no ExcelJS row.values)
function getColIndex(ws: ExcelJS.Worksheet) {
  const headers = ws.getRow(1).values as any[];
  const colIndex: Record<string, number> = {};
  headers.forEach((h, i) => {
    if (typeof h === "string" && h.trim()) colIndex[h.trim()] = i;
  });
  return colIndex;
}

function setByHeader(values: any[], colIndex: Record<string, number>, header: string, value: any) {
  const idx = colIndex[header];
  if (!idx) return;
  values[idx] = value;
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File | null;
    if (!file) return new NextResponse("Arquivo não enviado.", { status: 400 });

    const colorsPath = path.join(process.cwd(), "config", "colors.json");
    if (!fs.existsSync(colorsPath)) {
      return new NextResponse("Arquivo config/colors.json não encontrado.", { status: 400 });
    }
    const colors = JSON.parse(fs.readFileSync(colorsPath, "utf8")) as Color[];

    const templatePath = path.join(process.cwd(), "templates", "PLANILHA PADRÃO BLING.xlsx");
    if (!fs.existsSync(templatePath)) {
      return new NextResponse(
        "Template não encontrado em templates/PLANILHA PADRÃO BLING.xlsx",
        { status: 400 }
      );
    }

    const products = await readUploadedXlsx(file);

    const required = ["Código Pai", "Marca", "Peça", "Estampa", "Cores", "Tamanhos"];
    const missing = required.filter((k) => !(k in (products[0] || {})));
    if (missing.length) {
      return new NextResponse(`Colunas ausentes: ${missing.join(", ")}`, { status: 400 });
    }

    // Abre template
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.readFile(templatePath);
    const ws = outWb.worksheets[0];
    if (!ws) return new NextResponse("Template inválido (sem planilha).", { status: 400 });

    // Precisa ter linha 2 (pai) e 3 (variação) como modelo
    if (ws.rowCount < 3) {
      return new NextResponse(
        "Template precisa ter pelo menos 3 linhas: cabeçalho (1), modelo pai (2), modelo variação (3).",
        { status: 400 }
      );
    }

    const colIndex = getColIndex(ws);

    const baseParent = [...(ws.getRow(2).values as any[])];
    const baseVar = [...(ws.getRow(3).values as any[])];

    // Limpa tudo abaixo do header, mas mantém os modelos em memória
    ws.spliceRows(2, ws.rowCount - 1);

    let rowCursor = 2;

    for (const p of products) {
      const codePai = String(p["Código Pai"]).trim();
      const marca = String(p["Marca"]).trim();
      const peca = String(p["Peça"]).trim();
      const estampa = String(p["Estampa"]).trim();
      const coresRaw = String(p["Cores"]).trim();
      const tamanhosRaw = String(p["Tamanhos"]).trim();

      if (!codePai) {
        return new NextResponse("Encontrado produto sem 'Código Pai'.", { status: 400 });
      }

      const sizes = parseSizes(tamanhosRaw);

      const { out: colorsOk, unknown } = parseColors(coresRaw, colors);
      if (unknown.length) {
        return new NextResponse(
          `Cores não encontradas na base: ${unknown.join(", ")} (Código Pai ${codePai})`,
          { status: 400 }
        );
      }

      // ---- PAI (linha modelo 2) ----
      const parentTokens = {
        PPPP: codePai,
        MCC: marca,
        PECA: peca,
        ESTAMPA: estampa,
      };

      let parentVals = applyTokens([...baseParent], parentTokens);

      // reforço opcional por cabeçalho (se existirem)
      setByHeader(parentVals, colIndex, "Código", codePai);
      setByHeader(parentVals, colIndex, "Código Pai", "");
      // se você usa "Descrição" no template
      setByHeader(parentVals, colIndex, "Descrição", `${marca} - ${peca} - ${estampa}`);

      ws.insertRow(rowCursor++, parentVals);

      // ---- VARIAÇÕES (linha modelo 3) ----
      for (const c of colorsOk) {
        for (const size of sizes) {
          const varTokens = {
            PPPP: codePai,
            XXXX: c.code,
            CCCC: c.name,
            MCC: marca,
            PECA: peca,
            ESTAMPA: estampa,
            TAM: String(size),
          };

          let varVals = applyTokens([...baseVar], varTokens);

          // reforço opcional por cabeçalho
          setByHeader(varVals, colIndex, "Código", `${codePai}-${c.code}-${size}`); // padrão B
          setByHeader(varVals, colIndex, "Código Pai", codePai);

          ws.insertRow(rowCursor++, varVals);
        }
      }
    }

    const buffer = await outWb.xlsx.writeBuffer();
    return new NextResponse(Buffer.from(buffer), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="BLING_IMPORT.xlsx"',
      },
    });
  } catch (e: any) {
    console.error(e);
    return new NextResponse(e?.message ? String(e.message) : "Erro interno na geração", {
      status: 500,
    });
  }
}
