import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type Color = { name: string; code: string };

function norm(s: string) {
  return String(s || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseSizes(raw: string) {
  const t = String(raw || "").trim();
  if (!t) return ["Único"];

  const n = norm(t);
  if (n === "unico" || n === "único") return ["Único"];

  // "PP ao G3" / "36 ao 46"
  const m = t.match(/^(.+?)\s+ao\s+(.+)$/i);
  if (m) {
    const a = m[1].trim();
    const b = m[2].trim();

    // numérico (passo 2)
    if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
      const start = Number(a);
      const end = Number(b);
      const arr: string[] = [];
      for (let x = start; x <= end; x += 2) arr.push(String(x));
      return arr;
    }

    // adulto
    const adult = ["PP", "P", "M", "G", "GG", "XG", "G2", "G3", "G4", "G5", "G6", "G7"];
    const A = a.toUpperCase();
    const B = b.toUpperCase();
    const ia = adult.indexOf(A);
    const ib = adult.indexOf(B);
    if (ia !== -1 && ib !== -1 && ia <= ib) return adult.slice(ia, ib + 1);
  }

  // lista "P, M, G"
  if (t.includes(",")) {
    return t
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
  }

  return [t];
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

async function readUploadedXlsx(file: File) {
  const ab = await file.arrayBuffer();
  const buf = Buffer.from(ab);
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: "" });
}

// Substitui tokens em qualquer célula string
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

/**
 * ✅ Header -> índice 0-based (A=0, B=1, C=2...)
 * Observação: ws.getRow(1).values vem 1-based com dummy no [0]
 */
function getColIndex(ws: ExcelJS.Worksheet) {
  const headers = ws.getRow(1).values as any[];
  const colIndex: Record<string, number> = {};
  headers.forEach((h, i) => {
    if (typeof h === "string" && h.trim()) colIndex[h.trim()] = i - 1; // <- 0-based
  });
  return colIndex;
}

function setByHeader(values0: any[], colIndex: Record<string, number>, header: string, value: any) {
  const idx = colIndex[header];
  if (idx === undefined || idx < 0) return;
  values0[idx] = value;
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File | null;
    if (!file) return new NextResponse("Arquivo não enviado.", { status: 400 });

    // colors.json
    const colorsPath = path.join(process.cwd(), "config", "colors.json");
    if (!fs.existsSync(colorsPath)) {
      return new NextResponse("Arquivo config/colors.json não encontrado.", { status: 400 });
    }
    const colors = JSON.parse(fs.readFileSync(colorsPath, "utf8")) as Color[];

    // template
    const templatePath = path.join(process.cwd(), "templates", "PLANILHA PADRÃO BLING.xlsx");
    if (!fs.existsSync(templatePath)) {
      return new NextResponse("Template não encontrado em templates/PLANILHA PADRÃO BLING.xlsx", {
        status: 400,
      });
    }

    // produtos (upload)
    const products = await readUploadedXlsx(file);
    if (!products.length) return new NextResponse("Planilha de entrada vazia.", { status: 400 });

    const required = ["Código Pai", "Marca", "Peça", "Estampa", "Cores", "Tamanhos"];
    const missing = required.filter((k) => !(k in (products[0] || {})));
    if (missing.length) {
      return new NextResponse(`Colunas ausentes na entrada: ${missing.join(", ")}`, { status: 400 });
    }

    // abre template
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.readFile(templatePath);
    const ws = outWb.worksheets[0];
    if (!ws) return new NextResponse("Template inválido (sem worksheet).", { status: 400 });

    if (ws.rowCount < 3) {
      return new NextResponse(
        "Template precisa ter 3 linhas: cabeçalho (1), modelo pai (2), modelo variação (3).",
        { status: 400 }
      );
    }

    const colIndex = getColIndex(ws);

    // modelos (IMPORTANTE: slice(1) remove dummy e vira 0-based)
    const baseParentRow = ws.getRow(2);
    const baseVarRow = ws.getRow(3);

    const baseParent = (baseParentRow.values as any[]).slice(1);
    const baseVar = (baseVarRow.values as any[]).slice(1);

    // estilos (opcional)
    const parentStyle =
      (baseParentRow as any)._cells?.map((c: any) => c?.style || null) || [];
    const varStyle =
      (baseVarRow as any)._cells?.map((c: any) => c?.style || null) || [];

    // remove tudo abaixo do header (inclui modelos)
    ws.spliceRows(2, ws.rowCount - 1);

    let rowCursor = 2;

    // ✅ Inserção correta: row.values no ExcelJS deve ser 1-based ao SETAR.
    // Então a gente seta com [null, ...full0] (dummy + 0-based values)
    const addRowFromTemplate = (template0: any[], applied0: any[], styles: any[]) => {
      const r = ws.getRow(rowCursor++);

      const full0 = [...template0];
      for (let i = 0; i < applied0.length; i++) full0[i] = applied0[i];

      // 1-based para ExcelJS
      r.values = [null, ...full0];

      // estilos
      if (styles?.length) {
        // styles é 1-based de cells, então aplica em cell(i+1)
        for (let i = 0; i < styles.length; i++) {
          if (styles[i]) r.getCell(i + 1).style = styles[i];
        }
      }

      (r as any).commit?.();
    };

    for (const p of products) {
      const codePai = String(p["Código Pai"]).trim();
      const marca = String(p["Marca"]).trim();
      const peca = String(p["Peça"]).trim();
      const estampa = String(p["Estampa"]).trim();
      const coresRaw = String(p["Cores"]).trim();
      const tamanhosRaw = String(p["Tamanhos"]).trim();

      if (!codePai) {
        return new NextResponse("Encontrado item sem 'Código Pai' na planilha de entrada.", {
          status: 400,
        });
      }

      const sizes = parseSizes(tamanhosRaw);

      const { out: colorsOk, unknown } = parseColors(coresRaw, colors);
      if (unknown.length) {
        return new NextResponse(
          `Cores não encontradas na base: ${unknown.join(", ")} (Código Pai ${codePai})`,
          { status: 400 }
        );
      }

      // ===== PAI =====
      const parentTokens = { PPPP: codePai, MCC: marca, PECA: peca, ESTAMPA: estampa };
      let parent0 = applyTokens([...baseParent], parentTokens);

      // reforço por header
      setByHeader(parent0, colIndex, "Código", codePai);
      setByHeader(parent0, colIndex, "Descrição", `${marca} - ${peca} - ${estampa}`);
      setByHeader(parent0, colIndex, "Código Pai", "");

      addRowFromTemplate(baseParent, parent0, parentStyle);

      // ===== VARIAÇÕES =====
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

          let var0 = applyTokens([...baseVar], varTokens);

          // reforço por header
          setByHeader(var0, colIndex, "Código", `${codePai}-${c.code}-${size}`);
          setByHeader(var0, colIndex, "Código Pai", codePai);

          addRowFromTemplate(baseVar, var0, varStyle);
        }
      }
    }

    const buffer = await outWb.xlsx.writeBuffer();
    return new NextResponse(Buffer.from(buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
