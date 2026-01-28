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

  const m = t.match(/^(.+?)\s+ao\s+(.+)$/i);
  if (m) {
    const a = m[1].trim();
    const b = m[2].trim();

    if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
      const arr: string[] = [];
      for (let x = Number(a); x <= Number(b); x += 2) arr.push(String(x));
      return arr;
    }

    const adult = ["PP", "P", "M", "G", "GG", "XG", "G2", "G3", "G4", "G5", "G6", "G7"];
    const ia = adult.indexOf(a.toUpperCase());
    const ib = adult.indexOf(b.toUpperCase());
    if (ia !== -1 && ib !== -1) return adult.slice(ia, ib + 1);
  }

  if (t.includes(",")) {
    return t.split(",").map(x => x.trim()).filter(Boolean);
  }

  return [t];
}

function parseColors(raw: string, colors: Color[]) {
  const map = new Map(colors.map(c => [norm(c.name), c]));
  const out: Color[] = [];
  const unknown: string[] = [];

  for (const p of String(raw).split(",").map(s => s.trim())) {
    const c = map.get(norm(p));
    if (!c) unknown.push(p);
    else out.push(c);
  }

  return { out, unknown };
}

async function readUploadedXlsx(file: File) {
  const buf = Buffer.from(await file.arrayBuffer());
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: "" });
}

function applyTokens(values: any[], tokens: Record<string, string>) {
  return values.map(v => {
    if (typeof v !== "string") return v;
    let out = v;
    for (const k in tokens) out = out.split(k).join(tokens[k]);
    return out;
  });
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File;
    if (!file) return new NextResponse("Arquivo não enviado", { status: 400 });

    const colorsPath = path.join(process.cwd(), "config/colors.json");
    const templatePath = path.join(process.cwd(), "templates/PLANILHA PADRÃO BLING.xlsx");

    const colors = JSON.parse(fs.readFileSync(colorsPath, "utf8")) as Color[];
    const products = await readUploadedXlsx(file);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const ws = wb.worksheets[0];

    const parentTemplate = (ws.getRow(2).values as any[]).slice(1);
    const varTemplate = (ws.getRow(3).values as any[]).slice(1);

    ws.spliceRows(2, ws.rowCount - 1);

    let rowCursor = 2;

    const addRow = (values: any[]) => {
      ws.getRow(rowCursor++).values = values;
    };

    for (const p of products) {
      const codePai = p["Código Pai"];
      const marca = p["Marca"];
      const peca = p["Peça"];
      const estampa = p["Estampa"];

      const sizes = parseSizes(p["Tamanhos"]);
      const { out: cores, unknown } = parseColors(p["Cores"], colors);
      if (unknown.length) {
        return new NextResponse(`Cores inválidas: ${unknown.join(", ")}`, { status: 400 });
      }

      // ===== PRODUTO PAI =====
      let parent = applyTokens([...parentTemplate], {
        PPPP: codePai,
        MCC: marca,
        PECA: peca,
        ESTAMPA: estampa,
      });

      parent[1] = codePai;
      parent[2] = `${marca} - ${peca} - ${estampa}`;

      addRow(parent);

      // ===== VARIAÇÕES =====
      for (const c of cores) {
        for (const size of sizes) {
          let v = applyTokens([...varTemplate], {
            PPPP: codePai,
            XXXX: c.code,
            CCCC: c.name,
            TAM: size,
          });

          v[1] = `${codePai}-${c.code}-${size}`;
          v[2] = `Cor:${c.name};Tamanho:${size}`;

          addRow(v);
        }
      }
    }

    const buffer = await wb.xlsx.writeBuffer();
    return new NextResponse(Buffer.from(buffer), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="BLING_IMPORT.xlsx"',
      },
    });
  } catch (err: any) {
    console.error(err);
    return new NextResponse(err.message || "Erro interno", { status: 500 });
  }
}
