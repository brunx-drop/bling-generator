import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

/**
 * v1.1
 *
 * Alteração principal:
 * - Quando a planilha final passa de 1.000 linhas, o sistema divide automaticamente
 *   em partes de no máximo 1.000 linhas totais por arquivo.
 * - A divisão respeita blocos completos:
 *   Produto Pai + todos os Produtos Filhos/Variações.
 * - Nunca corta um produto pai no meio das variações.
 * - Não usa ZIP. Quando houver múltiplas partes, retorna JSON com os arquivos em base64
 *   para o frontend baixar um por um.
 */

type Color = {
  name: string;
  code: string;
};

type ProductRow = Record<string, any>;

type GeneratedRowType = "parent" | "variation";

type GeneratedRow = {
  type: GeneratedRowType;
  values: any[];
};

type ProductBlock = {
  codePai: string;
  rows: GeneratedRow[];
};

type RowStyleTemplate = {
  height?: number;
  cells: Partial<ExcelJS.Style>[];
};

const MAX_ROWS_PER_FILE = 1000;
const HEADER_ROWS_PER_FILE = 1;
const MAX_DATA_ROWS_PER_FILE = MAX_ROWS_PER_FILE - HEADER_ROWS_PER_FILE;

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

      for (let x = Number(a); x <= Number(b); x += 2) {
        arr.push(String(x));
      }

      return arr;
    }

    const adult = ["PP", "P", "M", "G", "GG", "XG", "G2", "G3", "G4", "G5", "G6", "G7"];
    const ia = adult.indexOf(a.toUpperCase());
    const ib = adult.indexOf(b.toUpperCase());

    if (ia !== -1 && ib !== -1) {
      return adult.slice(ia, ib + 1);
    }
  }

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
  const out: Color[] = [];
  const unknown: string[] = [];

  for (const p of String(raw)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)) {
    const c = map.get(norm(p));

    if (!c) {
      unknown.push(p);
    } else {
      out.push(c);
    }
  }

  return { out, unknown };
}

async function readUploadedXlsx(file: File) {
  const buf = Buffer.from(await file.arrayBuffer());
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];

  return XLSX.utils.sheet_to_json<ProductRow>(sheet, { defval: "" });
}

function applyTokens(values: any[], tokens: Record<string, any>) {
  return values.map((v) => {
    if (typeof v !== "string") return v;

    let out = v;

    for (const k in tokens) {
      out = out.split(k).join(String(tokens[k] ?? ""));
    }

    return out;
  });
}

function cloneStyle(style: Partial<ExcelJS.Style> | undefined): Partial<ExcelJS.Style> {
  if (!style) return {};

  return JSON.parse(JSON.stringify(style));
}

function captureRowStyle(row: ExcelJS.Row, columnCount: number): RowStyleTemplate {
  const cells: Partial<ExcelJS.Style>[] = [];

  for (let col = 1; col <= columnCount; col++) {
    cells[col] = cloneStyle(row.getCell(col).style);
  }

  return {
    height: row.height,
    cells,
  };
}

function applyRowStyle(row: ExcelJS.Row, template: RowStyleTemplate, columnCount: number) {
  row.height = template.height;

  for (let col = 1; col <= columnCount; col++) {
    row.getCell(col).style = cloneStyle(template.cells[col]);
  }
}

function buildProductBlocks(
  products: ProductRow[],
  colors: Color[],
  parentTemplate: any[],
  varTemplate: any[]
): ProductBlock[] {
  const blocks: ProductBlock[] = [];

  for (const p of products) {
    const codePai = String(p["Código Pai"] ?? "").trim();
    const marca = p["Marca"];
    const peca = p["Peça"];
    const estampa = p["Estampa"];

    const sizes = parseSizes(p["Tamanhos"]);
    const { out: cores, unknown } = parseColors(p["Cores"], colors);

    if (unknown.length) {
      throw new Error(`Cores inválidas: ${unknown.join(", ")}`);
    }

    if (!codePai) {
      throw new Error("Existe produto sem Código Pai na planilha de entrada.");
    }

    const rows: GeneratedRow[] = [];

    const parent = applyTokens([...parentTemplate], {
      PPPP: codePai,
      MCC: marca,
      PECA: peca,
      ESTAMPA: estampa,
    });

    parent[1] = codePai;
    parent[2] = `${marca} - ${peca} - ${estampa}`;

    rows.push({
      type: "parent",
      values: parent,
    });

    for (const c of cores) {
      for (const size of sizes) {
        const variation = applyTokens([...varTemplate], {
          PPPP: codePai,
          XXXX: c.code,
          CCCC: c.name,
          TAM: size,
        });

        variation[1] = `${codePai}-${c.code}-${size}`;
        variation[2] = `Cor:${c.name};Tamanho:${size}`;

        rows.push({
          type: "variation",
          values: variation,
        });
      }
    }

    blocks.push({
      codePai,
      rows,
    });
  }

  return blocks;
}

function splitBlocksIntoParts(blocks: ProductBlock[]) {
  const parts: ProductBlock[][] = [];
  let currentPart: ProductBlock[] = [];
  let currentRows = 0;

  for (const block of blocks) {
    const blockRows = block.rows.length;

    if (blockRows > MAX_DATA_ROWS_PER_FILE) {
      throw new Error(
        `O produto pai ${block.codePai} sozinho gera ${blockRows} linhas. ` +
          `Isso passa do limite de ${MAX_ROWS_PER_FILE} linhas por arquivo e não pode ser dividido sem quebrar a estrutura pai/filhos.`
      );
    }

    if (currentPart.length > 0 && currentRows + blockRows > MAX_DATA_ROWS_PER_FILE) {
      parts.push(currentPart);
      currentPart = [];
      currentRows = 0;
    }

    currentPart.push(block);
    currentRows += blockRows;
  }

  if (currentPart.length > 0) {
    parts.push(currentPart);
  }

  return parts;
}

async function createWorkbookBuffer(templatePath: string, blocks: ProductBlock[]) {
  const wb = new ExcelJS.Workbook();

  await wb.xlsx.readFile(templatePath);

  const ws = wb.worksheets[0];
  const columnCount = ws.columnCount;

  const parentStyle = captureRowStyle(ws.getRow(2), columnCount);
  const variationStyle = captureRowStyle(ws.getRow(3), columnCount);

  ws.spliceRows(2, Math.max(ws.rowCount - 1, 0));

  let rowCursor = 2;

  for (const block of blocks) {
    for (const generatedRow of block.rows) {
      const row = ws.getRow(rowCursor);

      row.values = generatedRow.values;

      if (generatedRow.type === "parent") {
        applyRowStyle(row, parentStyle, columnCount);
      } else {
        applyRowStyle(row, variationStyle, columnCount);
      }

      rowCursor++;
    }
  }

  return wb.xlsx.writeBuffer();
}

function getPartFileName(index: number) {
  return `BLING_IMPORT_parte_${String(index + 1).padStart(2, "0")}.xlsx`;
}

function bufferToBase64(buffer: unknown) {
  return Buffer.from(buffer as any).toString("base64");
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File;

    if (!file) {
      return new NextResponse("Arquivo não enviado", { status: 400 });
    }

    const colorsPath = path.join(process.cwd(), "config/colors.json");
    const templatePath = path.join(process.cwd(), "templates/PLANILHA PADRÃO BLING.xlsx");

    const colors = JSON.parse(fs.readFileSync(colorsPath, "utf8")) as Color[];
    const products = await readUploadedXlsx(file);

    const templateWb = new ExcelJS.Workbook();
    await templateWb.xlsx.readFile(templatePath);

    const templateWs = templateWb.worksheets[0];
    const parentTemplate = (templateWs.getRow(2).values as any[]).slice(1);
    const varTemplate = (templateWs.getRow(3).values as any[]).slice(1);

    const blocks = buildProductBlocks(products, colors, parentTemplate, varTemplate);
    const parts = splitBlocksIntoParts(blocks);

    if (parts.length === 1) {
      const buffer = await createWorkbookBuffer(templatePath, parts[0]);

      return new NextResponse(Buffer.from(buffer as any), {
        headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": 'attachment; filename="BLING_IMPORT.xlsx"',
          "X-Bling-Parts": "1",
        },
      });
    }

    const files = [];

    for (let i = 0; i < parts.length; i++) {
      const buffer = await createWorkbookBuffer(templatePath, parts[i]);

      files.push({
        fileName: getPartFileName(i),
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        base64: bufferToBase64(buffer),
      });
    }

    return NextResponse.json(
      {
        multiple: true,
        totalParts: files.length,
        files,
      },
      {
        headers: {
          "X-Bling-Parts": String(files.length),
        },
      }
    );
  } catch (err: any) {
    console.error(err);

    return new NextResponse(err.message || "Erro interno", { status: 500 });
  }
}
