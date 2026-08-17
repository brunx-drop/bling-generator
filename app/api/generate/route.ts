import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

/**
 * v1.5
 *
 * Alterações principais:
 * - Adiciona suporte a faixas de tamanhos agrupados:
 *   Exemplo: 33/34 ao 43/44
 *   Resultado: 33/34, 35/36, 37/38, 39/40, 41/42, 43/44
 * - Adiciona suporte a faixas decimais em passos de 0.25:
 *   Exemplo: 8.0 ao 8.5
 *   Resultado: 8.0, 8.25, 8.5
 * - Mantém suporte a faixas numéricas tradicionais em passos de 2:
 *   Exemplo: 36 ao 50
 *   Resultado: 36, 38, 40, 42, 44, 46, 48, 50
 * - Ajusta o tamanho usado no SKU:
 *   33/34 -> 3334
 *   8.0 -> 80
 *   8.25 -> 825
 *   8.5 -> 85
 * - O texto exibido da variação continua preservando o tamanho original.
 *
 * v1.4
 *
 * Alteração principal:
 * - Quando o tamanho for "Único", o SKU usa "0".
 *
 * v1.3
 *
 * Correção principal:
 * - Ajusta a dinâmica do SKU das variações.
 * - Antes: CódigoPai-CódigoCor-Tamanho
 *   Exemplo: 684172-0001-P
 * - Agora: CódigoPaiCódigoCorTamanho
 *   Exemplo: 6841720001P
 *
 * v1.2
 *
 * Correções principais:
 * - Filtra linhas vazias da planilha de entrada antes de gerar os produtos.
 * - Ignora lixo invisível em colunas fora do padrão, como espaços em H5, I20 etc.
 * - Mantém erro de "produto sem Código Pai" apenas quando a linha tem dados reais
 *   nas colunas esperadas, mas o Código Pai está vazio.
 * - Adiciona headers de no-store para reduzir risco de cache no navegador/Vercel.
 *
 * v1.1
 *
 * Alteração principal:
 * - Quando a planilha final passa de 1.000 linhas, o sistema divide automaticamente
 *   em partes de no máximo 1.000 linhas totais por arquivo.
 * - A divisão respeita blocos completos:
 *   Produto Pai + todos os Produtos Filhos/Variações.
 * - Nunca corta um produto pai no meio das variações.
 * - Não usa ZIP. Quando houver múltiplas partes, retorna JSON com os arquivos
 *   em base64 para o frontend baixar um por um.
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

const INPUT_COLUMNS = [
  "Código Pai",
  "Marca",
  "Peça",
  "Estampa",
  "Cores",
  "Tamanhos",
] as const;

function norm(s: string) {
  return String(s || "")
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function isBlankValue(value: any) {
  return String(value ?? "").trim() === "";
}

/**
 * Trata o tamanho especificamente para composição do SKU.
 *
 * Regra v1.5:
 * - "Único" -> "0"
 * - Remove "/" de tamanhos agrupados:
 *   33/34 -> 3334
 * - Remove "." de tamanhos decimais:
 *   8.0 -> 80
 *   8.25 -> 825
 *   8.5 -> 85
 * - Tamanhos alfabéticos continuam iguais:
 *   P -> P
 *   GG -> GG
 *   XG -> XG
 *
 * Importante:
 * O valor visual da variação não é alterado.
 * A limpeza acontece somente na montagem do SKU.
 */
function normalizeSizeForSku(size: string) {
  const rawSize = String(size ?? "").trim();

  if (norm(rawSize) === "unico") {
    return "0";
  }

  return rawSize.replace(/[/.]/g, "");
}

/**
 * Monta o SKU final da variação.
 *
 * Regra v1.5:
 * SKU = Código Pai + Código da Cor + Tamanho tratado para SKU
 *
 * Exemplos:
 * Código Pai: 85990003
 * Código Cor: 0001
 * Tamanho: 33/34
 * SKU final: 8599000300013334
 *
 * Código Pai: 85990006
 * Código Cor: 0000
 * Tamanho: 8.25
 * SKU final: 859900060000825
 *
 * Código Pai: 684172
 * Código Cor: 0001
 * Tamanho: Único
 * SKU final: 68417200010
 */
function buildVariationSku(codePai: string, colorCode: string, size: string) {
  const skuSize = normalizeSizeForSku(size);

  return `${String(codePai).trim()}${String(colorCode).trim()}${skuSize}`;
}

/**
 * Filtra linhas vazias reais da planilha de entrada.
 *
 * Importante:
 * O Excel/XLSX pode considerar linhas como "usadas" quando existe um espaço,
 * formatação antiga ou sujeira em colunas fora do padrão, por exemplo H5.
 *
 * Por isso, aqui olhamos somente as colunas oficiais da entrada:
 * Código Pai, Marca, Peça, Estampa, Cores e Tamanhos.
 *
 * - Linha 100% vazia nessas colunas: ignora.
 * - Linha com Marca/Peça/Cores/etc., mas sem Código Pai: mantém para a validação
 *   mostrar o erro correto.
 */
function removeEmptyInputRows(rows: ProductRow[]) {
  return rows.filter((row) => {
    return INPUT_COLUMNS.some((columnName) => !isBlankValue(row[columnName]));
  });
}

/**
 * Converte o campo "Tamanhos" da planilha em uma lista de variações.
 *
 * Formatos suportados:
 *
 * 1) Vazio
 *    -> ["Único"]
 *
 * 2) Único
 *    -> ["Único"]
 *
 * 3) Lista separada por vírgulas
 *    Exemplo:
 *    33/34, 35/36, 37/38
 *    -> ["33/34", "35/36", "37/38"]
 *
 * 4) Faixa numérica tradicional em passos de 2
 *    Exemplo:
 *    36 ao 50
 *    -> ["36", "38", "40", "42", "44", "46", "48", "50"]
 *
 * 5) Faixa agrupada
 *    Exemplo:
 *    33/34 ao 43/44
 *    -> ["33/34", "35/36", "37/38", "39/40", "41/42", "43/44"]
 *
 * 6) Faixa decimal em passos de 0.25
 *    Exemplo:
 *    8.0 ao 8.5
 *    -> ["8.0", "8.25", "8.5"]
 *
 * 7) Faixa alfabética adulta
 *    Exemplo:
 *    P ao GG
 *    -> ["P", "M", "G", "GG"]
 */
function parseSizes(raw: string) {
  const t = String(raw || "").trim();

  if (!t) return ["Único"];

  const n = norm(t);

  if (n === "unico" || n === "único") return ["Único"];

  /**
   * Lista explícita separada por vírgulas.
   *
   * Exemplos:
   * 33/34,35/36,37/38
   * 8.0,8.25,8.5
   * 36,38,40,42
   */
  if (t.includes(",")) {
    return t
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
  }

  const m = t.match(/^(.+?)\s+ao\s+(.+)$/i);

  if (m) {
    const a = m[1].trim();
    const b = m[2].trim();

    /**
     * Faixa agrupada:
     * 33/34 ao 43/44
     */
    const pairA = a.match(/^(\d+)\/(\d+)$/);
    const pairB = b.match(/^(\d+)\/(\d+)$/);

    if (pairA && pairB) {
      const a1 = Number(pairA[1]);
      const a2 = Number(pairA[2]);
      const b1 = Number(pairB[1]);
      const b2 = Number(pairB[2]);

      const isSequentialPair =
        a2 === a1 + 1 &&
        b2 === b1 + 1 &&
        b1 >= a1 &&
        (b1 - a1) % 2 === 0;

      if (isSequentialPair) {
        const arr: string[] = [];

        for (let x = a1; x <= b1; x += 2) {
          arr.push(`${x}/${x + 1}`);
        }

        return arr;
      }
    }

    /**
     * Faixa decimal em passos de 0.25:
     * 8.0 ao 8.5
     *
     * Pelo menos um dos lados precisa conter ponto decimal,
     * para não interferir na regra dos tamanhos inteiros.
     */
    const isDecimalRange =
      /^\d+(?:\.\d+)?$/.test(a) &&
      /^\d+(?:\.\d+)?$/.test(b) &&
      (a.includes(".") || b.includes("."));

    if (isDecimalRange) {
      const start = Number(a);
      const end = Number(b);

      if (Number.isFinite(start) && Number.isFinite(end) && end >= start) {
        const arr: string[] = [];

        /**
         * Trabalhamos em quartos para evitar erros de ponto flutuante.
         *
         * 8.00 -> 32 quartos
         * 8.25 -> 33 quartos
         * 8.50 -> 34 quartos
         */
        const startQuarter = Math.round(start * 4);
        const endQuarter = Math.round(end * 4);

        for (let q = startQuarter; q <= endQuarter; q++) {
          const value = q / 4;

          /**
           * Mantém uma casa decimal quando o valor for inteiro ou meio:
           * 8 -> 8.0
           * 8.5 -> 8.5
           *
           * Mantém duas casas quando for quarto:
           * 8.25 -> 8.25
           * 8.75 -> 8.75
           */
          if (Number.isInteger(value)) {
            arr.push(value.toFixed(1));
          } else if (Number.isInteger(value * 2)) {
            arr.push(value.toFixed(1));
          } else {
            arr.push(value.toFixed(2));
          }
        }

        return arr;
      }
    }

    /**
     * Faixa numérica inteira tradicional em passos de 2:
     * 36 ao 50
     */
    if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
      const arr: string[] = [];

      for (let x = Number(a); x <= Number(b); x += 2) {
        arr.push(String(x));
      }

      return arr;
    }

    /**
     * Faixa alfabética adulta:
     * PP ao G7
     */
    const adult = [
      "PP",
      "P",
      "M",
      "G",
      "GG",
      "XG",
      "G2",
      "G3",
      "G4",
      "G5",
      "G6",
      "G7",
    ];

    const ia = adult.indexOf(a.toUpperCase());
    const ib = adult.indexOf(b.toUpperCase());

    if (ia !== -1 && ib !== -1 && ib >= ia) {
      return adult.slice(ia, ib + 1);
    }
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

  const rawRows = XLSX.utils.sheet_to_json<ProductRow>(sheet, {
    defval: "",
  });

  return removeEmptyInputRows(rawRows);
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

function applyRowStyle(
  row: ExcelJS.Row,
  template: RowStyleTemplate,
  columnCount: number
) {
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

    if (!codePai) {
      throw new Error("Existe produto sem Código Pai na planilha de entrada.");
    }

    const sizes = parseSizes(p["Tamanhos"]);
    const { out: cores, unknown } = parseColors(p["Cores"], colors);

    if (unknown.length) {
      throw new Error(`Cores inválidas: ${unknown.join(", ")}`);
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

        /**
         * v1.3
         *
         * SKU sem hífen:
         * Código Pai + Código da Cor + Tamanho
         *
         * Exemplo:
         * 684172 + 0001 + P = 6841720001P
         */
        variation[1] = buildVariationSku(codePai, c.code, size);

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
        `O produto pai ${block.codePai} sozinho gera ${blockRows} linhas.\n` +
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

function noStoreHeaders(extraHeaders: Record<string, string> = {}) {
  return {
    "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
    Pragma: "no-cache",
    Expires: "0",
    ...extraHeaders,
  };
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File;

    if (!file) {
      return new NextResponse("Arquivo não enviado", {
        status: 400,
        headers: noStoreHeaders(),
      });
    }

    const colorsPath = path.join(process.cwd(), "config/colors.json");
    const templatePath = path.join(
      process.cwd(),
      "templates/PLANILHA PADRÃO BLING.xlsx"
    );

    const colors = JSON.parse(fs.readFileSync(colorsPath, "utf8")) as Color[];
    const products = await readUploadedXlsx(file);

    if (!products.length) {
      return new NextResponse("A planilha de entrada não possui produtos válidos.", {
        status: 400,
        headers: noStoreHeaders(),
      });
    }

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
        headers: noStoreHeaders({
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": 'attachment; filename="BLING_IMPORT.xlsx"',
          "X-Bling-Parts": "1",
        }),
      });
    }

    const files = [];

    for (let i = 0; i < parts.length; i++) {
      const buffer = await createWorkbookBuffer(templatePath, parts[i]);

      files.push({
        fileName: getPartFileName(i),
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        headers: noStoreHeaders({
          "X-Bling-Parts": String(files.length),
        }),
      }
    );
  } catch (err: any) {
    console.error(err);

    return new NextResponse(err.message || "Erro interno", {
      status: 500,
      headers: noStoreHeaders(),
    });
  }
}
