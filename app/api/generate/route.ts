import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type Color = { name: string; code: string };

function norm(s: string) {
  return s.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function parseColors(raw: string, colors: Color[]) {
  const map = new Map(colors.map((c) => [norm(c.name), c]));
  const parts = String(raw || "").split(",").map(s => s.trim()).filter(Boolean);
  const out: Color[] = [];
  for (const p of parts) {
    const found = map.get(norm(p));
    if (found) out.push(found);
  }
  return out;
}

export async function POST(req: Request) {
  const form = await req.formData();
  const file = form.get("file") as File | null;
  if (!file) return new NextResponse("Arquivo não enviado", { status: 400 });

  const colors = JSON.parse(fs.readFileSync(path.join(process.cwd(), "config/colors.json"), "utf8")) as Color[];
  const template = path.join(process.cwd(), "templates", "PLANILHA PADRÃO BLING.xlsx");

  const ab = await file.arrayBuffer();
  const wbIn = XLSX.read(Buffer.from(ab), { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json<any>(wbIn.Sheets[wbIn.SheetNames[0]]);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(template);
  const ws = wb.worksheets[0];

  ws.spliceRows(2, ws.rowCount);

  let r = 2;
  for (const p of rows) {
    ws.insertRow(r++, [p["Código Pai"], `${p["Marca"]} - ${p["Peça"]} - ${p["Estampa"]}`]);
    for (const c of parseColors(p["Cores"], colors)) {
      ws.insertRow(r++, [`${p["Código Pai"]}${c.code}`, `Cor:${c.name}`]);
    }
  }

  const out = await wb.xlsx.writeBuffer();
  return new NextResponse(Buffer.from(out), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": "attachment; filename=BLING_IMPORT.xlsx",
    },
  });
}
