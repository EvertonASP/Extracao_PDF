#!/usr/bin/env node
/**
 * Conversor de PDF (DARF Comprovante de Arrecadação) -> Excel
 * Uso: node converter-darf-para-excel.js <entrada.pdf|glob> <saida.xlsx>
 *
 * Dependências: pdf-parse, exceljs
 */

const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');

// Utilidades -------------------------------------------------------------

const brToNumber = (s) => {
  if (s == null) return null;
  const clean = String(s).replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const n = Number(clean);
  return Number.isFinite(n) ? n : null;
};

const normalizeSpaces = (s) => s.replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ').trim();

const joinWrappedLines = (lines) => {
  // Junta linhas que pertencem ao mesmo item: um item novo começa com 4 dígitos (código)
  const out = [];
  for (const raw of lines) {
    const line = normalizeSpaces(raw);
    if (!line) continue;
    if (/^\d{4}\b/.test(line)) {
      out.push(line);
    } else {
      // linha de continuação (descrição quebrada)
      if (out.length === 0) {
        out.push(line);
      } else {
        out[out.length - 1] = normalizeSpaces(out[out.length - 1] + ' ' + line);
      }
    }
  }
  return out;
};

const extractBetween = (text, startPattern, endPattern) => {
  const start = text.search(startPattern);
  if (start === -1) return null;
  const after = text.slice(start);
  const end = endPattern ? after.search(endPattern) : -1;
  return end === -1 ? after : after.slice(0, end);
};

// Parsing por página ------------------------------------------------------

function parsePageText(pageText) {
  const text = pageText.replace(/\r/g, '').replace(/[ \t]+\n/g, '\n').trim();

  // Cabeçalho principal
  const headerRegex =
    /CNPJ\s*([0-9.\-\/]+)\s+([^\n]+?)\s+Período Apuração\s*([0-9\/]+)\s+Data de Vencimento\s*([0-9\/]+)\s+Número do Documento\s*([0-9]+)/i;

  const headerMatch = text.match(headerRegex);

  const header = {
    cnpj: headerMatch?.[1]?.trim() || '',
    razaoSocial: headerMatch?.[2]?.trim() || '',
    periodoApuracao: headerMatch?.[3]?.trim() || '',
    dataVencimento: headerMatch?.[4]?.trim() || '',
    numeroDocumento: headerMatch?.[5]?.trim() || '',
  };

  // Bloco "Composição do Documento de Arrecadação" até "Totais"
  const compBlock = extractBetween(
    text,
    /Composição do Documento de Arrecadação/i,
    /Totais/i
  );

  const itens = [];
  if (compBlock) {
    const compClean = compBlock
      .replace(/Composição do Documento de Arrecadação/i, '')
      .replace(/\b(Código|Descrição|Principal|Multa|Juros|Total)\b/gi, ' ')
      .replace(/[ \t]+/g, ' ')
      .replace(/\n{2,}/g, '\n')
      .trim();

    const lines = compClean.split('\n').map((l) => l.trim()).filter(Boolean);
    const merged = joinWrappedLines(lines);

    for (const line of merged) {
      // Esperado: "1082 DESCRIÇÃO ... 1.649,18 [Multa?] [Juros?] 1.649,18"
      const m = line.match(/^(\d{4})\s+(.+?)\s+((?:\d{1,3}\.)*\d+,\d{2})(?:\s+((?:\d{1,3}\.)*\d+,\d{2}))?(?:\s+((?:\d{1,3}\.)*\d+,\d{2}))?\s+((?:\d{1,3}\.)*\d+,\d{2})$/);
      if (m) {
        const [, codigo, descricaoRaw, v1, v2, v3, vLast] = m;
        let principal = v1, multa = null, juros = null, total = vLast;
        const countVals = [v1, v2, v3, vLast].filter(Boolean).length;

        if (countVals === 2) {
          multa = '0,00';
          juros = '0,00';
        } else if (countVals === 4) {
          principal = v1;
          multa = v2;
          juros = v3;
        } else if (countVals === 3) {
          principal = v1;
          multa = '0,00';
          juros = v2;
        }

        itens.push({
          codigo: codigo.trim(),
          descricao: normalizeSpaces(descricaoRaw),
          principal: brToNumber(principal),
          multa: brToNumber(multa),
          juros: brToNumber(juros),
          total: brToNumber(total),
        });
      }
    }
  }

  // Totais
  const totalsMatch = text.match(/Totais\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})/i);
  const totals = totalsMatch
    ? {
        principal: brToNumber(totalsMatch[1]),
        multa: brToNumber(totalsMatch[2]),
        juros: brToNumber(totalsMatch[3]),
        total: brToNumber(totalsMatch[4]),
      }
    : { principal: null, multa: null, juros: null, total: null };

  // Banco / Data de Arrecadação / Agência / Estabelecimento / Valor Reservado / Referência
  const bankRegex = /Banco\s+(.+?)\s+Data de Arrecadação\s+([0-9\/]+)\s+Agência\s+(\d+)\s+Estabelecimento\s+(\d+)\s+Valor Reservado\/Restituído\s+([0-9\.-,]+)(?:\s+Referência\s+([^\n]+))?/i;
  const bankMatch = text.match(bankRegex);
  const arrec = bankMatch
    ? {
        banco: normalizeSpaces(bankMatch[1] || ''),
        dataArrecadacao: (bankMatch[2] || '').trim(),
        agencia: (bankMatch[3] || '').trim(),
        estabelecimento: (bankMatch[4] || '').trim(),
        valorReservado: brToNumber(bankMatch[5]),
        referencia: (bankMatch[6] || '').trim(),
      }
    : {
        banco: '',
        dataArrecadacao: '',
        agencia: '',
        estabelecimento: '',
        valorReservado: null,
        referencia: '',
      };

  return { header, itens, totals, arrec };
}

// Leitura do(s) PDF(s) ---------------------------------------------------

async function parsePdfFile(filePath) {
  const dataBuffer = fs.readFileSync(filePath);

  const options = {
    pagerender: (pageData) => {
      const render_options = {
        normalizeWhitespace: true,
        disableCombineTextItems: false,
      };
      return pageData.getTextContent(render_options).then((textContent) => {
        const strings = textContent.items.map((it) => it.str);
        return strings.join('\n');
      });
    },
  };

  const data = await require('pdf-parse')(dataBuffer, options);
  const pages = data.text.split('\f').map((p) => p.trim()).filter(Boolean);

  const registros = [];
  const itensAll = [];

  for (const pageText of pages) {
    const { header, itens, totals, arrec } = parsePageText(pageText);

    registros.push({
      arquivo: path.basename(filePath),
      cnpj: header.cnpj,
      razaoSocial: header.razaoSocial,
      periodoApuracao: header.periodoApuracao,
      dataVencimento: header.dataVencimento,
      numeroDocumento: header.numeroDocumento,
      banco: arrec.banco,
      dataArrecadacao: arrec.dataArrecadacao,
      agencia: arrec.agencia,
      estabelecimento: arrec.estabelecimento,
      valorReservado: arrec.valorReservado,
      referencia: arrec.referencia,
      totalPrincipal: totals.principal,
      totalMulta: totals.multa,
      totalJuros: totals.juros,
      totalGeral: totals.total,
    });

    for (const it of itens) {
      itensAll.push({
        numeroDocumento: header.numeroDocumento,
        codigo: it.codigo,
        descricao: it.descricao,
        principal: it.principal,
        multa: it.multa,
        juros: it.juros,
        total: it.total,
      });
    }
  }

  return { registros, itensAll };
}

// Escrita do Excel -------------------------------------------------------

async function writeExcel(saidaXlsx, registros, itensAll) {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();

  const wsResumo = wb.addWorksheet('Resumo', { views: [{ state: 'frozen', ySplit: 1 }] });
  wsResumo.columns = [
    { header: 'Arquivo', key: 'arquivo', width: 35 },
    { header: 'CNPJ', key: 'cnpj', width: 20 },
    { header: 'Razão Social', key: 'razaoSocial', width: 45 },
    { header: 'Período Apuração', key: 'periodoApuracao', width: 16 },
    { header: 'Vencimento', key: 'dataVencimento', width: 12 },
    { header: 'Número do Documento', key: 'numeroDocumento', width: 22 },
    { header: 'Banco', key: 'banco', width: 28 },
    { header: 'Data de Arrecadação', key: 'dataArrecadacao', width: 16 },
    { header: 'Agência', key: 'agencia', width: 10 },
    { header: 'Estabelecimento', key: 'estabelecimento', width: 14 },
    { header: 'Valor Reservado/Restituído', key: 'valorReservado', width: 18, style: { numFmt: '#,##0.00' } },
    { header: 'Ref.', key: 'referencia', width: 10 },
    { header: 'Total Principal', key: 'totalPrincipal', width: 16, style: { numFmt: '#,##0.00' } },
    { header: 'Total Multa', key: 'totalMulta', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Total Juros', key: 'totalJuros', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Total Geral', key: 'totalGeral', width: 16, style: { numFmt: '#,##0.00' } },
  ];
  wsResumo.addRows(registros);

  const wsItens = wb.addWorksheet('Itens', { views: [{ state: 'frozen', ySplit: 1 }] });
  wsItens.columns = [
    { header: 'Número do Documento', key: 'numeroDocumento', width: 22 },
    { header: 'Código', key: 'codigo', width: 10 },
    { header: 'Descrição', key: 'descricao', width: 50 },
    { header: 'Principal', key: 'principal', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Multa', key: 'multa', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Juros', key: 'juros', width: 14, style: { numFmt: '#,##0.00' } },
    { header: 'Total', key: 'total', width: 14, style: { numFmt: '#,##0.00' } },
  ];
  wsItens.addRows(itensAll);

  [wsResumo, wsItens].forEach((ws) => {
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { vertical: 'middle' };
  });

  await wb.xlsx.writeFile(saidaXlsx);
}

// CLI --------------------------------------------------------------------

(async () => {
  try {
    const [, , ...args] = process.argv;
    if (args.length < 2) {
      console.error('Uso: node converter-darf-para-excel.js <arquivo(s).pdf> <saida.xlsx>');
      process.exit(1);
    }
    const outFile = args[args.length - 1];
    const inputs = args.slice(0, -1);

    let allRegistros = [];
    let allItens = [];

    for (const input of inputs) {
      const stats = fs.existsSync(input) ? fs.statSync(input) : null;
      if (!stats) {
        console.warn(`Aviso: arquivo não encontrado: ${input}`);
        continue;
      }
      if (stats.isDirectory()) {
        const files = fs.readdirSync(input)
          .filter(f => f.toLowerCase().endsWith('.pdf'))
          .map(f => path.join(input, f));
        for (const f of files) {
          const { registros, itensAll } = await parsePdfFile(f);
          allRegistros = allRegistros.concat(registros);
          allItens = allItens.concat(itensAll);
        }
      } else {
        const { registros, itensAll } = await parsePdfFile(input);
        allRegistros = allRegistros.concat(registros);
        allItens = allItens.concat(itensAll);
      }
    }

    if (allRegistros.length === 0) {
      console.error('Nenhum registro encontrado. Verifique se o(s) PDF(s) correspondem ao layout esperado.');
      process.exit(2);
    }

    // Garante pasta de saída
    const outDir = path.dirname(outFile);
    if (outDir && outDir !== '.' && !fs.existsSync(outDir)) {
      fs.mkdirSync(outDir, { recursive: true });
    }

    await writeExcel(outFile, allRegistros, allItens);
    console.log(`✅ Excel gerado: ${outFile}`);
  } catch (err) {
    console.error('Falha ao converter:', err);
    process.exit(3);
  }
})();
