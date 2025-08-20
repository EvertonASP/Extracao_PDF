<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <title>Extração DARF (PDF → Excel)</title>
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <!-- SheetJS (XLSX) - CDN oficial -->
  <ttps://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
  <!-- PDF.js (UMD) - ajuste a versão se quiser -->
  https://cdn.jsdelivr.net/npm/pdfjs-dist@4.6.82/build/pdf.min.js
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
    .card { max-width: 920px; margin: auto; padding: 24px; border: 1px solid #e5e7eb; border-radius: 10px; }
    h1 { margin-top: 0; font-size: 1.4rem; }
    input[type=file] { padding: 8px; }
    button { margin-top: 12px; padding: 10px 16px; border: 0; border-radius: 8px; background: #2563eb; color: #fff; cursor: pointer; }
    button:disabled { background: #9ca3af; cursor: not-allowed; }
    pre { background: #0b1020; color: #9fe870; padding: 12px; border-radius: 8px; max-height: 240px; overflow: auto; }
    .tip { color: #374151; font-size: .95rem; }
  </style>
</head>
<body>
<div class="card">
  <h1>Extração DARF (PDF → Excel)</h1>
  <p class="tip">Selecione o(s) PDF(s) de <b>Comprovante de Arrecadação</b> da RFB. O processamento ocorre no seu navegador.</p>
  <input id="pdfInput" type="file" accept="application/pdf" multiple />
  <br />
  <button id="btn" disabled>Converter para Excel</button>
  <p class="tip" id="status">Aguardando arquivos…</p>
  <pre id="log"></pre>
</div>

<script>
  // Config do worker da PDF.js (obrigatória fora do viewer)
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.6.82/build/pdf.worker.min.js";

  const input = document.getElementById('pdfInput');
  const btn = document.getElementById('btn');
  const statusEl = document.getElementById('status');
  const logEl = document.getElementById('log');

  let files = [];

  input.addEventListener('change', () => {
    files = Array.from(input.files || []);
    btn.disabled = files.length === 0;
    statusEl.textContent = files.length
      ? `${files.length} arquivo(s) pronto(s) para processar`
      : 'Aguardando arquivos…';
  });

  const brToNumber = (s) => {
    if (s == null) return null;
    const clean = String(s).replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
    const n = Number(clean);
    return Number.isFinite(n) ? n : null;
  };

  const normalizeSpaces = (s) => s.replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ').trim();

  // Agrupa itens de texto por linha (por coordenada Y com tolerância), ordena por X e junta
  function textItemsToLines(textContent, yTolerance = 2) {
    const linesMap = new Map(); // key: yKey, val: array de itens
    for (const it of textContent.items) {
      const y = it.transform[5]; // posição Y
      const x = it.transform[4]; // posição X
      // acha uma chave de linha próxima (tolerância)
      let yKey = [...linesMap.keys()].find(k => Math.abs(k - y) <= yTolerance);
      if (yKey === undefined) yKey = y;
      if (!linesMap.has(yKey)) linesMap.set(yKey, []);
      linesMap.get(yKey).push({ x, str: it.str });
    }
    // ordena por Y (descendente: topo→baixo) e, em cada linha, por X (esquerda→direita)
    const sortedY = [...linesMap.keys()].sort((a,b) => b - a);
    const lines = [];
    for (const yKey of sortedY) {
      const items = linesMap.get(yKey).sort((a,b) => a.x - b.x);
      const line = items.map(it => it.str).join(' ');
      lines.push(normalizeSpaces(line));
    }
    return lines.filter(Boolean);
  }

  function joinWrappedLines(lines) {
    const out = [];
    for (const raw of lines) {
      const line = normalizeSpaces(raw);
      if (!line) continue;
      if (/^\d{4}\b/.test(line)) {
        out.push(line);
      } else {
        if (out.length === 0) out.push(line);
        else out[out.length - 1] = normalizeSpaces(out[out.length - 1] + ' ' + line);
      }
    }
    return out;
  }

  function extractBetween(text, startPattern, endPattern) {
    const start = text.search(startPattern);
    if (start === -1) return null;
    const after = text.slice(start);
    const end = endPattern ? after.search(endPattern) : -1;
    return end === -1 ? after : after.slice(0, end);
  }

  function parsePageText(pageText) {
    const text = pageText.replace(/\r/g, '').replace(/[ \t]+\n/g, '\n').trim();

    const headerRegex =
      /CNPJ\s*([0-9.\-\/]+)\s+([^\n]+?)\s+Período Apuração\s*([0-9/]+)\s+Data de Vencimento\s*([0-9/]+)\s+Número do Documento\s*([0-9]+)/i;
    const headerMatch = text.match(headerRegex);
    const header = {
      cnpj: headerMatch?.[1]?.trim() || '',
      razaoSocial: headerMatch?.[2]?.trim() || '',
      periodoApuracao: headerMatch?.[3]?.trim() || '',
      dataVencimento: headerMatch?.[4]?.trim() || '',
      numeroDocumento: headerMatch?.[5]?.trim() || '',
    };

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
        const m = line.match(/^(\d{4})\s+(.+?)\s+((?:\d{1,3}\.)*\d+,\d{2})(?:\s+((?:\d{1,3}\.)*\d+,\d{2}))?(?:\s+((?:\d{1,3}\.)*\d+,\d{2}))?\s+((?:\d{1,3}\.)*\d+,\d{2})$/);
        if (m) {
          const [, codigo, descricaoRaw, v1, v2, v3, vLast] = m;
          let principal = v1, multa = null, juros = null, total = vLast;
          const countVals = [v1, v2, v3, vLast].filter(Boolean).length;

          if (countVals === 2) { multa = '0,00'; juros = '0,00'; }
          else if (countVals === 4) { principal = v1; multa = v2; juros = v3; }
          else if (countVals === 3) { principal = v1; multa = '0,00'; juros = v2; }

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

    const totalsMatch = text.match(/Totais\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})/i);
    const totals = totalsMatch
      ? {
          principal: brToNumber(totalsMatch[1]),
          multa: brToNumber(totalsMatch[2]),
          juros: brToNumber(totalsMatch[3]),
          total: brToNumber(totalsMatch[4]),
        }
      : { principal: null, multa: null, juros: null, total: null };

    const bankRegex = /Banco\s+(.+?)\s+Data de Arrecadação\s+([0-9/]+)\s+Agência\s+(\d+)\s+Estabelecimento\s+(\d+)\s+Valor Reservado\/Restituído\s+([0-9\.\-,]+)(?:\s+Referência\s+([^\n]+))?/i;
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

  async function extractFromPdfFile(file) {
    const ab = await file.arrayBuffer();
    const loadingTask = pdfjsLib.getDocument({ data: ab });
    const pdf = await loadingTask.promise;

    const registros = [];
    const itensAll = [];

    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const textContent = await page.getTextContent({
        normalizeWhitespace: true,
        disableCombineTextItems: false
      });
      const lines = textItemsToLines(textContent, 2);
      const pageText = lines.join('\n');

      const { header, itens, totals, arrec } = parsePageText(pageText);

      registros.push({
        arquivo: file.name,
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

  function exportToXlsx(registros, itensAll) {
    const wb = XLSX.utils.book_new();

    const wsResumo = XLSX.utils.json_to_sheet(registros);
    const wsItens = XLSX.utils.json_to_sheet(itensAll);

    XLSX.utils.book_append_sheet(wb, wsResumo, 'Resumo');
    XLSX.utils.book_append_sheet(wb, wsItens, 'Itens');

    XLSX.writeFile(wb, 'saida.xlsx');
  }

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    logEl.textContent = '';
    statusEl.textContent = 'Processando…';

    try {
      let allRegistros = [];
      let allItens = [];

      for (const f of files) {
        statusEl.textContent = `Lendo: ${f.name}`;
        const { registros, itensAll } = await extractFromPdfFile(f);
        allRegistros = allRegistros.concat(registros);
        allItens = allItens.concat(itensAll);
      }

      // Log rápido no visor
      logEl.textContent = JSON.stringify({ registros: allRegistros, itens: allItens.slice(0, 10) }, null, 2);
      exportToXlsx(allRegistros, allItens);
      statusEl.textContent = `Concluído! Gerado "saida.xlsx".`;
    } catch (e) {
      console.error(e);
      statusEl.textContent = 'Falha na conversão.';
      logEl.textContent = String(e?.stack || e);
    } finally {
      btn.disabled = files.length === 0;
    }
  });
</script>
</body>
</html>
