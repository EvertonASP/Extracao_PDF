
# Conversor DARF (PDF) → Excel (.xlsx)

Este repositório converte **Comprovantes de Arrecadação (DARF)** em **planilhas Excel**, gerando:

- **Resumo**: 1 linha por comprovante (por página do PDF)
- **Itens**: 1 linha por código (vinculado ao *Número do Documento*)

> O parser está ajustado para o layout da RFB visto em 2025 (campos: CNPJ, Razão Social, Período, Vencimento, Número do Documento, bloco **Composição do Documento de Arrecadação**, e bloco **Banco/Data de Arrecadação/Agência/Estabelecimento/Valor Reservado/Referência**). Se o layout mudar, talvez seja necessário adaptar as *regex* no script.

## 🚀 Uso local

Pré‑requisitos: **Node.js 18+**

```bash
npm i
npm run build
```

- Coloque seus PDFs em `pdfs/` (ex.: `pdfs/424 - 35247621000159.pdf`).
- A saída será gerada em `dist/saida.xlsx`.

## ⚙️ GitHub Actions (automático)

O workflow `convert-darf.yml` roda quando:

- você faz **push** de arquivos `pdfs/**/*.pdf`, ou
- aciona manualmente via **Actions → Converter DARF → Run workflow**.

### Artifact de saída

Ao final da execução, faça o download do **artifact** `saida-excel` para obter `dist/saida.xlsx`.

### Comitar a saída no repositório (opcional)

Se preferir versionar a planilha, você pode criar um segundo workflow com *permissions: contents: write* e comitar `outputs/saida.xlsx`. Um exemplo está em comentários no YAML.

## 🧩 Estrutura

```
conversor-darf-template/
├─ pdfs/                 # coloque seus PDFs aqui
├─ dist/                 # saída do Excel (gerado)
├─ converter-darf-para-excel.js
├─ package.json
├─ .gitignore
└─ .github/
   └─ workflows/
      └─ convert-darf.yml
```

## 🔐 Privacidade

- Recomenda‑se repositório **privado** (contém dados sensíveis: CNPJ, valores, etc.).
- Artifacts no GitHub Actions têm retenção limitada (configurável nas políticas do repositório/organização).
- PDFs muito grandes (>100 MB) exigem Git LFS.

## 🛠️ Ajustes comuns

- **Um .xlsx por PDF**: altere o script para iterar sobre cada PDF e salvar `dist/<numeroDocumento>.xlsx`.
- **Consolidação por mês**: filtre por `Período de Apuração` antes de escrever a planilha.

## 📄 Licença

MIT — adapte livremente.
