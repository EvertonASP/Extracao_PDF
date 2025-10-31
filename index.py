<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Processador de Documentos PDF - Corporativo</title>
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300,400,700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary-color: #1F4E78; /* Azul corporativo, do cabeçalho do Excel */
            --secondary-color: #337ab7; /* Azul mais claro para detalhes */
            --text-color: #333;
            --bg-color: #f4f4f9;
            --card-bg: #fff;
            --code-bg: #eee;
        }
        body {
            font-family: 'Roboto', sans-serif;
            line-height: 1.6;
            color: var(--text-color);
            background-color: var(--bg-color);
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        header {
            background-color: var(--primary-color);
            color: white;
            padding: 40px 0;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        header h1 {
            margin: 0;
            font-weight: 700;
            font-size: 2.5em;
        }
        header p {
            margin-top: 10px;
            font-weight: 300;
            font-size: 1.2em;
        }
        section {
            padding: 60px 0;
            border-bottom: 1px solid #ddd;
        }
        section:last-child {
            border-bottom: none;
        }
        h2 {
            color: var(--primary-color);
            text-align: center;
            margin-bottom: 40px;
            font-weight: 700;
            font-size: 2em;
            position: relative;
        }
        h2::after {
            content: '';
            display: block;
            width: 50px;
            height: 3px;
            background-color: var(--secondary-color);
            margin: 10px auto 0;
        }
        .card-grid {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
            justify-content: center;
        }
        .card {
            background-color: var(--card-bg);
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            flex: 1 1 300px;
            transition: transform 0.3s, box-shadow 0.3s;
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.1);
        }
        .card h3 {
            color: var(--secondary-color);
            margin-top: 0;
            font-weight: 700;
        }
        .code-block {
            background-color: var(--code-bg);
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            font-family: monospace;
            white-space: pre;
            margin-bottom: 20px;
        }
        .interactive-element {
            text-align: center;
            margin-top: 20px;
        }
        .btn {
            background-color: var(--primary-color);
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1em;
            transition: background-color 0.3s;
        }
        .btn:hover {
            background-color: var(--secondary-color);
        }
        .hidden-content {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.5s ease-in-out;
            padding: 0 20px;
        }
        .hidden-content.show {
            max-height: 500px; /* Valor grande o suficiente para o conteúdo */
            padding: 20px;
        }
        footer {
            text-align: center;
            padding: 20px 0;
            background-color: var(--primary-color);
            color: white;
            font-size: 0.9em;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        table, th, td {
            border: 1px solid #ddd;
        }
        th, td {
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: var(--primary-color);
            color: white;
            font-weight: 700;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>

    <header>
        <div class="container">
            <h1>Processador de Documentos PDF</h1>
            <p>Solução corporativa para extração e consolidação automatizada de dados.</p>
        </div>
    </header>

    <section id="visao-geral">
        <div class="container">
            <h2>Visão Geral da Solução</h2>
            <p style="text-align: center; margin-bottom: 40px;">Este script Python é uma ferramenta robusta projetada para automatizar o processo de extração de dados de múltiplos arquivos PDF, utilizando processamento paralelo e inteligência artificial (via API Tela) para estruturar as informações em um relatório Excel formatado.</p>

            <div class="card-grid">
                <div class="card">
                    <h3>Processamento Paralelo</h3>
                    <p>Utiliza <code>ThreadPoolExecutor</code> para processar a leitura de múltiplos PDFs simultaneamente, garantindo alta performance e eficiência.</p>
                </div>
                <div class="card">
                    <h3>Extração Inteligente</h3>
                    <p>Integração com a API Tela (<code>canvas_id="496722d2-bb2f-4f12-962d-4f34ba7d9db8"</code>) para transformar o texto bruto do PDF em dados estruturados e consolidados.</p>
                </div>
                <div class="card">
                    <h3>Relatório Profissional</h3>
                    <p>Geração de um arquivo Excel (<code>.xlsx</code>) com formatação corporativa, incluindo cabeçalhos estilizados, filtros automáticos e ajuste de largura de colunas.</p>
                </div>
            </div>
        </div>
    </section>

    <section id="funcionalidades">
        <div class="container">
            <h2>Principais Funcionalidades</h2>
            <div class="card-grid">
                <div class="card">
                    <h3>Seleção de Arquivos</h3>
                    <p>Interface gráfica (Tkinter) para seleção fácil e intuitiva de múltiplos arquivos PDF.</p>
                </div>
                <div class="card">
                    <h3>Consolidação de Dados</h3>
                    <p>Uso da biblioteca Pandas para unificar os dados extraídos de todos os PDFs em um único DataFrame.</p>
                </div>
                <div class="card">
                    <h3>Controle de Tempo</h3>
                    <p>Medição e exibição do tempo total de processamento para monitoramento de performance.</p>
                </div>
            </div>
        </div>
    </section>

    <section id="estrutura-de-dados">
        <div class="container">
            <h2>Estrutura do Relatório Final</h2>
            <p style="text-align: center;">O script gera um arquivo <code>Resultado_Extração_PDF.xlsx</code> com a seguinte estrutura de colunas (baseado na saída do DataFrame):</p>

            <table>
                <thead>
                    <tr>
                        <th>Coluna</th>
                        <th>Descrição</th>
                        <th>Observações</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>NumeroDocumento</td>
                        <td>Identificador único do documento.</td>
                        <td>Usado para agrupar e corrigir a paginação.</td>
                    </tr>
                    <tr>
                        <td>NumeroPagina</td>
                        <td>Número sequencial da página dentro do documento.</td>
                        <td>Corrigido pelo script (<code>.cumcount() + 1</code>).</td>
                    </tr>
                    <!-- Adicionar mais colunas se a estrutura do JSON da API for conhecida -->
                    <tr>
                        <td>Outras Colunas</td>
                        <td>Dados extraídos pela API Tela.</td>
                        <td>Depende da configuração do Canvas ID <code>496722d2-bb2f-4f12-962d-4f34ba7d9db8</code>.</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </section>

    <section id="codigo-fonte">
        <div class="container">
            <h2>Código Fonte (Python)</h2>
            <p style="text-align: center;">O código a seguir demonstra a lógica principal do processamento. Clique para expandir.</p>

            <div class="interactive-element">
                <button class="btn" onclick="toggleCode()">Visualizar Código Completo</button>
            </div>

            <div id="code-details" class="hidden-content">
                <div class="code-block">
<pre>
def main():
    import os
    import time
    import pandas as pd
    from concurrent.futures import ThreadPoolExecutor
    from tkinter import Tk, filedialog, messagebox
    
    # ... (Funções selecionar_pdfs e extract_text_from_pdf) ...

    # Seleciona PDFs
    pdf_paths = selecionar_pdfs()
    if not pdf_paths:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
        return

    inicio = time.time()

    # Processa PDFs em paralelo
    with ThreadPoolExecutor() as executor:
        resultados = list(executor.map(extract_text_from_pdf, pdf_paths))

    # Cria cliente Tela (importação tardia)
    from tela import create_tela_client
    tela = create_tela_client(api_key="b688428d-e6e1-46db-b604-22d78aa0604a") # ATENÇÃO: Chave de API exposta!

    # Consolida dados
    dados_consolidados = []
    for _, conteudo_pdf in resultados:
        data = tela.completions.create(
            canvas_id="496722d2-bb2f-4f12-962d-4f34ba7d9db8",
            variables={"arquivo": conteudo_pdf},
        )
        resposta = data.choices[0].message.content["Composição do documento de arrecadação"]
        dados_consolidados.extend(resposta)

    # Cria DataFrame e salva Excel
    df = pd.DataFrame(dados_consolidados)
    
    # ... (Correção de NumeroPagina e salvamento com xlsxwriter) ...

    # Mensagem final e abertura de pasta
    # ...

if __name__ == "__main__":
    main()
</pre>
                </div>
                <p style="color: red; font-weight: bold; text-align: center;">ATENÇÃO: A chave de API está exposta no código. Para uso em produção, considere armazená-la em variáveis de ambiente ou um gerenciador de segredos.</p>
            </div>
        </div>
    </section>

    <footer>
        <div class="container">
            <p>&copy; 2025 Solução de Processamento de Dados. Todos os direitos reservados.</p>
        </div>
    </footer>

    <script>
        function toggleCode() {
            const content = document.getElementById('code-details');
            const button = document.querySelector('#codigo-fonte .btn');
            
            if (content.classList.contains('show')) {
                content.classList.remove('show');
                button.textContent = 'Visualizar Código Completo';
            } else {
                content.classList.add('show');
                button.textContent = 'Ocultar Código';
            }
        }

        // Smooth scrolling para navegação futura (se adicionarmos links de navegação)
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                document.querySelector(this.getAttribute('href')).scrollIntoView({
                    behavior: 'smooth'
                });
            });
        });
    </script>
</body>
</html>
