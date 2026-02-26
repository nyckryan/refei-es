<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cafeteria Access Report – PDC Campinas</title>
  <style>
    :root {
      --primary-color: #0b3d2e;
      --bg-color: #ffffff;
      --light-gray: #f4f4f4;
      --border-color: #e0e0e0;
      --text-color: #1a1a1a;
    }

    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      background-color: var(--bg-color);
      color: var(--text-color);
      line-height: 1.6;
      margin: 0;
      padding: 40px;
      max-width: 900px;
      margin: auto;
    }

    h1, h2, h3 {
      color: var(--primary-color);
    }

    h1 {
      border-bottom: 3px solid var(--primary-color);
      padding-bottom: 10px;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    code {
      background-color: var(--light-gray);
      padding: 3px 6px;
      border-radius: 4px;
      font-family: 'Consolas', monospace;
      font-size: 0.9em;
      color: #d63384;
    }

    pre {
      background-color: var(--light-gray);
      padding: 15px;
      border-radius: 8px;
      overflow-x: auto;
      border: 1px solid var(--border-color);
    }

    .box {
      border: 1px solid var(--border-color);
      border-left: 5px solid var(--primary-color);
      padding: 20px;
      margin: 25px 0;
      background-color: #fcfcfc;
      box-shadow: 2px 2px 5px rgba(0,0,0,0.02);
    }

    .box strong {
      color: var(--primary-color);
      text-transform: uppercase;
      font-size: 0.9em;
    }

    ul {
      padding-left: 20px;
    }

    li {
      margin-bottom: 8px;
    }

    .file-tree {
      font-family: 'Consolas', monospace;
      color: #333;
    }

    footer {
      margin-top: 60px;
      font-size: 0.85em;
      color: #666;
      border-top: 1px solid var(--border-color);
      padding-top: 20px;
      text-align: center;
    }

    .badge {
      background: var(--primary-color);
      color: white;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 0.8em;
      vertical-align: middle;
    }
  </style>
</head>
<body>

  <h1>🍽️ Cafeteria Access Report – PDC Campinas</h1>

  <p>
    Este projeto é uma aplicação web desenvolvida em <strong>Python + Streamlit</strong> 
    focada no <strong>tratamento, padronização e filtragem</strong> dos relatórios de acesso 
    da cafeteria do <strong>PDC / SAPDC Campinas</strong>.
  </p>

  <div class="box">
    <strong>🎯 Objetivo Principal</strong><br>
    Automatizar a limpeza do relatório bruto exportado do SAP, preservando o layout oficial 
    e gerando um arquivo final pronto para auditoria, RH e integração com Power BI.
  </div>

  <h2>✅ Funcionalidades</h2>
  <ul>
    <li><strong>Leitura Inteligente:</strong> Processa o Excel original do SAP ignorando linhas de metadados iniciais.</li>
    <li><strong>Detecção de Cabeçalho:</strong> Identifica dinamicamente a linha de início da tabela.</li>
    <li><strong>Filtro Regional:</strong> Mantém apenas registros pertinentes ao <strong>PDC / SAPDC Campinas</strong>.</li>
    <li><strong>Formatação Automática:</strong> 
      <ul>
        <li>Geração do título <code>Cafeteria Access Report</code>.</li>
        <li>Cálculo de <code>START/END DATE</code> baseado no nome do arquivo original.</li>
        <li>Conversão do intervalo de dados em <strong>Tabela do Excel (Ctrl + T)</strong>.</li>
      </ul>
    </li>
    <li><strong>Output Profissional:</strong> Arquivo limpo, com fundo branco e colunas padronizadas.</li>
  </ul>

  <h2>📂 Padrão de Nomenclatura</h2>
  <p>O sistema processa as datas automaticamente a partir do nome do arquivo:</p>
  
  <pre><code>Exemplo: 16A22FEV2026.xlsx</code></pre>

  <p>O cabeçalho interno será preenchido como:</p>
  <pre>START DATE: 2/16/2026 12:00:00 AM
END DATE:   2/22/2026 11:59:59 PM</pre>

  <h2>📄 Arquivo de Saída</h2>
  <p>O relatório final é gerado com o seguinte padrão de nome:</p>
  <code>Relatorio de Controle do Restaurante 16.02 a 22.02.xlsx</code>

  <h2>🖥️ Como Executar</h2>
  <p>Instale as dependências e inicie o servidor local:</p>
  <pre><code># Instalação (se necessário)
pip install streamlit pandas openpyxl

# Execução
streamlit run limpador.py</code></pre>

  <h2>📦 Estrutura do Repositório</h2>
  <pre class="file-tree">
refeicoes/
├── limpador.py          # Aplicação Streamlit
├── requirements.txt     # Dependências do projeto
└── README.html          # Documentação
  </pre>

  <h2>🔒 Notas de Segurança e Auditoria</h2>
  <ul>
    <li>O layout segue rigorosamente as normas corporativas.</li>
    <li>Apenas dados do PDC Campinas são retidos para garantir a conformidade dos dados.</li>
    <li>Ideal para processos de <strong>HR Data Analytics</strong>.</li>
  </ul>

  <footer>
    <p>
      Desenvolvido para automação e padronização de relatórios internos.<br>
      <strong>HR / Data Analytics – PDC Campinas</strong>
    </p>
  </footer>

</body>
</html>
