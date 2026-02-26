<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>Cafeteria Access Report – PDC Campinas</title>
  <style>
    body {
      font-family: Arial, Helvetica, sans-serif;
      background-color: #ffffff;
      color: #1a1a1a;
      line-height: 1.6;
      margin: 0;
      padding: 40px;
    }

    h1, h2, h3 {
      color: #0b3d2e;
    }

    h1 {
      border-bottom: 3px solid #0b3d2e;
      padding-bottom: 10px;
    }

    code {
      background-color: #f4f4f4;
      padding: 3px 6px;
      border-radius: 4px;
      font-size: 0.95em;
    }

    pre {
      background-color: #f4f4f4;
      padding: 15px;
      border-radius: 6px;
      overflow-x: auto;
    }

    .box {
      border: 1px solid #e0e0e0;
      border-left: 5px solid #0b3d2e;
      padding: 15px;
      margin: 20px 0;
      background-color: #fafafa;
    }

    ul {
      margin-left: 20px;
    }

    footer {
      margin-top: 50px;
      font-size: 0.9em;
      color: #555;
      border-top: 1px solid #ddd;
      padding-top: 15px;
    }
  </style>
</head>
<body>

  <h1>🍽️ Cafeteria Access Report – PDC Campinas</h1>

  <p>
    Este projeto é uma aplicação web desenvolvida em <strong>Python + Streamlit</strong>
    para <strong>tratamento, padronização e filtragem</strong> dos relatórios de acesso
    da cafeteria do <strong>PDC / SAPDC Campinas</strong>.
  </p>

  <div class="box">
    <strong>Objetivo principal:</strong><br>
    Automatizar a limpeza do relatório bruto exportado do SAP, mantendo o layout oficial
    e gerando um arquivo final pronto para auditoria, RH e Power BI.
  </div>

  <h2>✅ Funcionalidades</h2>
  <ul>
    <li>Leitura automática do Excel original do SAP (mesmo com linhas antes da tabela)</li>
    <li>Identificação dinâmica do cabeçalho real da tabela</li>
    <li>Filtro exclusivo para registros do <strong>PDC / SAPDC Campinas</strong></li>
    <li>Padronização e ordem fixa dos cabeçalhos</li>
    <li>Geração automática do cabeçalho:
      <ul>
        <li><code>Cafeteria Access Report</code></li>
        <li><code>QUERY: START DATE / END DATE</code> (a partir do nome do arquivo)</li>
      </ul>
    </li>
    <li>Criação automática de <strong>Tabela do Excel (Ctrl + T)</strong></li>
    <li>Arquivo final com layout limpo, fundo branco e pronto para uso</li>
  </ul>

  <h2>📂 Padrão do nome do arquivo</h2>
  <p>
    O sistema espera arquivos com o seguinte padrão:
  </p>

  <pre><code>16A22FEV2026.xlsx</code></pre>

  <p>
    A partir desse nome, o app gera automaticamente:
  </p>

  <pre><code>
START DATE:  2/16/2026 12:00:00 AM
END DATE:    2/22/2026 11:59:59 PM
  </code></pre>

  <h2>📄 Nome do arquivo gerado</h2>
  <p>
    O relatório final é exportado automaticamente com o nome:
  </p>

  <pre><code>
Relatorio de Controle do Restaurante 16.02 a 22.02.xlsx
  </code></pre>

  <h2>🖥️ Como executar o projeto</h2>

  <pre><code>
streamlit run limpador.py
  </code></pre>

  <h2>📦 Estrutura do repositório</h2>

  <pre><code>
refeicoes/
│
├── limpador.py
├── requirements.txt
└── README.html
  </code></pre>

  <h2>🔒 Observações importantes</h2>
  <ul>
    <li>O layout final do Excel segue o padrão original do SAP</li>
    <li>Somente registros do PDC Campinas são mantidos</li>
    <li>O projeto foi pensado para uso corporativo e auditoria</li>
  </ul>

  <footer>
    <p>
      Desenvolvido para automação e padronização de relatórios internos.<br>
      <strong>HR / Data Analytics – PDC Campinas</strong>
    </p>
  </footer>

</body>
</html>
