<h1>🍽️ Cafeteria Access Report – PDC Campinas</h1>

<p>
  Este projeto é uma aplicação web desenvolvida em <b>Python + Streamlit</b> 
  focada no <b>tratamento, padronização e filtragem</b> dos relatórios de acesso 
  da cafeteria do <b>PDC / SAPDC Campinas</b>.
</p>

<blockquote>
  <b>🎯 Objetivo Principal</b><br>
  Automatizar a limpeza do relatório bruto exportado do SAP, preservando o layout oficial 
  e gerando um arquivo final pronto para auditoria, RH e integração com Power BI.
</blockquote>

<hr>

<h2>✅ Funcionalidades</h2>
<ul>
  <li><b>Leitura Inteligente:</b> Processa o Excel original do SAP ignorando linhas iniciais.</li>
  <li><b>Detecção de Cabeçalho:</b> Identifica automaticamente a linha de início da tabela.</li>
  <li><b>Filtro Regional:</b> Mantém apenas registros do <b>PDC / SAPDC Campinas</b>.</li>
  <li><b>Formatação Automática:</b> 
    <ul>
      <li>Geração do título <code>Cafeteria Access Report</code>.</li>
      <li>Cálculo de <code>START/END DATE</code> baseado no nome do arquivo.</li>
      <li>Conversão dos dados em <b>Tabela do Excel (Ctrl + T)</b>.</li>
    </ul>
  </li>
  <li><b>Output Profissional:</b> Arquivo limpo e colunas padronizadas.</li>
</ul>

<h2>📂 Padrão de Nomenclatura</h2>
<p>O sistema processa as datas automaticamente a partir do nome do arquivo:</p>

<code>Exemplo: 16A22FEV2026.xlsx</code>

<p>O cabeçalho interno será preenchido como:</p>
<pre>
START DATE: 2/16/2026 12:00:00 AM
END DATE:   2/22/2026 11:59:59 PM
</pre>

<h2>📄 Arquivo de Saída</h2>
<p>O relatório final é gerado com o seguinte padrão:</p>
<code>Relatorio de Controle do Restaurante 16.02 a 22.02.xlsx</code>

<h2>🖥️ Como Executar</h2>
<p>Instale as dependências e inicie o servidor local:</p>

<pre>
# Instalação
pip install streamlit pandas openpyxl

# Execução
streamlit run limpador.py
</pre>

<h2>📦 Estrutura do Repositório</h2>
<pre>
refeicoes/
├── limpador.py          # Aplicação Streamlit
├── requirements.txt     # Dependências
└── README.md            # Documentação
</pre>

<hr>

<p align="center">
  <sub>Desenvolvido para automação e padronização de relatórios internos.</sub><br>
  <b>HR / Data Analytics – PDC Campinas</b>
</p>
