# 🌱 CrediRural Tech — Gestão de Crédito Agrícola

![Python](https://img.shields.io/badge/python-3.9+-blue.svg)
![Streamlit](https://img.shields.io/badge/streamlit-%23FF4B4B.svg?style=flat&logo=Streamlit&logoColor=white)
![SQLite3](https://img.shields.io/badge/sqlite-%2307405e.svg?style=flat&logo=sqlite&logoColor=white)

O **CrediRural Tech** é uma plataforma inteligente e responsiva voltada para a gestão de crédito, custeio e financiamento no ecossistema de cooperativas de crédito rural. Unindo a solidez do agronegócio com a inovação do movimento **AgTech**, o sistema entrega controle de risco financeiro, planos de amortização dinâmicos e automações operacionais em uma interface de alta performance.

---

## 🚀 Funcionalidades Principais

### 📊 1. Inteligência Financeira (Dashboard)
* **Visão Global Integrada:** Monitoramento em tempo real do capital total alocado, valores liquidados em caixa, saldo devedor da carteira e índice crítico de inadimplência.
* **Filtros Dinâmicos por Safra:** Cruzamento inteligente de dados por Produtor/Cooperado, Ano de Referência (Safra) e Períodos Mensais simultâneos através de controles segmentados.
* **Gráficos Adaptativos:** Projeção faturada vs. caixa realizado e gráficos de distribuição de contratos por status.

### 💼 2. Ciclo de Crédito e Amortização
* **Plano de Amortização Fixo:** Geração automática do fluxo de parcelas com cálculo integrado de taxas de juros subsidiadas para custeio.
* **Amortização Extra de Safra:** Recálculo automático do saldo devedor residual ao detectar pagamentos acima do valor requerido, diluindo as próximas parcelas proporcionalmente.
* **Liquidação Parcial de Títulos:** Controle preciso de recebimentos parciais, mantendo o saldo cobrável restante na mesma cota.

### 🔔 3. Mitigação de Risco e Ferramentas Operacionais
* **Monitoramento Preventivo:** Alertas visuais críticos na barra lateral indicando produtores em atraso e vencimentos nos próximos 5 dias.
* **Régua de Cobrança Integrada:** Link automatizado para notificação direta via WhatsApp com mensagens personalizadas e codificadas para cada cooperado.
* **Segurança de Dados:** Sistema de backup automatizado que gera relatórios em Excel a cada inicialização segura da plataforma.

---

## 🎨 Identidade Visual (AgTech UI)
A interface foi projetada utilizando injeção de estilos CSS customizados no Streamlit, adotando uma paleta corporativa e tecnológica:
* **Verde Esmeralda Profundo:** Representa a terra, a lavoura e o posicionamento sustentável da cooperativa.
* **Dourado/Âmbar:** Alusão à colheita, ao trigo e à prosperidade financeira.
* **Grafite Escuro (Slate):** Estruturação minimalista que remete a softwares modernos e robustos.

---

## 🛠️ Arquitetura Técnica e Engenharia de Dados

Para garantir estabilidade em ambiente local e paralelismo do Streamlit, o sistema implementa:
* **Gerenciamento de Contexto (`with` statement):** Isolamento total de conexões ao banco de dados SQLite, eliminando o risco de corrupção ou travamentos por concorrência de memória (*multi-threading*).
* **Processamento Vetorizado:** Manipulação de dados financeiros de alta performance utilizando a biblioteca **Pandas**.

### Stack Tecnológica
* **Linguagem:** Python
* **Interface:** Streamlit
* **Banco de Dados:** SQLite3
* **Manipulação e Relatórios:** Pandas / OpenPyXL

---

## 🔧 Como Executar o Projeto

### Pré-requisitos
Certifique-se de ter o Python 3.9 ou superior instalado em sua máquina.

### 1. Clonar o Repositório
```bash
git clone [https://github.com/seu-usuario/credirural-tech.git](https://github.com/seu-usuario/credirural-tech.git)
cd credirural-tech
