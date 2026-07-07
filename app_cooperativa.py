import streamlit as st
import pandas as pd
import sqlite3
from datetime import date, datetime, timedelta
import io
import os
import urllib.parse
import re

DB_PATH = 'cooperativa.db'

# ==========================================
# 1. CONFIGURAÇÃO E MIGRAÇÃO AUTOMÁTICA
# ==========================================
def realizar_backup_seguranca():
    if os.path.exists(DB_PATH):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clientes'")
                if cursor.fetchone():
                    df_clientes = pd.read_sql_query("SELECT id, nome, telefone, email, limite FROM clientes", conn)
                    df_parcelas = pd.read_sql_query("SELECT * FROM parcelas", conn)
                    
                    if not df_clientes.empty or not df_parcelas.empty:
                        if not os.path.exists('backups_automaticos'):
                            os.makedirs('backups_automaticos')
                        
                        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                        caminho_backup = f"backups_automaticos/Backup_Sistema_{timestamp}.xlsx"
                        
                        with pd.ExcelWriter(caminho_backup, engine='openpyxl') as writer:
                            df_clientes.to_excel(writer, index=False, sheet_name='Clientes')
                            df_parcelas.to_excel(writer, index=False, sheet_name='Parcelas_Emprestimos')
        except Exception as e:
            pass

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS clientes 
                     (id INTEGER PRIMARY KEY, nome TEXT, telefone TEXT, email TEXT, limite REAL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS parcelas 
                     (id INTEGER PRIMARY KEY, cliente_nome TEXT, valor REAL, 
                      vencimento DATE, status TEXT, valor_pago REAL, metodo TEXT, 
                      n_parcela TEXT, emprestimo_id TEXT, taxa_utilizada REAL)''')
        
        try:
            c.execute("ALTER TABLE parcelas ADD COLUMN emprestimo_id TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            c.execute("ALTER TABLE parcelas ADD COLUMN taxa_utilizada REAL")
        except sqlite3.OperationalError:
            pass
            
        conn.commit()

if 'backup_inicial_rodado' not in st.session_state:
    realizar_backup_seguranca()
    st.session_state['backup_inicial_rodado'] = True

# Inicializa o banco de dados garantindo as tabelas prontas
init_db()

# ==========================================
# 2. FUNÇÕES DE LÓGICA DE NEGÓCIO
# ==========================================
def calcular_juros_atraso(valor_original, data_vencimento):
    hoje = date.today()
    venc = datetime.strptime(data_vencimento, '%Y-%m-%d').date()
    
    if hoje > venc:
        dias_atraso = (hoje - venc).days
        multa_fixa = valor_original * 0.02 
        juros_dia = valor_original * 0.003 * dias_atraso 
        return valor_original + multa_fixa + juros_dia, dias_atraso
    return valor_original, 0

def limpar_telefone(telefone_str):
    return re.sub(r'\D', '', str(telefone_str))

# ==========================================
# 3. INTERFACE DO USUÁRIO (STREAMLIT UI/UX AGTECH)
# ==========================================
st.set_page_config(page_title="Cia Menezes - Gestão de Crédito", page_icon="🌱", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
    }
    
    .wa-button {
        display: inline-flex; align-items: center; justify-content: center;
        background: linear-gradient(135deg, #25D366 0%, #1CBD5D 100%); 
        color: white !important; padding: 10px 20px;
        border-radius: 8px; text-decoration: none; font-weight: 600; 
        margin-bottom: 15px; box-shadow: 0 4px 12px rgba(37, 211, 102, 0.2);
        transition: all 0.3s ease;
    }
    .wa-button:hover { 
        transform: translateY(-2px); 
        box-shadow: 0 6px 15px rgba(37, 211, 102, 0.3); 
    }
    
    .ficha-cliente { 
        background: #ffffff; 
        padding: 24px; 
        border-radius: 12px; 
        border: 1px solid #e2e8f0; 
        box-shadow: 0 4px 20px rgba(0,0,0,0.04);
        margin-bottom: 20px; 
    }
    
    .alerta-exclusao { 
        background-color: #fff5f5; 
        padding: 15px; 
        border-left: 5px solid #e53e3e; 
        border-radius: 6px; 
        margin-top: 10px; 
    }
    
    .header-banner {
        background: linear-gradient(135deg, #0B4624 0%, #1E5631 50%, #111827 100%);
        padding: 40px;
        border-radius: 16px;
        color: white;
        margin-bottom: 30px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 8px 30px rgba(11, 70, 36, 0.15);
    }
    .header-banner::after {
        content: '';
        position: absolute;
        right: -10%; top: -30%;
        width: 400px; height: 400px;
        background: rgba(212, 175, 55, 0.08);
        border-radius: 50%;
        pointer-events: none;
    }
    .header-banner h1 { color: #FFFFFF !important; font-weight: 700; margin: 0; font-size: 2.2rem; }
    .header-banner p { color: #A7F3D0 !important; margin: 5px 0 0 0; opacity: 0.9; font-size: 1.1rem; }
    
    .ag-card {
        background: white;
        padding: 22px;
        border-radius: 14px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
        border: 1px solid #f1f5f9;
        transition: transform 0.3s ease;
    }
    .ag-card:hover { transform: translateY(-3px); }
    .ag-title { font-size: 12px; font-weight: 700; letter-spacing: 1px; color: #64748B; margin-bottom: 8px; }
    .ag-value { font-size: 26px; font-weight: 700; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

menu = st.sidebar.selectbox("🗺️ Navegação Operacional", 
                            ["Dashboard", "Cadastrar Cliente", "Consultar Clientes", "Novo Empréstimo", "Recebimentos e Baixas", "Relatórios"])

st.sidebar.markdown("---")
st.sidebar.subheader("🔔 Monitoramento de Risco")

# Acesso seguro ao banco para renderizar alertas na barra lateral
with sqlite3.connect(DB_PATH) as conn:
    df_alertas = pd.read_sql_query("SELECT * FROM parcelas WHERE status != 'Pago'", conn)

if not df_alertas.empty:
    df_alertas['vencimento_dt'] = pd.to_datetime(df_alertas['vencimento']).dt.date
    hoje = date.today()
    limite_5_dias = hoje + timedelta(days=5)
    
    qtd_atrasados = len(df_alertas[df_alertas['vencimento_dt'] < hoje])
    qtd_proximos = len(df_alertas[(df_alertas['vencimento_dt'] >= hoje) & (df_alertas['vencimento_dt'] <= limite_5_dias)])
    
    if qtd_atrasados > 0:
        st.sidebar.error(f"🚨 **{qtd_atrasados}** Produtores em Atraso")
    else:
        st.sidebar.success("✅ Carteira 100% Adimplente")
        
    if qtd_proximos > 0:
        st.sidebar.warning(f"⏳ **{qtd_proximos}** Vencimentos em 5 dias")
else:
    st.sidebar.info("Nenhuma pendência na carteira.")

st.markdown("""
    <div class="header-banner">
        <h1>Cia Menezes</h1>
        <p>Plataforma Inteligente de Gestão de Crédito e Financiamento</p>
    </div>
    """, unsafe_allow_html=True)

# ------------------------------------------
# TELA 1: DASHBOARD (VISÃO GLOBAL POR PADRÃO)
# ------------------------------------------
if menu == "Dashboard":
    st.markdown("### 📊 Inteligência Financeira e Distribuição de Safra")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_parcelas = pd.read_sql_query("SELECT * FROM parcelas", conn)
    
    if df_parcelas.empty:
        st.info("💡 Nenhum dado financeiro registrado ainda.")
    else:
        df_parcelas['vencimento_dt'] = pd.to_datetime(df_parcelas['vencimento']).dt.date
        df_parcelas['ano_int'] = pd.to_datetime(df_parcelas['vencimento']).dt.year
        df_parcelas['mes_int'] = pd.to_datetime(df_parcelas['vencimento']).dt.month
        
        with st.container(border=True):
            st.markdown("##### 🔍 Parâmetros Dinâmicos de Análise")
            col_f1, col_f2 = st.columns(2)
            
            lista_clientes = ["Todos os Cooperados"] + sorted(df_parcelas['cliente_nome'].unique().tolist())
            filtro_cliente = col_f1.selectbox("Filtrar por Produtor / Cooperado", lista_clientes)
            
            anos_disponiveis = sorted(df_parcelas['ano_int'].unique().tolist())
            lista_anos = ["Todos os Anos"] + [str(ano) for ano in anos_disponiveis]
            filtro_ano = col_f2.selectbox("Safra / Ano de Referência", lista_anos, index=0)
            
            st.write("")
            st.markdown("**Filtro por Período Mensal (Deixe vazio para ver o ano completo):**")
            meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
            mapa_meses = {"Jan": 1, "Fev": 2, "Mar": 3, "Abr": 4, "Mai": 5, "Jun": 6, "Jul": 7, "Ago": 8, "Set": 9, "Out": 10, "Nov": 11, "Dez": 12}
            mapa_nomes_meses = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
            
            filtro_meses = st.segmented_control("Meses", options=meses_nomes, selection_mode="multi", label_visibility="collapsed")

        df_dash = df_parcelas.copy()
        hoje = date.today()
        
        if filtro_ano != "Todos os Anos":
            df_dash = df_dash[df_dash['ano_int'] == int(filtro_ano)]
        
        if filtro_meses:
            numeros_meses = [mapa_meses[m] for m in filtro_meses]
            df_dash = df_dash[df_dash['mes_int'].isin(numeros_meses)]
        
        if filtro_cliente != "Todos os Cooperados":
            df_dash = df_dash[df_dash['cliente_nome'] == filtro_cliente]
            
        st.write("")
        
        total_recebido = df_dash[df_dash['status'] == 'Pago']['valor_pago'].sum()
        em_aberto = df_dash[df_dash['status'] != 'Pago']['valor'].sum()
        capital_total_envolvido = total_recebido + em_aberto

        valor_atrasado = 0.0
        for index, row in df_dash[df_dash['status'] != 'Pago'].iterrows():
            if hoje > row['vencimento_dt']:
                valor_atrasado += row['valor']

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #475569;">
                    <div class="ag-title">CAPITAL TOTAL ALOCADO</div>
                    <div class="ag-value" style="color: #1E293B;">R$ {capital_total_envolvido:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #0B4624; background-color: #F6FBF7;">
                    <div class="ag-title" style="color: #0B4624;">LIQUIDADO EM CAIXA</div>
                    <div class="ag-value" style="color: #0B4624;">R$ {total_recebido:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid #D4AF37; background-color: #FDFBF4;">
                    <div class="ag-title" style="color: #B48A04;">SALDO DEVEDOR CARTEIRA</div>
                    <div class="ag-value" style="color: #B48A04;">R$ {em_aberto:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with col4:
            border_c = "#DC2626" if valor_atrasado > 0 else "#CBD5E1"
            bg_c = "#FFF5F5" if valor_atrasado > 0 else "#FFFFFF"
            text_c = "#DC2626" if valor_atrasado > 0 else "#64748B"
            st.markdown(f"""
                <div class="ag-card" style="border-left: 5px solid {border_c}; background-color: {bg_c};">
                    <div class="ag-title" style="color: {text_c};">INADIMPLÊNCIA CRÍTICA</div>
                    <div class="ag-value" style="color: {text_c};">R$ {valor_atrasado:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)

        st.write("")
        taxa_inadimplencia = (valor_atrasado / capital_total_envolvido * 100) if capital_total_envolvido > 0 else 0.0
        st.caption(f"⚡ **Análise de Risco Interno:** Índice de inadimplência posicionado em **{taxa_inadimplencia:.1f}%** sobre os contratos ativos analisados no filtro.")

        st.write("---")
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("##### 📈 Projeção Faturada vs Caixa Realizado (R$)")
            if not df_dash.empty:
                df_grafico = df_dash.copy()
                
                if filtro_ano == "Todos os Anos":
                    df_grafico['Eixo_X'] = df_grafico['ano_int'].astype(str)
                else:
                    df_grafico['Eixo_X'] = df_grafico['mes_int'].map(mapa_nomes_meses)
                
                df_agrupado = df_grafico.groupby('Eixo_X').agg({
                    'valor': 'sum',
                    'valor_pago': 'sum'
                }).rename(columns={'valor': 'Saldo Pendente', 'valor_pago': 'Recebido em Caixa'})
                
                if filtro_ano != "Todos os Anos":
                    ordem_meses_filtrados = [m for m in meses_nomes if m in df_agrupado.index]
                    df_agrupado = df_agrupado.reindex(ordem_meses_filtrados)
                
                st.bar_chart(df_agrupado, height=320, color=["#D4AF37", "#0B4624"])
            else:
                st.info("Sem dados suficientes para construir o gráfico.")
                
        with col_g2:
            st.markdown("##### 🔄 Distribuição de Contratos por Status")
            if not df_dash.empty:
                df_status_grafico = df_dash.copy()
                
                def mapear_status_grafico(row):
                    if row['status'] == 'Pago':
                        return 'Liquidado'
                    elif hoje > row['vencimento_dt']:
                        return 'Em Atraso'
                    else:
                        return 'Vencimento Futuro'
                        
                df_status_grafico['Situação'] = df_status_grafico.apply(mapear_status_grafico, axis=1)
                df_status_final = df_status_grafico.groupby('Situação').size().reset_index(name='Quantidade de Faturas')
                df_status_final = df_status_final.set_index('Situação')
                
                st.bar_chart(df_status_final, height=320, color="#1E5631")

# ------------------------------------------
# TELA 2: CADASTRAR CLIENTE
# ------------------------------------------
elif menu == "Cadastrar Cliente":
    st.header("🌱 Ficha de Matrícula do Cooperado")
    with st.form("form_cliente"):
        nome = st.text_input("Nome Completo do Produtor / Razão Social")
        telefone = st.text_input("WhatsApp para Cobrança Automática (Ex: 16991645755)")
        email = st.text_input("E-mail Eletrônico")
        limite = st.number_input("Cota Máxima de Crédito Pré-Aprovado (R$)", min_value=0.0, step=500.0)
        submit = st.form_submit_button("Efetivar Cadastro")
        
        if submit and nome:
            telefone_limpo = limpar_telefone(telefone)
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                c.execute("INSERT INTO clientes (nome, telefone, email, limite) VALUES (?, ?, ?, ?)", (nome, telefone_limpo, email, limite))
                conn.commit()
            st.success(f"Cooperado {nome} integrado com sucesso à cooperativa!")

# ------------------------------------------
# TELA 3: CONSULTAR E EDITAR CLIENTES
# ------------------------------------------
elif menu == "Consultar Clientes":
    st.header("📋 Cadastro Geral de Cooperados")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_clientes = pd.read_sql_query("SELECT id, nome AS [Nome], telefone AS [WhatsApp], email AS [E-mail], limite AS [Limite de Crédito] FROM clientes ORDER BY nome ASC", conn)
    
    if df_clientes.empty:
        st.info("Nenhum cliente cadastrado até o momento.")
    else:
        st.subheader("🔍 Localizador de Produtores")
        st.dataframe(df_clientes[['Nome', 'WhatsApp', 'E-mail', 'Limite de Crédito']], use_container_width=True, hide_index=True)
        
        st.write("---")
        st.subheader("👤 Auditoria de Limite e Informações Básicas")
        cliente_selecionado = st.selectbox("Selecione um cliente para gerenciar", ["Escolha um cliente..."] + df_clientes['Nome'].tolist())
        
        if cliente_selecionado != "Escolha um cliente...":
            info_cli = df_clientes[df_clientes['Nome'] == cliente_selecionado].iloc[0]
            cliente_id = info_cli['id']
            nome_original = info_cli['Nome']
            
            with sqlite3.connect(DB_PATH) as conn:
                df_part = pd.read_sql_query("SELECT * FROM parcelas WHERE cliente_nome = ?", conn, params=(cliente_selecionado,))
                
            total_contratado = df_part['valor'].sum() if not df_part.empty else 0.0
            total_pago = df_part[df_part['status'] == 'Pago']['valor_pago'].sum() if not df_part.empty else 0.0
            saldo_devedor = df_part[df_part['status'] != 'Pago']['valor'].sum() if not df_part.empty else 0.0
            
            st.markdown(f"""
            <div class="ficha-cliente">
                <h3 style='margin-top:0; color:#0B4624;'>📋 DADOS CADASTRAIS: {nome_original.upper()}</h3>
                <p><strong>📱 WhatsApp Corporativo:</strong> {info_cli['WhatsApp']}</p>
                <p><strong>✉️ E-mail Ativo:</strong> {info_cli['E-mail']}</p>
                <p><strong>💳 Limite Operacional Máximo:</strong> R$ {info_cli['Limite de Crédito']:,.2f}</p>
                <hr style='border-color: #e2e8f0;'/>
                <h4 style='color:#111827; margin-bottom:5px;'>📊 Linha de Crédito na Instituição:</h4>
                <span style='color:#1e40af; font-weight:600;'>💵 Total Contratado: R$ {total_contratado:,.2f}</span> | 
                <span style='color:#166534; font-weight:600;'>🟢 Retorno em Caixa: R$ {total_pago:,.2f}</span> | 
                <span style='color:#991b1b; font-weight:600;'>🔴 Exposição ao Risco (Saldo Devedor): R$ {saldo_devedor:,.2f}</span>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("✏️ Atualizar Dados de Matrícula"):
                with st.form("form_edicao"):
                    st.write("**Modificar Registro:**")
                    novo_nome = st.text_input("Nome / Razão Social", value=nome_original)
                    novo_tel = st.text_input("WhatsApp", value=info_cli['WhatsApp'])
                    novo_email = st.text_input("E-mail", value=info_cli['E-mail'])
                    novo_limite = st.number_input("Limite Operacional (R$)", value=float(info_cli['Limite de Crédito']), step=500.0)
                    
                    submit_edit = st.form_submit_button("Confirmar Atualizações")
                    if submit_edit:
                        tel_limpo = limpar_telefone(novo_tel)
                        with sqlite3.connect(DB_PATH) as conn:
                            c = conn.cursor()
                            if novo_nome != nome_original:
                                c.execute("UPDATE parcelas SET cliente_nome = ? WHERE cliente_nome = ?", (novo_nome, nome_original))
                            c.execute("UPDATE clientes SET nome=?, telefone=?, email=?, limite=? WHERE id=?", 
                                      (novo_nome, tel_limpo, novo_email, novo_limite, cliente_id))
                            conn.commit()
                        st.success("Dados cadastrais reestruturados!")
                        st.rerun()

                st.write("---")
                st.markdown("### ⚠️ Revogação de Acesso")
                chave_confirmar_del = f"confirmar_del_cli_{cliente_id}"
                if chave_confirmar_del not in st.session_state:
                    st.session_state[chave_confirmar_del] = False
                
                if not st.session_state[chave_confirmar_del]:
                    if st.button("🗑️ Desvincular Cooperado do Sistema", type="primary"):
                        st.session_state[chave_confirmar_del] = True
                        st.rerun()
                else:
                    st.markdown(f'<div class="alerta-exclusao"><strong>🚨 DESTRUIÇÃO DE REGISTRO:</strong> Esta ação expurgará o cooperado <b>{nome_original}</b> e todo o seu histórico financeiro do banco de dados de forma irreversível!</div>', unsafe_allow_html=True)
                    digitado_cli = st.text_input(f"Confirme digitando o nome exato do registro para exclusão ({nome_original}):")
                    
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("CONFIRMAR DELEÇÃO PERMANENTE 💀", type="primary", disabled=(digitado_cli != nome_original)):
                            with sqlite3.connect(DB_PATH) as conn:
                                c = conn.cursor()
                                c.execute("DELETE FROM clientes WHERE id=?", (cliente_id,))
                                c.execute("DELETE FROM parcelas WHERE cliente_nome=?", (nome_original,))
                                conn.commit()
                            st.session_state[chave_confirmar_del] = False
                            st.rerun()
                    with col_b2:
                        if st.button("Abortar"):
                            st.session_state[chave_confirmar_del] = False
                            st.rerun()

# ------------------------------------------
# TELA 4: NOVO EMPRÉSTIMO
# ------------------------------------------
elif menu == "Novo Empréstimo":
    st.header("🌾 Concessão de Crédito e Financiamento")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_clientes = pd.read_sql_query("SELECT nome FROM clientes", conn)
    
    if df_clientes.empty:
        st.warning("Efetue o cadastro de um produtor antes de emitir crédito.")
    else:
        with st.form("form_emprestimo"):
            cliente = st.selectbox("Selecione o Beneficiário", df_clientes['nome'].tolist())
            taxa_utilizada = st.number_input("Taxa de Juros Subsidiada / Custeio (%)", min_value=0.0, value=10.0, step=0.5)
            valor_total_solicitado = st.number_input("Aporte Solicitado (R$)", min_value=0.0, step=1000.0)
            qtd_parcelas = st.number_input("Número de Parcelas Finais", min_value=1, value=2)
            primeiro_vencimento = st.date_input("Vencimento do Primeiro Bloco")
            
            submit_emp = st.form_submit_button("Gerar Plano de Amortização Fixo")
            
            if submit_emp and valor_total_solicitado > 0:
                emp_id = f"EMP_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                amortizacao_fixa = valor_total_solicitado / qtd_parcelas
                saldo_devedor_aux = valor_total_solicitado
                acumulador_total_com_juros = 0.0
                
                for i in range(1, int(qtd_parcelas) + 1):
                    juros_mes = saldo_devedor_aux * (taxa_utilizada / 100)
                    acumulador_total_com_juros += (amortizacao_fixa + juros_mes)
                    saldo_devedor_aux -= amortizacao_fixa
                
                valor_parcela_fixa = round(acumulador_total_com_juros / qtd_parcelas, 2)
                
                with sqlite3.connect(DB_PATH) as conn:
                    c = conn.cursor()
                    for i in range(1, int(qtd_parcelas) + 1):
                        data_venc = primeiro_vencimento if i == 1 else primeiro_vencimento + timedelta(days=(i-1)*30)
                        c.execute("""INSERT INTO parcelas (cliente_nome, valor, vencimento, status, valor_pago, metodo, n_parcela, emprestimo_id, taxa_utilizada) 
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                                  (cliente, valor_parcela_fixa, data_venc.strftime('%Y-%m-%d'), 'Pendente', 0.0, '', f"{i}/{qtd_parcelas}", emp_id, taxa_utilizada))
                    conn.commit()
                st.success(f"✅ Cronograma registrado com o Token de Rastreamento: {emp_id}")

# ------------------------------------------
# TELA 5: RECEBIMENTOS E BAIXAS
# ------------------------------------------
elif menu == "Recebimentos e Baixas":
    st.header("📥 Central de Liquidação de Títulos")
    
    query = """
        SELECT p.id, p.cliente_nome, p.valor, p.vencimento, p.status, p.n_parcela, p.emprestimo_id, p.taxa_utilizada, p.valor_pago, c.telefone, c.email 
        FROM parcelas p
        LEFT JOIN clientes c ON p.cliente_nome = c.nome
        ORDER BY p.cliente_nome ASC, p.vencimento ASC
    """
    with sqlite3.connect(DB_PATH) as conn:
        df_geral = pd.read_sql_query(query, conn)
    
    if df_geral.empty:
        st.success("Nenhum registro encontrado no sistema.")
    else:
        clientes_unicos = sorted(df_geral['cliente_nome'].unique().tolist())
        
        cliente_sel = st.selectbox(
            "🎯 Selecione o Cooperado para gerenciar faturas:", 
            ["-- Escolha um Cliente --"] + clientes_unicos,
            key="sb_cliente_baixa_persistente"
        )
        
        if cliente_sel != "-- Escolha um Cliente --":
            df_cli = df_geral[df_geral['cliente_nome'] == cliente_sel]
            emprestimos_unicos = df_cli['emprestimo_id'].unique()
            
            for emp in emprestimos_unicos:
                contrato_valido = pd.notna(emp) and str(emp).strip() != "" and str(emp).strip().lower() != "nan"
                emp_label = str(emp).strip() if contrato_valido else "CONTRATO ANTIGO"
                
                df_emp = df_cli[df_cli['emprestimo_id'] == emp]
                total_restante_contrato = df_emp[df_emp['status'] != 'Pago']['valor'].sum()
                taxa_exibida = df_emp['taxa_utilizada'].iloc[0] if pd.notna(df_emp['taxa_utilizada'].iloc[0]) else 0.0
                
                st.markdown(f"""
                <div style='background-color: #ffffff; padding: 18px; border-radius: 10px; margin: 15px 0; border: 1px solid #e2e8f0; border-left: 6px solid #0B4624; box-shadow: 0 4px 12px rgba(0,0,0,0.02);'>
                    <span style='font-size: 14px; color: #64748B;'>📑 <b>Lote de Contrato:</b> <code style='color:#0B4624; font-weight:bold;'>{emp_label}</code> | <b>Taxa Aplicada:</b> {taxa_exibida}%</span><br/>
                    <div style='color: #B48A04; font-size: 19px; font-weight: bold; margin-top: 5px;'>💵 Saldo Remanescente: R$ {total_restante_contrato:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
                
                dados_tabela = []
                parcelas_pendentes_opcoes = {}
                
                for index, row in df_emp.iterrows():
                    valor_atualizado, dias_atraso = calcular_juros_atraso(row['valor'], row['vencimento'])
                    venc_pt = datetime.strptime(row['vencimento'], '%Y-%m-%d').strftime('%d/%m/%Y')
                    
                    if row['status'] == 'Pago':
                        status_txt = "🟢 Liquidado"
                        v_base = row['valor']
                        v_pago_exibir = row['valor_pago']
                    elif dias_atraso > 0:
                        status_txt = f"🔴 Vencido (+{dias_atraso} dias)"
                        v_base = row['valor']
                        v_pago_exibir = valor_atualizado
                    else:
                        status_txt = "🟡 Em Aberto"
                        v_base = row['valor']
                        v_pago_exibir = row['valor']
                        
                    dados_tabela.append({
                        "Cota": row['n_parcela'],
                        "Data Vencimento": venc_pt,
                        "Valor Base (R$)": f"R$ {v_base:,.2f}",
                        "Ajustado / Requerido": f"R$ {v_pago_exibir:,.2f}",
                        "Status Atual": status_txt
                    })
                    
                    if row['status'] != 'Pago':
                        label_combo = f"Cota {row['n_parcela']} (Vence em {venc_pt} | Requerido: R$ {v_pago_exibir:.2f})"
                        parcelas_pendentes_opcoes[label_combo] = (row, v_pago_exibir)
                
                df_visualizacao = pd.DataFrame(dados_tabela)
                st.dataframe(df_visualizacao, use_container_width=True, hide_index=True)
                
                if parcelas_pendentes_opcoes:
                    with st.container(border=True):
                        st.markdown(f"##### 📥 Registrar Amortização / Baixa - `Contrato {emp_label}`")
                        
                        col_sel_p, col_forma, col_v_pago = st.columns([2, 1, 1])
                        
                        opcao_selecionada = col_sel_p.selectbox(
                            "Selecione a cota correspondente:", 
                            options=list(parcelas_pendentes_opcoes.keys()),
                            key=f"sel_p_{emp_label}_{emp}"
                        )
                        
                        row_selecionada, valor_requerido = parcelas_pendentes_opcoes[opcao_selecionada]
                        p_id = row_selecionada['id']
                        
                        m_pago = col_forma.selectbox("Via de Entrada", ["Pix", "Dinheiro", "Cartão"], key=f"m_{p_id}")
                        v_pago = col_v_pago.number_input("Montante Pago (R$)", value=float(valor_requerido), key=f"v_{p_id}", step=50.0)
                        
                        dif = v_pago - valor_requerido
                        
                        if row_selecionada['telefone']:
                            venc_pt_sel = datetime.strptime(row_selecionada['vencimento'], '%Y-%m-%d').strftime('%d/%m/%Y')
                            msg_wsp = f"Olá {row_selecionada['cliente_nome']}, lembrete da sua parcela {row_selecionada['n_parcela']} no valor de R$ {valor_requerido:.2f} com vencimento em {venc_pt_sel}."
                            texto_codificado = urllib.parse.quote(msg_wsp)
                            link_final = f"https://wa.me/55{row_selecionada['telefone']}?text={texto_codificado}"
                            st.markdown(f'<a href="{link_final}" target="_blank" class="wa-button">💬 Notificar Produtor via WhatsApp</a>', unsafe_allow_html=True)
                        
                        if dif > 0.01:
                            if not contrato_valido:
                                st.error("⚠️ **Bloqueio de Amortização:** Este empréstimo é antigo e não possui ID.")
                            else:
                                with sqlite3.connect(DB_PATH) as conn:
                                    c_aux = conn.cursor()
                                    c_aux.execute("SELECT id, valor FROM parcelas WHERE emprestimo_id=? AND status != 'Pago' AND id != ?", (str(emp), p_id))
                                    proximas_parc = c_aux.fetchall()
                                
                                if proximas_parc:
                                    qtd_r = len(proximas_parc)
                                    total_r = sum(p[1] for p in proximas_parc)
                                    novo_total_r = max(0.0, total_r - dif)
                                    novo_val = round(novo_total_r / qtd_r, 2)
                                    
                                    st.success(f"📈 **Amortização de Safra Detectada! (+ R$ {dif:.2f})**")
                                    st.info(f"✨ **RECALCULANDO PARCELAS:** Ao confirmar, o saldo residual será diluído. As próximas **{qtd_r}** faturas cairão para **R$ {novo_val:.2f}** cada!")
                                else:
                                    st.warning("⚠️ Esta é a última parcela em aberto deste contrato.")
                        elif dif < -0.01:
                            st.warning(f"📉 Pagamento parcial. Saldo cobrável restante nesta parcela será de R$ {abs(dif):.2f}")
                        
                        if st.button("Autenticar Transação ✅", key=f"btn_baixa_{p_id}", type="primary"):
                            with sqlite3.connect(DB_PATH) as conn:
                                c = conn.cursor()
                                
                                if v_pago < valor_requerido:
                                    resto = valor_requerido - v_pago
                                    c.execute("UPDATE parcelas SET valor=? WHERE id=?", (resto, p_id))
                                    c.execute("""INSERT INTO parcelas (cliente_nome, valor, vencimento, status, valor_pago, metodo, n_parcela, emprestimo_id, taxa_utilizada) 
                                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                              (row_selecionada['cliente_nome'], v_pago, row_selecionada['vencimento'], 'Pago', v_pago, m_pago, f"{row_selecionada['n_parcela']} (Parcial)", str(emp), row_selecionada['taxa_utilizada']))
                                else:
                                    c.execute("UPDATE parcelas SET status='Pago', valor_pago=?, metodo=? WHERE id=?", (v_pago, m_pago, p_id))
                                    
                                    if dif > 0.01 and contrato_valido:
                                        c.execute("SELECT id, valor FROM parcelas WHERE emprestimo_id=? AND status != 'Pago' AND id != ?", (str(emp), p_id))
                                        parcelas_restantes = c.fetchall()
                                        
                                        if parcelas_restantes:
                                            qtd_restantes = len(parcelas_restantes)
                                            total_restante = sum(p[1] for p in parcelas_restantes)
                                            novo_total_restante = max(0.0, total_restante - dif)
                                            novo_valor_parcela = round(novo_total_restante / qtd_restantes, 2)
                                            
                                            for p_restante in parcelas_restantes:
                                                c.execute("UPDATE parcelas SET valor=? WHERE id=?", (novo_valor_parcela, p_restante[0]))
                                conn.commit()
                            st.toast("💰 Baixa processada com sucesso!", icon="✅")
                            st.rerun()
                else:
                    st.success("🎉 Todas as faturas deste contrato foram totalmente liquidadas!")
                st.write("---")

# ------------------------------------------
# TELA 6: RELATÓRIOS
# ------------------------------------------
elif menu == "Relatórios":
    st.header("📊 Inteligência e Exportação de Relatórios")
    
    with sqlite3.connect(DB_PATH) as conn:
        df_todas = pd.read_sql_query("SELECT * FROM parcelas", conn)
    
    st.subheader("🔍 Filtros de Exportação")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    lista_clientes = ["TODOS"] + df_todas['cliente_nome'].unique().tolist()
    filtro_cli = col_f1.selectbox("Filtrar por Cooperado", lista_clientes)
    
    data_min_banco = pd.to_datetime(df_todas['vencimento']).min().date() if not df_todas.empty else date.today()
    data_max_banco = pd.to_datetime(df_todas['vencimento']).max().date() if not df_todas.empty else date.today()
    
    dt_inicio = col_f2.date_input("Vencimento a partir de", data_min_banco)
    dt_fim = col_f3.date_input("Vencimento até", data_max_banco)
    
    df_filtrado = df_todas.copy()
    df_filtrado['vencimento'] = pd.to_datetime(df_filtrado['vencimento']).dt.date
    
    if filtro_cli != "TODOS":
        df_filtrado = df_filtrado[df_filtrado['cliente_nome'] == filtro_cli]
        
    df_filtrado = df_filtrado[(df_filtrado['vencimento'] >= dt_inicio) & (df_filtrado['vencimento'] <= dt_fim)]
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Relatorio_Filtrado')
    
    st.download_button(
        label="📊 Exportar Relatório de Carteira para Excel",
        data=buffer.getvalue(),
        file_name=f"Relatorio_Financeiro_Cia_Menezes.xlsx",
        mime="application/vnd.ms-excel",
        type="primary"
    )
