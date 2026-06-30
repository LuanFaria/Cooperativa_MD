import streamlit as st
import pandas as pd
import sqlite3
from datetime import date, datetime, timedelta
import io
import os
import urllib.parse

# ==========================================
# 1. CONFIGURAÇÃO E MIGRAÇÃO AUTOMÁTICA
# ==========================================
def realizar_backup_seguranca():
    if os.path.exists('cooperativa.db'):
        try:
            conn = sqlite3.connect('cooperativa.db')
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
            conn.close()
        except Exception as e:
            print(f"Erro ao gerar backup automático: {e}")

def init_db():
    conn = sqlite3.connect('cooperativa.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS clientes 
                 (id INTEGER PRIMARY KEY, nome TEXT, telefone TEXT, email TEXT, limite REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS parcelas 
                 (id INTEGER PRIMARY KEY, cliente_nome TEXT, valor REAL, 
                  vencimento DATE, status TEXT, valor_pago REAL, metodo TEXT, 
                  n_parcela TEXT, emprestimo_id TEXT, taxa_utilizada REAL)''')
    conn.commit()
    return conn

if 'backup_inicial_rodado' not in st.session_state:
    realizar_backup_seguranca()
    st.session_state['backup_inicial_rodado'] = True

conn = init_db()

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

# ==========================================
# 3. INTERFACE DO USUÁRIO (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Gestão de Crédito", page_icon="💰", layout="wide")

st.markdown("""
    <style>
    .wa-button {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background-color: #25D366;
        color: white !important;
        padding: 8px 16px;
        border-radius: 5px;
        text-decoration: none;
        font-weight: bold;
        text-align: center;
        margin-bottom: 15px;
    }
    .wa-button:hover { background-color: #1ebe5d; text-decoration: none; }
    .box-info { background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef; margin-bottom: 10px; }
    .ficha-cliente { background-color: #f1f5f9; padding: 20px; border-radius: 8px; border: 1px solid #cbd5e1; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("💰 Sistema da Cooperativa de Crédito")

# Adicionado "Consultar Clientes" no Menu
menu = st.sidebar.selectbox("Menu Principal", 
                            ["Dashboard", "Cadastrar Cliente", "Consultar Clientes", "Novo Empréstimo", "Recebimentos e Baixas", "Relatórios"])

# ------------------------------------------
# TELA 1: DASHBOARD
# ------------------------------------------
if menu == "Dashboard":
    st.markdown("## 📊 Visão Geral do Negócio")
    st.write("---")
    df_parcelas = pd.read_sql_query("SELECT * FROM parcelas", conn)
    
    if not df_parcelas.empty:
        total_emprestado = df_parcelas['valor'].sum()
        total_recebido = df_parcelas[df_parcelas['status'] == 'Pago']['valor_pago'].sum()
        em_aberto = df_parcelas[df_parcelas['status'].str.contains('Pendente', na=False)]['valor'].sum()
        
        df_pagas = df_parcelas[df_parcelas['status'] == 'Pago']
        lucro_juros = df_pagas['valor_pago'].sum() - df_pagas['valor'].sum()
        if lucro_juros < 0: lucro_juros = 0.0

        hoje = date.today()
        valor_atrasado = 0.0
        for index, row in df_parcelas[df_parcelas['status'].str.contains('Pendente', na=False)].iterrows():
            venc_date = datetime.strptime(row['vencimento'], '%Y-%m-%d').date()
            if hoje > venc_date:
                valor_atrasado += row['valor']

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"<div style='background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #29b5e8;'><p style='margin:0; font-size:14px; color:#555; font-weight:bold;'>TOTAL EM CARTEIRA (GERADO)</p><h2 style='margin:0; color:#1f2937;'>R$ {total_emprestado:,.2f}</h2></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div style='background-color: #e6f4ea; padding: 20px; border-radius: 10px; border-left: 5px solid #137333;'><p style='margin:0; font-size:14px; color:#137333; font-weight:bold;'>TOTAL RECEBIDO (CAIXA)</p><h2 style='margin:0; color:#137333;'>R$ {total_recebido:,.2f}</h2></div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div style='background-color: #e8f0fe; padding: 20px; border-radius: 10px; border-left: 5px solid #1a73e8;'><p style='margin:0; font-size:14px; color:#1a73e8; font-weight:bold;'>LUCRO EXTRA REALIZADO</p><h2 style='margin:0; color:#1a73e8;'>R$ {lucro_juros:,.2f}</h2></div>", unsafe_allow_html=True)

        st.write("")
        col4, col5 = st.columns(2)
        with col4:
            st.markdown(f"<div style='background-color: #fef7e0; padding: 20px; border-radius: 10px; border-left: 5px solid #f9ab00;'><p style='margin:0; font-size:14px; color:#b06000; font-weight:bold;'>VALOR RESTANTE A RECEBER</p><h2 style='margin:0; color:#b06000;'>R$ {em_aberto:,.2f}</h2></div>", unsafe_allow_html=True)
        with col5:
            bg_atrasado = "#fce8e6" if valor_atrasado > 0 else "#f0f2f6"
            txt_atrasado = "#c5221f" if valor_atrasado > 0 else "#555"
            border_atrasado = "#c5221f" if valor_atrasado > 0 else "#ccc"
            st.markdown(f"<div style='background-color: {bg_atrasado}; padding: 20px; border-radius: 10px; border-left: 5px solid {border_atrasado};'><p style='margin:0; font-size:14px; color:{txt_atrasado}; font-weight:bold;'>⚠️ VALOR EM ATRASO CRÍTICO</p><h2 style='margin:0; color:{txt_atrasado};'>R$ {valor_atrasado:,.2f}</h2></div>", unsafe_allow_html=True)
    else:
        st.info("💡 Nenhum dado financeiro registrado ainda.")

# ------------------------------------------
# TELA 2: CADASTRAR CLIENTE
# ------------------------------------------
elif menu == "Cadastrar Cliente":
    st.header("Novo Cliente")
    with st.form("form_cliente"):
        nome = st.text_input("Nome Completo")
        telefone = st.text_input("WhatsApp (Somente números, ex: 11940028922)")
        email = st.text_input("E-mail")
        limite = st.number_input("Valor Máximo de Crédito (R$)", min_value=0.0, step=100.0)
        submit = st.form_submit_button("Salvar Cliente")
        
        if submit and nome:
            c = conn.cursor()
            c.execute("INSERT INTO clientes (nome, telefone, email, limite) VALUES (?, ?, ?, ?)", (nome, telefone, email, limite))
            conn.commit()
            st.success(f"Cliente {nome} cadastrado com sucesso!")

# ------------------------------------------
# TELA NOVA: CONSULTAR CLIENTES
# ------------------------------------------
elif menu == "Consultar Clientes":
    st.header("📋 Cadastro e Base de Clientes")
    
    df_clientes = pd.read_sql_query("SELECT id AS [ID], nome AS [Nome], telefone AS [WhatsApp], email AS [E-mail], limite AS [Limite de Crédito] FROM clientes ORDER BY nome ASC", conn)
    
    if df_clientes.empty:
        st.info("Nenhum cliente cadastrado até o momento.")
    else:
        st.subheader("🔍 Lista Geral")
        st.dataframe(df_clientes, use_container_width=True, hide_index=True)
        
        st.write("---")
        st.subheader("👤 Ficha Detalhada do Cliente")
        
        cliente_selecionado = st.selectbox("Selecione um cliente para ver o histórico de crédito", ["Escolha um cliente..."] + df_clientes['Nome'].tolist())
        
        if cliente_selecionado != "Escolha um cliente...":
            # Puxa informações básicas do cliente
            info_cli = df_clientes[df_clientes['Nome'] == cliente_selecionado].iloc[0]
            
            # Puxa dados financeiros deste cliente
            df_part = pd.read_sql_query("SELECT * FROM parcelas WHERE cliente_nome = ?", conn, params=(cliente_selecionado,))
            
            total_contratado = df_part['valor'].sum() if not df_part.empty else 0.0
            total_pago = df_part[df_part['status'] == 'Pago']['valor_pago'].sum() if not df_part.empty else 0.0
            saldo_devedor = df_part[df_part['status'].str.contains('Pendente', na=False)]['valor'].sum() if not df_part.empty else 0.0
            
            st.markdown(f"""
            <div class="ficha-cliente">
                <h3 style='margin-top:0; color:#1e293b;'>📋 Dados Cadastrais: {info_cli['Nome'].upper()}</h3>
                <p><strong>📱 WhatsApp / Telefone:</strong> {info_cli['WhatsApp']}</p>
                <p><strong>✉️ E-mail:</strong> {info_cli['E-mail']}</p>
                <p><strong>💳 Limite Pré-Aprovado:</strong> R$ {info_cli['Limite de Crédito']:,.2f}</p>
                <hr style='border-color: #cbd5e1;'/>
                <h4 style='color:#0f172a; margin-bottom:5px;'>📊 Resumo Financeiro na Cooperativa:</h4>
                <span style='color:#1e40af;'>💵 Total Gerado em Contratos: R$ {total_contratado:,.2f}</span> | 
                <span style='color:#166534;'>🟢 Total Pago: R$ {total_pago:,.2f}</span> | 
                <span style='color:#991b1b;'>🔴 Saldo Devedor Restante: R$ {saldo_devedor:,.2f}</span>
            </div>
            """, unsafe_allow_html=True)
            
            if not df_part.empty:
                st.write("**Histórico de Parcelas Recorrentes:**")
                # Exibe de forma limpa as parcelas vinculadas a ele
                df_part_exibicao = df_part[['emprestimo_id', 'n_parcela', 'valor', 'vencimento', 'status']].copy()
                df_part_exibicao.columns = ['ID do Contrato', 'Nº Parcela', 'Valor Original', 'Data Vencimento', 'Situação']
                st.dataframe(df_part_exibicao, use_container_width=True, hide_index=True)
            else:
                st.caption("💡 Este cliente não possui nenhum contrato de empréstimo ativo ou finalizado.")

# ------------------------------------------
# TELA 3: NOVO EMPRÉSTIMO
# ------------------------------------------
elif menu == "Novo Empréstimo":
    st.header("Registrar Contrato de Empréstimo (Parcelas Fixas)")
    df_clientes = pd.read_sql_query("SELECT nome FROM clientes", conn)
    
    if df_clientes.empty:
        st.warning("Cadastre um cliente primeiro.")
    else:
        with st.form("form_emprestimo"):
            cliente = st.selectbox("Selecione o Cliente", df_clientes['nome'].tolist())
            taxa_utilizada = st.number_input("Taxa de Juros para este Empréstimo (%)", min_value=0.0, max_value=100.0, value=10.0, step=0.5)
            valor_total_solicitado = st.number_input("Valor Total Solicitado (R$)", min_value=0.0, step=100.0)
            qtd_parcelas = st.number_input("Quantidade de Parcelas", min_value=1, max_value=72, value=2, step=1)
            primeiro_vencimento = st.date_input("Data do Primeiro Vencimento")
            
            submit_emp = st.form_submit_button("Gerar Cronograma Fixo")
            
            if submit_emp and valor_total_solicitado > 0:
                c = conn.cursor()
                emp_id = f"EMP_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                
                amortizacao_fixa = valor_total_solicitado / qtd_parcelas
                saldo_devedor_aux = valor_total_solicitado
                acumulador_total_com_juros = 0.0
                
                for i in range(1, int(qtd_parcelas) + 1):
                    juros_mes = saldo_devedor_aux * (taxa_utilizada / 100)
                    acumulador_total_com_juros += (amortizacao_fixa + juros_mes)
                    saldo_devedor_aux -= amortizacao_fixa
                
                valor_parcela_fixa = round(acumulador_total_com_juros / qtd_parcelas, 2)
                
                for i in range(1, int(qtd_parcelas) + 1):
                    if i == 1:
                        data_venc = primeiro_vencimento
                    else:
                        data_venc = primeiro_vencimento + timedelta(days=(i-1)*30)
                        
                    c.execute("""INSERT INTO parcelas (cliente_nome, valor, vencimento, status, valor_pago, metodo, n_parcela, emprestimo_id, taxa_utilizada) 
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                              (cliente, valor_parcela_fixa, data_venc.strftime('%Y-%m-%d'), 'Pendente', 0.0, '', f"{i}/{qtd_parcelas}", emp_id, taxa_utilizada))
                    
                conn.commit()
                st.success(f"✅ Empréstimo registrado! {qtd_parcelas} parcelas fixas de R$ {valor_parcela_fixa:.2f} (Total com juros: R$ {acumulador_total_com_juros:.2f})")

# ------------------------------------------
# TELA 4: RECEBIMENTOS E BAIXAS
# ------------------------------------------
elif menu == "Recebimentos e Baixas":
    st.header("Controle por Clientes e Contratos")
    
    query = """
        SELECT p.id, p.cliente_nome, p.valor, p.vencimento, p.status, p.n_parcela, p.emprestimo_id, p.taxa_utilizada, c.telefone, c.email 
        FROM parcelas p
        LEFT JOIN clientes c ON p.cliente_nome = c.nome
        ORDER BY p.cliente_nome ASC, p.vencimento ASC
    """
    df_geral = pd.read_sql_query(query, conn)
    
    if df_geral.empty:
        st.success("Nenhum registro encontrado no sistema.")
    else:
        clientes_unicos = df_geral['cliente_nome'].unique()
        for cli in clientes_unicos:
            df_cli = df_geral[df_geral['cliente_nome'] == cli]
            
            with st.expander(f"👤 CLIENTE: {cli.upper()}", expanded=False):
                emprestimos_unicos = df_cli['emprestimo_id'].unique()
                for emp in emprestimos_unicos:
                    df_emp = df_cli[df_cli['emprestimo_id'] == emp]
                    
                    total_restante_contrato = df_emp[df_emp['status'].str.contains('Pendente', na=False)]['valor'].sum()
                    taxa_exibida = df_emp['taxa_utilizada'].iloc[0]
                    
                    st.markdown(f"""
                    <div style='background-color: #eef2f7; padding: 12px; border-radius: 6px; margin: 10px 0; border-left: 5px solid #1a73e8;'>
                        <strong>📑 Contrato:</strong> <code style='color:#1a73e8;'>{emp}</code> | <strong>Taxa base:</strong> {taxa_exibida}% <br/>
                        <span style='color: #c5221f; font-size: 16px; font-weight: bold;'>💰 Saldo Devedor Atual: R$ {total_restante_contrato:,.2f}</span>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    for index, row in df_emp.iterrows():
                        valor_atualizado, dias_atraso = calcular_juros_atraso(row['valor'], row['vencimento'])
                        venc_pt = datetime.strptime(row['vencimento'], '%Y-%m-%d').strftime('%d/%m/%Y')
                        
                        if row['status'] == 'Pago':
                            status_badge = "🟢 Pago"
                            label_botao = f"Parc. {row['n_parcela']} | R$ {row['valor']:.2f} | {status_badge}"
                        elif dias_atraso > 0:
                            status_badge = f"🔴 Atrasado ({dias_atraso} dias)"
                            label_botao = f"⚠️ Parc. {row['n_parcela']} | R$ {valor_atualizado:.2f} | {status_badge}"
                        else:
                            status_badge = "🟡 Pendente"
                            label_botao = f"Parc. {row['n_parcela']} | R$ {row['valor']:.2f} | {status_badge}"
                        
                        with st.expander(label_botao):
                            st.markdown(f"### Detalhes da Parcela {row['n_parcela']}")
                            
                            if row['status'] == 'Pago':
                                st.success("✨ Esta parcela já foi totalmente liquidada!")
                            else:
                                if row['telefone']:
                                    msg_wsp = f"Olá {row['cliente_nome']}, lembrete da sua parcela {row['n_parcela']} no valor de R$ {valor_atualizado:.2f} com vencimento em {venc_pt}."
                                    texto_codificado = urllib.parse.quote(msg_wsp)
                                    link_final = f"https://wa.me/55{row['telefone']}?text={texto_codificado}"
                                    st.markdown(f'<a href="{link_final}" target="_blank" class="wa-button">💬 Enviar Cobrança no WhatsApp</a>', unsafe_allow_html=True)
                                
                                col_m, col_val = st.columns(2)
                                m_pago = col_m.selectbox("Forma de Pagamento", ["Pix", "Dinheiro", "Cartão"], key=f"m_{row['id']}")
                                v_pago = col_val.number_input("Valor Recebido (R$)", value=float(valor_atualizado), key=f"v_{row['id']}", step=10.0)
                                
                                dif = v_pago - valor_atualizado
                                st.markdown("---")
                                st.write("**Resumo da Operação:**")
                                st.write(f"- Valor Original: R$ {row['valor']:.2f}")
                                st.write(f"- Valor Com Juros de Atraso: **R$ {valor_atualizado:.2f}**")
                                
                                if dif > 0:
                                    st.success(f"📈 Amortização maior detectada! (+ R$ {dif:.2f})")
                                elif dif < 0:
                                    st.warning(f"📉 Pagamento parcial. O sistema criará um resíduo de R$ {abs(dif):.2f}")
                                
                                if st.button("Confirmar Baixa ✅", key=f"b_btn_{row['id']}", type="primary"):
                                    c = conn.cursor()
                                    if v_pago < valor_atualizado:
                                        resto = valor_atualizado - v_pago
                                        c.execute("UPDATE parcelas SET valor=? WHERE id=?", (resto, row['id']))
                                        c.execute("""INSERT INTO parcelas (cliente_nome, valor, vencimento, status, valor_pago, metodo, n_parcela, emprestimo_id, taxa_utilizada) 
                                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                                  (row['cliente_nome'], v_pago, row['vencimento'], 'Pago', v_pago, m_pago, f"{row['n_parcela']} (Parcial)", row['emprestimo_id'], row['taxa_utilizada']))
                                    else:
                                        c.execute("UPDATE parcelas SET status='Pago', valor_pago=?, metodo=? WHERE id=?", (v_pago, m_pago, row['id']))
                                    
                                    conn.commit()
                                    st.rerun()

# ------------------------------------------
# TELA 5: RELATÓRIOS
# ------------------------------------------
elif menu == "Relatórios":
    st.header("Relatório Financeiro Customizado")
    
    df_todas = pd.read_sql_query("SELECT * FROM parcelas", conn)
    
    st.subheader("🔍 Filtros de Exportação")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    lista_clientes = ["TODOS"] + df_todas['cliente_nome'].unique().tolist()
    filtro_cli = col_f1.selectbox("Filtrar por Cliente", lista_clientes)
    
    data_min_banco = pd.to_datetime(df_todas['vencimento']).min().date() if not df_todas.empty else date.today()
    data_max_banco = pd.to_datetime(df_todas['vencimento']).max().date() if not df_todas.empty else date.today()
    
    dt_inicio = col_f2.date_input("Vencimento a partir de", data_min_banco)
    dt_fim = col_f3.date_input("Vencimento até", data_max_banco)
    
    df_filtrado = df_todas.copy()
    df_filtrado['vencimento'] = pd.to_datetime(df_filtrado['vencimento']).dt.date
    
    if filtro_cli != "TODOS":
        df_filtrado = df_filtrado[df_filtrado['cliente_nome'] == filtro_cli]
        
    df_filtrado = df_filtrado[(df_filtrado['vencimento'] >= dt_inicio) & (df_filtrado['vencimento'] <= dt_fim)]
    
    st.write(f"Registros encontrados com os filtros aplicados: **{len(df_filtrado)}**")
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Relatorio_Filtrado')
    
    st.download_button(
        label="📊 Exportar Relatório Filtrado para Excel",
        data=buffer.getvalue(),
        file_name=f"Relatorio_Financeiro_Cooperativa.xlsx",
        mime="application/vnd.ms-excel",
        type="primary"
    )