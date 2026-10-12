import time
from datetime import datetime
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="WINFUT - Dashboard de Fluxo", layout="wide")

# --- SISTEMA DE AUTENTICAÇÃO NATIVO SEGURO ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.markdown("<h2 style='text-align: center;'>Login - WINFUT Flow</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        with st.form("form_login"):
            usuario = st.text_input("Username")
            senha = st.text_input("Password", type="password")
            submit = st.form_submit_button("Entrar")
            
            if submit:
                if (usuario == "sergio" and senha == "123456") or (usuario == "convidado" and senha == "trade2026"):
                    st.session_state["autenticado"] = True
                    st.session_state["nome_usuario"] = "Sérgio Oliveira" if usuario == "sergio" else "Convidado Trader"
                    st.rerun()
                else:
                    st.error("Utilizador ou senha incorretos.")
    st.stop()

# --- ÁREA DO DASHBOARD APÓS LOGIN ---
st.sidebar.success(f"Bem-vindo ao sistema, *{st.session_state['nome_usuario']}*!")
if st.sidebar.button("Sair"):
    st.session_state["autenticado"] = False
    st.rerun()

dashboard_placeholder = st.empty()

MAX_HISTORICO = 10000
h_tempos_acm = []
h_shadow_acm, h_converge_acm = [], []
h_feeders_acm, h_equalizers_acm = [], []
h_disruption_acm, h_rebounders_acm = [], []

blocos_30m = []
bloco_atual_dados = {'tempos': [], 'shadow': [], 'converge': [], 'feeders': [], 'equalizers': [], 'disruption': [], 'rebounders': []}

ultimo_bloco_minuto = None
base_shadow, base_converge = 0, 0
base_feeders, base_equalizers = 0, 0
base_disruption, base_rebounders = 0, 0

base_shadow_5m, base_converge_5m, base_feeders_5m, base_equalizers_5m, base_disruption_5m, base_rebounders_5m = 0, 0, 0, 0, 0, 0
base_shadow_60m, base_converge_60m, base_feeders_60m, base_equalizers_60m, base_disruption_60m, base_rebounders_60m = 0, 0, 0, 0, 0, 0
base_shadow_dia, base_converge_dia, base_feeders_dia, base_equalizers_dia, base_disruption_dia, base_rebounders_dia = None, None, None, None, None, None

ultimo_5m, ultimo_60m, ultimo_dia = None, None, None
pico_shadow_5m, pico_converge_5m, pico_feeders_5m = 1.0, 1.0, 1.0
pico_equalizers_5m, pico_disruption_5m, pico_rebounders_5m = 1.0, 1.0, 1.0
pico_shadow_dia, pico_converge_dia, pico_feeders_dia = 1.0, 1.0, 1.0
pico_equalizers_dia, pico_disruption_dia, pico_rebounders_dia = 1.0, 1.0, 1.0

grupo_1 = ['XP', 'BTG', 'Genial', 'Toro', 'CM']
grupo_2 = ['UBS', 'Ágora', 'Agora', 'Itaú', 'Itau']
grupo_3 = ['Goldman', 'Morgan', 'JP Morgan', 'JPMorgan']

estado_visibilidade = {}
contador_pisca = 0

def suavizar(valores):
    if len(valores) == 0:
        return []
    return pd.Series(valores).rolling(window=3, min_periods=1).mean().tolist()

def obter_texto_condicao(valor):
    if 0 <= valor <= 49.999:
        return "ALTA"
    elif 50 <= valor <= 100:
        return "TENDENCIA DE ALTA"
    elif -49.999 <= valor < 0:
        return "BAIXA"
    elif -100 <= valor <= -50:
        return "TENDENCIA DE BAIXA"
    elif valor > 100:
        return "TENDENCIA DE ALTA"
    elif valor < -100:
        return "TENDENCIA DE BAIXA"
    return "ALTA"

# Tenta carregar o xlwings apenas se estiver no Windows local
xlwings_disponivel = False
try:
    import xlwings as xw
    xlwings_disponivel = True
except ImportError:
    pass

wb = None
sheet = None

while True:
    try:
        # Leitura dos dados (Excel local ou modo demonstração na nuvem)
        if xlwings_disponivel:
            try:
                if wb is None:
                    wb = xw.Book('testepython.xlsx')
                    sheet = wb.sheets[0]
                valor_contrato = float(sheet.range('D151').value or 1.0)
                df = sheet.range('A3:J120').options(pd.DataFrame, header=False, index=False).value
                df.columns = ['Corretora', 'Qtd', 'PrecoMedio', 'QtdCpa', 'QtdVda', 'AgrCpa', 'AgrVda', 'SldAgr', 'PLBruto', 'SldPassivo']
            except:
                # Fallback se o excel não estiver aberto localmente
                raise Exception("Aguardando Excel local...")
        else:
            # Modo Nuvem / Demonstração (Simula valores de fluxo em tempo real)
            import random
            valor_contrato = 1.0
            data_sim = []
            corretoras_ex = grupo_1 + grupo_2 + grupo_3 + ['Bradesco', 'Santander', 'Ideal']
            for corr in corretoras_ex:
                data_sim.append([corr, 100, 125000, 50, 50, random.randint(-1000, 1000), random.randint(-1000, 1000), random.randint(-5000, 5000), 100.0, random.randint(-5000, 5000)])
            df = pd.DataFrame(data_sim, columns=['Corretora', 'Qtd', 'PrecoMedio', 'QtdCpa', 'QtdVda', 'AgrCpa', 'AgrVda', 'SldAgr', 'PLBruto', 'SldPassivo'])

        df['SldPassivo'] = pd.to_numeric(df['SldPassivo'], errors='coerce').fillna(0)
        df['SldAgr'] = pd.to_numeric(df['SldAgr'], errors='coerce').fillna(0)
        
        pat_g1 = '|'.join(grupo_1)
        pat_g2 = '|'.join(grupo_2)
        pat_g3 = '|'.join(grupo_3)
        
        df_g1 = df[df['Corretora'].astype(str).str.contains(pat_g1, case=False, na=False)]
        df_g2 = df[df['Corretora'].astype(str).str.contains(pat_g2, case=False, na=False)]
        df_g3 = df[df['Corretora'].astype(str).str.contains(pat_g3, case=False, na=False)]
        
        raw_shadow = (df_g1['SldPassivo'].sum() * (-1)) * valor_contrato
        raw_converge = df_g1['SldAgr'].sum() * valor_contrato
        raw_feeders = df_g2['SldAgr'].sum() * valor_contrato
        raw_equalizers = (df_g2['SldPassivo'].sum() * (-1)) * valor_contrato
        raw_disruption = df_g3['SldAgr'].sum() * valor_contrato
        raw_rebounders = df_g3['SldPassivo'].sum() * valor_contrato
        
        agora = datetime.now()
        bloco_atual = agora.minute // 30  
        
        bloco_5m_id = agora.hour * 12 + agora.minute // 5
        bloco_30m_id = agora.hour * 2 + agora.minute // 30
        bloco_60m_id = agora.hour
        bloco_dia_id = agora.date()

        if base_shadow_dia is None or bloco_dia_id != ultimo_dia:
            ultimo_dia = bloco_dia_id
            base_shadow_dia = raw_shadow
            base_converge_dia = raw_converge
            base_feeders_dia = raw_feeders
            base_equalizers_dia = raw_equalizers
            base_disruption_dia = raw_disruption
            base_rebounders_dia = raw_rebounders
            pico_shadow_dia, pico_converge_dia, pico_feeders_dia = 1.0, 1.0, 1.0
            pico_equalizers_dia, pico_disruption_dia, pico_rebounders_dia = 1.0, 1.0, 1.0

        if ultimo_5m is None or bloco_5m_id != ultimo_5m:
            ultimo_5m = bloco_5m_id
            base_shadow_5m = raw_shadow
            base_converge_5m = raw_converge
            base_feeders_5m = raw_feeders
            base_equalizers_5m = raw_equalizers
            base_disruption_5m = raw_disruption
            base_rebounders_5m = raw_rebounders
            pico_shadow_5m, pico_converge_5m, pico_feeders_5m = 1.0, 1.0, 1.0
            pico_equalizers_5m, pico_disruption_5m, pico_rebounders_5m = 1.0, 1.0, 1.0

        if ultimo_60m is None or bloco_60m_id != ultimo_60m:
            ultimo_60m = bloco_60m_id
            base_shadow_60m = raw_shadow
            base_converge_60m = raw_converge
            base_feeders_60m = raw_feeders
            base_equalizers_60m = raw_equalizers
            base_disruption_60m = raw_disruption
            base_rebounders_60m = raw_rebounders

        if ultimo_bloco_minuto is None or bloco_atual != ultimo_bloco_minuto:
            if ultimo_bloco_minuto is not None and len(bloco_atual_dados['tempos']) > 0:
                blocos_30m.append(bloco_atual_dados)
                if len(blocos_30m) > 12:
                    blocos_30m.pop(0)
            
            ultimo_bloco_minuto = bloco_atual
            base_shadow = raw_shadow
            base_converge = raw_converge
            base_feeders = raw_feeders
            base_equalizers = raw_equalizers
            base_disruption = raw_disruption
            base_rebounders = raw_rebounders
            
            bloco_atual_dados = {
                'tempos': [], 'shadow': [], 'converge': [],
                'feeders': [], 'equalizers': [], 'disruption': [], 'rebounders': []
            }
        
        h_tempos_acm.append(agora)
        h_shadow_acm.append(raw_shadow)
        h_converge_acm.append(raw_converge)
        h_feeders_acm.append(raw_feeders)
        h_equalizers_acm.append(raw_equalizers)
        h_disruption_acm.append(raw_disruption)
        h_rebounders_acm.append(raw_rebounders)
        
        if len(h_tempos_acm) > MAX_HISTORICO:
            h_tempos_acm.pop(0)
            h_shadow_acm.pop(0)
            h_converge_acm.pop(0)
            h_feeders_acm.pop(0)
            h_equalizers_acm.pop(0)
            h_disruption_acm.pop(0)
            h_rebounders_acm.pop(0)
        
        val_shadow_5m = raw_shadow - base_shadow_5m
        val_converge_5m = raw_converge - base_converge_5m
        val_feeders_5m = raw_feeders - base_feeders_5m
        val_equalizers_5m = raw_equalizers - base_equalizers_5m
        val_disruption_5m = raw_disruption - base_disruption_5m
        val_rebounders_5m = raw_rebounders - base_rebounders_5m

        val_shadow_30m = raw_shadow - base_shadow
        val_converge_30m = raw_converge - base_converge
        val_feeders_30m = raw_feeders - base_feeders
        val_equalizers_30m = raw_equalizers - base_equalizers
        val_disruption_30m = raw_disruption - base_disruption
        val_rebounders_30m = raw_rebounders - base_rebounders

        val_shadow_60m = raw_shadow - base_shadow_60m
        val_converge_60m = raw_converge - base_converge_60m
        val_feeders_60m = raw_feeders - base_feeders_60m
        val_equalizers_60m = raw_equalizers - base_equalizers_60m
        val_disruption_60m = raw_disruption - base_disruption_60m
        val_rebounders_60m = raw_rebounders - base_rebounders_60m

        val_shadow_dia = raw_shadow - base_shadow_dia
        val_converge_dia = raw_converge - base_converge_dia
        val_feeders_dia = raw_feeders - base_feeders_dia
        val_equalizers_dia = raw_equalizers - base_equalizers_dia
        val_disruption_dia = raw_disruption - base_disruption_dia
        val_rebounders_dia = raw_rebounders - base_rebounders_dia
        
        bloco_atual_dados['tempos'].append(agora)
        bloco_atual_dados['shadow'].append(val_shadow_30m)
        bloco_atual_dados['converge'].append(val_converge_30m)
        bloco_atual_dados['feeders'].append(val_feeders_30m)
        bloco_atual_dados['equalizers'].append(val_equalizers_30m)
        bloco_atual_dados['disruption'].append(val_disruption_30m)
        bloco_atual_dados['rebounders'].append(val_rebounders_30m)
        
        if len(h_tempos_acm) > 0:
            fig = make_subplots(
                rows=7, cols=2,
                shared_xaxes=False,
                vertical_spacing=0.03,
                column_widths=[0.88, 0.12],
                specs=[
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "xy"}],
                    [{"type": "xy"}, {"type": "indicator"}]
                ],
                subplot_titles=(
                    '<b>winfut (30m) :: disruption & rebounders</b>', '',
                    '<b>winfut (30m) :: feeders & equalizers</b>', '',
                    '<b>winfut (30m) :: shadow & converge</b>', '',
                    '<b>winfut (Acumulado) :: disruption & rebounders</b>', '',
                    '<b>winfut (Acumulado) :: feeders & equalizers</b>', '',
                    '<b>winfut (Acumulado) :: shadow & converge</b>', '',
                    '', ''
                )
            )
            
            def add_tr(row, x, y, color, name, unique_id):
                vis = estado_visibilidade.get(unique_id, True)
                fig.add_trace(go.Scatter(
                    x=x, y=suavizar(y), mode='lines',
                    line=dict(color=color, width=3),
                    name=name,
                    visible=vis,
                    showlegend=False,
                    hovertemplate=f'<b>{name}</b><br>Hora: %{{x|%H:%M:%S}}<br>Valor: %{{y:,.2f}}<extra></extra>'
                ), row=row, col=1)
                fig.add_hline(y=0, line=dict(color='white', width=1.5), row=row, col=1)

            todos_b = blocos_30m + [bloco_atual_dados]
            for b in todos_b:
                if b['tempos']:
                    add_tr(1, b['tempos'], b['disruption'], 'limegreen', 'Disruption (30m)', '30m_disruption')
                    add_tr(1, b['tempos'], b['rebounders'], 'white', 'Rebounders (30m)', '30m_rebounders')
                    add_tr(2, b['tempos'], b['feeders'], 'purple', 'Feeders (30m)', '30m_feeders')
                    add_tr(2, b['tempos'], b['equalizers'], 'deeppink', 'Equalizers (30m)', '30m_equalizers')
                    add_tr(3, b['tempos'], b['shadow'], 'deepskyblue', 'Shadow (30m)', '30m_shadow')
                    add_tr(3, b['tempos'], b['converge'], 'gold', 'Converge (30m)', '30m_converge')
            
            add_tr(4, h_tempos_acm, h_disruption_acm, 'limegreen', 'Disruption (Acm)', 'acm_disruption')
            add_tr(4, h_tempos_acm, h_rebounders_acm, 'white', 'Rebounders (Acm)', 'acm_rebounders')
            add_tr(5, h_tempos_acm, h_feeders_acm, 'purple', 'Feeders (Acm)', 'acm_feeders')
            add_tr(5, h_tempos_acm, h_equalizers_acm, 'deeppink', 'Equalizers (Acm)', 'acm_equalizers')
            add_tr(6, h_tempos_acm, h_shadow_acm, 'deepskyblue', 'Shadow (Acm)', 'acm_shadow')
            add_tr(6, h_tempos_acm, h_converge_acm, 'gold', 'Converge (Acm)', 'acm_converge')

            def calcula_pct_30m(lista_valores, valor_atual):
                if not lista_valores:
                    return 0.0
                pico_max = max([abs(v) for v in lista_valores])
                if pico_max == 0:
                    return 0.0
                pct = (valor_atual / pico_max) * 100
                return max(-100, min(100, pct))

            p_shadow_30m = calcula_pct_30m(bloco_atual_dados['shadow'], val_shadow_30m)
            p_converge_30m = calcula_pct_30m(bloco_atual_dados['converge'], val_converge_30m)
            p_feeders_30m = calcula_pct_30m(bloco_atual_dados['feeders'], val_feeders_30m)
            p_disruption_30m = calcula_pct_30m(bloco_atual_dados['disruption'], val_disruption_30m)
            p_equalizers_30m = calcula_pct_30m(bloco_atual_dados['equalizers'], val_equalizers_30m)
            p_rebounders_30m = calcula_pct_30m(bloco_atual_dados['rebounders'], val_rebounders_30m)

            def add_bar(row, y_val, name):
                bar_color = 'limegreen' if y_val >= 0 else 'crimson'
                fig.add_trace(go.Bar(
                    x=[y_val], y=[name], orientation='h',
                    marker=dict(color=bar_color),
                    width=0.45,
                    showlegend=False,
                    hovertemplate=f'<b>{name} (30m)</b>: %{{x:.1f}}%<extra></extra>'
                ), row=row, col=2)

            add_bar(1, p_shadow_30m, 'Shadow')
            add_bar(2, p_converge_30m, 'Converge')
            add_bar(3, p_feeders_30m, 'Feeders')
            add_bar(4, p_disruption_30m, 'Disruption')
            add_bar(5, p_equalizers_30m, 'Equalizers')
            add_bar(6, p_rebounders_30m, 'Rebounders')

            nomes_indicadores = ['Shadow', 'Converge', 'Feeders', 'Disruption', 'Equalizers', 'Rebounders']
            
            valores_5m = [val_shadow_5m, val_converge_5m, val_feeders_5m, val_disruption_5m, val_equalizers_5m, val_rebounders_5m]
            valores_30m = [val_shadow_30m, val_converge_30m, val_feeders_30m, val_disruption_30m, val_equalizers_30m, val_rebounders_30m]
            valores_60m = [val_shadow_60m, val_converge_60m, val_feeders_60m, val_disruption_60m, val_equalizers_60m, val_rebounders_60m]
            
            cor_textos_5m = ['#00FF7F' if v >= 0 else '#FF4D4D' for v in valores_5m]
            cor_textos_30m = ['#00FF7F' if v >= 0 else '#FF4D4D' for v in valores_30m]
            cor_textos_60m = ['#00FF7F' if v >= 0 else '#FF4D4D' for v in valores_60m]

            fmt_5m = [f"{v:,.2f}" for v in valores_5m]
            fmt_30m = [f"{v:,.2f}" for v in valores_30m]
            fmt_60m = [f"{v:,.2f}" for v in valores_60m]

            pos_5m = sum(1 for v in valores_5m if v >= 0)
            neg_5m = len(valores_5m) - pos_5m
            placar_5m_str = str(pos_5m) + 'x' + str(neg_5m)

            pos_30m = sum(1 for v in valores_30m if v >= 0)
            neg_30m = len(valores_30m) - pos_30m
            placar_30m_str = str(pos_30m) + 'x' + str(neg_30m)

            pos_60m = sum(1 for v in valores_60m if v >= 0)
            neg_60m = len(valores_60m) - pos_60m
            placar_60m_str = str(pos_60m) + 'x' + str(neg_60m)

            fig.add_trace(go.Table(
                columnwidth=[70, 110, 120, 120],
                header=dict(
                    values=["<b>Indicador</b>", f"<b>Valor 5m ({placar_5m_str})</b>", f"<b>Valor 30m ({placar_30m_str})</b>", f"<b>Valor 60m ({placar_60m_str})</b>"],
                    fill_color='rgba(30,30,30,0.9)',
                    font=dict(size=11, color='white'),
                    align=['left', 'right', 'right', 'right']
                ),
                cells=dict(
                    values=[nomes_indicadores, fmt_5m, fmt_30m, fmt_60m],
                    fill_color='rgba(15,15,15,0.9)',
                    font=dict(size=11, color=['white', cor_textos_5m, cor_textos_30m, cor_textos_60m]),
                    align=['left', 'right', 'right', 'right'],
                    height=22
                ),
                domain=dict(x=[0.0, 0.72], y=[0.0, 0.12])
            ))

            pico_shadow_5m = max(pico_shadow_5m, abs(val_shadow_5m))
            pico_converge_5m = max(pico_converge_5m, abs(val_converge_5m))
            pico_feeders_5m = max(pico_feeders_5m, abs(val_feeders_5m))
            pico_disruption_5m = max(pico_disruption_5m, abs(val_disruption_5m))
            pico_equalizers_5m = max(pico_equalizers_5m, abs(val_equalizers_5m))
            pico_rebounders_5m = max(pico_rebounders_5m, abs(val_rebounders_5m))

            p_shadow_5m = max(-100, min(100, (val_shadow_5m / pico_shadow_5m) * 100)) if pico_shadow_5m > 0 else 0
            p_converge_5m = max(-100, min(100, (val_converge_5m / pico_converge_5m) * 100)) if pico_converge_5m > 0 else 0
            p_feeders_5m = max(-100, min(100, (val_feeders_5m / pico_feeders_5m) * 100)) if pico_feeders_5m > 0 else 0
            p_disruption_5m = max(-100, min(100, (val_disruption_5m / pico_disruption_5m) * 100)) if pico_disruption_5m > 0 else 0
            p_equalizers_5m = max(-100, min(100, (val_equalizers_5m / pico_equalizers_5m) * 100)) if pico_equalizers_5m > 0 else 0
            p_rebounders_5m = max(-100, min(100, (val_rebounders_5m / pico_rebounders_5m) * 100)) if pico_rebounders_5m > 0 else 0

            percentuais_barras_5m = [p_shadow_5m, p_converge_5m, p_feeders_5m, p_disruption_5m, p_equalizers_5m, p_rebounders_5m]
            media_barras_5m = sum(percentuais_barras_5m) / len(percentuais_barras_5m) if percentuais_barras_5m else 0.0
            cor_ponteiro_5m = "#00FF7F" if media_barras_5m >= 0 else "#FF4D4D"

            percentuais_barras_30m = [p_shadow_30m, p_converge_30m, p_feeders_30m, p_disruption_30m, p_equalizers_30m, p_rebounders_30m]
            media_barras_30m = sum(percentuais_barras_30m) / len(percentuais_barras_30m) if percentuais_barras_30m else 0.0
            cor_ponteiro_30m = "#00FF7F" if media_barras_30m >= 0 else "#FF4D4D"

            pico_shadow_dia = max(pico_shadow_dia, abs(val_shadow_dia))
            pico_converge_dia = max(pico_converge_dia, abs(val_converge_dia))
            pico_feeders_dia = max(pico_feeders_dia, abs(val_feeders_dia))
            pico_disruption_dia = max(pico_disruption_dia, abs(val_disruption_dia))
            pico_equalizers_dia = max(pico_equalizers_dia, abs(val_equalizers_dia))
            pico_rebounders_dia = max(pico_rebounders_dia, abs(val_rebounders_dia))

            p_shadow_dia = max(-100, min(100, (val_shadow_dia / pico_shadow_dia) * 100)) if pico_shadow_dia > 0 else 0
            p_converge_dia = max(-100, min(100, (val_converge_dia / pico_converge_dia) * 100)) if pico_converge_dia > 0 else 0
            p_feeders_dia = max(-100, min(100, (val_feeders_dia / pico_feeders_dia) * 100)) if pico_feeders_dia > 0 else 0
            p_disruption_dia = max(-100, min(100, (val_disruption_dia / pico_disruption_dia) * 100)) if pico_disruption_dia > 0 else 0
            p_equalizers_dia = max(-100, min(100, (val_equalizers_dia / pico_equalizers_dia) * 100)) if pico_equalizers_dia > 0 else 0
            p_rebounders_dia = max(-100, min(100, (val_rebounders_dia / pico_rebounders_dia) * 100)) if pico_rebounders_dia > 0 else 0

            percentuais_barras_dia = [p_shadow_dia, p_converge_dia, p_feeders_dia, p_disruption_dia, p_equalizers_dia, p_rebounders_dia]
            media_barras_dia = sum(percentuais_barras_dia) / len(percentuais_barras_dia) if percentuais_barras_dia else 0.0
            cor_ponteiro_dia = "#00FF7F" if media_barras_dia >= 0 else "#FF4D4D"
            
            fig.add_trace(go.Indicator(
                mode="gauge+number",
                value=media_barras_5m,
                domain=dict(x=[0.74, 0.90], y=[0.086, 0.124]),
                number={"font": {"color": cor_ponteiro_5m, "size": 20}, "suffix": "%"},
                gauge={"axis": {"range": [-100, 100], "tickcolor": "white", "tickvals": [-100, 0, 100]}, "bar": {"color": cor_ponteiro_5m}, "bgcolor": "rgba(20,20,20,0.9)", "borderwidth": 1, "bordercolor": "gray", "steps": [{"range": [-100, 0], "color": "rgba(255, 77, 77, 0.15)"}, {"range": [0, 100], "color": "rgba(0, 255, 127, 0.15)"}]}
            ))

            fig.add_trace(go.Indicator(
                mode="gauge+number",
                value=media_barras_30m,
                domain=dict(x=[0.74, 0.90], y=[0.043, 0.081]),
                number={"font": {"color": cor_ponteiro_30m, "size": 20}, "suffix": "%"},
                gauge={"axis": {"range": [-100, 100], "tickcolor": "white", "tickvals": [-100, 0, 100]}, "bar": {"color": cor_ponteiro_30m}, "bgcolor": "rgba(20,20,20,0.9)", "borderwidth": 1, "bordercolor": "gray", "steps": [{"range": [-100, 0], "color": "rgba(255, 77, 77, 0.15)"}, {"range": [0, 100], "color": "rgba(0, 255, 127, 0.15)"}]}
            ))

            fig.add_trace(go.Indicator(
                mode="gauge+number",
                value=media_barras_dia,
                domain=dict(x=[0.74, 0.90], y=[0.0, 0.038]),
                number={"font": {"color": cor_ponteiro_dia, "size": 20}, "suffix": "%"},
                gauge={"axis": {"range": [-100, 100], "tickcolor": "white", "tickvals": [-100, 0, 100]}, "bar": {"color": cor_ponteiro_dia}, "bgcolor": "rgba(20,20,20,0.9)", "borderwidth": 1, "bordercolor": "gray", "steps": [{"range": [-100, 0], "color": "rgba(255, 77, 77, 0.15)"}, {"range": [0, 100], "color": "rgba(0, 255, 127, 0.15)"}]}
            ))

            relogio_str = agora.strftime('%H:%M:%S')
            texto_5m = obter_texto_condicao(media_barras_5m)
            texto_30m = obter_texto_condicao(media_barras_30m)
            texto_dia = obter_texto_condicao(media_barras_dia)

            if raw_disruption > 0 and raw_shadow < 0 and raw_converge < 0 and raw_converge < raw_shadow:
                frase_sinal = "VAI PRA CIMA, PROCURAR NÍVEIS DE FIBONACCI PARA COMPRA"
                cor_sinal = "#00FF7F"
            elif raw_disruption < 0 and raw_shadow > 0 and raw_converge > 0 and raw_converge > raw_shadow:
                frase_sinal = "VAI PRA BAIXO, PROCURAR NÍVEIS DE FIBONACCI PARA VENDA"
                cor_sinal = "#FF4D4D"
            elif raw_disruption > 0 and raw_shadow > 0 and raw_converge > 0:
                frase_sinal = "ROMPIMENTO DE TOPO"
                cor_sinal = "#00FF7F"
            elif raw_disruption < 0 and raw_shadow < 0 and raw_converge < 0:
                frase_sinal = "ROMPIMENTO DE FUNDO"
                cor_sinal = "#FF4D4D"
            else:
                frase_sinal = ""
                cor_sinal = "white"

            s_pos = raw_shadow >= 0
            c_pos = raw_converge >= 0
            d_pos = raw_disruption >= 0
            f_pos = raw_feeders >= 0

            if s_pos and c_pos and (not d_pos) and (not f_pos):
                frase_condicao_extra = "GRINGO VENDIDO, PRA CIMA SO PRA PEGAR LIQUIDEZ"
                cor_condicao_extra = "#FF4D4D"
            elif (not s_pos) and (not c_pos) and d_pos and f_pos:
                frase_condicao_extra = "GRINGO COMPRADO, PRA BAIXO SO PRA PEGAR LIQUIDEZ"
                cor_condicao_extra = "#00FF7F"
            elif s_pos == c_pos == d_pos == f_pos:
                frase_condicao_extra = "TENDÊNCIA FORTE"
                cor_condicao_extra = "yellow"
            elif (s_pos == c_pos) and (s_pos == d_pos):
                frase_condicao_extra = "SWINGTRADE DOMINANDO"
                cor_condicao_extra = "cyan"
            elif (s_pos == c_pos) and (s_pos == f_pos):
                frase_condicao_extra = "DAYTRADE DOMINANDO"
                cor_condicao_extra = "magenta"
            else:
                frase_condicao_extra = ""
                cor_condicao_extra = "white"

            contador_pisca += 1
            mostrar_texto = (contador_pisca % 2 == 0)

            def get_annotation_color(texto, cor_padrao):
                if "TENDENCIA" in texto:
                    return cor_padrao if mostrar_texto else "rgba(0,0,0,0)"
                return cor_padrao

            cor_anot_5m = get_annotation_color(texto_5m, cor_ponteiro_5m)
            cor_anot_30m = get_annotation_color(texto_30m, cor_ponteiro_30m)
            cor_anot_dia = get_annotation_color(texto_dia, cor_ponteiro_dia)

            fig.update_layout(
                template='plotly_dark',
                height=1700,
                autosize=True,
                showlegend=False,
                margin=dict(l=10, r=10, t=30, b=10),
                font=dict(size=14, color='white', family='Arial, sans-serif'),
                annotations=[
                    dict(text=f"<b>RELÓGIO: {relogio_str}</b>", xref="paper", yref="paper", x=1.0, y=1.02, showarrow=False, font=dict(size=18, color="cyan", family="Courier New"), bgcolor="rgba(0,0,0,0.8)", bordercolor="cyan", borderwidth=2, borderpad=4),
                    dict(text=f"<b>{frase_sinal}</b>" if frase_sinal else "", xref="paper", yref="paper", x=0.25, y=0.015, showarrow=False, xanchor="center", font=dict(size=14, color=cor_sinal, family="Arial Black")),
                    dict(text=f"<b>{frase_condicao_extra}</b>" if frase_condicao_extra else "", xref="paper", yref="paper", x=0.25, y=-0.005, showarrow=False, xanchor="center", font=dict(size=14, color=cor_condicao_extra, family="Arial Black")),
                    dict(text=f"<b>{texto_5m}</b>", xref="paper", yref="paper", x=0.92, y=0.105, showarrow=False, xanchor="left", font=dict(size=10, color=cor_anot_5m, family="Arial Black")),
                    dict(text=f"<b>{texto_30m}</b>", xref="paper", yref="paper", x=0.92, y=0.062, showarrow=False, xanchor="left", font=dict(size=10, color=cor_anot_30m, family="Arial Black")),
                    dict(text=f"<b>{texto_dia}</b>", xref="paper", yref="paper", x=0.92, y=0.019, showarrow=False, xanchor="left", font=dict(size=10, color=cor_anot_dia, family="Arial Black"))
                ]
            )

            for idx in range(1, 7):
                fig.update_yaxes(side="right", tickfont=dict(size=14, color='white', family='Arial Black'), showgrid=False, zeroline=True, zerolinecolor='white', zerolinewidth=1.5, col=1, row=idx)
                fig.update_xaxes(tickfont=dict(size=14, color='white', family='Arial Black'), showgrid=False, col=1, row=idx)
            
            nomes_barras = ['Shadow', 'Converge', 'Feeders', 'Disruption', 'Equalizers', 'Rebounders']
            cores_barras = ['deepskyblue', 'gold', 'purple', 'limegreen', 'deeppink', 'white']

            for r in range(1, 7):
                fig.update_xaxes(showgrid=False, zeroline=True, zerolinecolor='white', range=[-105, 105], tickvals=[-100, -50, 0, 50, 100], ticktext=['-100', '-50', '0', '50', '100'], autorange=False, col=2, row=r)
                fig.update_yaxes(showgrid=False, showticklabels=True, side="left", tickvals=[nomes_barras[r-1]], ticktext=[f"<b style='color:{cores_barras[r-1]}'>{nomes_barras[r-1]}</b>"], tickfont=dict(size=12, family='Arial Black'), col=2, row=r)

            for r in range(1, 7):
                fig.update_xaxes(showline=True, linewidth=1, linecolor='gray', mirror=True, row=r, col=1)
                fig.update_yaxes(showline=True, linewidth=1, linecolor='gray', mirror=True, row=r, col=1)

            dashboard_placeholder.plotly_chart(fig, use_container_width=True)
          
    except Exception as e:
        time.sleep(1)
       
    time.sleep(1)
    st.rerun()
