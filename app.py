import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from io import BytesIO

st.set_page_config(
    page_title="Superhar RH — Dashboard R&S",
    page_icon="🔵",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
  .main { background: #F4F5F7; }
  .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
  .header-bar {
    background: #185FA5; color: white; padding: 12px 24px;
    border-radius: 10px; margin-bottom: 18px;
    display: flex; justify-content: space-between; align-items: center;
  }
  .stTabs [data-baseweb="tab-list"] { gap: 4px; }
  .stTabs [data-baseweb="tab"] {
    background: white; border-radius: 8px 8px 0 0;
    padding: 8px 20px; font-weight: 500;
  }
  .stTabs [aria-selected="true"] { background: #185FA5 !important; color: white !important; }
  div[data-testid="stMetricValue"] { font-size: 2rem !important; color: #185FA5; }
  div[data-testid="stMetricDelta"] { font-size: 0.8rem; }
  .stDataFrame { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ── PALETA ───────────────────────────────────
AZUL    = "#185FA5"
AZUL2   = "#378ADD"
VERDE   = "#1D9E75"
VERDE2  = "#0F6E56"
AMBER   = "#BA7517"
ROXO    = "#534AB7"
CORAL   = "#D85A30"
ROSA    = "#D4537E"
CINZA   = "#B4B2A9"
PAL     = [AZUL, VERDE, AMBER, ROXO, CORAL, ROSA, CINZA, AZUL2]

# ── HEADER ───────────────────────────────────
st.markdown("""
<div class="header-bar">
  <div>
    <span style="font-size:18px;font-weight:700;letter-spacing:-0.3px">SUPERHAR</span>
    <span style="font-size:12px;opacity:0.7;margin-left:8px">Recursos Humanos</span>
  </div>
  <div style="font-size:12px;opacity:0.85">Dashboard de Recrutamento &amp; Seleção</div>
</div>
""", unsafe_allow_html=True)

# ── UPLOAD ───────────────────────────────────
with st.expander("📂  Carregar base de dados (.xlsx)", expanded=True):
    uploaded = st.file_uploader(
        "Selecione o arquivo Excel (aba: Base Relatorio Executivo)",
        type=["xlsx","xls"], label_visibility="collapsed"
    )
    st.caption("Os dados ficam apenas na sua sessão e são apagados ao fechar o navegador.")

# ── LEITURA ──────────────────────────────────
def carregar(file_bytes):
    buf = BytesIO(file_bytes)
    # Detecta aba automaticamente
    xl = pd.ExcelFile(buf)
    abas = xl.sheet_names
    aba_alvo = next((a for a in abas if "base" in a.lower()), abas[0])
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=aba_alvo, header=1)
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]

    # ── RENOMEAR: mapeamento dos nomes ORIGINAIS do Excel → nomes padronizados
    # (os nomes originais têm \n, tabs e espaços extras)
    rename_map = {}
    for c in df.columns:
        cs = str(c)  # nome original com todos os caracteres especiais
        cu = cs.upper().replace('\n',' ').replace('\t',' ')
        cu = ' '.join(cu.split())  # colapsa espaços múltiplos

        if cu == '(PCD)':                                         rename_map[c] = 'PCD'
        elif cu == 'DATA ESCOLHA FINALISTA':                      rename_map[c] = 'DATA ESCOLHA FINALISTA'
        elif cu == 'DATA LONG LIST (REALIZADO)':                  rename_map[c] = 'DATA LONG LIST REALIZADO'
        elif cu == 'DATA LONG LIST (PREVISTA)':                   rename_map[c] = 'DATA LONG LIST PREVISTA'
        elif cu == 'DATA SHORT LIST (PREVISTA)':                  rename_map[c] = 'DATA SHORT LIST PREVISTA'
        elif cu == 'DATA SHORT LIST (REALIZADA)':                 rename_map[c] = 'DATA SHORT LIST REALIZADA'
        elif cu == 'NOVA DATA LONG LIST (PREVISTA)':              rename_map[c] = 'NOVA DATA LONG LIST PREVISTA'
        elif cu == 'NOVA DATA LONG LIST (REALIZADO)':             rename_map[c] = 'NOVA DATA LONG LIST REALIZADO'
        elif cu == 'NOVA DATA SHORT LIST (PREVISTA)':             rename_map[c] = 'NOVA DATA SHORT LIST PREVISTA'
        elif cu == 'NOVA DATA SHORT LIST (REALIZADA)':            rename_map[c] = 'NOVA DATA SHORT LIST REALIZADA'
        elif cu == 'TOTAL DE INSCRITOS':                          rename_map[c] = 'TOTAL INSCRITOS'
        elif cu == 'TOTAL DE CANDIDATOS ABORDADOS':               rename_map[c] = 'TOTAL ABORDADOS'
        elif cu == 'TOTAL DE CANDIDATOS APRESENTADAS EM LONG LIST': rename_map[c] = 'TOTAL LONG LIST'
        elif cu == 'TOTAL DE CANDIDATOS APROVADOS EM LONG':       rename_map[c] = 'TOTAL APROVADOS LONG'
        elif cu == 'TOTAL CANDIDATOS EM SELECAO SUPERHAR' or 'SELEÇÃO SUPERHAR' in cu: rename_map[c] = 'TOTAL SELECAO SUPERHAR'
        elif cu == 'TOTAL DE CANDIDATOS SHORT LIST':              rename_map[c] = 'TOTAL SHORT LIST'
        elif cu == 'TOTAL CANDIDATOS EM ENTREVISTA C/GESTOR':     rename_map[c] = 'TOTAL ENTREVISTA GESTOR'
        elif cu == 'TOTAL DE CANDIDATOS REPROVADOS EM MEDICINA':  rename_map[c] = 'TOTAL REPROVADOS MEDICINA'
        elif 'SALARIO ATUAL' in cu or 'SALÁRIO ATUAL' in cu:     rename_map[c] = 'SALARIO ATUAL'
        elif 'PRETENSAO SALARIAL' in cu or 'PRETENSÃO SALARIAL' in cu: rename_map[c] = 'PRETENSAO SALARIAL'
        elif 'ORIGEM' in cu and 'RECRUTAMENTO' in cu:            rename_map[c] = 'ORIGEM RECRUTAMENTO'

    df = df.rename(columns=rename_map)
    # Limpar espaços nas bordas dos nomes restantes
    df.columns = [str(c).strip() for c in df.columns]
    # Remover colunas duplicadas: manter a primeira ocorrência
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    # ── DATAS
    for c in ['DATA RECEBIMENTO','DATA DE ABERTURA (Inicio Cronograma)',
              'DATA LONG LIST REALIZADO','DATA SHORT LIST PREVISTA',
              'DATA SHORT LIST REALIZADA','DATA ESCOLHA FINALISTA','DATA ADMISSÃO',
              'DATA ALINHAMENTO']:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors='coerce')

    # ── NUMÉRICOS — funil
    for c in ['TOTAL INSCRITOS','TOTAL ABORDADOS','TOTAL LONG LIST',
              'TOTAL APROVADOS LONG','TOTAL SELECAO SUPERHAR',
              'TOTAL SHORT LIST','TOTAL ENTREVISTA GESTOR','TOTAL REPROVADOS MEDICINA']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # ── NUMÉRICOS — salário
    for c in ['SALÁRIO PREVISTO (Minimo)','SALÁRIO PREVISTO (Maximo)',
              'SALARIO ADMISSÃO','SALARIO ATUAL','PRETENSAO SALARIAL']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # ── DERIVADAS (com .get para não quebrar se coluna ausente)
    col_fin = 'DATA ESCOLHA FINALISTA'
    col_ab  = 'DATA DE ABERTURA (Inicio Cronograma)'
    if col_fin in df.columns and col_ab in df.columns:
        df['time_to_fill'] = (df[col_fin] - df[col_ab]).dt.days
    else:
        df['time_to_fill'] = pd.NA

    col_sl_r = 'DATA SHORT LIST REALIZADA'
    col_sl_p = 'DATA SHORT LIST PREVISTA'
    if col_sl_r in df.columns and col_sl_p in df.columns:
        df['sla_ok'] = df[col_sl_r] <= df[col_sl_p]
    else:
        df['sla_ok'] = False

    df['admitido'] = df['DATA ADMISSÃO'].notna() if 'DATA ADMISSÃO' in df.columns else False

    if col_ab in df.columns:
        df['mes_abertura'] = df[col_ab].dt.to_period('M').astype(str)
    else:
        df['mes_abertura'] = 'Desconhecido'

    # ── GÊNERO normalizado
    if 'GÊNERO' in df.columns:
        df['genero_norm'] = df['GÊNERO'].astype(str).str.strip().replace({
            'Homem cisgênero':'Masculino','Homem Cisgênero':'Masculino',
            'Mulher Cisgênero':'Feminino','Mulher cisgênero':'Feminino',
            'nan':'Não informado','':'Não informado'
        })
    else:
        df['genero_norm'] = 'Não informado'

    # ── PCD normalizado
    pcd_col = 'PCD' if 'PCD' in df.columns else None
    if pcd_col:
        df['pcd_norm'] = df[pcd_col].astype(str).str.strip().str.lower().map(
            {'sim':'Sim','não':'Não','nao':'Não','nan':'Não informado'}
        ).fillna('Não informado')
    else:
        df['pcd_norm'] = 'Não informado'

    # ── NÍVEL normalizado
    if 'NÍVEL' in df.columns:
        df['NÍVEL'] = df['NÍVEL'].astype(str).str.strip()

    return df

if uploaded:
    with st.spinner("Carregando base de dados…"):
        df_raw = carregar(uploaded.getvalue())
    st.success(f"✅  {len(df_raw)} registros | {df_raw['COD.VAGA SUPERHAR'].nunique()} vagas | carregados com sucesso")
else:
    st.info("⬆️  Carregue o arquivo Excel para visualizar os dados reais.")
    st.stop()

# ── ABAS ─────────────────────────────────────
aba1, aba2, aba3 = st.tabs(["📊  Gerencial", "👩‍💼  Consultoras", "🤝  Visão Cliente"])

# ══════════════════════════════════════════════
# ABA 1 — GERENCIAL
# ══════════════════════════════════════════════
with aba1:

    # Filtros
    cf1, cf2, cf3, cf4 = st.columns([1,1,1,2])
    with cf1:
        empresas = ["Todas"] + sorted(df_raw['EMPRESA'].dropna().unique().tolist())
        f_emp = st.selectbox("Empresa / Cliente", empresas)
    with cf2:
        niveis = ["Todos"] + sorted(df_raw['NÍVEL'].dropna().unique().tolist())
        f_niv = st.selectbox("Nível", niveis)
    with cf3:
        meses = ["Todos"] + sorted(df_raw['mes_abertura'].dropna().unique().tolist())
        f_mes = st.selectbox("Período (abertura)", meses)
    with cf4:
        consultoras_lista = ["Todas"] + sorted(df_raw['CONSULTOR (A) SELEÇÃO'].dropna().unique().tolist())
        f_cons = st.selectbox("Consultora Seleção", consultoras_lista)

    dg = df_raw.copy()
    if f_emp  != "Todas": dg = dg[dg['EMPRESA']==f_emp]
    if f_niv  != "Todos": dg = dg[dg['NÍVEL']==f_niv]
    if f_mes  != "Todos": dg = dg[dg['mes_abertura']==f_mes]
    if f_cons != "Todas": dg = dg[dg['CONSULTOR (A) SELEÇÃO']==f_cons]

    # Vagas únicas (para KPIs de processo)
    dg_vaga = dg.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')

    total_vagas   = dg_vaga['COD.VAGA SUPERHAR'].nunique()
    concluidas    = dg_vaga['Status Samarco'].isin(['Concluído','Concluído C/Interno','Admissão']).sum()
    em_andamento  = total_vagas - concluidas
    tx_concl      = round(concluidas / total_vagas * 100) if total_vagas > 0 else 0
    ttf_med       = int(dg_vaga['time_to_fill'].median()) if dg_vaga['time_to_fill'].notna().any() else 0
    admitidos     = int(dg['admitido'].sum())
    pcd_sim       = int((dg['pcd_norm']=='Sim').sum())
    sla_pct       = round(dg_vaga['sla_ok'].mean()*100) if dg_vaga['sla_ok'].notna().any() else 0

    # KPIs
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Total de Vagas",     total_vagas)
    k2.metric("Concluídas",         concluidas)
    k3.metric("Em Andamento",       em_andamento)
    k4.metric("Taxa de Conclusão",  f"{tx_concl}%")
    k5.metric("Mediana Time to Fill", f"{ttf_med} dias")
    k6.metric("Admitidos",          admitidos)

    st.markdown("---")

    # LINHA 2 — Clientes | Funil | Origem
    c1,c2,c3 = st.columns(3)

    with c1:
        st.markdown("**Vagas por empresa / cliente**")
        vc = dg_vaga['EMPRESA'].value_counts().reset_index()
        vc.columns = ['Empresa','Vagas']
        fig = px.pie(vc, values='Vagas', names='Empresa', hole=0.4,
                     color_discrete_sequence=PAL)
        fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=230,
                          legend=dict(font_size=10, orientation='v'))
        fig.update_traces(textinfo='percent+value', textfont_size=10)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("**Funil de recrutamento — totais**")
        ins  = int(dg_vaga['TOTAL INSCRITOS'].sum())
        aord = int(dg_vaga['TOTAL ABORDADOS'].sum())
        ll   = int(dg_vaga['TOTAL LONG LIST'].sum())
        sl   = int(dg_vaga['TOTAL SHORT LIST'].sum())
        eg   = int(dg_vaga['TOTAL ENTREVISTA GESTOR'].sum())
        adm  = admitidos
        fig_f = go.Figure(go.Funnel(
            y=['Inscritos','Abordados','Long List','Short List','Entrev. Gestor','Admitidos'],
            x=[ins, aord, ll, sl, eg, adm],
            textinfo='value+percent initial',
            marker=dict(color=[AZUL,'#2B74BE',AZUL2,VERDE,VERDE2,'#085041']),
            connector=dict(line=dict(color='#E2E4E9',width=1))
        ))
        fig_f.update_layout(margin=dict(t=0,b=0,l=0,r=80), height=230, font=dict(size=10))
        st.plotly_chart(fig_f, use_container_width=True)

    with c3:
        st.markdown("**Origem do recrutamento (candidatos)**")
        ori = dg['ORIGEM RECRUTAMENTO'].value_counts().reset_index()
        ori.columns = ['Canal','Qtd']
        fig = px.bar(ori, x='Qtd', y='Canal', orientation='h',
                     color_discrete_sequence=[AZUL])
        fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=230,
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # LINHA 3 — Mensal | Diversidade | Mini-indicadores
    c4,c5,c6 = st.columns([1.4,1,0.6])

    with c4:
        st.markdown("**Vagas abertas por mês**")
        dg_vaga2 = df_raw.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first').copy()
        dg_vaga2['mes'] = dg_vaga2['DATA DE ABERTURA (Inicio Cronograma)'].dt.to_period('M').astype(str)
        mens = dg_vaga2.groupby('mes').size().reset_index(name='Vagas').sort_values('mes')
        fig = px.bar(mens, x='mes', y='Vagas', color_discrete_sequence=[AZUL])
        fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        fig.update_traces(marker_line_width=0)
        st.plotly_chart(fig, use_container_width=True)

    with c5:
        st.markdown("**Diversidade dos candidatos**")
        cg1, cg2 = st.columns(2)
        with cg1:
            st.caption("Gênero")
            gen = dg[dg['genero_norm']!='Não informado']['genero_norm'].value_counts().reset_index()
            gen.columns = ['Gênero','Qtd']
            if len(gen) > 0:
                fig_g = px.pie(gen, values='Qtd', names='Gênero', hole=0.4,
                               color_discrete_sequence=[ROSA, AZUL, ROXO])
                fig_g.update_layout(margin=dict(t=0,b=20,l=0,r=0), height=160, showlegend=False)
                fig_g.update_traces(textinfo='percent', textfont_size=9)
                st.plotly_chart(fig_g, use_container_width=True)
            else:
                st.caption("Sem dados de gênero")
        with cg2:
            st.caption("PcD")
            pcd_df = dg[dg['pcd_norm']!='Não informado']['pcd_norm'].value_counts().reset_index()
            pcd_df.columns = ['PcD','Qtd']
            if len(pcd_df) > 0:
                fig_p = px.pie(pcd_df, values='Qtd', names='PcD', hole=0.4,
                               color_discrete_sequence=[AMBER, CINZA])
                fig_p.update_layout(margin=dict(t=0,b=20,l=0,r=0), height=160, showlegend=False)
                fig_p.update_traces(textinfo='percent', textfont_size=9)
                st.plotly_chart(fig_p, use_container_width=True)
            else:
                st.caption("Sem dados de PcD")

    with c6:
        st.markdown("**Indicadores**")
        ttf_max = int(dg_vaga['time_to_fill'].max()) if dg_vaga['time_to_fill'].notna().any() else 0
        med_rep  = int(dg_vaga['TOTAL REPROVADOS MEDICINA'].sum())
        for label, val, cor, barpct in [
            ("SLA cumprido",        f"{sla_pct}%",    VERDE,  sla_pct),
            ("Candidatos PcD",      f"{pcd_sim}",      AZUL,   min(pcd_sim*3,100)),
            ("Maior TTF (dias)",    f"{ttf_max}",      AMBER,  min(ttf_max//2,100)),
            ("Reprov. medicina",    f"{med_rep}",      CORAL,  min(med_rep,100)),
        ]:
            st.markdown(f"""
            <div style="background:white;border-radius:8px;padding:10px 12px;
                        border:1px solid #E2E4E9;margin-bottom:8px;text-align:center">
              <div style="font-size:10px;color:#6B7280;font-weight:600;
                          text-transform:uppercase;letter-spacing:0.04em">{label}</div>
              <div style="font-size:22px;font-weight:700;color:{cor};margin:2px 0">{val}</div>
              <div style="height:4px;background:#E2E4E9;border-radius:3px;overflow:hidden">
                <div style="width:{barpct}%;height:100%;background:{cor};border-radius:3px"></div>
              </div>
            </div>""", unsafe_allow_html=True)

    # Tabela síntese por empresa
    st.markdown("---")
    st.markdown("**Síntese por empresa**")
    dg_vg = dg.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
    sint = dg_vg.groupby('EMPRESA').agg(
        Vagas         = ('COD.VAGA SUPERHAR','count'),
        Concluídas    = ('Status Samarco', lambda x: x.isin(['Concluído','Concluído C/Interno','Admissão']).sum()),
        Inscritos     = ('TOTAL INSCRITOS','sum'),
        Short_List    = ('TOTAL SHORT LIST','sum'),
        TTF_Mediana   = ('time_to_fill','median'),
    ).reset_index()
    sint['Conclusão %'] = (sint['Concluídas']/sint['Vagas']*100).round(0).astype(int)
    sint['TTF_Mediana'] = sint['TTF_Mediana'].fillna(0).astype(int)
    sint = sint.rename(columns={'EMPRESA':'Empresa','TTF_Mediana':'Mediana TTF (dias)'})
    st.dataframe(sint[['Empresa','Vagas','Concluídas','Conclusão %','Inscritos','Short_List','Mediana TTF (dias)']],
                 use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# ABA 2 — CONSULTORAS
# ══════════════════════════════════════════════
with aba2:

    cf1, cf2, _ = st.columns([1,1,2])
    with cf1:
        emp2 = ["Todas"] + sorted(df_raw['EMPRESA'].dropna().unique().tolist())
        f_emp2 = st.selectbox("Empresa", emp2, key="cons_emp")
    with cf2:
        niv2 = ["Todos"] + sorted(df_raw['NÍVEL'].dropna().unique().tolist())
        f_niv2 = st.selectbox("Nível", niv2, key="cons_niv")

    dc = df_raw.copy()
    if f_emp2 != "Todas": dc = dc[dc['EMPRESA']==f_emp2]
    if f_niv2 != "Todos": dc = dc[dc['NÍVEL']==f_niv2]

    dc_vaga = dc.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')

    grp = dc_vaga.groupby('CONSULTOR (A) SELEÇÃO').agg(
        Vagas       = ('COD.VAGA SUPERHAR','count'),
        Concluídas  = ('Status Samarco', lambda x: x.isin(['Concluído','Concluído C/Interno','Admissão']).sum()),
        TTF_med     = ('time_to_fill','median'),
        SLA_ok      = ('sla_ok','mean'),
        Inscritos   = ('TOTAL INSCRITOS','sum'),
        Short_List  = ('TOTAL SHORT LIST','sum'),
        Entrev      = ('TOTAL ENTREVISTA GESTOR','sum'),
    ).reset_index()
    grp['Conclusão %'] = (grp['Concluídas']/grp['Vagas']*100).round(0).fillna(0).astype(int)
    grp['TTF_med']     = grp['TTF_med'].fillna(0).round(0).astype(int)
    grp['SLA %']       = (grp['SLA_ok']*100).round(0).fillna(0).astype(int)
    grp['Conv %']      = np.where(grp['Inscritos']>0,
                                  (grp['Short_List']/grp['Inscritos']*100).round(1), 0)

    # KPIs destaques
    if len(grp) > 0:
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Maior volume",
                  grp.loc[grp['Vagas'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0],
                  f"{grp['Vagas'].max()} vagas")
        k2.metric("Mais concluídas",
                  grp.loc[grp['Concluídas'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0],
                  f"{grp['Concluídas'].max()} concluídas")
        ttf_min_idx = grp[grp['TTF_med']>0]['TTF_med'].idxmin() if (grp['TTF_med']>0).any() else grp['TTF_med'].idxmin()
        k3.metric("Menor TTF",
                  grp.loc[ttf_min_idx,'CONSULTOR (A) SELEÇÃO'].split()[0],
                  f"{grp.loc[ttf_min_idx,'TTF_med']} dias")
        k4.metric("Melhor SLA",
                  grp.loc[grp['SLA %'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0],
                  f"SLA {grp['SLA %'].max()}%")

    st.markdown("---")

    # Tabela comparativa
    st.markdown("**Desempenho por consultora**")
    tabela = grp.rename(columns={
        'CONSULTOR (A) SELEÇÃO':'Consultora',
        'Vagas':'Vagas','Concluídas':'Concluídas','Conclusão %':'Conclusão %',
        'TTF_med':'TTF mediana (dias)','SLA %':'SLA %','Conv %':'Conv. Inscrito→SL %',
        'Inscritos':'Inscritos','Short_List':'Short List','Entrev':'Entrevistas Gestor'
    })
    st.dataframe(
        tabela[['Consultora','Vagas','Concluídas','Conclusão %','TTF mediana (dias)',
                'SLA %','Inscritos','Short List','Entrevistas Gestor']],
        use_container_width=True, hide_index=True
    )

    st.markdown("---")

    c1,c2,c3 = st.columns(3)
    nomes_c = grp['CONSULTOR (A) SELEÇÃO'].apply(lambda x: ' '.join(x.split()[:2]))

    with c1:
        st.markdown("**Vagas por consultora**")
        fig = px.bar(x=grp['Vagas'], y=nomes_c, orientation='h',
                     color_discrete_sequence=[AZUL])
        fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("**Vagas concluídas por consultora**")
        fig = px.bar(x=grp['Concluídas'], y=nomes_c, orientation='h',
                     color_discrete_sequence=[VERDE])
        fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)

    with c3:
        st.markdown("**Mediana TTF por consultora (dias)**")
        grp_ttf = grp[grp['TTF_med']>0]
        nomes_ttf = grp_ttf['CONSULTOR (A) SELEÇÃO'].apply(lambda x: ' '.join(x.split()[:2]))
        fig = px.bar(x=grp_ttf['TTF_med'], y=nomes_ttf, orientation='h',
                     color_discrete_sequence=[AMBER])
        fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                          xaxis_title='dias', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)

    # Funil por consultora
    st.markdown("---")
    st.markdown("**Funil detalhado por consultora**")
    funil_cons = dc_vaga.groupby('CONSULTOR (A) SELEÇÃO').agg(
        Inscritos   = ('TOTAL INSCRITOS','sum'),
        Abordados   = ('TOTAL ABORDADOS','sum'),
        Long_List   = ('TOTAL LONG LIST','sum'),
        Short_List  = ('TOTAL SHORT LIST','sum'),
        Entrevistas = ('TOTAL ENTREVISTA GESTOR','sum'),
    ).reset_index()
    funil_cons.columns = ['Consultora','Inscritos','Abordados','Long List','Short List','Entrevistas Gestor']
    st.dataframe(funil_cons, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# ABA 3 — VISÃO CLIENTE
# ══════════════════════════════════════════════
with aba3:

    # Seletor de vaga
    vagas_lista = df_raw.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
    vagas_lista = vagas_lista[vagas_lista['EMPRESA'].notna() & vagas_lista['CARGO'].notna()]
    vagas_lista['label'] = vagas_lista['CARGO'].str.strip() + " — " + vagas_lista['EMPRESA'].str.strip()

    sel_label = st.selectbox("Selecione o processo:", vagas_lista['label'].tolist())
    vaga_row = vagas_lista[vagas_lista['label']==sel_label].iloc[0]

    # Candidatos desta vaga
    cand_vaga = df_raw[df_raw['COD.VAGA SUPERHAR']==vaga_row['COD.VAGA SUPERHAR']]

    st.markdown("---")

    # KPIs da vaga
    status_v   = str(vaga_row.get('Status Samarco','—'))
    cons_v     = str(vaga_row.get('CONSULTOR (A) SELEÇÃO','—'))
    inscritos  = int(vaga_row['TOTAL INSCRITOS']) if pd.notna(vaga_row.get('TOTAL INSCRITOS')) else 0
    sl_v       = int(vaga_row['TOTAL SHORT LIST']) if pd.notna(vaga_row.get('TOTAL SHORT LIST')) else 0
    eg_v       = int(vaga_row['TOTAL ENTREVISTA GESTOR']) if pd.notna(vaga_row.get('TOTAL ENTREVISTA GESTOR')) else 0
    adm_v      = int(cand_vaga['admitido'].sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Status", status_v)
    k2.metric("Inscritos", inscritos)
    k3.metric("Short List", sl_v)
    k4.metric("Entrev. Gestor", eg_v)
    k5.metric("Admitido(s)", adm_v)

    st.markdown("---")

    # Timeline + Funil
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Linha do tempo do processo**")
        etapas_tl = [
            ("Alinhamento",   vaga_row.get('DATA ALINHAMENTO')),
            ("Abertura",      vaga_row.get('DATA DE ABERTURA (Inicio Cronograma)')),
            ("Long List",     vaga_row.get('DATA LONG LIST REALIZADO')),
            ("Short List",    vaga_row.get('DATA SHORT LIST REALIZADA')),
            ("Finalista",     vaga_row.get('DATA ESCOLHA FINALISTA')),
            ("Admissão",      cand_vaga['DATA ADMISSÃO'].dropna().min() if cand_vaga['DATA ADMISSÃO'].notna().any() else None),
        ]
        concl_tl = sum(1 for _,d in etapas_tl if pd.notna(d))
        html_tl = '<div style="display:flex;align-items:flex-start;overflow-x:auto;padding:10px 0">'
        for i,(nome,data) in enumerate(etapas_tl):
            done    = pd.notna(data)
            current = i == concl_tl and not done
            dot_bg  = VERDE if done else (AZUL if current else "#E2E4E9")
            dot_txt = "✓" if done else ("●" if current else "○")
            nc      = VERDE2 if done else (AZUL if current else "#6B7280")
            data_s  = pd.Timestamp(data).strftime('%d/%m/%y') if done else "—"
            line    = f'<div style="flex:1;height:2px;background:{"#1D9E75" if i<concl_tl-1 else "#E2E4E9"};margin-top:13px;min-width:18px"></div>' if i<len(etapas_tl)-1 else ""
            html_tl += f'''
            <div style="display:flex;flex-direction:column;align-items:center;min-width:75px">
              <div style="width:26px;height:26px;border-radius:50%;background:{dot_bg};
                          display:flex;align-items:center;justify-content:center;
                          font-size:12px;color:white;font-weight:700">{dot_txt}</div>
              <div style="font-size:9px;font-weight:600;color:{nc};text-align:center;margin-top:5px;line-height:1.3">{nome}</div>
              <div style="font-size:9px;color:#6B7280;margin-top:1px">{data_s}</div>
            </div>{line}'''
        html_tl += '</div>'
        st.markdown(html_tl, unsafe_allow_html=True)

    with c2:
        st.markdown("**Funil desta vaga**")
        ll_v = int(vaga_row['TOTAL LONG LIST']) if pd.notna(vaga_row.get('TOTAL LONG LIST')) else 0
        fig_f = go.Figure(go.Funnel(
            y=['Inscritos','Long List','Short List','Entrev. Gestor','Admitidos'],
            x=[inscritos, ll_v, sl_v, eg_v, adm_v],
            textinfo='value+percent initial',
            marker=dict(color=[AZUL, AZUL2, VERDE, VERDE2,'#085041']),
        ))
        fig_f.update_layout(margin=dict(t=0,b=0,l=0,r=80), height=200, font=dict(size=11))
        st.plotly_chart(fig_f, use_container_width=True)

    st.markdown("---")

    # Benchmarking salarial + Diversidade
    c3,c4,c5 = st.columns(3)

    with c3:
        st.markdown("**Benchmarking salarial**")
        sal_prev_min = vaga_row.get('SALÁRIO PREVISTO (Minimo)')
        sal_prev_max = vaga_row.get('SALÁRIO PREVISTO (Maximo)')
        sal_adm      = cand_vaga['SALARIO ADMISSÃO'].dropna().mean() if cand_vaga['SALARIO ADMISSÃO'].notna().any() else None
        sal_pret     = cand_vaga.get('PRETENSAO SALARIAL', cand_vaga.get('PRETENSAO SALARIAL', pd.Series(dtype=float))).dropna().mean() if cand_vaga.get('PRETENSAO SALARIAL', cand_vaga.get('PRETENSAO SALARIAL', pd.Series(dtype=float))).notna().any() else None

        labels_sal, vals_sal, cors_sal = [], [], []
        try:
            if pd.notna(sal_prev_min) and float(sal_prev_min) > 0:
                labels_sal.append('Previsto (min)')
                vals_sal.append(float(sal_prev_min))
                cors_sal.append('#B5D4F4')
            if pd.notna(sal_prev_max) and float(sal_prev_max) > 0:
                labels_sal.append('Previsto (max)')
                vals_sal.append(float(sal_prev_max))
                cors_sal.append('#378ADD')
            if sal_pret and float(sal_pret) > 0:
                labels_sal.append('Pretensao media')
                vals_sal.append(float(sal_pret))
                cors_sal.append('#BA7517')
            if sal_adm and float(sal_adm) > 0:
                labels_sal.append('Admissao media')
                vals_sal.append(float(sal_adm))
                cors_sal.append('#1D9E75')
        except Exception:
            pass

        if vals_sal:
            txt_sal = ['R$ {:,.0f}'.format(v).replace(',','.') for v in vals_sal]
            fig_s = go.Figure(go.Bar(
                x=labels_sal, y=vals_sal,
                marker=dict(color=cors_sal),
                text=txt_sal, textposition='outside'
            ))
            fig_s.update_layout(margin=dict(t=20,b=0,l=0,r=0), height=200,
                                 yaxis=dict(showticklabels=False), font=dict(size=10))
            st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.caption("Sem dados salariais para esta vaga.")

    with c4:
        st.markdown("**Gênero dos candidatos**")
        gen_v = cand_vaga[cand_vaga['genero_norm']!='Não informado']['genero_norm'].value_counts().reset_index()
        gen_v.columns = ['Gênero','Qtd']
        if len(gen_v) > 0:
            fig_g = px.pie(gen_v, values='Qtd', names='Gênero', hole=0.4,
                           color_discrete_sequence=[ROSA, AZUL, ROXO])
            fig_g.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                                 legend=dict(font_size=10, orientation='h', y=-0.15))
            fig_g.update_traces(textinfo='percent+label', textfont_size=10)
            st.plotly_chart(fig_g, use_container_width=True)
        else:
            st.caption("Dados de gênero não informados nesta vaga.")

    with c5:
        st.markdown("**PcD dos candidatos**")
        pcd_v = cand_vaga[cand_vaga['pcd_norm']!='Não informado']['pcd_norm'].value_counts().reset_index()
        pcd_v.columns = ['PcD','Qtd']
        if len(pcd_v) > 0:
            fig_p = px.pie(pcd_v, values='Qtd', names='PcD', hole=0.4,
                           color_discrete_sequence=[AMBER, CINZA])
            fig_p.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                                 legend=dict(font_size=10, orientation='h', y=-0.15))
            fig_p.update_traces(textinfo='percent+label', textfont_size=10)
            st.plotly_chart(fig_p, use_container_width=True)
        else:
            st.caption("Dados de PcD não informados nesta vaga.")

    st.markdown("---")

    # Candidatos desta vaga
    st.markdown("**Candidatos deste processo**")
    cols_cand = ['NOME CANDIDATO','GÊNERO','RAÇA','PCD','ORIGEM RECRUTAMENTO',
                 'ULTIMA EMPRESA / ATUAL','SALÁRIO ATUAL','PRETENSAO SALARIAL',
                 'SALARIO ADMISSÃO','DATA ADMISSÃO','INTERNO OU EXTERNO']
    cols_ex = [c for c in cols_cand if c in cand_vaga.columns]
    df_cand_show = cand_vaga[cols_ex].copy()
    # Formata datas
    if 'DATA ADMISSÃO' in df_cand_show.columns:
        df_cand_show['DATA ADMISSÃO'] = pd.to_datetime(df_cand_show['DATA ADMISSÃO'], errors='coerce').dt.strftime('%d/%m/%Y')
    st.dataframe(df_cand_show, use_container_width=True, hide_index=True)

# ── FOOTER ───────────────────────────────────
st.markdown("---")
st.markdown(f"""
<div style="text-align:center;font-size:11px;color:#6B7280;padding:6px">
  Superhar Recursos Humanos &nbsp;|&nbsp; superhar.com.br &nbsp;|&nbsp;
  (31) 3017-6729 &nbsp;|&nbsp; Relatório gerado em {date.today().strftime('%d/%m/%Y')}
</div>
""", unsafe_allow_html=True)
