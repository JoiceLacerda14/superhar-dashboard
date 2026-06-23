
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import date, datetime
from io import BytesIO
 
st.set_page_config(page_title="Superhar RH — Dashboard R&S", page_icon="🔵",
                   layout="wide", initial_sidebar_state="collapsed")
 
st.markdown("""
<style>
  .main { background:#F4F5F7; }
  .block-container { padding-top:1.2rem; padding-bottom:2rem; }
  .header-bar { background:#185FA5; color:white; padding:12px 24px; border-radius:10px;
    margin-bottom:18px; display:flex; justify-content:space-between; align-items:center; }
  .stTabs [data-baseweb="tab-list"] { gap:4px; }
  .stTabs [data-baseweb="tab"] { background:white; border-radius:8px 8px 0 0;
    padding:8px 20px; font-weight:500; }
  .stTabs [aria-selected="true"] { background:#185FA5 !important; color:white !important; }
  div[data-testid="stMetricValue"] { font-size:1.8rem !important; color:#185FA5; }
</style>
""", unsafe_allow_html=True)
 
AZUL="#185FA5"; AZUL2="#378ADD"; VERDE="#1D9E75"; VERDE2="#0F6E56"
AMBER="#BA7517"; ROXO="#534AB7"; CORAL="#D85A30"; ROSA="#D4537E"
CINZA="#B4B2A9"; PAL=[AZUL,VERDE,AMBER,ROXO,CORAL,ROSA,CINZA,AZUL2]
 
st.markdown("""
<div class="header-bar">
  <div><span style="font-size:18px;font-weight:700">SUPERHAR</span>
  <span style="font-size:12px;opacity:0.7;margin-left:8px">Recursos Humanos</span></div>
  <div style="font-size:12px;opacity:0.85">Dashboard de Recrutamento &amp; Seleção</div>
</div>""", unsafe_allow_html=True)
 
with st.expander("📂  Carregar base de dados (.xlsx)", expanded=True):
    uploaded = st.file_uploader("Selecione o arquivo Excel",
                                type=["xlsx","xls"], label_visibility="collapsed")
    st.caption("Os dados ficam apenas na sua sessão e são apagados ao fechar o navegador.")
 
def carregar(file_bytes):
    buf = BytesIO(file_bytes)
    xl = pd.ExcelFile(buf)
    aba_alvo = next((a for a in xl.sheet_names if "base" in a.lower()), xl.sheet_names[0])
    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=aba_alvo, header=1)
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
 
    # Renomear por posição-semântica (conteúdo normalizado)
    rename_map = {}
    for c in df.columns:
        cu = ' '.join(str(c).upper().replace('\n',' ').replace('\t',' ').split())
        if cu == '(PCD)' and 'PCD' not in [str(x).strip() for x in df.columns]:
            rename_map[c] = 'PCD'
        elif cu == 'DATA ESCOLHA FINALISTA':                          rename_map[c] = 'DATA ESCOLHA FINALISTA'
        elif cu == 'DATA LONG LIST (REALIZADO)':                      rename_map[c] = 'DATA LONG LIST REALIZADO'
        elif cu == 'DATA LONG LIST (PREVISTA)':                       rename_map[c] = 'DATA LONG LIST PREVISTA'
        elif cu == 'DATA SHORT LIST (PREVISTA)':                      rename_map[c] = 'DATA SHORT LIST PREVISTA'
        elif cu == 'DATA SHORT LIST (REALIZADA)':                     rename_map[c] = 'DATA SHORT LIST REALIZADA'
        elif cu == 'NOVA DATA LONG LIST (PREVISTA)' or cu == 'NOVA  DATA LONG LIST (PREVISTA)':  rename_map[c] = 'NOVA DATA LONG LIST PREVISTA'
        elif cu == 'NOVA DATA LONG LIST (REALIZADO)' or cu == 'NOVA  DATA LONG LIST (REALIZADO)': rename_map[c] = 'NOVA DATA LONG LIST REALIZADO'
        elif cu == 'NOVA DATA SHORT LIST (PREVISTA)' or cu == 'NOVA DATA  SHORT LIST (PREVISTA)': rename_map[c] = 'NOVA DATA SHORT LIST PREVISTA'
        elif cu == 'NOVA DATA SHORT LIST (REALIZADA)' or cu == 'NOVA  DATA SHORT LIST (REALIZADA)': rename_map[c] = 'NOVA DATA SHORT LIST REALIZADA'
        elif cu == 'TOTAL DE INSCRITOS':                              rename_map[c] = 'TOTAL INSCRITOS'
        elif cu == 'TOTAL DE CANDIDATOS ABORDADOS':                   rename_map[c] = 'TOTAL ABORDADOS'
        elif cu == 'TOTAL DE CANDIDATOS APRESENTADAS EM LONG LIST':   rename_map[c] = 'TOTAL LONG LIST'
        elif cu == 'TOTAL DE CANDIDATOS APROVADOS EM LONG':           rename_map[c] = 'TOTAL APROVADOS LONG'
        elif 'SELE' in cu and 'SUPERHAR' in cu:                      rename_map[c] = 'TOTAL SELECAO SUPERHAR'
        elif cu == 'TOTAL DE CANDIDATOS SHORT LIST':                  rename_map[c] = 'TOTAL SHORT LIST'
        elif 'ENTREVISTA' in cu and 'GESTOR' in cu:                  rename_map[c] = 'TOTAL ENTREVISTA GESTOR'
        elif 'REPROVADOS' in cu and 'MEDICINA' in cu:                rename_map[c] = 'TOTAL REPROVADOS MEDICINA'
        elif 'SALARIO ATUAL' in cu or 'SALÁRIO ATUAL' in cu:         rename_map[c] = 'SALARIO ATUAL'
        elif 'PRETENS' in cu and 'SALARIAL' in cu:                   rename_map[c] = 'PRETENSAO SALARIAL'
        elif 'ORIGEM' in cu and 'RECRUTAMENTO' in cu:                rename_map[c] = 'ORIGEM RECRUTAMENTO'
 
    df = df.rename(columns=rename_map)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
 
    # Garantir Status Samarco
    if 'Status Samarco' not in df.columns:
        scols = [c for c in df.columns if 'status' in str(c).lower()]
        if scols: df = df.rename(columns={scols[0]: 'Status Samarco'})
 
    # Datas
    date_cols = ['DATA RECEBIMENTO','DATA ALINHAMENTO',
                 'DATA DE ABERTURA (Inicio Cronograma)',
                 'DATA LONG LIST PREVISTA','DATA LONG LIST REALIZADO',
                 'DATA SHORT LIST PREVISTA','DATA SHORT LIST REALIZADA',
                 'DATA ESCOLHA FINALISTA','DATA ADMISSÃO',
                 'NOVA DATA LONG LIST PREVISTA','NOVA DATA LONG LIST REALIZADO',
                 'NOVA DATA SHORT LIST PREVISTA','NOVA DATA SHORT LIST REALIZADA']
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors='coerce')
 
    # Numéricos funil
    for c in ['TOTAL INSCRITOS','TOTAL ABORDADOS','TOTAL LONG LIST','TOTAL APROVADOS LONG',
              'TOTAL SELECAO SUPERHAR','TOTAL SHORT LIST','TOTAL ENTREVISTA GESTOR',
              'TOTAL REPROVADOS MEDICINA','IDADE']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
 
    # Numéricos salário
    for c in ['SALÁRIO PREVISTO (Minimo)','SALÁRIO PREVISTO (Maximo)',
              'SALARIO ADMISSÃO','SALARIO ATUAL','PRETENSAO SALARIAL']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
 
    col_ab  = 'DATA DE ABERTURA (Inicio Cronograma)'
    col_fin = 'DATA ESCOLHA FINALISTA'
    col_slr = 'DATA SHORT LIST REALIZADA'
    col_slp = 'DATA SHORT LIST PREVISTA'
    col_adm = 'DATA ADMISSÃO'
 
    df['time_to_fill'] = (df[col_fin] - df[col_ab]).dt.days if (col_fin in df.columns and col_ab in df.columns) else pd.NA
    df['days_abertura_sl'] = (df[col_slr] - df[col_ab]).dt.days if (col_slr in df.columns and col_ab in df.columns) else pd.NA
    df['days_sl_contrat']  = (df[col_adm] - df[col_slr]).dt.days if (col_adm in df.columns and col_slr in df.columns) else pd.NA
    df['sla_ok']     = (df[col_slr] <= df[col_slp]) if (col_slr in df.columns and col_slp in df.columns) else False
    df['admitido']   = df[col_adm].notna() if col_adm in df.columns else False
    df['mes_abertura'] = df[col_ab].dt.to_period('M').dt.to_timestamp() if col_ab in df.columns else pd.NaT
 
    # Faixa etária
    if 'IDADE' in df.columns:
        bins   = [0,25,30,35,40,45,50,120]
        labels = ['Até 25','26–30','31–35','36–40','41–45','46–50','50+']
        df['faixa_etaria'] = pd.cut(df['IDADE'], bins=bins, labels=labels, right=True)
 
    # Gênero
    if 'GÊNERO' in df.columns:
        df['genero_norm'] = df['GÊNERO'].astype(str).str.strip().replace({
            'Homem cisgênero':'Masculino','Homem Cisgênero':'Masculino',
            'Mulher Cisgênero':'Feminino','Mulher cisgênero':'Feminino',
            'nan':'Não informado','':'Não informado'})
    else:
        df['genero_norm'] = 'Não informado'
 
    # PCD
    pcd_col = 'PCD' if 'PCD' in df.columns else None
    if pcd_col:
        df['pcd_norm'] = df[pcd_col].astype(str).str.strip().str.lower().map(
            {'sim':'Sim','não':'Não','nao':'Não','nan':'Não informado'}
        ).fillna('Não informado')
    else:
        df['pcd_norm'] = 'Não informado'
 
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
 
# ── HELPER: filtro de período ─────────────────
def filtro_periodo(df, key_prefix):
    col_ab = 'DATA DE ABERTURA (Inicio Cronograma)'
    datas_validas = df[col_ab].dropna()
    min_d = datas_validas.min().date() if len(datas_validas) else date(2024,1,1)
    max_d = datas_validas.max().date() if len(datas_validas) else date.today()
    c1, c2 = st.columns(2)
    with c1:
        inicio = st.date_input("Início", value=min_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_ini")
    with c2:
        fim = st.date_input("Fim", value=max_d, min_value=min_d, max_value=max_d, key=f"{key_prefix}_fim")
    mask = (df[col_ab].dt.date >= inicio) & (df[col_ab].dt.date <= fim)
    return df[mask], inicio, fim
 
# ═══════════════════════════════════════════════
# ABAS
# ═══════════════════════════════════════════════
aba1, aba2, aba3 = st.tabs(["📊  Gerencial", "👩‍💼  Consultoras", "🤝  Visão Cliente"])
 
# ══════════════════════════════════════════════
# ABA 1 — GERENCIAL
# ══════════════════════════════════════════════
with aba1:
    st.markdown("**Período de abertura da vaga**")
    dg, ini_g, fim_g = filtro_periodo(df_raw, 'ger')
 
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        emps = ["Todas"] + sorted(dg['EMPRESA'].dropna().unique().tolist())
        f_emp = st.selectbox("Empresa / Cliente", emps, key="g_emp")
    with fc2:
        nivs = ["Todos"] + sorted(dg['NÍVEL'].dropna().unique().tolist())
        f_niv = st.selectbox("Nível", nivs, key="g_niv")
    with fc3:
        cons_s = ["Todas"] + sorted(dg['CONSULTOR (A) SELEÇÃO'].dropna().unique().tolist())
        f_cons = st.selectbox("Consultora Seleção", cons_s, key="g_cons")
 
    if f_emp  != "Todas": dg = dg[dg['EMPRESA']==f_emp]
    if f_niv  != "Todos": dg = dg[dg['NÍVEL']==f_niv]
    if f_cons != "Todas": dg = dg[dg['CONSULTOR (A) SELEÇÃO']==f_cons]
 
    dg_vaga = dg.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
    STATUS_COL = 'Status Samarco' if 'Status Samarco' in dg_vaga.columns else None
 
    total_vagas  = dg_vaga['COD.VAGA SUPERHAR'].nunique()
    concluidas   = dg_vaga[STATUS_COL].isin(['Concluído','Concluído C/Interno','Admissão']).sum() if STATUS_COL else 0
    em_andamento = total_vagas - concluidas
    tx_concl     = round(concluidas/total_vagas*100) if total_vagas>0 else 0
    ttf_med      = int(dg_vaga['time_to_fill'].median()) if dg_vaga['time_to_fill'].notna().any() else 0
    admitidos    = int(dg['admitido'].sum())
    sla_pct      = round(dg_vaga['sla_ok'].mean()*100) if dg_vaga['sla_ok'].notna().any() else 0
 
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Total de Vagas",      total_vagas)
    k2.metric("Concluídas",          concluidas)
    k3.metric("Em Andamento",        em_andamento)
    k4.metric("Taxa de Conclusão",   f"{tx_concl}%")
    k5.metric("Mediana Time to Fill",f"{ttf_med} dias")
    k6.metric("Admitidos",           admitidos)
 
    st.markdown("---")
 
    # LINHA 2: Vagas por cliente  |  Vagas abertas por mês
    c1, c2 = st.columns(2)
 
    with c1:
        st.markdown("**Vagas por empresa / cliente**")
        vc = dg_vaga['EMPRESA'].value_counts().reset_index()
        vc.columns = ['Empresa','Vagas']
        fig = px.pie(vc, values='Vagas', names='Empresa', hole=0.4,
                     color_discrete_sequence=PAL)
        fig.update_traces(textinfo='percent+value', textfont_size=11)
        fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=260,
                          legend=dict(font_size=10))
        st.plotly_chart(fig, use_container_width=True)
 
    with c2:
        st.markdown("**Vagas abertas por mês**")
        dv2 = df_raw.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first').copy()
        dv2['mes'] = dv2['DATA DE ABERTURA (Inicio Cronograma)'].dt.to_period('M').dt.to_timestamp()
        mens = dv2.groupby('mes').size().reset_index(name='Vagas').sort_values('mes')
        mens['mes_label'] = mens['mes'].dt.strftime('%b/%Y')
        fig2 = px.bar(mens, x='mes_label', y='Vagas',
                      color_discrete_sequence=[AZUL], text='Vagas')
        fig2.update_traces(textposition='outside', textfont_size=11)
        fig2.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=260,
                           xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig2, use_container_width=True)
 
    st.markdown("---")
 
    # LINHA 3: Canal de origem  |  Diversidade (gênero + PcD + pizza Tipo Cronograma)
    c3, c4 = st.columns(2)
 
    with c3:
        st.markdown("**Candidatos por canal de origem**")
        if 'ORIGEM RECRUTAMENTO' in dg.columns:
            ori = dg['ORIGEM RECRUTAMENTO'].value_counts().reset_index()
            ori.columns = ['Canal','Qtd']
            fig3 = px.bar(ori, x='Qtd', y='Canal', orientation='h',
                          color_discrete_sequence=[AZUL], text='Qtd')
            fig3.update_traces(textposition='outside', textfont_size=10)
            fig3.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=260,
                               xaxis_title='', yaxis_title='', font=dict(size=10))
            st.plotly_chart(fig3, use_container_width=True)
 
    with c4:
        st.markdown("**Diversidade e Fluxo**")
        cg1, cg2, cg3 = st.columns(3)
 
        with cg1:
            st.caption("Gênero")
            gen = dg[dg['genero_norm']!='Não informado']['genero_norm'].value_counts().reset_index()
            gen.columns = ['Gênero','Qtd']
            if len(gen):
                fg = px.pie(gen, values='Qtd', names='Gênero', hole=0.35,
                            color_discrete_sequence=[ROSA,AZUL,ROXO])
                fg.update_traces(textinfo='percent+value', textfont_size=9)
                fg.update_layout(margin=dict(t=0,b=20,l=0,r=0), height=180, showlegend=False)
                st.plotly_chart(fg, use_container_width=True)
 
        with cg2:
            st.caption("PcD")
            pcd = dg[dg['pcd_norm']!='Não informado']['pcd_norm'].value_counts().reset_index()
            pcd.columns = ['PcD','Qtd']
            if len(pcd):
                fp = px.pie(pcd, values='Qtd', names='PcD', hole=0.35,
                            color_discrete_sequence=[AMBER,CINZA])
                fp.update_traces(textinfo='percent+value', textfont_size=9)
                fp.update_layout(margin=dict(t=0,b=20,l=0,r=0), height=180, showlegend=False)
                st.plotly_chart(fp, use_container_width=True)
 
        with cg3:
            st.caption("Fluxo do Processo")
            if 'Tipo Cronograma' in dg_vaga.columns:
                tc = dg_vaga['Tipo Cronograma'].value_counts().reset_index()
                tc.columns = ['Tipo','Qtd']
                ft = px.pie(tc, values='Qtd', names='Tipo', hole=0.35,
                            color_discrete_sequence=[VERDE,AZUL2,AMBER,CORAL])
                ft.update_traces(textinfo='percent+value', textfont_size=9)
                ft.update_layout(margin=dict(t=0,b=20,l=0,r=0), height=180, showlegend=False)
                st.plotly_chart(ft, use_container_width=True)
 
    st.markdown("---")
 
    # SÍNTESE POR EMPRESA
    st.markdown("**Síntese por empresa**")
    dg_vg = dg.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
    sc = 'Status Samarco' if 'Status Samarco' in dg_vg.columns else None
    sint = dg_vg.groupby('EMPRESA').agg(
        Vagas       = ('COD.VAGA SUPERHAR','count'),
        Concluídas  = ('Status Samarco', lambda x: x.isin(['Concluído','Concluído C/Interno','Admissão']).sum()) if sc else ('COD.VAGA SUPERHAR', lambda x: 0),
        Inscritos   = ('TOTAL INSCRITOS','sum'),
        Short_List  = ('TOTAL SHORT LIST','sum'),
        TTF_Mediana = ('time_to_fill','median'),
    ).reset_index()
    sint['Conclusão %']    = (sint['Concluídas']/sint['Vagas']*100).round(0).fillna(0).astype(int)
    sint['TTF_Mediana']    = sint['TTF_Mediana'].fillna(0).astype(int)
    sint = sint.rename(columns={'EMPRESA':'Empresa','TTF_Mediana':'Mediana TTF (dias)'})
    st.dataframe(sint[['Empresa','Vagas','Concluídas','Conclusão %','Inscritos','Short_List','Mediana TTF (dias)']],
                 use_container_width=True, hide_index=True)
 
 
# ══════════════════════════════════════════════
# ABA 2 — CONSULTORAS
# ══════════════════════════════════════════════
with aba2:
    st.markdown("**Período de abertura da vaga**")
    dc, ini_c, fim_c = filtro_periodo(df_raw, 'cons')
 
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        emps2 = ["Todas"] + sorted(dc['EMPRESA'].dropna().unique().tolist())
        f_emp2 = st.selectbox("Empresa", emps2, key="c_emp")
    with fc2:
        # Consultora CAPTAÇÃO (coluna Y)
        capt_list = ["Todas"] + sorted(dc['CONSULTOR (A) CAPTAÇÃO'].dropna().unique().tolist()) if 'CONSULTOR (A) CAPTAÇÃO' in dc.columns else ["Todas"]
        f_capt = st.selectbox("Consultora Captação", capt_list, key="c_capt")
    with fc3:
        niv2 = ["Todos"] + sorted(dc['NÍVEL'].dropna().unique().tolist())
        f_niv2 = st.selectbox("Nível", niv2, key="c_niv")
    with fc4:
        # Filtro por código de vaga
        vagas_lista = ["Todas"] + sorted(dc['COD.VAGA SUPERHAR'].dropna().unique().tolist())
        f_vaga_c = st.selectbox("Código da Vaga", vagas_lista, key="c_vaga")
 
    if f_emp2  != "Todas": dc = dc[dc['EMPRESA']==f_emp2]
    if f_capt  != "Todas" and 'CONSULTOR (A) CAPTAÇÃO' in dc.columns:
        dc = dc[dc['CONSULTOR (A) CAPTAÇÃO']==f_capt]
    if f_niv2  != "Todos": dc = dc[dc['NÍVEL']==f_niv2]
    if f_vaga_c != "Todas": dc = dc[dc['COD.VAGA SUPERHAR']==f_vaga_c]
 
    dc_vaga = dc.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
 
    # KPIs destaque
    grp = dc_vaga.groupby('CONSULTOR (A) SELEÇÃO').agg(
        Vagas      = ('COD.VAGA SUPERHAR','count'),
        Concluídas = ('Status Samarco', lambda x: x.isin(['Concluído','Concluído C/Interno','Admissão']).sum()) if 'Status Samarco' in dc_vaga.columns else ('COD.VAGA SUPERHAR', lambda x: 0),
        TTF_med    = ('time_to_fill','median'),
        TTF_abertura_sl = ('days_abertura_sl','median'),
        SLA_ok     = ('sla_ok','mean'),
        Inscritos  = ('TOTAL INSCRITOS','sum'),
        Short_List = ('TOTAL SHORT LIST','sum'),
        Entrev     = ('TOTAL ENTREVISTA GESTOR','sum'),
    ).reset_index()
    grp['Conclusão %'] = (grp['Concluídas']/grp['Vagas']*100).round(0).fillna(0).astype(int)
    grp['TTF_med']     = grp['TTF_med'].fillna(0).round(0).astype(int)
    grp['TTF_ab_sl']   = grp['TTF_abertura_sl'].fillna(0).round(0).astype(int)
    grp['SLA %']       = (grp['SLA_ok']*100).round(0).fillna(0).astype(int)
 
    if len(grp)>0:
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Maior Volume",       grp.loc[grp['Vagas'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0], f"{grp['Vagas'].max()} vagas")
        k2.metric("Mais Concluídas",    grp.loc[grp['Concluídas'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0], f"{grp['Concluídas'].max()} conc.")
        tidx = grp[grp['TTF_med']>0]['TTF_med'].idxmin() if (grp['TTF_med']>0).any() else grp['TTF_med'].idxmin()
        k3.metric("Menor TTF",          grp.loc[tidx,'CONSULTOR (A) SELEÇÃO'].split()[0], f"{grp.loc[tidx,'TTF_med']} dias")
        k4.metric("Melhor SLA",         grp.loc[grp['SLA %'].idxmax(),'CONSULTOR (A) SELEÇÃO'].split()[0], f"SLA {grp['SLA %'].max()}%")
 
    st.markdown("---")
 
    # FUNIL por consultora
    st.markdown("**Funil de recrutamento — totais do período filtrado**")
    f1, f2 = st.columns(2)
    with f1:
        ins  = int(dc_vaga['TOTAL INSCRITOS'].sum())
        ab   = int(dc_vaga['TOTAL ABORDADOS'].sum())
        ll   = int(dc_vaga['TOTAL LONG LIST'].sum())
        sl   = int(dc_vaga['TOTAL SHORT LIST'].sum())
        eg   = int(dc_vaga['TOTAL ENTREVISTA GESTOR'].sum())
        med  = int(dc_vaga['TOTAL REPROVADOS MEDICINA'].sum()) if 'TOTAL REPROVADOS MEDICINA' in dc_vaga.columns else 0
        adm2 = int(dc['admitido'].sum())
 
        fig_f = go.Figure(go.Funnel(
            y=['Inscritos','Abordados','Long List','Short List','Entrev. Gestor','Reprov. Medicina','Admitidos'],
            x=[ins,ab,ll,sl,eg,med,adm2],
            textinfo='value+percent initial',
            marker=dict(color=[AZUL,'#2B74BE',AZUL2,VERDE,VERDE2,AMBER,'#085041']),
            connector=dict(line=dict(color='#E2E4E9',width=1))
        ))
        fig_f.update_layout(margin=dict(t=0,b=0,l=0,r=80), height=280, font=dict(size=11))
        st.plotly_chart(fig_f, use_container_width=True)
 
    with f2:
        # Tempo de fechamento: abertura → Short List Realizada
        st.markdown("**Tempo médio: Abertura → Short List (dias)**")
        grp_ttf = grp[grp['TTF_ab_sl']>0].copy()
        nomes = grp_ttf['CONSULTOR (A) SELEÇÃO'].apply(lambda x: ' '.join(x.split()[:2]))
        fig_tt = px.bar(x=grp_ttf['TTF_ab_sl'], y=nomes, orientation='h',
                        color_discrete_sequence=[AMBER], text=grp_ttf['TTF_ab_sl'])
        fig_tt.update_traces(textposition='outside', textfont_size=10)
        fig_tt.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=280,
                             xaxis_title='dias', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig_tt, use_container_width=True)
 
    st.markdown("---")
 
    # TABELA COMPARATIVA
    st.markdown("**Desempenho por consultora (Seleção)**")
    tab_show = grp.rename(columns={
        'CONSULTOR (A) SELEÇÃO':'Consultora','Vagas':'Vagas','Concluídas':'Conc.',
        'Conclusão %':'Conc.%','TTF_med':'TTF Total (dias)','TTF_ab_sl':'Aber.→SL (dias)',
        'SLA %':'SLA %','Inscritos':'Inscritos','Short_List':'Short List','Entrev':'Entrevistas'
    })
    st.dataframe(tab_show[['Consultora','Vagas','Conc.','Conc.%','TTF Total (dias)',
                            'Aber.→SL (dias)','SLA %','Inscritos','Short List','Entrevistas']],
                 use_container_width=True, hide_index=True)
 
    st.markdown("---")
 
    # GRÁFICOS COMPARATIVOS
    cc1, cc2, cc3 = st.columns(3)
    nomes_c = grp['CONSULTOR (A) SELEÇÃO'].apply(lambda x: ' '.join(x.split()[:2]))
 
    with cc1:
        st.markdown("**Vagas por consultora**")
        fig = px.bar(x=grp['Vagas'], y=nomes_c, orientation='h',
                     color_discrete_sequence=[AZUL], text=grp['Vagas'])
        fig.update_traces(textposition='outside', textfont_size=10)
        fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)
 
    with cc2:
        st.markdown("**Vagas concluídas por consultora**")
        fig = px.bar(x=grp['Concluídas'], y=nomes_c, orientation='h',
                     color_discrete_sequence=[VERDE], text=grp['Concluídas'])
        fig.update_traces(textposition='outside', textfont_size=10)
        fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                          xaxis_title='', yaxis_title='', font=dict(size=10))
        st.plotly_chart(fig, use_container_width=True)
 
    with cc3:
        st.markdown("**Status dos processos**")
        if 'Status Samarco' in dc_vaga.columns:
            stat = dc_vaga['Status Samarco'].value_counts().reset_index()
            stat.columns = ['Status','Qtd']
            fig = px.pie(stat, values='Qtd', names='Status', hole=0.35,
                         color_discrete_sequence=PAL)
            fig.update_traces(textinfo='percent+value', textfont_size=10)
            fig.update_layout(height=240, margin=dict(t=0,b=0,l=0,r=0),
                              legend=dict(font_size=9))
            st.plotly_chart(fig, use_container_width=True)
 
    # FUNIL POR CONSULTORA — TABELA
    st.markdown("---")
    st.markdown("**Funil detalhado por consultora**")
    funil_cons = dc_vaga.groupby('CONSULTOR (A) SELEÇÃO').agg(
        Inscritos   = ('TOTAL INSCRITOS','sum'),
        Long_List   = ('TOTAL LONG LIST','sum'),
        Short_List  = ('TOTAL SHORT LIST','sum'),
        Entrevistas = ('TOTAL ENTREVISTA GESTOR','sum'),
        Rep_Med     = ('TOTAL REPROVADOS MEDICINA','sum'),
    ).reset_index()
    funil_cons.columns = ['Consultora','Inscritos','Long List','Short List','Entrev. Gestor','Reprov. Medicina']
    st.dataframe(funil_cons, use_container_width=True, hide_index=True)
 
 
# ══════════════════════════════════════════════
# ABA 3 — VISÃO CLIENTE
# ══════════════════════════════════════════════
with aba3:
    # FILTROS
    ff1, ff2, ff3 = st.columns(3)
    with ff1:
        clientes_vc = ["Todos"] + sorted(df_raw['EMPRESA'].dropna().unique().tolist())
        f_cli_vc = st.selectbox("Cliente", clientes_vc, key="vc_cli")
    dvc = df_raw.copy()
    if f_cli_vc != "Todos": dvc = dvc[dvc['EMPRESA']==f_cli_vc]
 
    with ff2:
        vagas_vc = dvc.drop_duplicates(subset='COD.VAGA SUPERHAR', keep='first')
        vagas_vc = vagas_vc[vagas_vc['CARGO'].notna()]
        vagas_vc['label_vaga'] = vagas_vc['COD.VAGA SUPERHAR'].astype(str) + ' — ' + vagas_vc['CARGO'].astype(str).str.strip()
        sel_vaga = st.selectbox("Vaga (Código + Cargo)", vagas_vc['label_vaga'].tolist(), key="vc_vaga")
 
    vaga_row = vagas_vc[vagas_vc['label_vaga']==sel_vaga].iloc[0] if len(vagas_vc) else None
 
    if vaga_row is None:
        st.warning("Nenhuma vaga encontrada para o filtro selecionado.")
        st.stop()
 
    cand_vaga = df_raw[df_raw['COD.VAGA SUPERHAR']==vaga_row['COD.VAGA SUPERHAR']]
 
    with ff3:
        etapas_vc = ["Todas"]
        if 'Etapa' in cand_vaga.columns:
            etapas_vc += sorted(cand_vaga['Etapa'].dropna().unique().tolist())
        f_etapa = st.selectbox("Etapa (coluna AK)", etapas_vc, key="vc_etapa")
    if f_etapa != "Todas" and 'Etapa' in cand_vaga.columns:
        cand_vaga = cand_vaga[cand_vaga['Etapa']==f_etapa]
 
    st.markdown("---")
 
    # KPI CARDS — fixos conforme solicitado
    status_vc   = str(vaga_row.get('Status Samarco','—'))
    cons_capt   = str(vaga_row.get('CONSULTOR (A) CAPTAÇÃO','—'))  # captação conforme solicitado
    inscritos_v = int(vaga_row['TOTAL INSCRITOS']) if pd.notna(vaga_row.get('TOTAL INSCRITOS')) else 0
    sl_v        = int(vaga_row['TOTAL SHORT LIST']) if pd.notna(vaga_row.get('TOTAL SHORT LIST')) else 0
 
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Status",                    status_vc)
    k2.metric("Candidatos Mapeados",       inscritos_v,   "Long list")
    k3.metric("Apresentados ao Cliente",   sl_v,          "Short list")
    k4.metric("Responsável (Captação)",    cons_capt)
 
    st.markdown("---")
 
    # LINHA DO TEMPO + DIAS
    col_tl, col_dias = st.columns([2,1])
 
    with col_tl:
        st.markdown("**Linha do tempo do processo**")
        # Nomenclaturas solicitadas: Alinhamento, Long List, Entrevistas, Short list, Oferta, Contratação
        etapas_tl = [
            ("Alinhamento",  vaga_row.get('DATA ALINHAMENTO')),
            ("Long List",    vaga_row.get('DATA LONG LIST REALIZADO')),
            ("Entrevistas",  vaga_row.get('DATA ESCOLHA FINALISTA')),
            ("Short List",   vaga_row.get('DATA SHORT LIST REALIZADA')),
            ("Oferta",       None),  # não há coluna de oferta — mostra pendente
            ("Contratação",  cand_vaga['DATA ADMISSÃO'].dropna().min() if cand_vaga['DATA ADMISSÃO'].notna().any() else None),
        ]
        concl_tl = sum(1 for _,d in etapas_tl if pd.notna(d))
        html_tl = '<div style="display:flex;align-items:flex-start;overflow-x:auto;padding:10px 0">'
        for i,(nome,data) in enumerate(etapas_tl):
            done    = pd.notna(data)
            current = i == concl_tl and not done
            dot_bg  = VERDE if done else (AZUL if current else '#E2E4E9')
            dot_txt = '✓' if done else ('●' if current else '○')
            nc      = VERDE2 if done else (AZUL if current else '#6B7280')
            data_s  = pd.Timestamp(data).strftime('%d/%m/%y') if done else '—'
            line = f'<div style="flex:1;height:2px;background:{"#1D9E75" if i<concl_tl-1 else "#E2E4E9"};margin-top:13px;min-width:20px"></div>' if i<len(etapas_tl)-1 else ''
            html_tl += f'''<div style="display:flex;flex-direction:column;align-items:center;min-width:80px">
              <div style="width:26px;height:26px;border-radius:50%;background:{dot_bg};display:flex;
                align-items:center;justify-content:center;font-size:12px;color:white;font-weight:700">{dot_txt}</div>
              <div style="font-size:9px;font-weight:600;color:{nc};text-align:center;margin-top:5px;line-height:1.3">{nome}</div>
              <div style="font-size:9px;color:#6B7280;margin-top:1px">{data_s}</div>
            </div>{line}'''
        html_tl += '</div>'
        st.markdown(html_tl, unsafe_allow_html=True)
 
    with col_dias:
        st.markdown("**Tempo em dias**")
        d_aber_sl  = int(vaga_row.get('days_abertura_sl', 0) or 0)
        d_sl_cont  = int(vaga_row.get('days_sl_contrat', 0) or 0)
 
        for label, val, cor in [
            ("Alinhamento → Short List", f"{d_aber_sl} dias", AZUL),
            ("Short List → Contratação", f"{d_sl_cont} dias", VERDE),
        ]:
            st.markdown(f"""
            <div style="background:white;border-radius:8px;padding:10px 14px;
                        border:1px solid #E2E4E9;margin-bottom:10px;text-align:center">
              <div style="font-size:10px;color:#6B7280;font-weight:600;
                text-transform:uppercase;letter-spacing:0.04em">{label}</div>
              <div style="font-size:26px;font-weight:700;color:{cor};margin:4px 0">{val}</div>
            </div>""", unsafe_allow_html=True)
 
    st.markdown("---")
 
    # FUNIL + BENCHMARKING
    cf1, cf2 = st.columns(2)
 
    with cf1:
        st.markdown("**Funil de candidatos desta vaga**")
        ins_v  = int(vaga_row['TOTAL INSCRITOS'])    if pd.notna(vaga_row.get('TOTAL INSCRITOS'))    else 0
        ll_v   = int(vaga_row['TOTAL LONG LIST'])    if pd.notna(vaga_row.get('TOTAL LONG LIST'))    else 0
        sl_v2  = int(vaga_row['TOTAL SHORT LIST'])   if pd.notna(vaga_row.get('TOTAL SHORT LIST'))   else 0
        eg_v   = int(vaga_row['TOTAL ENTREVISTA GESTOR']) if pd.notna(vaga_row.get('TOTAL ENTREVISTA GESTOR')) else 0
        med_v  = int(vaga_row['TOTAL REPROVADOS MEDICINA']) if pd.notna(vaga_row.get('TOTAL REPROVADOS MEDICINA')) else 0
        adm_v  = int(cand_vaga['admitido'].sum())
 
        fig_fv = go.Figure(go.Funnel(
            y=['Inscritos','Long List','Short List','Entrev. Gestor','Reprov. Medicina','Admitidos'],
            x=[ins_v, ll_v, sl_v2, eg_v, med_v, adm_v],
            textinfo='value+percent initial',
            marker=dict(color=[AZUL,AZUL2,VERDE,VERDE2,AMBER,'#085041']),
        ))
        fig_fv.update_layout(margin=dict(t=0,b=0,l=0,r=80), height=260, font=dict(size=11))
        st.plotly_chart(fig_fv, use_container_width=True)
 
    with cf2:
        st.markdown("**Benchmarking salarial**")
        sal_min  = vaga_row.get('SALÁRIO PREVISTO (Minimo)')
        sal_max  = vaga_row.get('SALÁRIO PREVISTO (Maximo)')
        sal_pret = cand_vaga['PRETENSAO SALARIAL'].dropna().mean() if 'PRETENSAO SALARIAL' in cand_vaga.columns and cand_vaga['PRETENSAO SALARIAL'].notna().any() else None
        sal_adm  = cand_vaga['SALARIO ADMISSÃO'].dropna().mean() if 'SALARIO ADMISSÃO' in cand_vaga.columns and cand_vaga['SALARIO ADMISSÃO'].notna().any() else None
 
        labels_s, vals_s, cors_s = [], [], []
        try:
            if pd.notna(sal_min)  and float(sal_min)>0:  labels_s.append('Previsto (mín)'); vals_s.append(float(sal_min));  cors_s.append('#B5D4F4')
            if pd.notna(sal_max)  and float(sal_max)>0:  labels_s.append('Previsto (máx)'); vals_s.append(float(sal_max));  cors_s.append(AZUL2)
            if sal_pret and float(sal_pret)>0:            labels_s.append('Pretensão média'); vals_s.append(float(sal_pret)); cors_s.append(AMBER)
            if sal_adm  and float(sal_adm)>0:            labels_s.append('Admissão média'); vals_s.append(float(sal_adm));  cors_s.append(VERDE)
        except: pass
 
        if vals_s:
            txt_s = ['R$ {:,.0f}'.format(v).replace(',','.') for v in vals_s]
            fig_s = go.Figure(go.Bar(x=labels_s, y=vals_s,
                                     marker=dict(color=cors_s),
                                     text=txt_s, textposition='outside'))
            fig_s.update_layout(margin=dict(t=20,b=0,l=0,r=0), height=240,
                                yaxis=dict(showticklabels=False), font=dict(size=10))
            st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.caption("Sem dados salariais para esta vaga.")
 
    st.markdown("---")
 
    # DIVERSIDADE: Gênero + PcD + Faixa etária
    st.markdown("**Diversidade da short list**")
    cd1, cd2, cd3 = st.columns(3)
 
    with cd1:
        st.caption("Gênero")
        gv = cand_vaga[cand_vaga['genero_norm']!='Não informado']['genero_norm'].value_counts().reset_index()
        gv.columns = ['Gênero','Qtd']
        if len(gv):
            fg = px.pie(gv, values='Qtd', names='Gênero', hole=0.35,
                        color_discrete_sequence=[ROSA,AZUL,ROXO])
            fg.update_traces(textinfo='percent+value', textfont_size=10)
            fg.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                             legend=dict(font_size=9, orientation='h', y=-0.15))
            st.plotly_chart(fg, use_container_width=True)
        else:
            st.caption("Sem dados de gênero.")
 
    with cd2:
        st.caption("PcD")
        pv = cand_vaga[cand_vaga['pcd_norm']!='Não informado']['pcd_norm'].value_counts().reset_index()
        pv.columns = ['PcD','Qtd']
        if len(pv):
            fp = px.pie(pv, values='Qtd', names='PcD', hole=0.35,
                        color_discrete_sequence=[AMBER,CINZA])
            fp.update_traces(textinfo='percent+value', textfont_size=10)
            fp.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                             legend=dict(font_size=9, orientation='h', y=-0.15))
            st.plotly_chart(fp, use_container_width=True)
        else:
            st.caption("Sem dados de PcD.")
 
    with cd3:
        st.caption("Faixa etária")
        if 'faixa_etaria' in cand_vaga.columns:
            fe = cand_vaga['faixa_etaria'].value_counts().sort_index().reset_index()
            fe.columns = ['Faixa','Qtd']
            fe = fe[fe['Qtd']>0]
            if len(fe):
                ff_fig = px.pie(fe, values='Qtd', names='Faixa', hole=0.35,
                                color_discrete_sequence=[AZUL,VERDE,AMBER,ROXO,CORAL,ROSA,CINZA])
                ff_fig.update_traces(textinfo='percent+value', textfont_size=10)
                ff_fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=200,
                                     legend=dict(font_size=9, orientation='h', y=-0.15))
                st.plotly_chart(ff_fig, use_container_width=True)
            else:
                st.caption("Sem dados de idade.")
        else:
            st.caption("Sem dados de faixa etária.")
 
    st.markdown("---")
 
    # CANDIDATOS FINALISTAS
    st.markdown("**Candidatos deste processo**")
    cols_fin = ['NOME CANDIDATO','GÊNERO','RAÇA','pcd_norm','faixa_etaria',
                'ORIGEM RECRUTAMENTO','ULTIMA EMPRESA / ATUAL',
                'SALARIO ATUAL','PRETENSAO SALARIAL','SALARIO ADMISSÃO',
                'DATA ADMISSÃO','INTERNO OU EXTERNO','Etapa']
    cols_ex = [c for c in cols_fin if c in cand_vaga.columns]
    df_fin  = cand_vaga[cols_ex].copy()
    if 'DATA ADMISSÃO' in df_fin.columns:
        df_fin['DATA ADMISSÃO'] = pd.to_datetime(df_fin['DATA ADMISSÃO'], errors='coerce').dt.strftime('%d/%m/%Y')
    rename_fin = {'pcd_norm':'PcD','faixa_etaria':'Faixa Etária',
                  'ORIGEM RECRUTAMENTO':'Canal','ULTIMA EMPRESA / ATUAL':'Empresa Atual',
                  'SALARIO ATUAL':'Sal. Atual','PRETENSAO SALARIAL':'Pretensão',
                  'SALARIO ADMISSÃO':'Sal. Admissão'}
    df_fin = df_fin.rename(columns=rename_fin)
    st.dataframe(df_fin, use_container_width=True, hide_index=True)
 
# ── FOOTER
st.markdown("---")
st.markdown(f"""<div style="text-align:center;font-size:11px;color:#6B7280;padding:6px">
  Superhar Recursos Humanos &nbsp;|&nbsp; superhar.com.br &nbsp;|&nbsp;
  (31) 3017-6729 &nbsp;|&nbsp; Relatório gerado em {date.today().strftime('%d/%m/%Y')}
</div>""", unsafe_allow_html=True)
 
