import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
import openpyxl
from streamlit.components.v1 import html as st_html

st.set_page_config(page_title="Planilha de Ganhos", layout="wide")

# Tabela de lookup da diária (igual ao VLOOKUP do Excel)
DIARIA_LOOKUP = [
    (0, 240), (101, 260), (126, 280), (151, 300), (176, 320),
    (201, 340), (226, 360), (251, 380), (276, 400), (301, 420),
    (326, 440), (351, 460), (376, 480)
]


def calc_diaria(km, is_sunday: bool):
    if pd.isna(km):
        return None
    try:
        km = float(km)
    except (ValueError, TypeError):
        return None
    if km <= 0:
        return 0
    val = 0
    for thr, v in DIARIA_LOOKUP:
        if km >= thr:
            val = v
        else:
            break
    if is_sunday:
        val += 48
    return val


def calc_pct(pacotes, entregas):
    try:
        p = float(pacotes)
        e = float(entregas)
        if p == 0:
            return None
        return e / p
    except (TypeError, ValueError):
        return None


def calc_bonus(pct):
    if pct is None:
        return 0
    if pct == 1.0:
        return 20
    if 0.98 <= pct < 1.0:
        return 10
    return 0


def dia_semana_br(dt):
    if pd.isna(dt):
        return ""
    dias = [
        "segunda-feira", "terca-feira", "quarta-feira", "quinta-feira",
        "sexta-feira", "sabado", "domingo"
    ]
    return dias[dt.weekday()]


MESES = {
    "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12,
}


def importar_excel_original():
    wb = openpyxl.load_workbook("Planilha de Ganhos.xlsx", data_only=True)
    dados = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        mes = sheet.strip().lower()
        if mes not in MESES:
            continue
        rows = []
        empties = 0
        for r in range(4, 100):
            v_data = ws.cell(row=r, column=2).value
            if v_data is None:
                empties += 1
                if empties >= 10:
                    break
                continue
            empties = 0
            if isinstance(v_data, str):
                try:
                    v_data = datetime.strptime(v_data, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
            rows.append({
                "data": v_data,
                "km": ws.cell(row=r, column=4).value,
                "pacotes": ws.cell(row=r, column=6).value,
                "entregas": ws.cell(row=r, column=7).value,
                "gasolina": ws.cell(row=r, column=15).value,
                "almoco": ws.cell(row=r, column=16).value,
                "besteira": ws.cell(row=r, column=17).value,
                "carro": ws.cell(row=r, column=18).value,
            })
        df = pd.DataFrame(rows)
        for c in ["km", "pacotes", "entregas", "gasolina", "almoco", "besteira", "carro"]:
            if c not in df.columns:
                df[c] = None
        if df.empty:
            df = pd.DataFrame(columns=["data", "km", "pacotes", "entregas",
                                        "gasolina", "almoco", "besteira", "carro"])
        dados[mes] = df
    wb.close()
    return dados


def criar_vazio_para_mes(ano: int, mes_num: int):
    _, last_day = calendar.monthrange(ano, mes_num)
    dias = [datetime(ano, mes_num, d) for d in range(1, last_day + 1)]
    df = pd.DataFrame({
        "data": dias,
        "km": [None] * len(dias),
        "pacotes": [None] * len(dias),
        "entregas": [None] * len(dias),
        "gasolina": [None] * len(dias),
        "almoco": [None] * len(dias),
        "besteira": [None] * len(dias),
        "carro": [None] * len(dias),
    })
    return df


@st.cache_data(show_spinner=False)
def ler_planilha():
    try:
        return importar_excel_original()
    except Exception as e:
        st.toast(f"Erro ao importar Excel: {e}", icon="⚠️")
        return {}


def inicializar_dados():
    if "dados" not in st.session_state:
        dados_importados = ler_planilha()
        dados = {}
        for mes, num in MESES.items():
            if mes in dados_importados and not dados_importados[mes].empty:
                dados[mes] = dados_importados[mes].copy()
            else:
                dados[mes] = criar_vazio_para_mes(2026, num)
        st.session_state["dados"] = dados


def calcular_df(df):
    out = df.copy()
    for c in ["km", "pacotes", "entregas", "gasolina", "almoco", "besteira", "carro"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    out["dia_semana"] = out["data"].apply(dia_semana_br)
    out["is_sunday"] = out["data"].apply(lambda d: d.weekday() == 6 if pd.notna(d) else False)
    out["diaria"] = out.apply(lambda r: calc_diaria(r["km"], r["is_sunday"]), axis=1)
    out["pct_entrega"] = out.apply(lambda r: calc_pct(r["pacotes"], r["entregas"]), axis=1)
    out["bonus"] = out["pct_entrega"].apply(calc_bonus)
    diaria_num = pd.to_numeric(out["diaria"], errors="coerce").fillna(0)
    bonus_num = pd.to_numeric(out["bonus"], errors="coerce").fillna(0)
    out["total_ganho"] = diaria_num + bonus_num
    out["total_gasto"] = (
        pd.to_numeric(out["gasolina"], errors="coerce").fillna(0) +
        pd.to_numeric(out["almoco"], errors="coerce").fillna(0) +
        pd.to_numeric(out["besteira"], errors="coerce").fillna(0) +
        pd.to_numeric(out["carro"], errors="coerce").fillna(0)
    )
    out["lucro"] = out["total_ganho"] - out["total_gasto"]
    return out


def resumo(df_calc):
    n = len(df_calc)
    meio = min(15, n)
    df1 = df_calc.iloc[:meio]
    df2 = df_calc.iloc[meio:]

    pac1 = df1["pacotes"].sum()
    ent1 = df1["entregas"].sum()
    tax1 = ent1 / pac1 if pac1 else None
    dias1 = int(df1["km"].notna().sum())
    g1 = df1["total_ganho"].sum()
    b600_1 = 600 if (dias1 >= 13 and tax1 is not None and 0.98 <= tax1 <= 1.0) else 0
    t1 = g1 + b600_1

    pac2 = df2["pacotes"].sum()
    ent2 = df2["entregas"].sum()
    tax2 = ent2 / pac2 if pac2 else None
    dias2 = int(df2["km"].notna().sum())
    g2 = df2["total_ganho"].sum()
    b600_2 = 600 if (dias2 >= 13 and tax2 is not None and 0.98 <= tax2 <= 1.0) else 0
    t2 = g2 + b600_2

    total_g = t1 + t2
    total_s = df_calc["total_gasto"].sum()
    return {
        "ganho1": g1, "taxa1": tax1, "dias1": dias1, "bonus600_1": b600_1, "total1": t1,
        "ganho2": g2, "taxa2": tax2, "dias2": dias2, "bonus600_2": b600_2, "total2": t2,
        "total_ganho": total_g, "total_gasto": total_s, "total_lucro": total_g - total_s,
    }


def salvar_no_excel_original(dados_dict):
    wb = openpyxl.load_workbook("Planilha de Ganhos.xlsx", data_only=False)
    for sheet in wb.sheetnames:
        mes = sheet.strip().lower()
        if mes not in dados_dict:
            continue
        df = dados_dict[mes]
        ws = wb[sheet]
        # Cria um mapa data -> linha a partir do Excel
        data_para_linha = {}
        for r_excel in range(4, ws.max_row + 1):
            v_data = ws.cell(row=r_excel, column=2).value
            if isinstance(v_data, datetime):
                chave = v_data.strftime("%Y-%m-%d")
                data_para_linha[chave] = r_excel

        for _, row in df.iterrows():
            dt = row["data"]
            if pd.isna(dt):
                continue
            chave = dt.strftime("%Y-%m-%d")
            if chave not in data_para_linha:
                continue
            r = data_para_linha[chave]
            km = row["km"]
            ws.cell(row=r, column=4).value = int(km) if pd.notna(km) else None
            pac = row["pacotes"]
            ws.cell(row=r, column=6).value = int(pac) if pd.notna(pac) else None
            ent = row["entregas"]
            ws.cell(row=r, column=7).value = int(ent) if pd.notna(ent) else None
            gas = row["gasolina"]
            ws.cell(row=r, column=15).value = float(gas) if pd.notna(gas) else None
            alm = row["almoco"]
            ws.cell(row=r, column=16).value = float(alm) if pd.notna(alm) else None
            bes = row["besteira"]
            ws.cell(row=r, column=17).value = float(bes) if pd.notna(bes) else None
            car = row["carro"]
            ws.cell(row=r, column=18).value = float(car) if pd.notna(car) else None
    wb.save("Planilha de Ganhos.xlsx")
    wb.close()
    # Limpa cache para recarregar dados atualizados
    ler_planilha.clear()


# =================== MAIN ===================

inicializar_dados()

st.title("Planilha de Ganhos - AC2")

# Mes atual como padrao
mes_atual = datetime.now().month
mes_nomes = list(MESES.keys())
mes_index_padrao = 0
for i, (nome, num) in enumerate(MESES.items()):
    if num == mes_atual:
        mes_index_padrao = i
        break

mes_sel = st.selectbox("Mes", mes_nomes, index=mes_index_padrao)

tab_lanc, tab_res = st.tabs(["Lancamentos", "Resultado Calculado"])

with tab_lanc:
    st.subheader(f"Lancamentos - {mes_sel.capitalize()}/2026")

    # Monta lista de datas do mes (apenas dias do mes selecionado)
    df_mes = st.session_state["dados"][mes_sel].copy()
    mes_num = MESES[mes_sel]
    hoje = datetime.now()

    datas_opcoes = []
    idx_padrao = 0
    for i, row in df_mes.iterrows():
        dt = row["data"]
        if pd.notna(dt) and dt.month == mes_num:
            dia_sem = dia_semana_br(dt)
            label = f"{dt.strftime('%d/%m/%Y')} - {dia_sem}"
            datas_opcoes.append((i, label))
            if dt.day == hoje.day and mes_num == hoje.month:
                idx_padrao = len(datas_opcoes) - 1

    if not datas_opcoes:
        st.warning("Nenhuma data encontrada para este mes.")
        st.stop()

    sel_combo = st.selectbox(
        "Selecione a data",
        range(len(datas_opcoes)),
        index=idx_padrao,
        format_func=lambda i: datas_opcoes[i][1],
        key=f"sel_{mes_sel}"
    )

    sel_idx = datas_opcoes[sel_combo][0]

    # Pega valores atuais dessa linha
    linha = df_mes.iloc[sel_idx]

    with st.form(key=f"form_{mes_sel}_{sel_idx}"):
        c1, c2 = st.columns(2)
        with c1:
            km = st.number_input("Km Executado", min_value=0, value=int(linha["km"]) if pd.notna(linha["km"]) else 0, step=1)
            pacotes = st.number_input("Qtd. Pacotes", min_value=0, value=int(linha["pacotes"]) if pd.notna(linha["pacotes"]) else 0, step=1)
            entregas = st.number_input("Qtd. Entregas", min_value=0, value=int(linha["entregas"]) if pd.notna(linha["entregas"]) else 0, step=1)
        with c2:
            gasolina = st.number_input("Gasolina", min_value=0.0, value=float(linha["gasolina"]) if pd.notna(linha["gasolina"]) else 0.0, step=0.01)
            almoco = st.number_input("Almoço", min_value=0.0, value=float(linha["almoco"]) if pd.notna(linha["almoco"]) else 0.0, step=0.01)
            besteira = st.number_input("Besteira", min_value=0.0, value=float(linha["besteira"]) if pd.notna(linha["besteira"]) else 0.0, step=0.01)
            carro = st.number_input("Carro", min_value=0.0, value=float(linha["carro"]) if pd.notna(linha["carro"]) else 0.0, step=0.01)

        submitted = st.form_submit_button("Salvar")
        if submitted:
            st.session_state["dados"][mes_sel].at[sel_idx, "km"] = km if km > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "pacotes"] = pacotes if pacotes > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "entregas"] = entregas if entregas > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "gasolina"] = gasolina if gasolina > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "almoco"] = almoco if almoco > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "besteira"] = besteira if besteira > 0 else None
            st.session_state["dados"][mes_sel].at[sel_idx, "carro"] = carro if carro > 0 else None
            salvar_no_excel_original(st.session_state["dados"])
            st.success("Dados salvos na planilha!")
            st.rerun()

with tab_res:
    st.subheader("Resultado Calculado")
    calc = calcular_df(st.session_state["dados"][mes_sel].copy())

    # Resumo
    r = resumo(calc)

    # Cards superiores customizados
    st.markdown("---")
    col1, col2, col3, col4, col5 = st.columns(5)

    card_style = """
    <div style="background-color: {bg}; padding: 20px; border-radius: 12px; text-align: center; box-shadow: 2px 2px 8px rgba(0,0,0,0.1); margin-bottom: 10px;">
        <div style="font-size: 28px; margin-bottom: 8px;">{icon}</div>
        <div style="color: #666; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;">{label}</div>
        <div style="color: {fg}; font-size: 24px; font-weight: 700; margin-top: 4px;">{value}</div>
    </div>
    """

    with col1:
        st.markdown(card_style.format(
            bg="#FFF3E0", icon="📅", label="1ª Quinzena",
            fg="#E65100", value=f"R$ {r['total1']:,.0f}".replace(",", ".")
        ), unsafe_allow_html=True)

    with col2:
        st.markdown(card_style.format(
            bg="#E3F2FD", icon="📅", label="2ª Quinzena",
            fg="#1565C0", value=f"R$ {r['total2']:,.0f}".replace(",", ".")
        ), unsafe_allow_html=True)

    with col3:
        st.markdown(card_style.format(
            bg="#E8F5E9", icon="💰", label="Total de Ganhos",
            fg="#2E7D32", value=f"R$ {r['total_ganho']:,.0f}".replace(",", ".")
        ), unsafe_allow_html=True)

    with col4:
        st.markdown(card_style.format(
            bg="#FFEBEE", icon="💸", label="Total de Gastos",
            fg="#C62828", value=f"R$ {r['total_gasto']:,.2f}".replace(",", ".")
        ), unsafe_allow_html=True)

    with col5:
        lucro_bg = "#E8F5E9" if r['total_lucro'] >= 0 else "#FFEBEE"
        lucro_fg = "#2E7D32" if r['total_lucro'] >= 0 else "#C62828"
        st.markdown(card_style.format(
            bg=lucro_bg, icon="📊", label="Lucro Líquido",
            fg=lucro_fg, value=f"R$ {r['total_lucro']:,.2f}".replace(",", ".")
        ), unsafe_allow_html=True)

    # Tabela de detalhes
    st.markdown("---")

    # Formata para exibicao
    fmt = calc.copy()
    # Filtra apenas dias com dados
    fmt = fmt[fmt["km"].notna() & (fmt["km"] > 0)].reset_index(drop=True)

    # Substituir NaN/None em colunas de gasto para evitar "R$ nan"
    for col in ["gasolina", "almoco", "besteira", "carro"]:
        fmt[col] = fmt[col].fillna(0)

    def fmt_val(val, dec=0):
        if pd.isna(val):
            return ""
        return f"{val:.{dec}f}"

    def fmt_rs(val, dec=0):
        if pd.isna(val):
            return ""
        return f"R$ {val:,.{dec}f}".replace(",", "_").replace("_", ".").replace(".", ",", 1)

    # Monta tabela HTML customizada
    html = """
    <style>
    .tabela-ganhos { width: 100%; border-collapse: collapse; font-size: 13px; font-family: sans-serif; }
    .tabela-ganhos thead th {
        background-color: #1565C0;
        color: #ffffff;
        padding: 10px 6px;
        text-align: center;
        font-weight: 600;
        border: 1px solid #0d47a1;
    }
    .tabela-ganhos td {
        padding: 8px 6px;
        text-align: center;
        border-bottom: 1px solid #e0e0e0;
        border-right: 1px solid #e0e0e0;
    }
    .tabela-ganhos tr:nth-child(even) td { background-color: #f5f5f5; }
    .tabela-ganhos tr:hover td { background-color: #e3f2fd; }
    .tabela-ganhos .text-left { text-align: left; padding-left: 10px; }
    .tabela-ganhos .numero { font-weight: 600; white-space: nowrap; }
    .tabela-ganhos .ganho { background-color: #E8F5E9 !important; }
    .tabela-ganhos .gasto { background-color: #FFEBEE !important; }
    .tabela-ganhos .positivo { color: #2E7D32; font-weight: 700; }
    .tabela-ganhos .negativo { color: #C62828; font-weight: 700; }
    </style>
    <table class="tabela-ganhos">
    <thead>
    <tr>
        <th>Data</th>
        <th>Dia da Semana</th>
        <th>Km Executado</th>
        <th>Diária</th>
        <th>Qtd. Pacotes</th>
        <th>Qtd. Entregas</th>
        <th>% Entrega</th>
        <th>Bônus</th>
        <th>Total Ganhos</th>
        <th>Gasolina</th>
        <th>Almoço</th>
        <th>Besteira</th>
        <th>Carro</th>
        <th>Total Gastos</th>
        <th>Lucro Diário</th>
    </tr>
    </thead>
    <tbody>
    """

    for _, row in fmt.iterrows():
        lucro_cls = "positivo" if row["lucro"] >= 0 else "negativo"
        html += f"""<tr>
            <td class="text-left">{row['data'].strftime('%d/%m/%Y') if pd.notna(row['data']) else ''}</td>
            <td class="text-left">{row['dia_semana']}</td>
            <td class="numero">{int(row['km']) if pd.notna(row['km']) else ''}</td>
            <td class="numero">R$ {fmt_val(row['diaria'])}</td>
            <td class="numero">{int(row['pacotes']) if pd.notna(row['pacotes']) else ''}</td>
            <td class="numero">{int(row['entregas']) if pd.notna(row['entregas']) else ''}</td>
            <td class="numero">{fmt_val(row['pct_entrega']*100 if pd.notna(row['pct_entrega']) else 0, dec=0)}%</td>
            <td class="numero">R$ {fmt_val(row['bonus'])}</td>
            <td class="numero ganho">R$ {fmt_val(row['total_ganho'])}</td>
            <td class="numero">R$ {fmt_val(row['gasolina'], dec=2)}</td>
            <td class="numero">R$ {fmt_val(row['almoco'], dec=2)}</td>
            <td class="numero">R$ {fmt_val(row['besteira'], dec=2)}</td>
            <td class="numero">R$ {fmt_val(row['carro'], dec=2)}</td>
            <td class="numero gasto">R$ {fmt_val(row['total_gasto'], dec=2)}</td>
            <td class="numero {lucro_cls}">R$ {fmt_val(row['lucro'], dec=2)}</td>
        </tr>"""

    html += "</tbody></table>"

    st_html(html, height=600, scrolling=True)

st.divider()
if st.button("Exportar para Excel"):
    with pd.ExcelWriter("Planilha_de_Ganhos_Exportada.xlsx", engine="openpyxl") as writer:
        for mes, d in st.session_state["dados"].items():
            calcular_df(d).to_excel(writer, sheet_name=mes, index=False)
    st.success("Planilha exportada com sucesso!")
