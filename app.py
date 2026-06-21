from pathlib import Path
from io import BytesIO
import re
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =====================================================
# CONFIGURAÇÃO
# =====================================================
st.set_page_config(
    page_title="Dashboard SUS",
    page_icon="📊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
FILE_NAME = "CIA014_SUS_Prova_2.xlsx"
SHEET_NAME = "SUS_TABLE"

REGION_COLORS = {
    "Centro-Oeste": "#3B82F6",
    "Nordeste": "#F97316",
    "Norte": "#14B8A6",
    "Sudeste": "#E879F9",
    "Sul": "#84CC16",
}

st.markdown(
    """
    <style>
        .block-container {padding-top: 1.3rem; padding-bottom: 1.5rem;}
        .main-title {text-align:center; font-size: 1.7rem; font-weight: 800; margin-bottom: 0.2rem;}
        .author {font-size: 0.9rem; font-weight: 700; color: #374151; margin-bottom: 0.9rem;}
        .kpi-box {padding: 3.5rem 1rem; text-align: center; border-radius: 18px; background: #FFFFFF;}
        .kpi-label {font-size: 2rem; font-weight: 800; color: #4B5563; margin-bottom: 0.5rem;}
        .kpi-value {font-size: 5rem; font-weight: 900; line-height: 1; color: #3B82F6;}
        .sub-title {font-size: 1.25rem; font-weight: 800; color: #374151; margin-bottom: 0.2rem;}
        .footer-number {text-align:center; font-weight: 700; margin-top: 0.8rem;}
        div[data-testid="stMetricValue"] {color: #3B82F6;}
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================
def decode_excel_text(value):
    """Corrige textos importados com códigos do tipo _x0020_."""
    if not isinstance(value, str):
        return value
    value = re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m.group(1), 16)), value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def format_number_br(value, decimals=2):
    if pd.isna(value):
        value = 0
    text = f"{float(value):,.{decimals}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def format_money_br(value):
    return "R$ " + format_number_br(value, 2)


def compact_number(value):
    if pd.isna(value):
        return "0"
    value = float(value)
    abs_value = abs(value)

    if abs_value >= 1_000_000_000:
        number = value / 1_000_000_000
        suffix = "bi"
    elif abs_value >= 1_000_000:
        number = value / 1_000_000
        suffix = "mi"
    elif abs_value >= 1_000:
        number = value / 1_000
        suffix = "mil"
    else:
        number = value
        suffix = ""

    if abs(number) >= 100 or number.is_integer():
        formatted = format_number_br(number, 0)
    else:
        formatted = format_number_br(number, 1)

    return f"{formatted} {suffix}".strip()


def find_local_excel():
    """Procura a planilha na raiz, na pasta data ou em subpastas do projeto."""
    possible_paths = [
        BASE_DIR / FILE_NAME,
        BASE_DIR / "data" / FILE_NAME,
    ]

    for path in possible_paths:
        if path.exists():
            return path

    matches = list(BASE_DIR.rglob(FILE_NAME))
    if matches:
        return matches[0]

    return None


def read_excel_source(source):
    """Lê a aba SUS_TABLE. Se ela não existir, lê a primeira aba."""
    try:
        return pd.read_excel(source, sheet_name=SHEET_NAME)
    except ValueError:
        return pd.read_excel(source, sheet_name=0)


def prepare_data(df):
    text_cols = [
        "Codigo_Municipio_DV", "Codigo_Municipio", "Nome_Municipio",
        "Regiao_Codigo", "Regiao_Nome", "UF_Codigo", "UF", "UF_Nome",
        "Municipio_Capital", "Faixa_Populacao", "Faixa_Populacao_FPM",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(decode_excel_text)

    numeric_cols = [
        "ano", "mes", "Numero_Habitantes_Censo_2022", "QTD_Total", "VL_Total",
        "LATITUDE", "LONGITUDE", "qtd_01", "qtd_02", "qtd_03", "qtd_04",
        "qtd_05", "qtd_06", "qtd_07", "qtd_08", "vl_02", "vl_03", "vl_04",
        "vl_05", "vl_06", "vl_07", "vl_08",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Garante que as principais colunas existam mesmo se a planilha mudar um pouco.
    if "QTD_Total" not in df.columns:
        qtd_cols = [c for c in df.columns if c.lower().startswith("qtd_")]
        df["QTD_Total"] = df[qtd_cols].sum(axis=1) if qtd_cols else 0

    if "VL_Total" not in df.columns:
        vl_cols = [c for c in df.columns if c.lower().startswith("vl_")]
        df["VL_Total"] = df[vl_cols].sum(axis=1) if vl_cols else 0

    df["QTD_Total"] = df["QTD_Total"].fillna(0)
    df["VL_Total"] = df["VL_Total"].fillna(0)

    for col in ["Regiao_Nome", "UF", "UF_Nome", "Nome_Municipio", "Faixa_Populacao"]:
        if col in df.columns:
            df[col] = df[col].fillna("Não informado")

    return df


@st.cache_data(show_spinner="Carregando a planilha...")
def load_data_from_path(path_as_text):
    df = read_excel_source(path_as_text)
    return prepare_data(df)


@st.cache_data(show_spinner="Carregando a planilha enviada...")
def load_data_from_upload(file_bytes):
    df = read_excel_source(BytesIO(file_bytes))
    return prepare_data(df)


def page_header(title):
    st.markdown(f"<div class='main-title'>{title}</div>", unsafe_allow_html=True)
    st.markdown("<div class='author'>Made by Kaike Armond Costa</div>", unsafe_allow_html=True)


def footer(number):
    st.markdown(f"<div class='footer-number'>{number}</div>", unsafe_allow_html=True)


def build_uf_table(df):
    grouped = (
        df.groupby("UF", as_index=False)
        .agg(
            **{
                "Quantidade Total": ("QTD_Total", "sum"),
                "Valor Total": ("VL_Total", "sum"),
            }
        )
        .sort_values("UF")
    )

    total_qtd = grouped["Quantidade Total"].sum()
    total_valor = grouped["Valor Total"].sum()

    grouped["% Valor Gasto"] = np.where(total_valor > 0, grouped["Valor Total"] / total_valor * 100, 0)
    grouped["Valor médio dos procedimentos"] = np.where(
        grouped["Quantidade Total"] > 0,
        grouped["Valor Total"] / grouped["Quantidade Total"],
        0,
    )

    total = pd.DataFrame([
        {
            "UF": "Total",
            "Quantidade Total": total_qtd,
            "Valor Total": total_valor,
            "% Valor Gasto": 100 if total_valor > 0 else 0,
            "Valor médio dos procedimentos": total_valor / total_qtd if total_qtd > 0 else 0,
        }
    ])

    return pd.concat([total, grouped], ignore_index=True)


def show_filters(df):
    st.sidebar.title("Filtros")

    filtered = df.copy()

    if "ano" in filtered.columns:
        anos = sorted([int(x) for x in filtered["ano"].dropna().unique()])
        ano_sel = st.sidebar.multiselect("Ano", anos, default=anos)
        filtered = filtered[filtered["ano"].isin(ano_sel)]

    if "mes" in filtered.columns:
        meses = sorted([int(x) for x in filtered["mes"].dropna().unique()])
        mes_sel = st.sidebar.multiselect("Mês", meses, default=meses)
        filtered = filtered[filtered["mes"].isin(mes_sel)]

    if "Regiao_Nome" in filtered.columns:
        regioes = sorted(filtered["Regiao_Nome"].dropna().unique().tolist())
        regiao_sel = st.sidebar.multiselect("Região", regioes, default=regioes)
        filtered = filtered[filtered["Regiao_Nome"].isin(regiao_sel)]

    if "UF" in filtered.columns:
        ufs = sorted(filtered["UF"].dropna().unique().tolist())
        uf_sel = st.sidebar.multiselect("UF", ufs, default=ufs)
        filtered = filtered[filtered["UF"].isin(uf_sel)]

    return filtered


# =====================================================
# CARREGAMENTO DA BASE
# =====================================================
local_excel = find_local_excel()

if local_excel is None:
    st.warning(
        "A planilha não foi encontrada na pasta do projeto. "
        "Envie o arquivo CIA014_SUS_Prova_2.xlsx abaixo para carregar o dashboard."
    )
    uploaded_file = st.file_uploader("Enviar planilha Excel", type=["xlsx"])
    if uploaded_file is None:
        st.stop()
    data = load_data_from_upload(uploaded_file.getvalue())
else:
    data = load_data_from_path(str(local_excel))

filtered_data = show_filters(data)

if filtered_data.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")
    st.stop()

# =====================================================
# ABAS DO DASHBOARD
# =====================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "Quantidade e Valor Total",
    "Total de Procedimentos e Gastos",
    "Mapa",
    "Procedimentos Interativos",
])

# =====================================================
# TELA 1 - KPIs
# =====================================================
with tab1:
    page_header("Quantidade e Valor Total")

    qtd_total = filtered_data["QTD_Total"].sum()
    valor_total = filtered_data["VL_Total"].sum()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-label'>Quantidade Total</div>
                <div class='kpi-value'>{compact_number(qtd_total)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-label'>Valor Total</div>
                <div class='kpi-value'>{compact_number(valor_total)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.caption(
        f"Registros filtrados: {len(filtered_data):,} | "
        f"UFs: {filtered_data['UF'].nunique() if 'UF' in filtered_data.columns else 0} | "
        f"Municípios: {filtered_data['Codigo_Municipio'].nunique() if 'Codigo_Municipio' in filtered_data.columns else 0}"
    )
    footer(1)

# =====================================================
# TELA 2 - TABELA
# =====================================================
with tab2:
    page_header("Total de Procedimentos e Gastos")
    table = build_uf_table(filtered_data)

    table_view = table.copy()
    table_view["Quantidade Total"] = table_view["Quantidade Total"].apply(lambda x: format_number_br(x, 2))
    table_view["Valor Total"] = table_view["Valor Total"].apply(lambda x: format_number_br(x, 2))
    table_view["% Valor Gasto"] = table_view["% Valor Gasto"].apply(lambda x: format_number_br(x, 2) + "%")
    table_view["Valor médio dos procedimentos"] = table_view["Valor médio dos procedimentos"].apply(lambda x: format_number_br(x, 2))

    st.dataframe(table_view, hide_index=True, use_container_width=True, height=620)
    footer(2)

# =====================================================
# TELA 3 - MAPA
# =====================================================
with tab3:
    page_header("Mapa")
    st.markdown("<div class='sub-title'>Valor total de Procedimentos por Região</div>", unsafe_allow_html=True)

    needed_cols = {"Codigo_Municipio", "Nome_Municipio", "UF", "Regiao_Nome", "LATITUDE", "LONGITUDE"}
    if needed_cols.issubset(filtered_data.columns):
        map_data = (
            filtered_data.groupby(["Codigo_Municipio", "Nome_Municipio", "UF", "Regiao_Nome"], as_index=False)
            .agg(
                QTD_Total=("QTD_Total", "sum"),
                VL_Total=("VL_Total", "sum"),
                LATITUDE=("LATITUDE", "first"),
                LONGITUDE=("LONGITUDE", "first"),
            )
            .dropna(subset=["LATITUDE", "LONGITUDE"])
        )
        map_data = map_data[(map_data["QTD_Total"] > 0) | (map_data["VL_Total"] > 0)].copy()

        if map_data.empty:
            st.info("Não há coordenadas disponíveis para montar o mapa com os filtros atuais.")
        else:
            map_data["Valor"] = map_data["VL_Total"].apply(format_money_br)
            map_data["Quantidade"] = map_data["QTD_Total"].apply(lambda x: format_number_br(x, 0))

            fig_map = px.scatter_mapbox(
                map_data,
                lat="LATITUDE",
                lon="LONGITUDE",
                size="VL_Total",
                color="Regiao_Nome",
                color_discrete_map=REGION_COLORS,
                hover_name="Nome_Municipio",
                custom_data=["UF", "Regiao_Nome", "Quantidade", "Valor"],
                size_max=18,
                zoom=3.1,
                height=620,
            )
            fig_map.update_layout(
                mapbox_style="open-street-map",
                mapbox_center={"lat": -14.2, "lon": -52.8},
                margin=dict(l=0, r=0, t=10, b=0),
                legend=dict(orientation="h", y=-0.06, x=0.5, xanchor="center", title_text="Região"),
            )
            fig_map.update_traces(
                marker=dict(opacity=0.72),
                hovertemplate=(
                    "<b>%{hovertext}</b><br>"
                    "UF: %{customdata[0]}<br>"
                    "Região: %{customdata[1]}<br>"
                    "Quantidade Total: %{customdata[2]}<br>"
                    "Valor Total: %{customdata[3]}"
                    "<extra></extra>"
                ),
            )
            st.plotly_chart(fig_map, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info("A planilha não possui todas as colunas necessárias para o mapa.")

    footer(3)

# =====================================================
# TELA 4 - GRÁFICOS
# =====================================================
with tab4:
    page_header("Procedimentos Interativos")

    selected_metric = st.radio(
        "Indicador",
        ["Quantidade Total", "Valor Total"],
        horizontal=True,
    )
    metric_col = "QTD_Total" if selected_metric == "Quantidade Total" else "VL_Total"

    col1, col2, col3 = st.columns([1.1, 1.2, 1.2], gap="large")

    with col1:
        st.markdown(f"<div class='sub-title'>{selected_metric} por Região</div>", unsafe_allow_html=True)
        region_data = (
            filtered_data.groupby("Regiao_Nome", as_index=False)[metric_col]
            .sum()
            .sort_values(metric_col, ascending=False)
        )
        fig_donut = px.pie(
            region_data,
            values=metric_col,
            names="Regiao_Nome",
            hole=0.58,
            color="Regiao_Nome",
            color_discrete_map=REGION_COLORS,
        )
        fig_donut.update_traces(textinfo="none")
        fig_donut.update_layout(
            height=470,
            margin=dict(l=0, r=0, t=20, b=40),
            legend=dict(orientation="h", y=-0.06, x=0.5, xanchor="center", title_text="Região"),
            annotations=[dict(text=compact_number(region_data[metric_col].sum()), x=0.5, y=0.5, showarrow=False, font_size=31, font_color="#4B5563")],
        )
        st.plotly_chart(fig_donut, use_container_width=True, config={"displayModeBar": False})

    with col2:
        st.markdown(f"<div class='sub-title'>{selected_metric} por Faixa de Habitante</div>", unsafe_allow_html=True)
        if "Faixa_Populacao" in filtered_data.columns:
            faixa_data = (
                filtered_data.groupby("Faixa_Populacao", as_index=False)[metric_col]
                .sum()
                .sort_values(metric_col, ascending=False)
            )
            fig_tree = px.treemap(
                faixa_data,
                path=["Faixa_Populacao"],
                values=metric_col,
                color=metric_col,
                color_continuous_scale="Blues",
            )
            fig_tree.update_traces(texttemplate="<b>%{label}</b>", textfont_size=13)
            fig_tree.update_layout(height=470, margin=dict(l=0, r=0, t=20, b=0), coloraxis_showscale=False)
            st.plotly_chart(fig_tree, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Coluna Faixa_Populacao não encontrada.")

    with col3:
        st.markdown(f"<div class='sub-title'>{selected_metric} por Unidade Federativa</div>", unsafe_allow_html=True)
        uf_data = (
            filtered_data.groupby("UF", as_index=False)[metric_col]
            .sum()
            .sort_values(metric_col, ascending=True)
        )
        fig_bar = px.bar(
            uf_data,
            x=metric_col,
            y="UF",
            orientation="h",
            labels={metric_col: selected_metric, "UF": "UF"},
            height=500,
        )
        fig_bar.update_layout(
            showlegend=False,
            margin=dict(l=0, r=0, t=20, b=30),
            xaxis=dict(showgrid=True, gridcolor="#E5E7EB"),
            yaxis=dict(categoryorder="total ascending"),
        )
        fig_bar.update_traces(marker_color="#3B82F6")
        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

    footer(4)
