import base64
import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================
# Configurações principais
# =========================
st.set_page_config(
    page_title="SUS - Quantidade e Valor Total",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "data" / "CIA014_SUS_Prova_2.xlsx"
LOGO_PATH = BASE_DIR / "assets" / "iesb_logo.png"

AUTHOR = "Made by Kaike Armond Costa"

REGION_COLORS = {
    "Centro-Oeste": "#4C9BF5",
    "Nordeste": "#FF8A3D",
    "Norte": "#53B6A7",
    "Sudeste": "#E36CC7",
    "Sul": "#9AC640",
}

# =========================
# CSS para aproximar das telas enviadas
# =========================
st.markdown(
    """
    <style>
        .stApp {
            background: #ffffff;
            color: #2b313a;
        }

        div[data-testid="stToolbar"],
        div[data-testid="stDecoration"],
        div[data-testid="stStatusWidget"],
        #MainMenu,
        footer,
        header {
            visibility: hidden;
            height: 0%;
        }

        .block-container {
            padding-top: 1.1rem;
            padding-bottom: 1.5rem;
            max-width: 1280px;
        }

        .report-title {
            text-align: center;
            font-size: 1.45rem;
            font-weight: 800;
            margin: 0 0 0.25rem 0;
            color: #292f3a;
        }

        .made-by {
            font-size: 0.92rem;
            font-weight: 700;
            color: #2f3540;
            margin-top: 0.15rem;
            margin-bottom: 1rem;
        }

        .logo-box {
            margin-top: 1.1rem;
        }

        .kpi-wrap {
            min-height: 420px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .kpi-card {
            padding: 1.2rem 1.8rem;
            border-radius: 18px;
            background: #ffffff;
        }

        .kpi-label {
            font-size: 2.05rem;
            line-height: 1.05;
            font-weight: 800;
            color: #4a515e;
            margin-bottom: 0.35rem;
        }

        .kpi-value {
            font-size: 4.9rem;
            line-height: 1;
            font-weight: 800;
            color: #418ce8;
            letter-spacing: -0.08rem;
        }

        .small-note {
            color: #687181;
            font-size: 0.88rem;
        }

        .section-subtitle {
            font-size: 1.25rem;
            font-weight: 800;
            color: #2b313a;
            margin-top: 0.25rem;
            margin-bottom: 0.1rem;
        }

        .page-footer {
            text-align: center;
            color: #111827;
            font-weight: 700;
            margin-top: 0.5rem;
            font-size: 0.84rem;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            justify-content: center;
        }

        .stTabs [data-baseweb="tab"] {
            background-color: #f4f6fb;
            border-radius: 12px 12px 0 0;
            padding: 10px 16px;
            font-weight: 700;
        }

        .stTabs [aria-selected="true"] {
            background-color: #e7f0fe;
            color: #2367bb;
        }

        div[data-testid="stMetricValue"] {
            color: #418ce8;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 12px;
        }

        .filter-panel {
            border: 1px solid #e8edf5;
            background: #fbfcff;
            border-radius: 14px;
            padding: 0.8rem 1rem;
            margin-bottom: 0.8rem;
        }

        @media (max-width: 900px) {
            .kpi-value {font-size: 3.2rem;}
            .kpi-label {font-size: 1.45rem;}
            .kpi-wrap {min-height: 260px;}
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# Funções auxiliares
# =========================
def decode_excel_escapes(value):
    """Corrige textos que vieram no padrão _x0020_."""
    if not isinstance(value, str):
        return value
    value = re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m.group(1), 16)), value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def compact_number(value):
    """Exibe valores como 440 mi, 21 bi etc."""
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

    if abs(number) >= 100 or float(number).is_integer():
        formatted = f"{number:,.0f}"
    else:
        formatted = f"{number:,.1f}"

    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{formatted} {suffix}".strip()


def br_number(value, decimals=2):
    if pd.isna(value):
        value = 0
    text = f"{float(value):,.{decimals}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def br_money(value):
    return f"R$ {br_number(value, 2)}"


def br_percent(value):
    return f"{br_number(value, 2)}%"


@st.cache_data(show_spinner="Carregando base SUS...")
def load_data():
    df = pd.read_excel(DATA_PATH, sheet_name="SUS_TABLE")

    text_cols = [
        "Codigo_Municipio_DV",
        "Codigo_Municipio",
        "Nome_Municipio",
        "Regiao_Codigo",
        "Regiao_Nome",
        "UF_Codigo",
        "UF",
        "UF_Nome",
        "Municipio_Capital",
        "Faixa_Populacao",
        "Faixa_Populacao_FPM",
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(decode_excel_escapes)

    numeric_cols = [
        "ano",
        "mes",
        "Numero_Habitantes_Censo_2022",
        "QTD_Total",
        "VL_Total",
        "LATITUDE",
        "LONGITUDE",
        "qtd_01",
        "qtd_02",
        "qtd_03",
        "qtd_04",
        "qtd_05",
        "qtd_06",
        "qtd_07",
        "qtd_08",
        "vl_02",
        "vl_03",
        "vl_04",
        "vl_05",
        "vl_06",
        "vl_07",
        "vl_08",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["QTD_Total"] = df["QTD_Total"].fillna(0)
    df["VL_Total"] = df["VL_Total"].fillna(0)
    df["mes_nome"] = df["mes"].map(
        {
            1: "Jan",
            2: "Fev",
            3: "Mar",
            4: "Abr",
            5: "Mai",
            6: "Jun",
            7: "Jul",
            8: "Ago",
            9: "Set",
            10: "Out",
            11: "Nov",
            12: "Dez",
        }
    )

    # Ordem usada no gráfico de faixa populacional.
    df["Faixa_Populacao"] = df["Faixa_Populacao"].fillna("Não informado")
    return df


def header(title):
    left, center, right = st.columns([1.1, 4.2, 1.1])
    with left:
        st.markdown(f"<div class='made-by'>{AUTHOR}</div>", unsafe_allow_html=True)
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=78)
    with center:
        st.markdown(f"<div class='report-title'>{title}</div>", unsafe_allow_html=True)
    with right:
        st.write("")


def apply_filters(df):
    with st.expander("Filtros do painel", expanded=False):
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            anos = sorted([int(x) for x in df["ano"].dropna().unique()])
            ano_sel = st.multiselect("Ano", anos, default=anos)

        with c2:
            meses = sorted([int(x) for x in df["mes"].dropna().unique()])
            mes_sel = st.multiselect("Mês", meses, default=meses)

        with c3:
            regioes = sorted(df["Regiao_Nome"].dropna().unique().tolist())
            regiao_sel = st.multiselect("Região", regioes, default=regioes)

        with c4:
            ufs = sorted(df["UF"].dropna().unique().tolist())
            uf_sel = st.multiselect("UF", ufs, default=ufs)

    filtered = df[
        df["ano"].isin(ano_sel)
        & df["mes"].isin(mes_sel)
        & df["Regiao_Nome"].isin(regiao_sel)
        & df["UF"].isin(uf_sel)
    ].copy()

    return filtered


def metric_total(df, metric):
    return float(df[metric].sum())


def uf_table(df):
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

    grouped["%Valor Gasto"] = np.where(
        total_valor > 0, grouped["Valor Total"] / total_valor * 100, 0
    )
    grouped["Valor medio dos procedimentos"] = np.where(
        grouped["Quantidade Total"] > 0,
        grouped["Valor Total"] / grouped["Quantidade Total"],
        0,
    )

    total_row = pd.DataFrame(
        [
            {
                "UF": "Total",
                "Quantidade Total": total_qtd,
                "Valor Total": total_valor,
                "%Valor Gasto": 100.0 if total_valor > 0 else 0,
                "Valor medio dos procedimentos": total_valor / total_qtd
                if total_qtd > 0
                else 0,
            }
        ]
    )

    return pd.concat([total_row, grouped], ignore_index=True)


def styled_uf_table(table):
    styled = table.style.format(
        {
            "Quantidade Total": lambda x: br_number(x, 2),
            "Valor Total": lambda x: br_number(x, 2),
            "%Valor Gasto": lambda x: br_percent(x),
            "Valor medio dos procedimentos": lambda x: br_number(x, 2),
        }
    )
    styled = styled.set_properties(
        **{
            "font-size": "12px",
            "text-align": "right",
            "border-color": "#e7eaf0",
        },
        subset=[
            "Quantidade Total",
            "Valor Total",
            "%Valor Gasto",
            "Valor medio dos procedimentos",
        ],
    )
    styled = styled.set_properties(
        **{"font-size": "12px", "font-weight": "700", "text-align": "left"},
        subset=["UF"],
    )
    styled = styled.set_table_styles(
        [
            {
                "selector": "thead th",
                "props": [
                    ("background-color", "#ffffff"),
                    ("color", "#303641"),
                    ("font-weight", "800"),
                    ("border-bottom", "1px solid #d8dde6"),
                    ("font-size", "12px"),
                ],
            },
            {
                "selector": "tbody tr:nth-child(even)",
                "props": [("background-color", "#f5f6f8")],
            },
            {
                "selector": "tbody tr:nth-child(odd)",
                "props": [("background-color", "#ffffff")],
            },
        ]
    )
    return styled


def page_footer(number):
    st.markdown(f"<div class='page-footer'>{number}</div>", unsafe_allow_html=True)


# =========================
# Carregamento
# =========================
data = load_data()
filtered_data = apply_filters(data)

if filtered_data.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")
    st.stop()

# =========================
# Abas
# =========================
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "Quantidade e Valor Total",
        "Total de Procedimentos e Gastos",
        "Mapa",
        "Procedimentos Interativos",
    ]
)

# =========================
# Página 1 - KPIs
# =========================
with tab1:
    header("Quantidade e Valor Total")

    qtd_total = metric_total(filtered_data, "QTD_Total")
    valor_total = metric_total(filtered_data, "VL_Total")

    st.markdown("<div class='kpi-wrap'>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([0.3, 1.2, 1.2])

    with c2:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Quantidade Total</div>
                <div class="kpi-value">{compact_number(qtd_total)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Valor Total</div>
                <div class="kpi-value">{compact_number(valor_total)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.caption(
        f"Base filtrada: {len(filtered_data):,} registros | "
        f"{filtered_data['UF'].nunique()} UFs | "
        f"{filtered_data['Codigo_Municipio'].nunique()} municípios."
    )
    page_footer(1)

# =========================
# Página 2 - Tabela
# =========================
with tab2:
    header("Total de Procedimentos e Gastos")
    tabela = uf_table(filtered_data)

    st.dataframe(
        styled_uf_table(tabela),
        hide_index=True,
        use_container_width=True,
        height=610,
    )

    page_footer(2)

# =========================
# Página 3 - Mapa
# =========================
with tab3:
    header("Mapa")

    left, main = st.columns([0.08, 0.92])
    with main:
        st.markdown(
            "<div class='section-subtitle'>Valor total de Procedimentos por Região</div>",
            unsafe_allow_html=True,
        )

        mapa = (
            filtered_data.groupby(
                ["Codigo_Municipio", "Nome_Municipio", "UF", "UF_Nome", "Regiao_Nome"],
                as_index=False,
            )
            .agg(
                QTD_Total=("QTD_Total", "sum"),
                VL_Total=("VL_Total", "sum"),
                LATITUDE=("LATITUDE", "first"),
                LONGITUDE=("LONGITUDE", "first"),
            )
            .dropna(subset=["LATITUDE", "LONGITUDE"])
        )

        # Evita pontos sem valor e deixa o mapa mais limpo.
        mapa = mapa[(mapa["QTD_Total"] > 0) | (mapa["VL_Total"] > 0)].copy()

        mapa["Valor formatado"] = mapa["VL_Total"].apply(br_money)
        mapa["Quantidade formatada"] = mapa["QTD_Total"].apply(lambda x: br_number(x, 0))

        fig_map = px.scatter_mapbox(
            mapa,
            lat="LATITUDE",
            lon="LONGITUDE",
            size="VL_Total",
            color="Regiao_Nome",
            color_discrete_map=REGION_COLORS,
            hover_name="Nome_Municipio",
            custom_data=[
                "UF",
                "Regiao_Nome",
                "Quantidade formatada",
                "Valor formatado",
            ],
            zoom=3.1,
            height=620,
            size_max=20,
            labels={
                "Regiao_Nome": "Regiao_Nome",
                "VL_Total": "Valor Total",
            },
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

        fig_map.update_layout(
            mapbox_style="carto-positron",
            mapbox_center={"lat": -14.2, "lon": -52.8},
            margin=dict(l=0, r=0, t=8, b=0),
            legend=dict(
                orientation="h",
                y=-0.05,
                x=0.5,
                xanchor="center",
                title_text="Regiao_Nome",
            ),
        )

        st.plotly_chart(fig_map, use_container_width=True, config={"displayModeBar": False})

    page_footer(3)

# =========================
# Página 4 - Procedimentos Interativos
# =========================
with tab4:
    header("Procedimentos Interativos")

    metric_label = st.radio(
        "Indicador",
        ["Quantidade Total", "Valor Total"],
        horizontal=True,
        label_visibility="collapsed",
    )
    metric_col = "QTD_Total" if metric_label == "Quantidade Total" else "VL_Total"

    col1, col2, col3 = st.columns([1.05, 1.15, 1.2], gap="large")

    with col1:
        st.markdown(
            f"<div class='section-subtitle'>{metric_label} por Região</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='small-note'>{metric_label}</div>",
            unsafe_allow_html=True,
        )

        regiao = (
            filtered_data.groupby("Regiao_Nome", as_index=False)[metric_col]
            .sum()
            .sort_values(metric_col, ascending=False)
        )

        fig_donut = px.pie(
            regiao,
            values=metric_col,
            names="Regiao_Nome",
            hole=0.58,
            color="Regiao_Nome",
            color_discrete_map=REGION_COLORS,
        )
        fig_donut.update_traces(textinfo="none", hovertemplate="<b>%{label}</b><br>%{value:,.2f}<extra></extra>")
        fig_donut.update_layout(
            height=460,
            margin=dict(l=5, r=5, t=20, b=40),
            legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center", title_text="Regiao_Nome"),
            annotations=[
                dict(
                    text=compact_number(regiao[metric_col].sum()),
                    x=0.5,
                    y=0.5,
                    font_size=30,
                    showarrow=False,
                    font_color="#5a6372",
                )
            ],
        )
        st.plotly_chart(fig_donut, use_container_width=True, config={"displayModeBar": False})

    with col2:
        st.markdown(
            f"<div class='section-subtitle'>{metric_label} de Procedimento por Faixa de Habitante</div>",
            unsafe_allow_html=True,
        )

        faixa = (
            filtered_data.groupby("Faixa_Populacao", as_index=False)[metric_col]
            .sum()
            .sort_values(metric_col, ascending=False)
        )

        fig_tree = px.treemap(
            faixa,
            path=["Faixa_Populacao"],
            values=metric_col,
            color=metric_col,
            color_continuous_scale="Blues",
        )
        fig_tree.update_traces(
            texttemplate="<b>%{label}</b>",
            textfont_size=13,
            hovertemplate="<b>%{label}</b><br>%{value:,.2f}<extra></extra>",
        )
        fig_tree.update_layout(
            height=460,
            margin=dict(l=0, r=0, t=20, b=0),
            coloraxis_showscale=True,
        )
        st.plotly_chart(fig_tree, use_container_width=True, config={"displayModeBar": False})

    with col3:
        st.markdown(
            f"<div class='section-subtitle'>{metric_label} por Unidade Federativa</div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='small-note'>UF</div>", unsafe_allow_html=True)

        uf = (
            filtered_data.groupby("UF", as_index=False)[metric_col]
            .sum()
            .sort_values(metric_col, ascending=True)
        )

        fig_bar = px.bar(
            uf,
            x=metric_col,
            y="UF",
            orientation="h",
            color_discrete_sequence=["#3d95f6"],
            labels={metric_col: f"{metric_label} ({'milhões' if metric_col == 'QTD_Total' else 'R$'})", "UF": "UF"},
            height=500,
        )
        fig_bar.update_traces(hovertemplate="<b>%{y}</b><br>%{x:,.2f}<extra></extra>")
        fig_bar.update_layout(
            margin=dict(l=0, r=0, t=12, b=30),
            xaxis=dict(showgrid=True, gridcolor="#dce2ec"),
            yaxis=dict(categoryorder="total ascending"),
            showlegend=False,
        )
        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

    page_footer(4)
