st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 1.5rem;
        }

        .main-title {
            text-align: center;
            font-size: 1.7rem;
            font-weight: 800;
            margin-bottom: 0.2rem;
        }

        .author {
            font-size: 0.9rem;
            font-weight: 700;
            color: #374151;
            margin-bottom: 0.9rem;
        }

        .kpi-box {
            padding: 3.5rem 1rem;
            text-align: center;
            border-radius: 18px;
            background: #FFFFFF;
        }

        .kpi-label {
            font-size: 2rem;
            font-weight: 800;
            color: #4B5563;
            margin-bottom: 0.5rem;
        }

        .kpi-value {
            font-size: 5rem;
            font-weight: 900;
            line-height: 1;
            color: #3B82F6;
        }

        .sub-title {
            font-size: 1.25rem;
            font-weight: 800;
            color: #374151;
            margin-bottom: 0.2rem;
        }

        .footer-number {
            text-align: center;
            font-weight: 700;
            margin-top: 0.8rem;
        }

        div[data-testid="stMetricValue"] {
            color: #3B82F6;
        }

        /* ===== MELHORIA DAS ABAS ===== */
        div[data-testid="stTabs"] {
            margin-top: 0.3rem;
        }

        div[data-testid="stTabs"] [role="tablist"] {
            gap: 0.5rem;
            overflow-x: auto;
            padding-bottom: 0.3rem;
        }

        div[data-testid="stTabs"] [role="tab"] {
            height: auto;
            min-height: 52px;
            padding: 0.8rem 1.1rem;
            white-space: nowrap;
            border-radius: 12px 12px 0 0;
            border: 1px solid #334155;
            background-color: #0F172A;
            color: #E5E7EB;
        }

        div[data-testid="stTabs"] [role="tab"] p {
            font-size: 0.97rem;
            font-weight: 700;
            margin: 0;
        }

        div[data-testid="stTabs"] [role="tab"]:hover {
            background-color: #1E293B;
            color: #FFFFFF;
        }

        div[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
            background-color: #FFFFFF;
            border-color: #3B82F6;
        }

        div[data-testid="stTabs"] [role="tab"][aria-selected="true"] p {
            color: #111827 !important;
        }

        div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
            background-color: #3B82F6 !important;
            height: 3px !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)
