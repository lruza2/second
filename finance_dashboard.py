import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


DEFAULT_EXCEL_PATH = Path(r"G:\Meu Drive\Anotacoes\Budget e Impostos\2025 ldn\Gastos totais.xlsx")

DEFAULT_FIXED_CATEGORIES = [
    "Aluguel",
    "Home Box",
    "Transporte",
    "Internet",
    "Celular",
    "Casa",
    "Academia",
    "Mercado",
    "Restaurante",
]

PAYSLIP_NUMERIC_COLUMNS = [
    "ReferencePayRateAnnual",
    "ReferenceSal",
    "SalaryAdj",
    "WFHStipend",
    "Relocation",
    "Bonus",
    "PensionSalarySacrifice",
    "PensionBonusSacrifice",
    "PreSacrificeEarnings",
    "GrossPay",
    "MedTaxable",
    "DentalTaxable",
    "TravelTaxable",
    "TotalTaxableBenefits",
    "PAYETaxableGrossCurrent",
    "NIC",
    "PAYE",
    "Deductions",
    "NetPay",
    "TotalPensionSacrifice",
    "TakeHomeRateVsGross",
    "YTDTaxableGrossExclBenefits",
    "YTDTaxableBenefit",
    "YTDPAYETaxableGross",
    "YTDNITaxableGross",
    "YTDTotalTax",
    "YTDTotalNI",
    "YTDPension",
    "YTDPensionBonus",
    "YTDPensionMatch",
]

PAYSLIP_FIELD_HELP = {
    "ReferencePayRateAnnual": "Salário anual de referência informado no payslip.",
    "ReferenceSal": "Parcela mensal do salário de referência.",
    "SalaryAdj": "Ajuste de salário lançado no mês.",
    "WFHStipend": "Ajuda de custo de trabalho em casa.",
    "Relocation": "Pagamento relacionado a relocation.",
    "Bonus": "Bônus lançado no mês.",
    "PensionSalarySacrifice": "Sacrifício de salário destinado à pension.",
    "PensionBonusSacrifice": "Parcela do bônus destinada à pension.",
    "GrossPay": "Pagamento bruto após os sacrifícios de pension registrados no arquivo.",
    "TotalTaxableBenefits": "Total de benefícios tributáveis informados no payslip.",
    "PAYE": "Imposto PAYE do período. Valor negativo indica ajuste ou devolução no próprio registro.",
    "NIC": "National Insurance do período.",
    "Deductions": "Deduções em dinheiro registradas no payslip.",
    "NetPay": "Valor líquido recebido.",
    "TaxCode": "Tax code usado no período.",
    "TaxBasis": "Base de cálculo indicada no payslip.",
    "PayrollTaxYear": "Ano fiscal de payroll informado no arquivo.",
    "TaxPeriod": "Período fiscal dentro do ano de payroll.",
    "YTDPAYETaxableGross": "Bruto tributável acumulado no ano fiscal.",
    "YTDTotalTax": "PAYE acumulado no ano fiscal.",
    "YTDTotalNI": "NIC acumulado no ano fiscal.",
    "YTDPension": "Pension acumulada no ano fiscal.",
    "YTDPensionBonus": "Pension relacionada a bônus acumulada no ano fiscal.",
    "YTDPensionMatch": "Employer pension match acumulado, quando disponível.",
}


st.set_page_config(page_title="Finanças pessoais", page_icon="💷", layout="wide")


@st.cache_data
def load_excel_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    sheets = pd.read_excel(
        Path(path),
        sheet_name=["Mov_sem_internas", "Payslip"],
        engine="openpyxl",
    )
    return (
        sheets.get("Mov_sem_internas", pd.DataFrame()).copy(),
        sheets.get("Payslip", pd.DataFrame()).copy(),
    )


@st.cache_data
def load_uploaded_excel(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    sheets = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=["Mov_sem_internas", "Payslip"],
        engine="openpyxl",
    )
    return (
        sheets.get("Mov_sem_internas", pd.DataFrame()).copy(),
        sheets.get("Payslip", pd.DataFrame()).copy(),
    )


@st.cache_data
def load_uploaded_csvs(mov_bytes: bytes, payslip_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    mov = pd.read_csv(io.BytesIO(mov_bytes))
    payslip = pd.read_csv(io.BytesIO(payslip_bytes))
    return mov, payslip


def parse_mixed_dates(series: pd.Series) -> pd.Series:
    """Lê formatos como YYYY MM DD e YYYY MM DD HH MM SS na mesma coluna."""
    try:
        return pd.to_datetime(series, format="mixed", errors="coerce")
    except (TypeError, ValueError):
        return pd.to_datetime(series, errors="coerce")


@st.cache_data
def clean_mov_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    required = {"Date", "Amount"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Mov_sem_internas sem colunas obrigatórias: {sorted(missing)}")

    df["Date"] = parse_mixed_dates(df["Date"])
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    if "Direction" in df.columns:
        df = df[df["Direction"].astype(str).str.upper().eq("DEBIT")].copy()

    df = df.dropna(subset=["Date", "Amount"]).copy()
    df["Gasto"] = df["Amount"].abs()
    df["MonthKey"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    df["MonthLabel"] = df["Date"].dt.strftime("%Y %m")
    df["Weekday"] = df["Date"].dt.day_name()
    df["Day"] = df["Date"].dt.day

    for col in ["Categoria", "Description", "TxnType", "Source"]:
        if col not in df.columns:
            df[col] = "Não informado"
        else:
            df[col] = df[col].fillna("Não informado").astype(str)

    return df.sort_values("Date").reset_index(drop=True)


@st.cache_data
def clean_payslip_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "Date" not in df.columns:
        raise ValueError("Payslip sem coluna Date")

    df["Date"] = parse_mixed_dates(df["Date"])
    df = df.dropna(subset=["Date"]).copy()

    for column in PAYSLIP_NUMERIC_COLUMNS:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")

    for column in ["ReferenceSal", "SalaryAdj", "WFHStipend", "Relocation", "Bonus"]:
        if column not in df.columns:
            df[column] = 0.0

    for column in ["PensionSalarySacrifice", "PensionBonusSacrifice"]:
        if column not in df.columns:
            df[column] = 0.0

    for column in ["GrossPay", "NetPay", "PAYE", "NIC", "Deductions", "TotalTaxableBenefits"]:
        if column not in df.columns:
            df[column] = np.nan

    salary_pension = df["PensionSalarySacrifice"].fillna(0).abs()
    bonus_pension = df["PensionBonusSacrifice"].fillna(0).abs()
    pension_derived = salary_pension + bonus_pension

    if "TotalPensionSacrifice" in df.columns:
        total_pension_reported = pd.to_numeric(df["TotalPensionSacrifice"], errors="coerce").abs()
        df["PensionTotal"] = total_pension_reported.where(total_pension_reported.notna(), pension_derived)
    else:
        df["PensionTotal"] = pension_derived

    df["CashBeforePension"] = df["GrossPay"].fillna(0) + df["PensionTotal"].fillna(0)
    df["TaxAndNI"] = df["PAYE"].fillna(0) + df["NIC"].fillna(0)
    df["TaxableGrossDerived"] = df["GrossPay"].fillna(0) + df["TotalTaxableBenefits"].fillna(0)
    df["TakeHomeRateDerived"] = np.where(
        df["CashBeforePension"].abs() > 0,
        df["NetPay"] / df["CashBeforePension"],
        np.nan,
    )
    df["PensionRateDerived"] = np.where(
        df["CashBeforePension"].abs() > 0,
        df["PensionTotal"] / df["CashBeforePension"],
        np.nan,
    )
    df["TaxRateDerived"] = np.where(
        df["TaxableGrossDerived"].abs() > 0,
        df["TaxAndNI"] / df["TaxableGrossDerived"],
        np.nan,
    )
    df["MonthKey"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    df["MonthLabel"] = df["Date"].dt.strftime("%Y %m")

    return df.sort_values("Date").reset_index(drop=True)



MONTH_NAMES_PT = {
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


def add_month_display(df: pd.DataFrame, date_col: str = "MonthKey", label_col: str = "Mes") -> pd.DataFrame:
    """Cria rótulos mensais categóricos para evitar deslocamento visual de datas no Plotly."""
    out = df.copy()
    if date_col not in out.columns:
        out[label_col] = ""
        return out

    dates = pd.to_datetime(out[date_col], errors="coerce")
    out[label_col] = [
        f"{MONTH_NAMES_PT.get(d.month, '')} {d.year}" if pd.notna(d) else ""
        for d in dates
    ]
    return out

def money(value) -> str:
    if pd.isna(value):
        return "n/a"
    return f"£{value:,.0f}"


def money2(value) -> str:
    if pd.isna(value):
        return "n/a"
    return f"£{value:,.2f}"


def pct(value, decimals: int = 1) -> str:
    if pd.isna(value) or np.isinf(value):
        return "n/a"
    return f"{value:.{decimals}f}%"


def delta_pct(current, previous) -> str | None:
    if previous is None or pd.isna(previous) or previous == 0 or pd.isna(current):
        return None
    return f"{((current / previous) - 1) * 100:+.1f}%"


def filter_dates(df: pd.DataFrame, start_date, end_date) -> pd.DataFrame:
    start = pd.Timestamp(start_date)
    end = pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    return df[df["Date"].between(start, end)].copy()


def build_monthly_spend(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["MonthKey", "Spend"])
    return (
        df.groupby("MonthKey", as_index=False)["Gasto"]
        .sum()
        .rename(columns={"Gasto": "Spend"})
        .sort_values("MonthKey")
    )


def build_monthly_income(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["MonthKey", "NetPay", "GrossPay", "PensionTotal", "PAYE", "NIC"])

    columns = ["NetPay", "GrossPay", "PensionTotal", "PAYE", "NIC", "Bonus", "Relocation"]
    available = [c for c in columns if c in df.columns]
    return df.groupby("MonthKey", as_index=False)[available].sum().sort_values("MonthKey")


def build_monthly_comparison(mov: pd.DataFrame, payslip: pd.DataFrame) -> pd.DataFrame:
    spend = build_monthly_spend(mov)
    income = build_monthly_income(payslip)
    merged = spend.merge(income, on="MonthKey", how="outer").sort_values("MonthKey")

    for col in ["Spend", "NetPay", "GrossPay", "PensionTotal", "PAYE", "NIC", "Bonus", "Relocation"]:
        if col not in merged.columns:
            merged[col] = np.nan

    merged["Balance"] = merged["NetPay"] - merged["Spend"]
    merged["SavingsRate"] = np.where(
        merged["NetPay"] > 0,
        merged["Balance"] / merged["NetPay"] * 100,
        np.nan,
    )
    merged["SpendRate"] = np.where(
        merged["NetPay"] > 0,
        merged["Spend"] / merged["NetPay"] * 100,
        np.nan,
    )
    merged["CumulativeBalance"] = merged["Balance"].fillna(0).cumsum()
    return merged


def category_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Categoria", "Gasto", "Share"])
    out = (
        df.groupby("Categoria", as_index=False)["Gasto"]
        .sum()
        .sort_values("Gasto", ascending=False)
    )
    total = out["Gasto"].sum()
    out["Share"] = np.where(total > 0, out["Gasto"] / total * 100, 0)
    return out


def classify_costs(df: pd.DataFrame, fixed_categories: list[str]) -> pd.DataFrame:
    out = df.copy()
    out["CostType"] = np.where(out["Categoria"].isin(fixed_categories), "Fixo", "Variável")
    return out


def make_waterfall(row: pd.Series) -> go.Figure:
    before = float(row.get("CashBeforePension", 0) or 0)
    pension = float(row.get("PensionTotal", 0) or 0)
    gross = float(row.get("GrossPay", 0) or 0)
    paye = float(row.get("PAYE", 0) or 0)
    nic = float(row.get("NIC", 0) or 0)
    net = float(row.get("NetPay", 0) or 0)

    fig = go.Figure(
        go.Waterfall(
            measure=["absolute", "relative", "total", "relative", "relative", "total"],
            x=["Antes da pension", "Pension", "Gross pay", "PAYE", "NIC", "Net pay"],
            y=[before, -pension, 0, -paye, -nic, 0],
            text=[money(before), money(pension), money(gross), money(paye), money(nic), money(net)],
            textposition="outside",
            connector={"line": {"width": 1}},
        )
    )
    fig.update_layout(
        title="Do pagamento antes da pension até o líquido",
        yaxis_title="GBP",
        template="plotly_white",
        showlegend=False,
        margin=dict(l=20, r=20, t=70, b=20),
    )
    return fig


def style_figure(fig: go.Figure, height: int = 430) -> go.Figure:
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=20, r=20, t=65, b=30),
        hovermode="x unified",
        legend_title_text="",
    )
    return fig


def render_data_quality(raw_mov: pd.DataFrame, clean_mov: pd.DataFrame, raw_pay: pd.DataFrame, clean_pay: pd.DataFrame):
    with st.expander("Qualidade dos dados"):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Linhas gastos", f"{len(clean_mov):,}", delta=f"{len(clean_mov) - len(raw_mov):,} removidas")
        c2.metric("Linhas payslip", f"{len(clean_pay):,}", delta=f"{len(clean_pay) - len(raw_pay):,} removidas")
        c3.metric("Categorias", clean_mov["Categoria"].nunique())
        c4.metric("Meses com payslip", clean_pay["MonthKey"].nunique())
        st.caption("Linhas podem ser removidas quando Date ou Amount não podem ser interpretados.")


def render_summary(mov: pd.DataFrame, payslip: pd.DataFrame, fixed_categories: list[str]):
    comparison = build_monthly_comparison(mov, payslip)
    overlap = comparison.dropna(subset=["Spend", "NetPay"]).copy()

    total_spend = mov["Gasto"].sum()
    total_income = payslip["NetPay"].sum() if "NetPay" in payslip.columns else np.nan
    total_balance = total_income - total_spend if pd.notna(total_income) else np.nan
    savings_rate = total_balance / total_income * 100 if pd.notna(total_income) and total_income > 0 else np.nan
    avg_monthly_spend = build_monthly_spend(mov)["Spend"].mean() if not mov.empty else 0
    pension_total = payslip["PensionTotal"].sum() if "PensionTotal" in payslip.columns else np.nan

    st.subheader("Resumo")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Saídas", money(total_spend))
    c2.metric("Net pay", money(total_income))
    c3.metric("Saldo", money(total_balance))
    c4.metric("Taxa de poupança", pct(savings_rate))
    c5.metric("Gasto médio mensal", money(avg_monthly_spend))
    c6.metric("Pension", money(pension_total))

    if not overlap.empty:
        overlap_chart = add_month_display(overlap)
        month_order = overlap_chart["Mes"].tolist()
        chart = overlap_chart.melt(
            id_vars=["MonthKey", "Mes"],
            value_vars=["NetPay", "Spend", "Balance"],
            var_name="Serie",
            value_name="GBP",
        )
        fig = px.bar(
            chart,
            x="Mes",
            y="GBP",
            color="Serie",
            barmode="group",
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "GBP": "GBP", "Serie": ""},
            title="Renda líquida, gastos e saldo por mês",
        )
        st.plotly_chart(style_figure(fig), use_container_width=True)

        left, right = st.columns([1.5, 1])
        with left:
            fig_rate = px.line(
                overlap_chart,
                x="Mes",
                y="SavingsRate",
                markers=True,
                category_orders={"Mes": month_order},
                labels={"Mes": "Mês", "SavingsRate": "Taxa de poupança (%)"},
                title="Taxa de poupança mensal",
            )
            fig_rate.add_hline(y=0, line_width=1)
            st.plotly_chart(style_figure(fig_rate, 360), use_container_width=True)

        with right:
            table = overlap[["MonthKey", "NetPay", "Spend", "Balance", "SavingsRate"]].copy()
            table["MonthKey"] = table["MonthKey"].dt.strftime("%Y %m")
            st.dataframe(
                table.style.format(
                    {
                        "NetPay": "£{:,.0f}",
                        "Spend": "£{:,.0f}",
                        "Balance": "£{:,.0f}",
                        "SavingsRate": "{:.1f}%",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    st.markdown("#### Leituras rápidas")
    insights = []

    cat = category_summary(mov)
    if not cat.empty:
        top = cat.iloc[0]
        insights.append(f"Maior categoria no período: **{top['Categoria']}**, com {money(top['Gasto'])} e {top['Share']:.1f}% das saídas.")

    monthly = build_monthly_spend(mov)
    if not monthly.empty:
        highest = monthly.loc[monthly["Spend"].idxmax()]
        insights.append(f"Mês de maior gasto: **{highest['MonthKey'].strftime('%Y %m')}**, com {money(highest['Spend'])}.")
        if len(monthly) >= 2:
            latest = monthly.iloc[-1]
            previous = monthly.iloc[-2]
            change = delta_pct(latest["Spend"], previous["Spend"])
            if change:
                insights.append(f"Último mês: {money(latest['Spend'])}, variação de **{change}** contra o mês anterior.")

    classified = classify_costs(mov, fixed_categories)
    if not classified.empty:
        cost_split = classified.groupby("CostType")["Gasto"].sum()
        fixed = cost_split.get("Fixo", 0)
        share = fixed / total_spend * 100 if total_spend else 0
        insights.append(f"Categorias marcadas como fixas representam **{share:.1f}%** das saídas no período.")

    if not payslip.empty:
        latest_pay = payslip.iloc[-1]
        insights.append(
            f"Último payslip: net pay de **{money(latest_pay.get('NetPay'))}**, PAYE de {money(latest_pay.get('PAYE'))}, NIC de {money(latest_pay.get('NIC'))} e pension de {money(latest_pay.get('PensionTotal'))}."
        )

    for item in insights:
        st.write("• " + item)


def render_spending(mov: pd.DataFrame, fixed_categories: list[str], top_n: int):
    st.subheader("Gastos")

    categories = sorted(mov["Categoria"].dropna().unique())
    selected = st.multiselect(
        "Categorias para analisar",
        options=categories,
        default=categories,
        key="spend_categories",
    )
    selected_mov = mov[mov["Categoria"].isin(selected)].copy() if selected else mov.iloc[0:0].copy()
    classified = classify_costs(selected_mov, fixed_categories)

    total = selected_mov["Gasto"].sum()
    count = len(selected_mov)
    median_txn = selected_mov["Gasto"].median() if count else 0
    monthly = build_monthly_spend(selected_mov)
    monthly_avg = monthly["Spend"].mean() if not monthly.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total", money(total))
    c2.metric("Média mensal", money(monthly_avg))
    c3.metric("Mediana por transação", money2(median_txn))
    c4.metric("Transações", f"{count:,}")

    cats = category_summary(selected_mov).head(top_n)
    left, right = st.columns(2)

    with left:
        fig_cat = px.bar(
            cats.sort_values("Gasto"),
            x="Gasto",
            y="Categoria",
            orientation="h",
            text="Gasto",
            labels={"Gasto": "GBP", "Categoria": "Categoria"},
            title=f"Top {top_n} categorias",
        )
        fig_cat.update_traces(texttemplate="£%{text:,.0f}", textposition="outside")
        st.plotly_chart(style_figure(fig_cat), use_container_width=True)

    with right:
        split = classified.groupby("CostType", as_index=False)["Gasto"].sum()
        fig_split = px.pie(
            split,
            names="CostType",
            values="Gasto",
            hole=0.55,
            title="Fixos e variáveis",
        )
        st.plotly_chart(style_figure(fig_split), use_container_width=True)

    if not classified.empty:
        monthly_type = (
            classified.groupby(["MonthKey", "CostType"], as_index=False)["Gasto"].sum()
        )
        monthly_type = add_month_display(monthly_type)
        month_order = monthly_type.sort_values("MonthKey")["Mes"].drop_duplicates().tolist()
        fig_month_type = px.bar(
            monthly_type,
            x="Mes",
            y="Gasto",
            color="CostType",
            barmode="stack",
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "Gasto": "GBP", "CostType": ""},
            title="Fixos e variáveis por mês",
        )
        st.plotly_chart(style_figure(fig_month_type), use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        weekday = selected_mov.groupby("Weekday", as_index=False)["Gasto"].sum()
        weekday["Weekday"] = pd.Categorical(weekday["Weekday"], categories=weekday_order, ordered=True)
        weekday = weekday.sort_values("Weekday")
        fig_weekday = px.bar(
            weekday,
            x="Weekday",
            y="Gasto",
            text="Gasto",
            labels={"Weekday": "Dia da semana", "Gasto": "GBP"},
            title="Gasto por dia da semana",
        )
        fig_weekday.update_traces(texttemplate="£%{text:,.0f}", textposition="outside")
        st.plotly_chart(style_figure(fig_weekday, 360), use_container_width=True)

    with col_b:
        source = selected_mov.groupby("Source", as_index=False)["Gasto"].sum().sort_values("Gasto", ascending=False)
        fig_source = px.bar(
            source,
            x="Source",
            y="Gasto",
            text="Gasto",
            labels={"Source": "Conta", "Gasto": "GBP"},
            title="Gastos por origem",
        )
        fig_source.update_traces(texttemplate="£%{text:,.0f}", textposition="outside")
        st.plotly_chart(style_figure(fig_source, 360), use_container_width=True)

    st.markdown("#### Detalhe diário por categoria")
    detail_categories = st.multiselect(
        "Escolha uma ou mais categorias",
        options=categories,
        default=categories[:1] if categories else [],
        key="daily_categories",
    )

    detail = selected_mov[selected_mov["Categoria"].isin(detail_categories)].copy()
    if not detail.empty:
        detail_total = detail["Gasto"].sum()
        detail_count = len(detail)
        detail_days = detail["Date"].dt.normalize().nunique()
        detail_avg = detail["Gasto"].mean() if detail_count else 0.0

        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Total selecionado", money(detail_total))
        d2.metric("Transações", f"{detail_count:,}")
        d3.metric("Dias com gasto", f"{detail_days:,}")
        d4.metric("Média por transação", money2(detail_avg))

        # Calcula o maior intervalo sem lançamentos nas categorias selecionadas.
        period_start = selected_mov["Date"].min().normalize()
        period_end = selected_mov["Date"].max().normalize()
        calendar_days = pd.date_range(period_start, period_end, freq="D")
        category_activity = pd.Series(0, index=calendar_days, dtype="int64")
        observed_days = pd.DatetimeIndex(detail["Date"].dt.normalize().unique())
        category_activity.loc[category_activity.index.intersection(observed_days)] = 1

        zero_mask = category_activity.eq(0).to_numpy()
        longest_start = None
        longest_end = None
        longest_days = 0
        run_start = None

        for idx, is_zero in enumerate(zero_mask):
            if is_zero and run_start is None:
                run_start = idx
            is_last = idx == len(zero_mask) - 1
            if run_start is not None and ((not is_zero) or is_last):
                run_end = idx if is_zero and is_last else idx - 1
                run_length = run_end - run_start + 1
                if run_length > longest_days:
                    longest_days = run_length
                    longest_start = calendar_days[run_start]
                    longest_end = calendar_days[run_end]
                run_start = None

        if longest_days >= 3 and longest_start is not None:
            all_txns_in_gap = selected_mov[
                selected_mov["Date"].dt.normalize().between(longest_start, longest_end)
            ]
            st.info(
                f"Maior intervalo sem lançamentos nas categorias selecionadas: "
                f"{longest_days} dias, de {longest_start.strftime('%d/%m/%Y')} "
                f"a {longest_end.strftime('%d/%m/%Y')}. "
                f"Nesse mesmo intervalo existem {len(all_txns_in_gap)} transações "
                f"em todas as categorias, somando {money(all_txns_in_gap['Gasto'].sum())}."
            )

        st.markdown("##### Transações das categorias selecionadas")
        text_filter = st.text_input(
            "Filtrar descrição dentro da seleção",
            key="spend_text_filter",
        ).strip()

        detail_table = detail.copy()
        if text_filter:
            detail_table = detail_table[
                detail_table["Description"].str.contains(text_filter, case=False, na=False)
            ]

        detail_table = detail_table[
            ["Date", "Description", "Categoria", "Gasto", "TxnType", "Source"]
        ].sort_values(["Date", "Gasto"], ascending=[False, False])

        st.dataframe(
            detail_table.style.format({"Gasto": "£{:,.2f}", "Date": "{:%d/%m/%Y}"}),
            use_container_width=True,
            hide_index=True,
            height=360,
        )

        csv_detail = detail_table.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar transações filtradas em CSV",
            data=csv_detail,
            file_name="transacoes_categorias_selecionadas.csv",
            mime="text/csv",
            key="download_daily_category_detail",
        )

        st.markdown("##### Gasto diário")
        daily = (
            detail.assign(DateOnly=detail["Date"].dt.normalize())
            .groupby("DateOnly", as_index=False)["Gasto"]
            .sum()
            .rename(columns={"DateOnly": "Date"})
        )

        show_all_days = st.checkbox(
            "Incluir dias sem gasto na série diária",
            value=True,
            key="show_zero_spend_days",
        )

        if show_all_days:
            daily = (
                pd.DataFrame({"Date": calendar_days})
                .merge(daily, on="Date", how="left")
                .fillna({"Gasto": 0.0})
            )
            fig_daily = px.line(
                daily,
                x="Date",
                y="Gasto",
                markers=True,
                labels={"Date": "Data", "Gasto": "GBP"},
                title="Gasto diário nas categorias selecionadas",
            )
            fig_daily.update_traces(
                hovertemplate="%{x|%d/%m/%Y}<br>£%{y:,.2f}<extra></extra>"
            )
        else:
            fig_daily = px.bar(
                daily,
                x="Date",
                y="Gasto",
                text="Gasto",
                labels={"Date": "Data", "Gasto": "GBP"},
                title="Gasto diário nas categorias selecionadas",
            )
            fig_daily.update_traces(texttemplate="£%{text:,.0f}", textposition="outside")

        st.plotly_chart(style_figure(fig_daily, 380), use_container_width=True)

        with st.expander("Checar cobertura dos dados no período"):
            all_daily = (
                selected_mov.assign(DateOnly=selected_mov["Date"].dt.normalize())
                .groupby("DateOnly", as_index=False)
                .agg(Transacoes=("Gasto", "size"), Gasto=("Gasto", "sum"))
            )
            all_daily = (
                pd.DataFrame({"DateOnly": calendar_days})
                .merge(all_daily, on="DateOnly", how="left")
                .fillna({"Transacoes": 0, "Gasto": 0.0})
            )
            all_daily["Transacoes"] = all_daily["Transacoes"].astype(int)

            coverage = selected_mov.groupby("Source", as_index=False).agg(
                Primeira_data=("Date", "min"),
                Ultima_data=("Date", "max"),
                Transacoes=("Gasto", "size"),
                Gasto=("Gasto", "sum"),
            )

            st.write(
                f"No período há {len(selected_mov):,} transações em todas as categorias, "
                f"distribuídas por {selected_mov['Date'].dt.normalize().nunique():,} dias com movimento."
            )
            st.dataframe(
                coverage.style.format(
                    {
                        "Primeira_data": "{:%d/%m/%Y}",
                        "Ultima_data": "{:%d/%m/%Y}",
                        "Gasto": "£{:,.2f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

            fig_coverage = px.bar(
                all_daily,
                x="DateOnly",
                y="Transacoes",
                labels={"DateOnly": "Data", "Transacoes": "Número de transações"},
                title="Quantidade de transações por dia em todas as categorias",
            )
            st.plotly_chart(style_figure(fig_coverage, 300), use_container_width=True)
    elif detail_categories:
        st.info("Não há lançamentos nas categorias selecionadas dentro dos filtros atuais.")

    st.markdown("#### Maiores transações")
    largest = selected_mov.nlargest(20, "Gasto")[["Date", "Description", "Categoria", "Gasto", "Source"]]
    st.dataframe(largest.style.format({"Gasto": "£{:,.2f}"}), use_container_width=True, hide_index=True)


def render_payslip(payslip: pd.DataFrame):
    st.subheader("Payslip")

    if payslip.empty:
        st.info("Não há payslips no período selecionado.")
        return

    options = payslip["Date"].dt.strftime("%Y %m %d").tolist()
    selected_label = st.selectbox("Payslip para inspeção", options=options, index=len(options) - 1)
    selected_date = pd.to_datetime(selected_label, format="%Y %m %d")
    row = payslip.loc[payslip["Date"].dt.normalize() == selected_date.normalize()].iloc[-1]

    st.markdown("#### Pagamento selecionado")
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Antes da pension", money(row.get("CashBeforePension")))
    c2.metric("Pension", money(row.get("PensionTotal")))
    c3.metric("Gross pay", money(row.get("GrossPay")))
    c4.metric("PAYE", money(row.get("PAYE")))
    c5.metric("NIC", money(row.get("NIC")))
    c6.metric("Net pay", money(row.get("NetPay")))
    c7.metric("Benefícios tributáveis", money(row.get("TotalTaxableBenefits")))

    meta1, meta2, meta3, meta4 = st.columns(4)
    meta1.metric("Tax code", str(row.get("TaxCode", "n/a")))
    meta2.metric("Tax basis", str(row.get("TaxBasis", "n/a")))
    meta3.metric("Ano fiscal", str(row.get("PayrollTaxYear", "n/a")))
    meta4.metric("Tax period", str(row.get("TaxPeriod", "n/a")))

    st.plotly_chart(make_waterfall(row), use_container_width=True)

    c_left, c_right = st.columns(2)

    with c_left:
        comp_cols = ["ReferenceSal", "SalaryAdj", "WFHStipend", "Relocation", "Bonus"]
        comp = add_month_display(payslip[["MonthKey"] + comp_cols].copy())
        month_order = comp.sort_values("MonthKey")["Mes"].drop_duplicates().tolist()
        comp_long = comp.melt(["MonthKey", "Mes"], var_name="Componente", value_name="GBP")
        comp_long = comp_long[comp_long["GBP"].fillna(0) != 0]
        fig_comp = px.bar(
            comp_long,
            x="Mes",
            y="GBP",
            color="Componente",
            barmode="stack",
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "GBP": "GBP", "Componente": ""},
            title="Componentes da remuneração antes da pension",
        )
        st.plotly_chart(style_figure(fig_comp), use_container_width=True)

    with c_right:
        deductions = add_month_display(payslip[["MonthKey", "PAYE", "NIC", "PensionTotal"]].copy())
        deductions_long = deductions.melt(["MonthKey", "Mes"], var_name="Componente", value_name="GBP")
        fig_ded = px.bar(
            deductions_long,
            x="Mes",
            y="GBP",
            color="Componente",
            barmode="group",
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "GBP": "GBP", "Componente": ""},
            title="PAYE, NIC e pension por mês",
        )
        st.plotly_chart(style_figure(fig_ded), use_container_width=True)

    st.markdown("#### Tax code e salário de referência")
    timeline_cols = [
        "Date",
        "PayrollTaxYear",
        "TaxPeriod",
        "TaxCode",
        "TaxBasis",
        "ReferencePayRateAnnual",
        "GrossPay",
        "PAYE",
        "NIC",
        "NetPay",
        "PensionTotal",
    ]
    timeline_cols = [c for c in timeline_cols if c in payslip.columns]
    timeline = payslip[timeline_cols].copy().sort_values("Date", ascending=False)
    formatters = {c: "£{:,.0f}" for c in ["ReferencePayRateAnnual", "GrossPay", "PAYE", "NIC", "NetPay", "PensionTotal"] if c in timeline.columns}
    st.dataframe(timeline.style.format(formatters), use_container_width=True, hide_index=True)

    ytd_cols = [
        "PayrollTaxYear",
        "Date",
        "YTDPAYETaxableGross",
        "YTDTotalTax",
        "YTDTotalNI",
        "YTDPension",
        "YTDPensionBonus",
        "YTDPensionMatch",
    ]
    available_ytd = [c for c in ytd_cols if c in payslip.columns]
    if "PayrollTaxYear" in available_ytd:
        latest_ytd = (
            payslip.sort_values("Date")
            .groupby("PayrollTaxYear", as_index=False)
            .tail(1)[available_ytd]
            .sort_values("Date", ascending=False)
        )
        st.markdown("#### Acumulado por ano fiscal")
        ytd_formatters = {c: "£{:,.0f}" for c in available_ytd if c.startswith("YTD")}
        st.dataframe(latest_ytd.style.format(ytd_formatters), use_container_width=True, hide_index=True)

    st.markdown("#### Todos os campos do payslip selecionado")
    field_rows = []
    for field, value in row.items():
        if field in ["MonthKey", "MonthLabel"]:
            continue
        if isinstance(value, (float, np.floating)) and pd.notna(value):
            rendered = money2(value) if field in PAYSLIP_NUMERIC_COLUMNS or field in {"PensionTotal", "CashBeforePension", "TaxAndNI", "TaxableGrossDerived"} else str(value)
        elif isinstance(value, pd.Timestamp):
            rendered = value.strftime("%Y %m %d")
        else:
            rendered = "n/a" if pd.isna(value) else str(value)
        field_rows.append(
            {
                "Campo": field,
                "Valor": rendered,
                "Leitura": PAYSLIP_FIELD_HELP.get(field, "Campo presente no arquivo de origem."),
            }
        )
    st.dataframe(pd.DataFrame(field_rows), use_container_width=True, hide_index=True)


def render_income_vs_spend(mov: pd.DataFrame, payslip: pd.DataFrame):
    st.subheader("Renda e gastos")

    comparison = build_monthly_comparison(mov, payslip)
    comparison = comparison.dropna(subset=["Spend", "NetPay"]).copy()

    if comparison.empty:
        st.info("Não há meses com gastos e payslip ao mesmo tempo no período selecionado.")
        return

    total_net = comparison["NetPay"].sum()
    total_spend = comparison["Spend"].sum()
    total_balance = comparison["Balance"].sum()
    rate = total_balance / total_net * 100 if total_net > 0 else np.nan

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Net pay comparável", money(total_net))
    c2.metric("Gastos comparáveis", money(total_spend))
    c3.metric("Saldo comparável", money(total_balance))
    c4.metric("Taxa de poupança", pct(rate))

    comparison_chart = add_month_display(comparison)
    month_order = comparison_chart["Mes"].tolist()
    bar_data = comparison_chart.melt(
        id_vars=["MonthKey", "Mes"],
        value_vars=["NetPay", "Spend"],
        var_name="Serie",
        value_name="GBP",
    )
    fig = px.bar(
        bar_data,
        x="Mes",
        y="GBP",
        color="Serie",
        barmode="group",
        category_orders={"Mes": month_order},
        labels={"Mes": "Mês", "GBP": "GBP", "Serie": ""},
        title="Net pay e gastos por mês",
    )
    st.plotly_chart(style_figure(fig), use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        fig_balance = px.bar(
            comparison_chart,
            x="Mes",
            y="Balance",
            text="Balance",
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "Balance": "Saldo (GBP)"},
            title="Saldo mensal",
        )
        fig_balance.update_traces(texttemplate="£%{text:,.0f}", textposition="outside")
        fig_balance.add_hline(y=0, line_width=1)
        st.plotly_chart(style_figure(fig_balance, 380), use_container_width=True)

    with col_b:
        fig_spend_rate = px.line(
            comparison_chart,
            x="Mes",
            y="SpendRate",
            markers=True,
            category_orders={"Mes": month_order},
            labels={"Mes": "Mês", "SpendRate": "Gasto sobre net pay (%)"},
            title="Percentual do net pay consumido",
        )
        fig_spend_rate.add_hline(y=100, line_width=1)
        st.plotly_chart(style_figure(fig_spend_rate, 380), use_container_width=True)

    fig_cum = px.area(
        comparison_chart,
        x="Mes",
        y="CumulativeBalance",
        category_orders={"Mes": month_order},
        labels={"Mes": "Mês", "CumulativeBalance": "Saldo acumulado (GBP)"},
        title="Saldo acumulado no período comparável",
    )
    st.plotly_chart(style_figure(fig_cum, 380), use_container_width=True)

    events = comparison[(comparison["Bonus"].fillna(0).abs() > 0) | (comparison["Relocation"].fillna(0).abs() > 0)].copy()
    if not events.empty:
        st.markdown("#### Meses com remuneração fora do salário mensal")
        event_table = events[["MonthKey", "NetPay", "Bonus", "Relocation", "PensionTotal", "Balance"]].copy()
        event_table["MonthKey"] = event_table["MonthKey"].dt.strftime("%Y %m")
        st.dataframe(
            event_table.style.format({c: "£{:,.0f}" for c in ["NetPay", "Bonus", "Relocation", "PensionTotal", "Balance"]}),
            use_container_width=True,
            hide_index=True,
        )
        st.caption("Esses meses merecem leitura separada porque bônus, relocation ou pension podem distorcer a comparação com um mês salarial comum.")

    table = comparison[["MonthKey", "NetPay", "Spend", "Balance", "SavingsRate", "SpendRate"]].copy()
    table["MonthKey"] = table["MonthKey"].dt.strftime("%Y %m")
    st.markdown("#### Tabela mensal")
    st.dataframe(
        table.style.format(
            {
                "NetPay": "£{:,.0f}",
                "Spend": "£{:,.0f}",
                "Balance": "£{:,.0f}",
                "SavingsRate": "{:.1f}%",
                "SpendRate": "{:.1f}%",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


def render_data_explorer(mov: pd.DataFrame, payslip: pd.DataFrame):
    st.subheader("Dados")

    tab_mov, tab_pay = st.tabs(["Movimentos", "Payslip"])

    with tab_mov:
        text_filter = st.text_input("Buscar descrição", key="raw_mov_search").strip()
        display_mov = mov.copy()
        if text_filter:
            display_mov = display_mov[
                display_mov["Description"].str.contains(text_filter, case=False, na=False)
            ]
        st.dataframe(display_mov, use_container_width=True, hide_index=True)
        st.download_button(
            "Baixar gastos filtrados em CSV",
            data=display_mov.to_csv(index=False).encode("utf8"),
            file_name="gastos_filtrados.csv",
            mime="text/csv",
        )

    with tab_pay:
        st.dataframe(payslip, use_container_width=True, hide_index=True)
        st.download_button(
            "Baixar payslip filtrado em CSV",
            data=payslip.to_csv(index=False).encode("utf8"),
            file_name="payslip_filtrado.csv",
            mime="text/csv",
        )


def get_data_from_ui() -> tuple[pd.DataFrame, pd.DataFrame] | None:
    st.sidebar.header("Fonte dos dados")
    source = st.sidebar.radio(
        "Como carregar",
        ["Excel local", "Enviar Excel", "Enviar dois CSVs"],
        index=0,
    )

    try:
        if source == "Excel local":
            path = st.sidebar.text_input("Caminho do Excel", value=str(DEFAULT_EXCEL_PATH))
            if not Path(path).exists():
                st.sidebar.warning("Arquivo não encontrado. Ajuste o caminho ou use upload.")
                return None
            return load_excel_data(path)

        if source == "Enviar Excel":
            uploaded = st.sidebar.file_uploader("Arquivo Excel", type=["xlsx", "xlsm"], key="excel_upload")
            if uploaded is None:
                return None
            return load_uploaded_excel(uploaded.getvalue())

        mov_file = st.sidebar.file_uploader("CSV Mov_sem_internas", type=["csv"], key="mov_csv")
        pay_file = st.sidebar.file_uploader("CSV Payslip", type=["csv"], key="pay_csv")
        if mov_file is None or pay_file is None:
            return None
        return load_uploaded_csvs(mov_file.getvalue(), pay_file.getvalue())

    except Exception as exc:
        st.error(f"Erro ao carregar os dados: {exc}")
        return None


def main():
    st.title("💷 Finanças pessoais")
    st.caption("Gastos, payslip, pension, impostos e comparação entre renda líquida e saídas.")

    raw = get_data_from_ui()
    if raw is None:
        st.info("Selecione uma fonte de dados na barra lateral para abrir o dashboard.")
        st.stop()

    raw_mov, raw_pay = raw

    try:
        mov = clean_mov_data(raw_mov)
        payslip = clean_payslip_data(raw_pay)
    except Exception as exc:
        st.error(f"Erro na limpeza dos dados: {exc}")
        st.stop()

    if mov.empty:
        st.error("Não há gastos válidos após a limpeza dos dados.")
        st.stop()

    all_dates = pd.concat([mov["Date"], payslip["Date"]], ignore_index=True).dropna()
    min_date = all_dates.min().date()
    max_date = all_dates.max().date()

    st.sidebar.header("Filtros")
    date_range = st.sidebar.date_input(
        "Período",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )
    if len(date_range) != 2:
        st.sidebar.warning("Selecione data inicial e final.")
        st.stop()

    start_date, end_date = date_range
    mov_filtered = filter_dates(mov, start_date, end_date)
    pay_filtered = filter_dates(payslip, start_date, end_date)

    available_categories = sorted(mov["Categoria"].dropna().unique())
    default_fixed = [c for c in DEFAULT_FIXED_CATEGORIES if c in available_categories]
    fixed_categories = st.sidebar.multiselect(
        "Categorias tratadas como fixas",
        options=available_categories,
        default=default_fixed,
    )
    comparison_exclusions = st.sidebar.multiselect(
        "Retirar só da comparação com renda",
        options=available_categories,
        default=[],
        help="Útil para caução, transferências ou outras saídas que você não quer tratar como consumo mensal.",
    )
    top_n = st.sidebar.slider("Quantidade no ranking", min_value=5, max_value=25, value=12)

    mov_comparison = mov_filtered[
        ~mov_filtered["Categoria"].isin(comparison_exclusions)
    ].copy()

    render_data_quality(raw_mov, mov, raw_pay, payslip)

    tabs = st.tabs(["Resumo", "Gastos", "Payslip", "Renda e gastos", "Dados"])

    with tabs[0]:
        render_summary(mov_comparison, pay_filtered, fixed_categories)

    with tabs[1]:
        render_spending(mov_filtered, fixed_categories, top_n)

    with tabs[2]:
        render_payslip(pay_filtered)

    with tabs[3]:
        render_income_vs_spend(mov_comparison, pay_filtered)

    with tabs[4]:
        render_data_explorer(mov_filtered, pay_filtered)


if __name__ == "__main__":
    main()
