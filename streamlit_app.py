import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO
from typing import Optional

# =============================
# Streamlit Setup
# =============================
st.set_page_config(
    page_title="Excel → PDF Balkendiagramme – Big4 Style",
    layout="wide"
)

st.title("📊 Excel → PDF Balkendiagramme – Big4 Style")

st.write(
    "Dieses Tool erzeugt Balkendiagramme aus Excel-Dateien im **Corporate-Beratungslayout** "
    "(klar, reduziert, hochwertig). "
    "Lade unten deine Excel-Datei hoch, wähle ein Farbset und erhalte eine PDF mit allen Diagrammen."
)

# =============================
# Farbpaletten (Colourways)
# =============================
colourways = {
    "Blau/Grau (Standard Big4)": ["#002060", "#0050b3", "#7f7f7f"],
    "Grün/Türkis": ["#006400", "#009999", "#a6a6a6"],
    "Lila/Beere": ["#4B004B", "#732673", "#B380B3"],
    "Orange/Grau": ["#E46C0A", "#F4B183", "#7f7f7f"],
}

colour_choice = st.selectbox(
    "🎨 Farbset auswählen:",
    list(colourways.keys()),
    index=0
)
selected_colors = colourways[colour_choice]

# =============================
# Hilfsfunktionen
# =============================
def find_category_column(df: pd.DataFrame) -> Optional[str]:
    """Finde die erste nicht-numerische Spalte (als Kategorie)."""
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            return col
    return None


def render_sheet_to_pdf(df: pd.DataFrame, pdf: PdfPages, sheet_label: str, colors):
    """Erzeuge Balkendiagramme im Big4-Stil für ein Sheet."""
    cat_col = find_category_column(df)
    if cat_col is None:
        df = df.copy()
        df["Index"] = df.index.astype(str)
        cat_col = "Index"

    categories = df[cat_col].astype(str).tolist()
    value_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

    for col in value_cols:
        # Fehlerwerte suchen
        error_col = None
        for candidate in [f"{col}_Fehler", f"{col}_SD", f"{col}_Error"]:
            if candidate in df.columns:
                error_col = candidate
                break

        values = df[col].values
        errors = df[error_col].values if error_col else None

        fig, ax = plt.subplots(figsize=(7, 4.5), dpi=150)

        bars = ax.bar(
            categories,
            values,
            yerr=errors,
            capsize=4,
            error_kw=dict(ecolor="grey", lw=1, alpha=0.8),
            color=[colors[i % len(colors)] for i in range(len(categories))]
        )

        # --- Stil-Anpassungen ---
        ax.set_ylim(0, max(values) * 1.2 if len(values) > 0 else 1)
        ax.set_axisbelow(True)
        ax.yaxis.grid(True, color="#E0E0E0", linestyle="-", linewidth=0.8)
        ax.xaxis.grid(False)

        # Spines
        for spine in ["top", "right"]:
            ax.spines[spine].set_visible(False)
        ax.spines["left"].set_color("#999999")
        ax.spines["bottom"].set_color("#999999")

        # Achsenticks & Schrift
        ax.tick_params(axis="x", labelrotation=0, labelsize=10)
        ax.tick_params(axis="y", labelsize=9)

        # Titel
        ax.set_title(
            f"{sheet_label} – {col}",
            fontsize=12,
            fontweight="bold",
            loc="left",
            pad=12
        )

        # Legende (unterhalb, reduziert)
        ax.legend(
            bars,
            categories,
            loc="upper center",
            bbox_to_anchor=(0.5, -0.15),
            ncol=len(categories),
            frameon=False,
            fontsize=9
        )

        plt.tight_layout()
        pdf.savefig(fig)
        plt.close(fig)


# =============================
# Hauptlogik
# =============================
uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xls"])

if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
    except Exception as e:
        st.error(f"Fehler beim Einlesen der Datei: {e}")
        st.stop()

    st.write(f"**Gefundene Sheets:** {', '.join(xls.sheet_names)}")

    if st.button("📑 PDF generieren"):
        pdf_bytes = BytesIO()
        with PdfPages(pdf_bytes) as pdf:
            for sheet in xls.sheet_names:
                df = pd.read_excel(uploaded, sheet_name=sheet)
                if df.empty:
                    continue
                try:
                    render_sheet_to_pdf(df, pdf, sheet_label=sheet, colors=selected_colors)
                except Exception as e:
                    fig, ax = plt.subplots(figsize=(7, 5))
                    ax.axis("off")
                    ax.text(0.05, 0.95, f"Fehler in Sheet '{sheet}':\n{e}",
                            va="top", wrap=True)
                    pdf.savefig(fig)
                    plt.close(fig)

        st.success("✅ PDF erfolgreich erstellt!")
        st.download_button(
            "📥 PDF herunterladen",
            data=pdf_bytes.getvalue(),
            file_name="charts.pdf",
            mime="application/pdf"
        )
