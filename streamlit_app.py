# streamlit_insurance_letter.py
"""
Enhanced Streamlit app:
• Correct table‑ordering and formatting to mirror the official NOVIS letter
  layout.
• Any unmapped items fall back to the main cost table (T1).
• Header row is only printed when a visible title precedes the table, matching
  the reference PDF.
• Amounts are right‑aligned; total rows bolded.
• Tables rendered in the order specified by TABLE_CONFIG keys.
"""

# ---- Imports ------------------------------------------------------------
import locale
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
from babel.numbers import format_currency
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List   # if it was only used for those lists

# -------------------------------------------------------------------------
#  LOCALE SETUP
# -------------------------------------------------------------------------
try:
    locale.setlocale(locale.LC_ALL, "it_IT.utf8")
except locale.Error:
    pass  # Streamlit Cloud may lack the locale, but Babel still formats OK

# ╔══════════════════════════════════════════════════════════════════════╗
# ║  HARD‑CODED CONFIGURATION                                            ║
# ╚══════════════════════════════════════════════════════════════════════╝

ITEM_CONFIG = {
    # ────────────────  Table T1  ────────────────
    "Acquisition cost deduction from regular premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
        "pos": 1,
    },
    "Contract fee deduction from regular premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
        "pos": 1,                 # same position; rows will merge
    },
    "Acquisition cost deduction from single premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
        "pos": 1,
    },
    "Contract fee deduction from single premium": {
        "label": "Costi di emissione e gestione",
        "table": "T1",
        "pos": 1,
    },

    "Administrative deduction": {
        "label": "Costi di caricamento",
        "table": "T1",
        "pos": 2,
    },

    "Investment deduction": {
        "label": "Costi di investimento",
        "table": "T1",
        "pos": 3,
    },
    "Investment deduction from Regular Premium Balance": {
        "label": "Costi di investimento",
        "table": "T1",
        "pos": 3,
    },
    "Investment deduction from Single PremiumBalance": {
        "label": "Costi di investimento",
        "table": "T1",
        "pos": 3,
    },

    "Investment return from insurance funds": {
        "label": "Capitalizzazione",
        "table": "T1",
        "pos": 4,
    },

    "Paid Premium": {
        "label": "Pagamenti dei Premi identificati",
        "table": "T1",
        "pos": 5,
    },

    "Risk deduction - Death": {
        "label": "Trattenuta copertura rischio morte",
        "table": "T1",
        "pos": 6,
    },
    "Risk deduction - accident insurance deduction": {
        "label": "Trattenuta copertura rischio infortunio",
        "table": "T1",
        "pos": 7,
    },
    "Risk deduction - Illnesses and operations": {
        "label": "Trattenuta copertura rischio malattia, interventi chirurgici e assistenza",
        "table": "T1",
        "pos": 8,
    },
   "Risk deduction - Waiver of premium": {
        "label": "Esonero Pagamento Premi ITP",
        "table": "T1",
        "pos": 9,      # appears after ordered rows
    },
    "Partial surrender": {
         "label": "Riscatto (parziale) + Costi di riscatto",
         "table": "T1",
         "pos": 10,
    },

    # ────────────────  Table T2  ────────────────
    "Investment return of Novis Loyalty Bonus": {
        "label": "Rendimento Bonus Fedeltà NOVIS",
        "table": "T2",
        "pos": 1,
    },
    # If the raw file already contains the Italian string use it directly:
    "NOVIS Loyalty Bonus": {
        "label": "Bonus Fedeltà NOVIS",
        "table": "T2",
        "pos": 2,
    },
  
    # ────────────────  Table T3 (Special Bonus)  ────────────────
    "NOVIS Special Bonus": {
        "label": "NOVIS Special Bonus",
        "table": "T3",
        "pos": 1,        # only row in its table
    },
   }
 
LABEL_POS = {cfg["label"]: cfg.get("pos", 999) for cfg in ITEM_CONFIG.values()}

TABLE_CONFIG = {
    # title empty → no "Item / Importo" header row (as in template)
    "T1": {"title": "", "include_total": True, "total_label": "Somma totale (escluso Bonus Fedeltà NOVIS e Special Bonus)"},
    "T2": {"title": "", "include_total": True, "total_label": "Bonus Fedeltà NOVIS con rendimento"},
    "T3": {"title": "", "include_total": False},
}

LETTER_SUBJECT_TPL = (
    "Dettaglio costi per il valore della Sua posizione assicurativa polizza n. "
    "{contract_number} al {calc_date} con codice fiscale {cf}."
)
LETTER_BODY_HEADER_TPL = (
    "Egregio/a {client_name},\n\n"
    "siamo con la presente a trasmetterLe di seguito la tabella riportante il "
    "dettaglio dei costi applicati ai fini di calcolo del valore della Sua "
    "posizione assicurativa al {calc_date}."
)
OUTRO_PARAGRAPH = (
    "Qualora necessitasse di ulteriori informazioni in merito, La invitiamo "
    "gentilmente a riferirsi alla Tabella Costi contenuta nelle Condizioni di Assicurazione.\n\n"
    "Rimaniamo a disposizione per qualsiasi chiarimento e, ringraziando per la "
    "cortese attenzione, Le porgiamo i nostri più cordiali saluti."
)
GOODBYE_LINE = "Cordiali saluti,"
SIGNATURE_BLOCK = (
    "Il team NOVIS\n\n"
    "NOVIS Insurance Company,\n"
    "NOVIS Versicherungsgesellschaft,\n"
    "NOVIS Compagnia di Assicurazioni,\n"
    "NOVIS Poisťovňa a.s."
)

COLUMN_ALIASES = {
    "EntryDate": "Item date",
    "ValueDate": "Item date",
    "EntryType": "Item name",
    "Amount": "Item value",
}
EXPECTED_COLS = {"Item date", "Item name", "Item value"}

# -------------------------------------------------------------------------
#  HELPERS
# -------------------------------------------------------------------------

def _fmt(amount: float) -> str:
    return format_currency(amount, "EUR", locale="it_IT")


def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    if EXPECTED_COLS.issubset(df.columns):
        return df
    df = df.rename(columns={k: v for k, v in COLUMN_ALIASES.items() if k in df.columns})
    if not EXPECTED_COLS.issubset(df.columns):
        raise ValueError("Il file non contiene le colonne richieste.")
    return df


def aggregate_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df["Item value"] = pd.to_numeric(df["Item value"], errors="coerce")
    df = df.dropna(subset=["Item value"])
    df["Label"] = df["Item name"].apply(lambda x: ITEM_CONFIG.get(x, {}).get("label", x))
    df["Table"] = df["Item name"].apply(lambda x: ITEM_CONFIG.get(x, {}).get("table", "T1"))
    grouped = df.groupby(["Table", "Label"], as_index=False)["Item value"].sum()
    grouped.rename(columns={"Item value": "Amount"}, inplace=True)

    tables = {}
    for tid, g in grouped.groupby("Table"):
        order = TABLE_CONFIG.get(tid, {}).get("order", [])
        # custom order first, then alphabetic
        g["sort_key"] = g["Label"].apply(lambda x: LABEL_POS.get(x, 999))
        g = g.sort_values(["sort_key", "Label"]).drop(columns="sort_key")
        tables[tid] = g.drop(columns="Table")
    return tables

def doc_to_bytes(doc: Document) -> bytes:
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# -------------------------------------------------------------------------
#  DOCX BUILDER
# -------------------------------------------------------------------------

def build_doc(
    client_name: str,
    client_addr: str,
    cf: str,
    contract: str,
    calc_date: str,
    tables: dict[str, pd.DataFrame],
) -> Document:
    doc = Document()

    # address block
    for part in (client_name, client_addr, "", date.today().strftime("%d/%m/%Y"), ""):
        doc.add_paragraph(part)

    doc.add_paragraph(
        LETTER_SUBJECT_TPL.format(contract_number=contract, calc_date=calc_date, cf=cf)
    ).style = "Heading 2"
    doc.add_paragraph("")
    doc.add_paragraph(LETTER_BODY_HEADER_TPL.format(client_name=client_name, calc_date=calc_date))

    grand_total = 0
  
    # tables in predefined order
    for tid in [k for k in TABLE_CONFIG if k in tables]:
        cfg = TABLE_CONFIG[tid]
        df_tbl = tables[tid]

        if cfg["title"]:
            doc.add_paragraph(cfg["title"]).style = "Heading 3"

        header = bool(cfg["title"])
        rows = 1 if header else 0
        tbl = doc.add_table(rows=rows, cols=2, style="Table Grid")   # ▸ borders

        if header:
            tbl.rows[0].cells[0].text = "Item"
            hdr_imp = tbl.rows[0].cells[1]
            hdr_imp.text = "Importo"
            hdr_imp.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for _, row in df_tbl.iterrows():
            c1, c2 = tbl.add_row().cells
            c1.text = row["Label"]
            # bold the Special Bonus row
            if row["Label"] == "NOVIS Special Bonus":
                run = c1.paragraphs[0].runs[0]
                run.bold = True
            c2.text = _fmt(row["Amount"])
            c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        subtotal = df_tbl["Amount"].sum()
        if cfg.get("include_total"):
            c1, c2 = tbl.add_row().cells
            c1.text = cfg.get("total_label", "Totale")
            for r in c1.paragraphs[0].runs:
                r.bold = True
            c2.text = _fmt(subtotal)
            c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            c2.paragraphs[0].runs[0].bold = True

        grand_total += subtotal
        doc.add_paragraph("")

    # grand total line
    p = doc.add_paragraph()
    run1 = p.add_run("Valore della Sua posizione assicurativa ")
    run1.bold = True
    p.add_run("(incluso Bonus Fedeltà NOVIS e NOVIS Special Bonus) ")
    p.add_run(_fmt(grand_total))

    doc.add_paragraph("")           # spacer
    doc.add_paragraph(OUTRO_PARAGRAPH)doc.add_paragraph("")
    doc.add_paragraph(GOODBYE_LINE)
    doc.add_paragraph("")
    doc.add_paragraph(SIGNATURE_BLOCK)
    return doc

# -------------------------------------------------------------------------
#  STREAMLIT FRONT END
# -------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="Generatore Lettera Valorizzazione", layout="centered")
    st.title("📄 Generatore Lettera Valorizzazione")

    file = st.file_uploader("Carica file movimenti (XLS/XLSX)", type=["xls", "xlsx"])

    st.subheader("Dati cliente")
    name = st.text_input("Nome")
    addr = st.text_area("Indirizzo")
    cf = st.text_input("Codice fiscale")
    contract = st.text_input("Numero polizza")
    calc_date = st.date_input("Data valorizzazione", value=date.today()).strftime("%d/%m/%Y")

    if file is not None:
        try:
            df = standardise_columns(pd.read_excel(file))
        except Exception as e:
            st.error(f"Errore nel file: {e}")
            st.stop()
        tables = aggregate_tables(df)
        st.subheader("Anteprima")
        for tid in [k for k in TABLE_CONFIG if k in tables]:
            tbl_df = tables[tid]
            cfg = TABLE_CONFIG[tid]
            st.markdown(f"#### {cfg['title'] or 'Tabella costi'}")
            st.dataframe(tbl_df.assign(Importo=tbl_df["Amount"].apply(_fmt)).drop(columns="Amount"), use_container_width=True)
            if cfg.get("include_total"):
                st.markdown(f"**{cfg['total_label']}: {_fmt(tbl_df['Amount'].sum())}**")

        if all([name, addr, cf, contract]):
            doc = build_doc(name, addr, cf, contract, calc_date, tables)
            st.download_button(
                label="⬇️ Scarica Word",
                data=doc_to_bytes(doc),
                file_name=f"Valorizzazione_dettagliata_polizza_{contract}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
             ) 
        else:
            st.info("Compila tutti i campi cliente per generare la lettera.")


if __name__ == "__main__":
    main()
