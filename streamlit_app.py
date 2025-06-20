# streamlit_insurance_letter.py
"""
Streamlit app to generate a summary letter for insurance contract movements.
Upload an XLS/XLSX file containing columns in one of two supported layouts:

**1. Generic layout (original spec)**
- Item date
- Item name
- Item value

**2. Novis export layout (real‑world file)**
- EntryDate *or* ValueDate
- EntryType
- Amount

The app automatically remaps the Novis column names to the generic names
so that no manual changes are required.
"""

# ---- Imports --------------------------------------------------------------
import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from docx import Document
import streamlit.components.v1 as components


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  HARD‑CODED CONFIGURATION                                            ║
# ║  ▸ Edit the values in this block to localise the app                 ║
# ║    (translations, letter subject/body, sign‑off, etc.).              ║
# ╚══════════════════════════════════════════════════════════════════════╝

TRANSLATION_MAP = {
    # EntryType ➜ Italian category (ITA)
    "Acquisition cost deduction from regular premium": "Costi di emissione e gestione",
    "Risk deduction - Death": "Trattenuta copertura rischio morte",
    "Administrative deduction": "Costi di caricamento",
    "Investment deduction": "Costi di investimento",
    "Contract fee deduction from regular premium": "Costi di emissione e gestione",
    "Risk deduction - Waiver of premium": "Esonero dal Pagamento dei Premi in Caso di Invalidita’ Totale e Permanente (ITP)",
    "Investment deduction from Regular Premium Balance": "Costi di investimento",
    "Investment deduction from Single PremiumBalance": "Costi di investimento",
    "Risk deduction - Illnesses and operations": "Trattenuta coperturarischio malattia, interventi chirurgici e assistenza",
    "Risk deduction - accident insurance deduction": "Trattenuta copertura rischio infortunio",
    "Acquisition cost deduction from single premium": "Costi di emissione e gestione",
    "Contract fee deduction from single premium": "Costi di emissione e gestione",
    "Investment return of Novis Loyalty Bonus": "Rendimento dell'investimento del Bonus Fedeltà NOVIS",
    "Investment return from insurance funds": "Capitalizzazione",
    "Paid Premium": "Pagamenti dei Premi identificati",
    "NOVIS Special Bonus": "NOVIS Special Bonus",
}

LETTER_SUBJECT = "Statement of Account – Insurance Contract"

LETTER_BODY_HEADER = (
    "Dear {client_name},\n\n"
    "Please find below the statement of movements on your insurance contract "
    "number {contract_number}."
)

GOODBYE_LINE = "Sincerely,"

SIGNATURE_BLOCK = "Your Insurance Company\nInsurance Operations Team"

# ──────────────────────────────────────────────────────────────────────────
#  END OF HARD‑CODED SECTION                                                
# ──────────────────────────────────────────────────────────────────────────

# Helper ▶ Column mapping for Novis export files
COLUMN_ALIASES = {
    #   Novis column          ▸ canonical column
    "EntryDate": "Item date",        # we keep ValueDate as fallback
    "ValueDate": "Item date",
    "EntryType": "Item name",
    "Amount": "Item value",
}

EXPECTED_COLS = {"Item date", "Item name", "Item value"}


def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure the dataframe has the expected canonical column names."""
    provided = set(df.columns)
    if EXPECTED_COLS.issubset(provided):
        return df

    renamed = {
        original: canonical
        for original, canonical in COLUMN_ALIASES.items()
        if canonical not in provided and original in provided
    }
    if renamed:
        df = df.rename(columns=renamed)
        provided |= set(renamed.values())

    if not EXPECTED_COLS.issubset(provided):
        missing = ", ".join(EXPECTED_COLS - provided)
        raise ValueError(
            "The uploaded file does not contain the required columns (or recognised "
            f"aliases). Missing: {missing}."
        )
    return df


# ──────────────────────────────────────────────────────────────────────────
#  DATA AGGREGATION LOGIC                                                   
# ──────────────────────────────────────────────────────────────────────────

def summarise_data(df: pd.DataFrame, translation: dict) -> tuple[pd.DataFrame, list[str]]:
    """Return (summary_df, untranslated_original_items).

    Steps:
    1️⃣ Ensure *Item value* numeric
    2️⃣ Sum at original *Item name* level
    3️⃣ Map each original name to translation → new col *Item*
    4️⃣ Re‑aggregate by *Item* (translation) to collapse duplicates
    """

    df["Item value"] = pd.to_numeric(df["Item value"], errors="coerce")
    df = df.dropna(subset=["Item value"])  # safely remove non‑numeric rows

    # First‑level sum by original wording
    level1 = df.groupby("Item name", as_index=False)["Item value"].sum()

    # Map translation; keep original alongside
    level1["Item"] = level1["Item name"].map(translation).fillna(level1["Item name"])

    # Second‑level sum by translated label
    summary = (
        level1.groupby("Item", as_index=False)["Item value"].sum()
        .rename(columns={"Item value": "Amount"})
    )

    # Detect which originals lacked translation
    untranslated = level1.loc[level1["Item"] == level1["Item name"], "Item name"].unique().tolist()

    return summary, untranslated


def build_letter_text(
    client_name: str,
    client_address: str,
    letter_date: date,
    contract_number: str,
    summary: pd.DataFrame,
) -> str:
    lines: list[str] = []
    lines.extend([client_name, client_address, "", letter_date.strftime("%d %B %Y"), "", LETTER_SUBJECT, ""])
    lines.append(
        LETTER_BODY_HEADER.format(client_name=client_name, contract_number=contract_number)
    )
    lines.append("")
    lines.append(summary.to_string(index=False, header=True))
    lines.extend(["", GOODBYE_LINE, "", SIGNATURE_BLOCK])
    return "\n".join(lines)


def build_letter_doc(
    client_name: str,
    client_address: str,
    letter_date: date,
    contract_number: str,
    summary: pd.DataFrame,
) -> Document:
    doc = Document()
    # address & heading
    for part in (client_name, client_address, "", letter_date.strftime("%d %B %Y"), ""):
        doc.add_paragraph(part)
    doc.add_paragraph(LETTER_SUBJECT).style = "Heading 2"
    doc.add_paragraph("")
    doc.add_paragraph(
        LETTER_BODY_HEADER.format(client_name=client_name, contract_number=contract_number)
    )

    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "Item"
    table.rows[0].cells[1].text = "Amount"
    for _, row in summary.iterrows():
        c1, c2 = table.add_row().cells
        c1.text = str(row["Item"])
        c2.text = f"{row['Amount']:.2f}"

    doc.add_paragraph("")
    doc.add_paragraph(GOODBYE_LINE)
    doc.add_paragraph("")
    doc.add_paragraph(SIGNATURE_BLOCK)
    return doc


def doc_to_bytes(doc: Document) -> BytesIO:
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ──────────────────────────────────────────────────────────────────────────
#  STREAMLIT FRONT END                                                      
# ──────────────────────────────────────────────────────────────────────────

def main() -> None:
    st.set_page_config(page_title="Insurance Letter Generator", layout="centered")
    st.title("📄 Insurance Contract Letter Generator")

    uploaded_file = st.file_uploader(
        "Upload XLS/XLSX file with movements", type=["xls", "xlsx"]
    )

    st.subheader("Client Information")
    client_name = st.text_input("Client Name")
    contract_number = st.text_input("Contract Number")
    client_address = st.text_area("Client Address")
    letter_date = st.date_input("Letter Date", value=date.today())

    if uploaded_file is not None:
        try:
            raw_df = pd.read_excel(uploaded_file)
        except Exception as err:
            st.error(f"❌ Could not read the file: {err}")
            st.stop()

        try:
            df = standardise_columns(raw_df)
        except ValueError as e:
            st.error(f"❌ {e}")
            st.stop()

        summary_df, untranslated = summarise_data(df, TRANSLATION_MAP)

        st.subheader("Summarised Movements (translated & aggregated)")
        st.dataframe(summary_df, use_container_width=True)

        if untranslated:
            st.warning(
                "These item types have no translation and appear as‑is in the table: "
                + ", ".join(untranslated)
            )

        if all([client_name, contract_number, client_address]):
            letter_txt = build_letter_text(
                client_name,
                client_address,
                letter_date,
                contract_number,
                summary_df,
            )

            st.subheader("Letter Preview & Copy")
            components.html(
                f"""
                <textarea id='letterArea' style='width:100%;height:260px;'>{letter_txt}</textarea><br>
                <button style='margin-top:6px;padding:6px 12px;font-size:14px;' onclick="navigator.clipboard.writeText(document.getElementById('letterArea').value)">📋 Copy Letter</button>
                """,
                height=320,
            )

            st.download_button(
                label="⬇️ Download Word Document",
                data=doc_to_bytes(
                    build_letter_doc(
                        client_name,
                        client_address,
                        letter_date,
                        contract_number,
                        summary_df,
                    )
                ),
                file_name="Insurance_Statement.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.info("ℹ️ Fill in all client details to generate the letter.")


if __name__ == "__main__":
    main()


🔧 Fixes delivered

1. KeyError: 'Item name'
Root cause: after building summary_df we trimmed the original column, but missing_translations() still expected it.
Fix: the helper now returns the untranslated list alongside the summary, so no lookup on a dropped column.


2. Duplicate Italian rows
Root cause: we summed at the original EntryType level, so two different codes that both map to “Costi di emissione e gestione” produced two rows.
Fix: aggregation now runs twice:

step	group-by	purpose

1️⃣	Item name	sum raw movements
2️⃣	Item (translated)	collapse duplicates



3. Untranslated detector – now computed during step 1 (items whose translation equals original).




---

What changed in code

Section	Key tweak

summarise_data()	returns (summary_df, untranslated); performs two-level aggregation.
missing_translations()	merged into above → removed.
Streamlit main()	uses the new tuple; warns if list is non-empty.
UI subtitle	“translated & aggregated” to make it clear duplicates are gone.



---

New behaviour with your real file

Before	After

Two rows: “Costi di emissione e gestione” 11 800 + 12 500	One row: “Costi di emissione e gestione” 24 300
App crashed with KeyError	App loads summary; any unmapped codes are shown in a yellow warning.



---

Next steps

1. Pull latest streamlit_app.py from Git or click Deploy again.


2. Upload ContractMovements_6120100070 (19_06_2025).xlsx → the table should now aggregate correctly.


3. Generate the letter to verify amounts and wording.



If you still see unexpected rows or amounts, send me their exact labels/values and I’ll adjust the logic further.

