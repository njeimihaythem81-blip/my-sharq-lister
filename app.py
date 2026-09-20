"""
MY SHARQ LISTER
----------------
A Streamlit app with two roles:

1. ADMIN: uploads a master Excel file listing all known equipment.
2. GUEST: uploads their own request (a photo of handwriting, a PDF,
   a Word document, or a .txt file) written in Arabic, English, or
   French. The app sends that file to Google's Gemini API (a model
   with a genuinely free tier), which reads/OCRs it, extracts the
   requested items, and matches each one to the closest entry in the
   master list. The guest can then download the matched results as
   an Excel file.

See the README.md that comes with this app for full setup and
deployment instructions.
"""

import io
import json
import os
from datetime import datetime

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

APP_NAME = "MY SHARQ LISTER"

st.set_page_config(page_title=APP_NAME, page_icon="🫒", layout="wide")

MASTER_DIR = "master_data"
MASTER_FILE_PATH = os.path.join(MASTER_DIR, "master_list.xlsx")
ALIASES_FILE_PATH = os.path.join(MASTER_DIR, "aliases.xlsx")
ASSUMPTIONS_FILE_PATH = os.path.join(MASTER_DIR, "assumptions.txt")
MODEL_NAME = "gemini-3.5-flash-lite"  # gemini-2.5-flash-lite was retired for new users
MAX_FILE_SIZE_MB = 20  # Gemini's own inline-file limit

# Safety ceiling on how many master-list rows get sent to the AI in one
# call. gemini-3.5-flash-lite's ~1M token context comfortably fits many
# thousands of short catalog rows (a 6,200-row list is well under this),
# but this cap protects against an unexpectedly huge upload blowing the
# context window. If the admin's list ever exceeds this, only the first
# MAX_MASTER_ROWS_SENT_TO_LLM rows are used, and the app warns visibly
# rather than silently dropping rows.
MAX_MASTER_ROWS_SENT_TO_LLM = 20000

os.makedirs(MASTER_DIR, exist_ok=True)


# --------------------------------------------------------------------------
# Brand colors: pistachio green, clear grey, black text.
# .streamlit/config.toml sets the base theme; this CSS block polishes the
# specific pieces Streamlit's theme settings don't reach (tab labels,
# headers, dataframe, download button).
# --------------------------------------------------------------------------

PISTACHIO = "#8FBC6F"
PISTACHIO_DARK = "#6FA24F"
GREY_BG = "#F2F2EF"
GREY_PANEL = "#E3E3DE"
BLACK = "#111111"


def apply_branding():
    """Inject the pistachio/grey/black look and the app header."""
    st.markdown(
        f"""
        <style>
            .stApp {{
                background-color: {GREY_BG};
                color: {BLACK};
            }}
            h1, h2, h3, h4, h5, h6, p, label, span, div {{
                color: {BLACK};
            }}
            .app-header {{
                display: flex;
                align-items: center;
                gap: 0.6rem;
                padding: 0.75rem 0 0.25rem 0;
            }}
            .app-header-title {{
                font-size: 2.1rem;
                font-weight: 800;
                letter-spacing: 0.02em;
                color: {BLACK};
                margin: 0;
            }}
            .app-header-badge {{
                background-color: {PISTACHIO};
                color: {BLACK};
                padding: 0.15rem 0.6rem;
                border-radius: 999px;
                font-size: 0.85rem;
                font-weight: 600;
            }}
            .stTabs [data-baseweb="tab-list"] {{
                gap: 0.5rem;
                background-color: {GREY_PANEL};
                padding: 0.35rem;
                border-radius: 10px;
            }}
            .stTabs [data-baseweb="tab"] {{
                background-color: transparent;
                color: {BLACK};
                border-radius: 8px;
                font-weight: 600;
            }}
            .stTabs [aria-selected="true"] {{
                background-color: {PISTACHIO} !important;
                color: {BLACK} !important;
            }}
            .stButton > button, .stDownloadButton > button {{
                background-color: {PISTACHIO};
                color: {BLACK};
                border: 1px solid {PISTACHIO_DARK};
                font-weight: 700;
                border-radius: 8px;
            }}
            .stButton > button:hover, .stDownloadButton > button:hover {{
                background-color: {PISTACHIO_DARK};
                color: {BLACK};
                border: 1px solid {PISTACHIO_DARK};
            }}
            [data-testid="stFileUploader"], .stTextInput > div > div {{
                background-color: {GREY_PANEL};
                border-radius: 8px;
            }}
            [data-testid="stDataFrame"] {{
                background-color: white;
                border-radius: 8px;
            }}
        </style>
        <div class="app-header">
            <span style="font-size:2rem;">🫒</span>
            <p class="app-header-title">{APP_NAME}</p>
            <span class="app-header-badge">Equipment Matcher</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


apply_branding()


# --------------------------------------------------------------------------
# Gemini client
# --------------------------------------------------------------------------

def get_secret(name: str) -> str:
    """
    Read a value the app needs to keep private (an API key, the admin
    password) from Streamlit's Secrets. On Streamlit Community Cloud,
    these come from the "Secrets" box you fill in when deploying (App
    settings -> Secrets), written as TOML, e.g.:

        GEMINI_API_KEY = "your-real-key-here"
        ADMIN_PASSWORD = "your-chosen-password"
    """
    try:
        value = st.secrets.get(name)
        if value:
            return value
    except Exception:
        pass
    return os.environ.get(name, "")


def get_gemini_client():
    """Build a Gemini client using the API key from Streamlit Secrets."""
    from google import genai

    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        st.error(
            "No GEMINI_API_KEY found. Add it under this app's Settings -> "
            "Secrets on Streamlit Cloud."
        )
        st.stop()
    return genai.Client(api_key=api_key)


# --------------------------------------------------------------------------
# Master list storage
# --------------------------------------------------------------------------

def load_master_df():
    """Load the previously-saved master Excel file, if one exists."""
    if os.path.exists(MASTER_FILE_PATH):
        try:
            return pd.read_excel(MASTER_FILE_PATH)
        except Exception as e:
            st.error(f"Could not read the saved master file: {e}")
            return None
    return None


def save_master_file(uploaded_file):
    """Persist the admin's uploaded master file to disk."""
    with open(MASTER_FILE_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())


# --------------------------------------------------------------------------
# Custom term dictionary storage (optional)
# --------------------------------------------------------------------------

def load_aliases_df():
    """Load the previously-saved custom term dictionary, if one exists."""
    if os.path.exists(ALIASES_FILE_PATH):
        try:
            return pd.read_excel(ALIASES_FILE_PATH)
        except Exception as e:
            st.error(f"Could not read the saved term dictionary: {e}")
            return None
    return None


def save_aliases_file(uploaded_file):
    """Persist the admin's uploaded term dictionary to disk."""
    with open(ALIASES_FILE_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())


# --------------------------------------------------------------------------
# Business assumptions storage (optional free text)
# --------------------------------------------------------------------------

def load_assumptions() -> str:
    """Load the admin's saved business assumptions text, if any."""
    if os.path.exists(ASSUMPTIONS_FILE_PATH):
        try:
            with open(ASSUMPTIONS_FILE_PATH, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return ""
    return ""


def save_assumptions(text: str):
    """Persist the admin's business assumptions text to disk."""
    with open(ASSUMPTIONS_FILE_PATH, "w", encoding="utf-8") as f:
        f.write(text.strip())


# --------------------------------------------------------------------------
# Reading the guest's file into a format Gemini can understand
# --------------------------------------------------------------------------

def read_guest_file_as_parts(uploaded_file):
    """
    Turn the guest's uploaded file into a list of Gemini API "Part" objects.
    Images and PDFs are sent as raw bytes (Gemini reads them directly, no
    separate OCR step needed). Word and text files are converted to plain
    text first. Returns (parts, error_message); parts is None if there was
    a problem, in which case error_message explains what went wrong.
    """
    from google.genai import types

    name = uploaded_file.name.lower()
    data = uploaded_file.read()

    size_mb = len(data) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        return None, (
            f"That file is {size_mb:.1f}MB, which is over the "
            f"{MAX_FILE_SIZE_MB}MB limit. Please use a smaller or "
            "lower-resolution file."
        )

    if name.endswith((".png", ".jpg", ".jpeg", ".webp", ".gif")):
        if name.endswith(".png"):
            mime = "image/png"
        elif name.endswith(".webp"):
            mime = "image/webp"
        elif name.endswith(".gif"):
            mime = "image/gif"
        else:
            mime = "image/jpeg"
        return [types.Part.from_bytes(data=data, mime_type=mime)], None

    elif name.endswith(".pdf"):
        return [
            types.Part.from_bytes(data=data, mime_type="application/pdf")
        ], None

    elif name.endswith(".docx"):
        import docx  # python-docx

        doc = docx.Document(io.BytesIO(data))
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if not text.strip():
            return None, "That Word document appears to be empty."
        return [types.Part.from_text(text=text)], None

    elif name.endswith(".txt"):
        text = data.decode("utf-8", errors="ignore")
        if not text.strip():
            return None, "That text file appears to be empty."
        return [types.Part.from_text(text=text)], None

    else:
        return None, "Unsupported file type."


# --------------------------------------------------------------------------
# Talking to the LLM: extraction AND matching done together, with real
# semantic understanding, against the full master list
# --------------------------------------------------------------------------

def build_master_context(master_df: pd.DataFrame) -> tuple[str, bool]:
    """
    Turn the master list into CSV text to give the LLM as reference data.
    Returns (csv_text, was_truncated). gemini-3.5-flash-lite's large
    context window comfortably fits many thousands of short catalog rows,
    so this only trims if the list exceeds the defensive safety ceiling.
    """
    was_truncated = len(master_df) > MAX_MASTER_ROWS_SENT_TO_LLM
    trimmed = master_df.head(MAX_MASTER_ROWS_SENT_TO_LLM)
    return trimmed.to_csv(index=False), was_truncated


def build_aliases_context(aliases_df):
    """
    Turn the optional custom term dictionary into CSV text, or None if no
    dictionary has been uploaded. Column names are kept exactly as the
    admin wrote them (e.g. "Term" / "Correct Match"), since the AI just
    reads them as labeled examples rather than a fixed schema.
    """
    if aliases_df is None or aliases_df.empty:
        return None
    return aliases_df.to_csv(index=False)


def call_llm_for_matching(
    client, guest_parts, master_csv: str, aliases_csv=None, assumptions: str = ""
):
    """
    Send the guest's file plus the ENTIRE master list to Gemini in one
    call, and ask it to both extract the requested items and match each
    one using genuine understanding - not literal text overlap. It can
    bridge languages (an Arabic or French description matching an
    English-coded row), resolve abbreviations and technical synonyms, and
    use domain knowledge of electrical/industrial equipment (phase count,
    power rating, voltage, frame size, etc.) to judge whether a loosely
    worded description plausibly refers to a specific catalog row - the
    way an experienced parts-counter clerk would.

    If aliases_csv is given (the admin's custom term dictionary), those
    mappings are treated as trusted ground truth and take priority over
    the AI's own judgment whenever an extracted item matches one closely.

    If assumptions is given (free text the admin wrote, e.g. about the
    typical supply voltage/phase used), it's used as the basis for any
    electrical unit conversions (kW <-> A) needed to bridge a request
    that states power rather than current, or vice versa.
    Raises a plain Exception with a human-readable message on failure.
    """
    from google.genai import types

    aliases_section = ""
    if aliases_csv:
        aliases_section = (
            "\n\nKNOWN TERM MAPPINGS (trusted ground truth - written by the "
            "business that owns this master list). If an extracted item "
            "matches one of these terms closely (in any of the supported "
            "languages, even loosely worded or misspelled), always prefer "
            "the mapped item(s) over your own independent guess. The "
            "mapped value uses a simple wildcard convention with '*':\n"
            "- No '*' at all (e.g. 'S72M200'): names one specific "
            "master-list row directly - use it as the match, High "
            "confidence.\n"
            "- Trailing '*' (e.g. 'A9F791*'): a PRODUCT FAMILY of every "
            "row whose code STARTS WITH that text.\n"
            "- Leading '*' (e.g. '*25'): a PRODUCT FAMILY of every row "
            "whose code ENDS WITH that text.\n"
            "- '*' on both ends (e.g. '*25*'): a PRODUCT FAMILY of every "
            "row whose code CONTAINS that text anywhere.\n"
            "For any family match, pick whichever specific row in that "
            "family best fits any further details in the extracted item "
            "(size, pole count, amperage, rating, etc.); if the request "
            "gives no way to narrow it down further, pick the most "
            "plausible single row from the family and explain the choice "
            "briefly in the notes field.\n\n"
            f"{aliases_csv}"
        )

    conversion_section = (
        "\n\nUNIT CONVERSION: requests sometimes state a motor/device's "
        "power (kW or HP) while the catalog codes by rated current (A), or "
        "the other way around. You may bridge this using standard "
        "electrical engineering knowledge (the usual three-phase formula "
        "I = P / (sqrt(3) x V x power factor x efficiency), typical "
        "assumed power factor/efficiency for induction motors, and common "
        "manufacturer AC-3 motor-starter selection tables for this kind of "
        "equipment) rather than only literal text matching."
    )
    if assumptions.strip():
        conversion_section += (
            "\n\nBUSINESS ASSUMPTIONS (use these defaults unless the "
            f"request clearly states otherwise):\n{assumptions.strip()}"
        )
    else:
        conversion_section += (
            " If the request doesn't state voltage/phase and no business "
            "default is given, state the assumption you used in the notes "
            "field so it can be double-checked."
        )

    system_prompt = (
        "You are an expert equipment-catalog clerk. You read equipment "
        "requests that may be handwritten, scanned, typed, or contained in "
        "a document, written in Arabic, English, or French - sometimes as "
        "precise part codes, sometimes as loose or indirect descriptions "
        "(a translated term, a partial spec, or a description of what the "
        "item does, its size, or its rating rather than its exact code).\n\n"
        "Step 1: Extract every distinct piece of equipment mentioned, "
        "together with its requested quantity if one is given.\n\n"
        "Step 2: For each extracted item, find the single best-matching "
        "row in the MASTER LIST below. Use real understanding, not just "
        "literal text overlap: match across languages, resolve "
        "abbreviations and technical synonyms, and use your domain "
        "knowledge of electrical/industrial equipment to judge whether a "
        "loosely worded description plausibly refers to a specific "
        "catalog row. If no row is a reasonable match, say so rather than "
        "guessing."
        f"{aliases_section}"
        f"{conversion_section}\n\n"
        "Respond with a JSON array. Each array element must look like "
        "this:\n"
        "{\n"
        '  "extracted_text": "the item as it appeared in the request",\n'
        '  "detected_language": "Arabic" | "English" | "French" | "Other",\n'
        '  "quantity": "quantity found, or empty string if none was given",\n'
        '  "matched_master_row": { ...every column from the matching '
        "master row, copied exactly... } or null if nothing matches "
        "well,\n"
        '  "confidence": "High" | "Medium" | "Low",\n'
        '  "notes": "brief reasoning for the match (or why nothing '
        'matched), especially useful when confidence is not High"\n'
        "}\n\n"
        f"MASTER LIST (CSV format, every row is a candidate):\n{master_csv}"
    )

    contents = list(guest_parts)
    contents.append(
        types.Part.from_text(
            text=(
                "Extract the equipment items from this file and match each "
                "one to the master list using real understanding, exactly "
                "as instructed in the system prompt."
            )
        )
    )

    try:
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=contents,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                response_mime_type="application/json",
            ),
        )
    except Exception as e:
        msg = str(e)
        if "RESOURCE_EXHAUSTED" in msg or "429" in msg or "quota" in msg.lower():
            raise RuntimeError(
                "The free Gemini usage limit was reached. Please wait a "
                "minute (or try again tomorrow if the daily limit was hit) "
                "and try again."
            ) from e
        raise RuntimeError(f"The AI service returned an error: {msg}") from e

    raw_text = getattr(response, "text", None)
    if not raw_text:
        raise RuntimeError(
            "The AI returned an empty response. This can happen if the "
            "file's content was blocked by a safety filter, or was too "
            "unclear to read. Please try a clearer file."
        )

    parsed = json.loads(raw_text)
    if not isinstance(parsed, list):
        raise RuntimeError(
            "The AI's response wasn't in the expected list format. "
            "Please try again."
        )
    return parsed


# --------------------------------------------------------------------------
# Turning the LLM's answer into a downloadable Excel file
# --------------------------------------------------------------------------

def results_to_dataframe(results) -> pd.DataFrame:
    rows = []
    for r in results:
        row = {
            "Extracted Text": r.get("extracted_text", ""),
            "Detected Language": r.get("detected_language", ""),
            "Requested Quantity": r.get("quantity", ""),
            "Confidence": r.get("confidence", ""),
            "Notes": r.get("notes", ""),
        }
        matched = r.get("matched_master_row")
        if matched:
            for k, v in matched.items():
                row[f"Master: {k}"] = v
        else:
            row["Master: Match"] = "NOT FOUND"
        rows.append(row)
    return pd.DataFrame(rows)


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Matched Results")
    return buf.getvalue()


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

tab_guest, tab_admin = st.tabs(
    ["Guest: Submit a Request", "Admin: Manage Master List"]
)

# ---- Admin tab -------------------------------------------------------
with tab_admin:
    st.subheader("Admin - Upload Master Equipment List")
    admin_password_input = st.text_input("Admin password", type="password")
    correct_password = get_secret("ADMIN_PASSWORD")

    if admin_password_input and admin_password_input == correct_password:
        st.success("Access granted.")
        master_upload = st.file_uploader(
            "Upload master Excel file (.xls or .xlsx)", type=["xls", "xlsx"]
        )
        if master_upload is not None:
            save_master_file(master_upload)
            st.success("Master file saved. Guests can now submit requests.")

        current_master = load_master_df()
        if current_master is not None:
            st.write(f"Current master list ({len(current_master)} rows) preview:")
            st.dataframe(current_master.head(20))

        st.divider()
        st.subheader("Optional - Custom Term Dictionary")
        st.caption(
            "Teach the app specific words or phrases that should always "
            "point to a particular item, in any supported language. Prepare "
            "a simple Excel file with two columns: the term/phrase in the "
            "first column, and the matching value from your master list in "
            "the second column - an exact code (e.g. 'S72M200') for one "
            "specific item, or use '*' as a wildcard for a whole family of "
            "items: 'A9F791*' = codes starting with that, '*25' = codes "
            "ending with that, '*25*' = codes containing that anywhere. "
            "The AI then picks the best specific item within a family "
            "using the rest of the request's details. These mappings are "
            "trusted as ground truth and override the AI's own guess."
        )
        aliases_upload = st.file_uploader(
            "Upload term dictionary (.xls or .xlsx)",
            type=["xls", "xlsx"],
            key="aliases_uploader",
        )
        if aliases_upload is not None:
            save_aliases_file(aliases_upload)
            st.success("Term dictionary saved.")

        current_aliases = load_aliases_df()
        if current_aliases is not None:
            st.write(
                f"Current term dictionary ({len(current_aliases)} entries) "
                "preview:"
            )
            st.dataframe(current_aliases.head(20))

        st.divider()
        st.subheader("Optional - Business Assumptions")
        st.caption(
            "Free text the AI uses as default context for every request - "
            "most useful for electrical unit conversions (kW <-> A), e.g. "
            "'Assume 380V three-phase supply unless the request states "
            "otherwise.' Leave blank to let the AI state its own assumption "
            "per request in the notes column instead."
        )
        assumptions_input = st.text_area(
            "Business assumptions",
            value=load_assumptions(),
            height=100,
        )
        if st.button("Save assumptions"):
            save_assumptions(assumptions_input)
            st.success("Assumptions saved.")
    elif admin_password_input:
        st.error("Incorrect password.")

# ---- Guest tab --------------------------------------------------------
with tab_guest:
    st.subheader("Guest - Upload Your Request")
    master_df = load_master_df()

    if master_df is None:
        st.warning(
            "No master list has been uploaded yet. Please ask the "
            "administrator to upload one first (Admin tab)."
        )
    else:
        st.info(
            "Upload a photo of handwriting, a scanned or digital PDF, a Word "
            "document, or a text file (up to 20MB). Arabic, English, and "
            "French are all supported."
        )
        guest_upload = st.file_uploader(
            "Upload your file",
            type=["png", "jpg", "jpeg", "webp", "gif", "pdf", "docx", "txt"],
        )

        if guest_upload is not None and st.button("Process Request"):
            with st.spinner("Reading your file and matching items..."):
                try:
                    guest_parts, file_error = read_guest_file_as_parts(guest_upload)
                    if file_error:
                        st.error(file_error)
                    else:
                        client = get_gemini_client()
                        master_csv, was_truncated = build_master_context(
                            master_df
                        )
                        if was_truncated:
                            st.warning(
                                f"The master list has more than "
                                f"{MAX_MASTER_ROWS_SENT_TO_LLM} rows; only "
                                "the first "
                                f"{MAX_MASTER_ROWS_SENT_TO_LLM} were used "
                                "for matching."
                            )
                        aliases_csv = build_aliases_context(load_aliases_df())
                        results = call_llm_for_matching(
                            client,
                            guest_parts,
                            master_csv,
                            aliases_csv,
                            load_assumptions(),
                        )
                        result_df = results_to_dataframe(results)

                        st.success(f"Found {len(result_df)} item(s).")
                        st.dataframe(result_df)

                        excel_bytes = df_to_excel_bytes(result_df)
                        st.download_button(
                            "Download Results as Excel",
                            data=excel_bytes,
                            file_name=(
                                "matched_equipment_"
                                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet"
                            ),
                        )
                except json.JSONDecodeError:
                    st.error(
                        "The AI's response could not be read as structured data. "
                        "Please try again; if it keeps happening, try a clearer "
                        "photo or a shorter file."
                    )
                except RuntimeError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"Something went wrong: {e}")
