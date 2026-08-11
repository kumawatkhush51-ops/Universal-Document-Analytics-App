import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import re
import math
from difflib import get_close_matches

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except ImportError:
    SimpleDocTemplate = Paragraph = Spacer = Table = TableStyle = None
    A4 = None
    colors = None
    getSampleStyleSheet = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from docx import Document
except ImportError:
    Document = None


# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Universal Document Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CLEAN APP STYLING
# ============================================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }

        [data-testid="stMetric"] {
            border: 1px solid rgba(128,128,128,0.20);
            border-radius: 12px;
            padding: 12px;
            min-width: 0;
            overflow: visible !important;
        }

        /* Keep long metric values fully visible instead of showing ... */
        [data-testid="stMetricValue"],
        [data-testid="stMetricValue"] > div {
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            overflow-wrap: anywhere !important;
            word-break: normal !important;
            line-height: 1.12 !important;
            font-size: clamp(1.25rem, 2.3vw, 2.05rem) !important;
        }

        .responsive-metric {
            border: 1px solid rgba(128,128,128,0.20);
            border-radius: 12px;
            padding: 14px 14px 12px 14px;
            min-height: 104px;
            box-sizing: border-box;
            overflow: visible;
        }

        .responsive-metric-label {
            font-size: 0.82rem;
            font-weight: 600;
            margin-bottom: 8px;
            opacity: 0.9;
        }

        .responsive-metric-value {
            font-weight: 500;
            line-height: 1.12;
            white-space: normal;
            overflow-wrap: anywhere;
            word-break: normal;
        }

        .responsive-metric-value.short { font-size: 2rem; }
        .responsive-metric-value.medium { font-size: 1.65rem; }
        .responsive-metric-value.long { font-size: 1.35rem; }
        .responsive-metric-value.xlong { font-size: 1.12rem; }

        @media (max-width: 900px) {
            .responsive-metric-value.short { font-size: 1.65rem; }
            .responsive-metric-value.medium { font-size: 1.4rem; }
            .responsive-metric-value.long { font-size: 1.18rem; }
            .responsive-metric-value.xlong { font-size: 1rem; }
        }

        .app-subtitle {
            color: #9aa4b2;
            margin-top: -12px;
            margin-bottom: 25px;
        }

        .small-note {
            color: #9aa4b2;
            font-size: 0.88rem;
        }

        .insight-card {
            border: 1px solid rgba(128,128,128,0.20);
            border-radius: 14px;
            padding: 14px;
            margin-bottom: 10px;
            min-height: 92px;
        }

        .insight-card h4 {
            margin: 0 0 6px 0;
        }

        .insight-card p {
            color: #9aa4b2;
            margin: 4px 0 0 0;
        }

        .section-card {
            border: 1px solid rgba(128,128,128,0.20);
            border-radius: 14px;
            padding: 16px;
            margin-bottom: 12px;
        }

        .kpi-note {
            color: #9aa4b2;
            font-size: 0.84rem;
        }

        .status-good {
            border-left: 4px solid #21c55d;
            padding: 10px 14px;
            border-radius: 8px;
            background: rgba(33,197,93,0.08);
        }

        .status-warn {
            border-left: 4px solid #f59e0b;
            padding: 10px 14px;
            border-radius: 8px;
            background: rgba(245,158,11,0.08);
        }

        .answer-box {
            border: 1px solid rgba(128,128,128,0.20);
            border-radius: 12px;
            padding: 16px;
            margin-top: 8px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================
def clean_number(value):
    """Safely convert common financial/numeric text into a number."""
    try:
        if pd.isna(value):
            return np.nan
        text = str(value).strip()
        if not text:
            return np.nan
        text = (text.replace("₹", "").replace("$", "").replace("€", "")
                .replace("£", "").replace("%", "").replace(",", ""))
        match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
        return float(match.group(0)) if match else np.nan
    except Exception:
        return np.nan


def make_numeric(df):
    """Convert columns containing mostly numeric values into real numbers."""
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_numeric_dtype(out[col]):
            continue
        converted = out[col].map(clean_number)
        non_blank = int(out[col].notna().sum())
        if non_blank > 0:
            ratio = converted.notna().sum() / non_blank
            if ratio >= 0.50:
                out[col] = converted
    return out


def find_column(df, keywords):
    """Find a likely matching column without exposing technical details."""
    columns = list(df.columns)

    normalized = {
        str(column).strip().lower(): column
        for column in columns
    }

    # Exact / contains matching
    for keyword in keywords:
        keyword = keyword.lower()

        for normalized_name, original in normalized.items():
            if keyword == normalized_name or keyword in normalized_name:
                return original

    # Fuzzy matching
    for keyword in keywords:
        matches = get_close_matches(
            keyword.lower(),
            list(normalized.keys()),
            n=1,
            cutoff=0.55,
        )

        if matches:
            return normalized[matches[0]]

    return None


def money(value):
    """Format a value as INR without crashing on string/object values."""
    if pd.isna(value):
        return "N/A"
    numeric = clean_number(value)
    if pd.isna(numeric):
        return "N/A"
    return f"₹{numeric:,.2f}"


def responsive_metric(label, value, col=None):
    """Render a metric card that never truncates long values with an ellipsis."""
    display_value = str(value)
    length = len(display_value)
    size_class = (
        "short" if length <= 12
        else "medium" if length <= 16
        else "long" if length <= 22
        else "xlong"
    )

    html = f"""
    <div class="responsive-metric">
        <div class="responsive-metric-label">{label}</div>
        <div class="responsive-metric-value {size_class}">{display_value}</div>
    </div>
    """

    target = col if col is not None else st
    target.markdown(html, unsafe_allow_html=True)


def first_title(text, fallback="Document"):
    """Return a clean first meaningful line as the document title."""
    for line in (text or "").splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if line and not line.lower().startswith(("--- page", "--- table")):
            return line[:100]
    return fallback or "Document"


def quality_score(df):
    """Simple user-friendly quality score."""
    total_cells = max(df.size, 1)
    missing = int(df.isna().sum().sum())
    duplicates = int(df.duplicated().sum())

    missing_penalty = min(50, (missing / total_cells) * 100)
    duplicate_penalty = min(
        30,
        (duplicates / max(len(df), 1)) * 100,
    )

    score = 100 - missing_penalty - duplicate_penalty

    return max(0, round(score, 1))


def extract_pdf(raw):
    if pdfplumber is None:
        raise RuntimeError(
            "PDF support is not installed. Please install the requirements file."
        )

    pages = []

    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""

            if text.strip():
                pages.append(
                    f"--- Page {page_number} ---\n{text}"
                )

    return "\n\n".join(pages)


def extract_docx(raw):
    if Document is None:
        raise RuntimeError(
            "Word document support is not installed. Please install the requirements file."
        )

    doc = Document(io.BytesIO(raw))

    parts = [
        paragraph.text
        for paragraph in doc.paragraphs
        if paragraph.text.strip()
    ]

    for table_number, table in enumerate(doc.tables, start=1):
        parts.append(f"\n--- Table {table_number} ---")

        for row in table.rows:
            parts.append(
                " | ".join(
                    cell.text.strip()
                    for cell in row.cells
                )
            )

    return "\n".join(parts)


def read_excel_robust(raw, filename):
    """Read Excel files reliably in both local and Streamlit Cloud environments.

    Some workbooks do not identify their engine cleanly, and files created by
    different Excel-compatible programs can have unusual metadata. We try the
    correct engine first and then a safe fallback instead of failing immediately.
    """
    lower_name = filename.lower()
    errors = []

    # .xlsx / .xlsm are ZIP-based Office Open XML files.
    if lower_name.endswith((".xlsx", ".xlsm")):
        for engine in ("openpyxl", "xlrd"):
            try:
                df = pd.read_excel(io.BytesIO(raw), engine=engine)
                return normalize_table(df)
            except Exception as exc:
                errors.append(f"{engine}: {exc}")

    # Legacy .xls files require xlrd.
    if lower_name.endswith(".xls"):
        try:
            df = pd.read_excel(io.BytesIO(raw), engine="xlrd")
            return normalize_table(df)
        except Exception as exc:
            errors.append(f"xlrd: {exc}")

        # A few files are incorrectly named .xls but are actually .xlsx.
        try:
            df = pd.read_excel(io.BytesIO(raw), engine="openpyxl")
            return normalize_table(df)
        except Exception as exc:
            errors.append(f"openpyxl fallback: {exc}")

    details = " | ".join(errors[-2:])
    raise ValueError(
        "This Excel workbook could not be read. It may be damaged, password protected, "
        "or saved with an incorrect Excel extension. Open it in Excel and use "
        "Save As → Excel Workbook (.xlsx), then upload the saved copy."
        + (f" Technical details: {details}" if details else "")
    )


def read_uploaded_file(uploaded_file):
    """Read supported files and return kind, dataframe, text."""
    name = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if not raw:
        raise ValueError("The uploaded file is empty.")

    if name.endswith(".csv"):
        # utf-8 first, then common Windows encoding.
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(raw), encoding="cp1252")
        return "table", normalize_table(df), ""

    if name.endswith((".xlsx", ".xls", ".xlsm")):
        return "table", read_excel_robust(raw, name), ""

    if name.endswith(".json"):
        text = raw.decode("utf-8", errors="replace")
        obj = json.loads(text)
        if isinstance(obj, list) and obj and all(isinstance(item, dict) for item in obj):
            return "table", normalize_table(pd.json_normalize(obj)), text
        return "text", None, json.dumps(obj, indent=2, ensure_ascii=False)

    if name.endswith(".docx"):
        return extract_docx_structured(raw)

    if name.endswith(".pdf"):
        return extract_pdf_structured(raw)

    if name.endswith(".txt"):
        return "text", None, raw.decode("utf-8", errors="replace")

    raise ValueError(
        "This file type is not supported. Please upload Excel, CSV, JSON, PDF, Word, or TXT."
    )


def answer_table_question(df, question):
    """
    Local question answering for structured tables.
    No external AI service is used.
    """
    q = question.lower().strip()

    if not q:
        return "Please enter a question."

    ndf = make_numeric(df)

    sales_col = find_column(
        ndf,
        ["sales", "revenue", "amount", "total sales"],
    )

    quantity_col = find_column(
        ndf,
        ["quantity sold", "quantity", "qty", "units"],
    )

    price_col = find_column(
        ndf,
        ["unit price", "price", "selling price"],
    )

    product_col = find_column(
        ndf,
        ["product name", "product", "item"],
    )

    category_col = find_column(
        ndf,
        ["category", "type", "segment"],
    )

    month_col = find_column(
        ndf,
        ["month", "date", "period", "year month"],
    )

    if q in {"help", "what can i ask", "what can i ask?"}:
        return (
            "You can ask questions such as:\n\n"
            "- What is the total sales?\n"
            "- Which product has the highest sales?\n"
            "- Which category has the highest sales?\n"
            "- What is the average unit price?\n"
            "- Which month had the highest sales?\n"
            "- How many rows are there?\n"
            "- How many missing values are there?\n"
            "- How many duplicate rows are there?"
        )

    if "how many rows" in q or "number of rows" in q:
        return f"The file contains **{len(df):,} rows**."

    if "how many columns" in q or "number of columns" in q:
        return f"The file contains **{len(df.columns):,} columns**."

    if any(word in q for word in ["missing", "null", "blank"]):
        missing = int(df.isna().sum().sum())
        return f"The file contains **{missing:,} missing cells**."

    if "duplicate" in q:
        duplicates = int(df.duplicated().sum())
        return f"The file contains **{duplicates:,} duplicate rows**."

    if any(
        phrase in q
        for phrase in [
            "total sales",
            "total revenue",
            "total amount",
        ]
    ):
        if sales_col:
            return (
                f"Total **{sales_col}** is "
                f"**{money(ndf[sales_col].sum())}**."
            )

        return (
            "I could not identify a sales, revenue, "
            "or amount column."
        )

    if "total quantity" in q or "total units" in q:
        if quantity_col:
            return (
                f"Total **{quantity_col}** is "
                f"**{ndf[quantity_col].sum():,.0f}**."
            )

        return "I could not identify a quantity column."

    if "average" in q and "price" in q:
        if price_col:
            return (
                f"Average **{price_col}** is "
                f"**{money(ndf[price_col].mean())}**."
            )

        return "I could not identify a price column."

    if (
        product_col
        and sales_col
        and any(
            word in q
            for word in [
                "highest",
                "top",
                "best",
                "most",
            ]
        )
    ):
        grouped = (
            ndf.groupby(product_col)[sales_col]
            .sum()
            .sort_values(ascending=False)
        )

        if len(grouped):
            return (
                f"Top product by sales is "
                f"**{grouped.index[0]}** with "
                f"**{money(grouped.iloc[0])}**."
            )

    if (
        category_col
        and sales_col
        and "category" in q
        and any(
            word in q
            for word in [
                "highest",
                "top",
                "best",
                "most",
            ]
        )
    ):
        grouped = (
            ndf.groupby(category_col)[sales_col]
            .sum()
            .sort_values(ascending=False)
        )

        if len(grouped):
            return (
                f"Top category by sales is "
                f"**{grouped.index[0]}** with "
                f"**{money(grouped.iloc[0])}**."
            )

    if (
        month_col
        and sales_col
        and "month" in q
        and any(
            word in q
            for word in [
                "highest",
                "top",
                "best",
                "most",
            ]
        )
    ):
        grouped = (
            ndf.groupby(month_col)[sales_col]
            .sum()
            .sort_values(ascending=False)
        )

        if len(grouped):
            return (
                f"Highest-sales period is "
                f"**{grouped.index[0]}** with "
                f"**{money(grouped.iloc[0])}**."
            )

    # Direct numeric-column questions
    numeric_columns = list(
        ndf.select_dtypes(include=np.number).columns
    )

    for column in numeric_columns:
        column_name = str(column).lower()

        if column_name in q:
            if "average" in q or "mean" in q:
                return (
                    f"Average of **{column}** is "
                    f"**{ndf[column].mean():,.2f}**."
                )

            if "sum" in q or "total" in q:
                return (
                    f"Sum of **{column}** is "
                    f"**{ndf[column].sum():,.2f}**."
                )

    return (
        "I could not confidently answer that using the built-in "
        "analysis. Try asking about totals, averages, highest "
        "values, products, categories, months, missing values, "
        "duplicates, or type **help**."
    )



# ============================================================
# V5 DOCUMENT INTELLIGENCE HELPERS
# ============================================================
def normalize_table(df):
    if df is None:
        return None
    out = df.copy()
    out.columns = [
        re.sub(r"\s+", " ", str(c).replace("\n", " ").strip()) or f"Column {i+1}"
        for i, c in enumerate(out.columns)
    ]
    out = out.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return out.reset_index(drop=True)


def docx_table_to_df(table):
    rows = []
    for row in table.rows:
        rows.append([cell.text.strip() for cell in row.cells])
    if len(rows) < 2:
        return None

    header = [str(x).strip() or f"Column {i+1}" for i, x in enumerate(rows[0])]
    if len(set(x.lower() for x in header)) < max(2, int(len(header) * 0.5)):
        return None

    data = []
    for row in rows[1:]:
        vals = [str(x).strip() for x in row]
        vals = (vals + [""] * len(header))[:len(header)]
        if any(vals):
            data.append(vals)
    if not data:
        return None

    return normalize_table(pd.DataFrame(data, columns=header))


def extract_docx_structured(raw):
    if Document is None:
        raise RuntimeError("Word document support is not installed.")
    doc = Document(io.BytesIO(raw))

    paragraphs = [
        p.text.strip()
        for p in doc.paragraphs
        if p.text and p.text.strip()
    ]

    tables = []
    for table in doc.tables:
        df = docx_table_to_df(table)
        if df is not None and len(df.columns) >= 2:
            tables.append(df)

    text = "\n".join(paragraphs)
    if tables:
        largest = max(tables, key=lambda x: x.shape[0] * x.shape[1])
        return "table", largest, text
    return "text", None, text


def pdf_table_to_df(rows):
    if not rows or len(rows) < 2:
        return None
    cleaned = []
    for row in rows:
        vals = ["" if x is None else str(x).strip() for x in row]
        if any(vals):
            cleaned.append(vals)
    if len(cleaned) < 2:
        return None
    header = [x or f"Column {i+1}" for i, x in enumerate(cleaned[0])]
    data = []
    for row in cleaned[1:]:
        vals = (row + [""] * len(header))[:len(header)]
        if any(vals):
            data.append(vals)
    if not data:
        return None
    return normalize_table(pd.DataFrame(data, columns=header))


def extract_pdf_structured(raw):
    if pdfplumber is None:
        raise RuntimeError("PDF support is not installed.")
    text_parts = []
    candidates = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            if page_text.strip():
                text_parts.append(f"Page {page_no}\n{page_text}")
            try:
                for rows in page.extract_tables() or []:
                    df = pdf_table_to_df(rows)
                    if df is not None:
                        candidates.append(df)
            except Exception:
                pass
    text = "\n\n".join(text_parts)
    if candidates:
        return "table", max(candidates, key=lambda x: x.shape[0] * x.shape[1]), text
    return "text", None, text


def detect_document_domain(df=None, text=""):
    content = ""
    if df is not None:
        content += " ".join(str(c) for c in df.columns).lower()
        content += " " + " ".join(
            str(x).lower() for x in df.head(5).astype(str).values.flatten()
        )
    content += " " + (text or "")[:12000].lower()

    groups = {
        "Sales": ["sales", "revenue", "order", "unit price", "quantity sold",
                  "gross profit", "product", "customer"],
        "Finance": ["profit", "loss", "balance sheet", "cash flow", "expense",
                    "asset", "liability", "ebitda", "financial"],
        "Marketing": ["campaign", "impressions", "clicks", "conversion",
                      "ctr", "roas", "advertising", "lead", "marketing"],
        "Inventory": ["inventory", "stock", "warehouse", "sku", "reorder",
                      "supplier", "on hand"],
        "HR": ["employee", "staff", "salary", "department", "attendance",
               "attrition", "headcount", "hire"],
        "Customers": ["customer", "client", "satisfaction", "retention",
                      "churn", "segment", "lifetime value"],
        "Operations": ["production", "machine", "downtime", "defect",
                       "capacity", "delivery", "lead time", "operations"],
    }
    scores = {name: sum(content.count(k) for k in keys)
              for name, keys in groups.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] else "General"


def text_word_count(text):
    return len(re.findall(r"\b\w+\b", text or "", flags=re.UNICODE))


def text_numbers(text):
    return re.findall(
        r"(?<![\w])[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?(?![\w])",
        text or ""
    )


def text_money_values(text):
    return re.findall(
        r"(?:₹|\$|€|£)\s?\d[\d,]*(?:\.\d+)?(?:\s?(?:crore|cr|lakh|lac|million|billion|k|m|bn))?",
        text or "",
        flags=re.I
    )


def text_highlights(text, limit=6):
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text or "")
    keywords = [
        "important", "key", "total", "revenue", "sales", "profit", "loss",
        "growth", "increase", "decrease", "decline", "investment", "cost",
        "customer", "employee", "risk", "target", "forecast", "plan",
        "result", "performance", "market", "strategy", "expected"
    ]
    ranked = []
    for sentence in sentences:
        s = re.sub(r"\s+", " ", sentence).strip()
        if len(s) < 30 or len(s) > 320:
            continue
        score = sum(s.lower().count(k) for k in keywords)
        if any(ch.isdigit() for ch in s):
            score += 2
        ranked.append((score, s))
    ranked.sort(key=lambda x: x[0], reverse=True)
    seen = set()
    output = []
    for _, s in ranked:
        if s.lower() not in seen:
            seen.add(s.lower())
            output.append(s)
        if len(output) >= limit:
            break
    return output


def text_topics(text, limit=8):
    stop = {
        "the","and","for","that","with","this","from","are","was","were",
        "have","has","had","will","would","there","their","about","into",
        "than","then","they","them","your","you","our","which","what",
        "when","where","who","how","why","also","been","being","company",
        "document","data","analysis","table","page","year","month"
    }
    words = re.findall(r"[A-Za-z][A-Za-z'-]{3,}", text or "")
    freq = {}
    for word in words:
        key = word.lower()
        if key in stop:
            continue
        freq[key] = freq.get(key, 0) + 1
    return [x.title() for x, _ in sorted(freq.items(), key=lambda z: z[1], reverse=True)[:limit]]


def text_answer(text, question):
    q = question.lower().strip()
    if not q:
        return "Please enter a question."

    if any(x in q for x in ["summary", "summarize", "main point", "overview"]):
        points = text_highlights(text, 5)
        return "\n\n".join(f"- {p}" for p in points) if points else "I could not find enough readable content for a reliable summary."

    if "topic" in q or "keyword" in q:
        topics = text_topics(text, 10)
        return "Key topics detected: **" + ", ".join(topics) + "**."

    if "how many words" in q:
        return f"The document contains approximately **{text_word_count(text):,} words**."

    if "how many numbers" in q:
        return f"I detected approximately **{len(text_numbers(text)):,} numeric values**."

    query_words = [
        w for w in re.findall(r"[a-zA-Z]{4,}", q)
        if w not in {"what","which","where","when","about","this","that","give","show","tell"}
    ]
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text or "")
    ranked = []
    for sentence in sentences:
        s = re.sub(r"\s+", " ", sentence).strip()
        if len(s) < 20:
            continue
        score = sum(1 for w in query_words if w in s.lower())
        if score:
            ranked.append((score, s))
    ranked.sort(key=lambda x: x[0], reverse=True)

    if ranked:
        return "\n\n".join(f"- {s}" for _, s in ranked[:4])

    return "I could not find a confident answer with the built-in document search. Try asking for a summary, key topics, important points, or include a specific term from the document."


def table_recommendations_v5(df, cols):
    ndf = make_numeric(df)
    recs = []
    sales = cols.get("sales")
    product = cols.get("product")
    category = cols.get("category")
    period = cols.get("month")

    if sales and product:
        grouped = ndf.groupby(product)[sales].sum().sort_values()
        if len(grouped) >= 2:
            recs.append(
                f"Focus on **{grouped.index[-1]}**, the strongest product by sales, and review **{grouped.index[0]}** for possible improvement."
            )
    if sales and category:
        grouped = ndf.groupby(category)[sales].sum().sort_values()
        if len(grouped) >= 2:
            recs.append(
                f"Compare the leading category **{grouped.index[-1]}** with **{grouped.index[0]}** to understand the performance gap."
            )
    if sales and period:
        grouped = ndf.groupby(period)[sales].sum()
        if len(grouped) >= 2:
            change = ((grouped.iloc[-1] - grouped.iloc[0]) / grouped.iloc[0] * 100) if grouped.iloc[0] else 0
            if change < 0:
                recs.append("The latest period is below the first period. Investigate the recent decline before increasing spending.")
            else:
                recs.append("The latest period is stronger than the first. Identify the products or categories driving the improvement.")
    return recs[:4]



# ============================================================
# V6 INTELLIGENCE HELPERS
# ============================================================
def safe_pct(a, b):
    try:
        return (float(a) / float(b) * 100) if float(b) else 0.0
    except Exception:
        return 0.0


def detect_columns(ndf):
    return {
        "sales": find_column(ndf, ["sales", "revenue", "amount", "total sales", "net sales"]),
        "quantity": find_column(ndf, ["quantity sold", "quantity", "qty", "units", "units sold"]),
        "price": find_column(ndf, ["unit price", "price", "selling price", "average price"]),
        "cost": find_column(ndf, ["unit cost", "cost", "purchase cost", "cogs"]),
        "product": find_column(ndf, ["product name", "product", "item", "sku"]),
        "category": find_column(ndf, ["category", "type", "segment", "department"]),
        "period": find_column(ndf, ["month", "date", "period", "year month", "year"]),
        "customer": find_column(ndf, ["customer", "client", "customer name"]),
    }


def calculate_table_metrics(df):
    ndf = make_numeric(df)
    cols = detect_columns(ndf)
    metrics = {"domain": detect_document_domain(df, "")}
    sales, qty, price, cost = cols["sales"], cols["quantity"], cols["price"], cols["cost"]
    metrics["rows"] = len(df)
    metrics["columns"] = len(df.columns)
    metrics["missing"] = int(df.isna().sum().sum())
    metrics["duplicates"] = int(df.duplicated().sum())
    if sales:
        metrics["sales"] = float(ndf[sales].sum())
    if qty:
        metrics["quantity"] = float(ndf[qty].sum())
    if price:
        metrics["avg_price"] = float(ndf[price].mean())
    if sales and cost and qty:
        metrics["profit"] = float((ndf[sales] - ndf[cost] * ndf[qty]).sum())
        metrics["margin"] = safe_pct(metrics["profit"], metrics["sales"])
    return metrics, ndf, cols


def trend_insight(ndf, sales_col, period_col):
    if not sales_col or not period_col or len(ndf) < 2:
        return None
    grouped = ndf.groupby(period_col)[sales_col].sum()
    if len(grouped) < 2:
        return None
    first, last = float(grouped.iloc[0]), float(grouped.iloc[-1])
    change = safe_pct(last - first, abs(first))
    direction = "increased" if change > 0 else "decreased" if change < 0 else "remained stable"
    return f"Sales {direction} by **{abs(change):.1f}%** from the first available period to the last available period."


def generate_table_insights(df, cols):
    ndf = make_numeric(df)
    insights, warnings, recommendations = [], [], []
    sales, qty, product, category, period = cols["sales"], cols["quantity"], cols["product"], cols["category"], cols["period"]
    if sales:
        total = float(ndf[sales].sum())
        insights.append(f"Total **{sales}** is **{money(total)}** across {len(df):,} records.")
        if product:
            g = ndf.groupby(product)[sales].sum().sort_values(ascending=False)
            if len(g):
                share = safe_pct(g.iloc[0], g.sum())
                insights.append(f"**{g.index[0]}** is the top product with **{money(g.iloc[0])}**, contributing **{share:.1f}%** of sales.")
                if len(g) >= 2 and g.iloc[-1] > 0:
                    recommendations.append(f"Review the lower-performing product **{g.index[-1]}** and compare its price, demand, and cost with the top performer.")
        if category:
            g = ndf.groupby(category)[sales].sum().sort_values(ascending=False)
            if len(g):
                share = safe_pct(g.iloc[0], g.sum())
                insights.append(f"**{g.index[0]}** is the leading category at **{share:.1f}%** of total sales.")
                if share >= 50:
                    warnings.append(f"Sales are highly concentrated: the leading category contributes **{share:.1f}%** of sales.")
        t = trend_insight(ndf, sales, period)
        if t: insights.append(t)
    if qty:
        insights.append(f"Total quantity is **{ndf[qty].sum():,.0f} units**.")
    if cols["cost"] and sales and qty:
        profit = float((ndf[sales] - ndf[cols["cost"]] * ndf[qty]).sum())
        margin = safe_pct(profit, ndf[sales].sum())
        insights.append(f"Calculated profit is **{money(profit)}** with an estimated margin of **{margin:.1f}%**.")
        if margin < 10:
            warnings.append(f"Estimated profit margin is only **{margin:.1f}%**; cost control may need attention.")
        recommendations.append("Compare high-sales products with their unit costs to protect profitable growth.")
    if int(df.isna().sum().sum()) > 0:
        warnings.append(f"There are **{int(df.isna().sum().sum()):,} missing cells** that may affect analysis.")
    if int(df.duplicated().sum()) > 0:
        warnings.append(f"There are **{int(df.duplicated().sum()):,} duplicate rows** that should be reviewed.")
    if not recommendations:
        recommendations.append("Use the Analytics tab to compare products, categories, quantities, and periods before making decisions.")
    return insights[:8], warnings[:6], recommendations[:6]


def analytics_type_answers(df, ndf, cols, metrics, recommendations):
    """Return actual answers for Descriptive, Diagnostic, Predictive and Prescriptive analytics."""
    sales = cols.get("sales")
    qty = cols.get("quantity")
    price = cols.get("price")
    cost = cols.get("cost")
    product = cols.get("product")
    category = cols.get("category")
    period = cols.get("period")

    # Descriptive: what happened?
    descriptive = []
    if sales:
        descriptive.append(f"Total sales/revenue is **{money(ndf[sales].sum())}** across **{len(df):,} records**.")
    if qty:
        descriptive.append(f"Total quantity is **{ndf[qty].sum():,.0f} units**.")
    if price:
        descriptive.append(f"Average unit price is **{money(ndf[price].mean())}**.")
    if sales and product:
        g = ndf.groupby(product)[sales].sum().sort_values(ascending=False)
        if len(g):
            descriptive.append(f"The top product is **{g.index[0]}** with **{money(g.iloc[0])}** in sales.")
    if sales and category:
        g = ndf.groupby(category)[sales].sum().sort_values(ascending=False)
        if len(g):
            descriptive.append(f"The leading category is **{g.index[0]}**, contributing **{safe_pct(g.iloc[0], g.sum()):.1f}%** of sales.")
    if not descriptive:
        descriptive.append(f"The file contains **{len(df):,} records** and **{len(df.columns):,} fields**, but no standard business measure was confidently detected.")

    # Diagnostic: why did it happen?
    diagnostic = []
    if sales and product:
        g = ndf.groupby(product)[sales].sum().sort_values(ascending=False)
        if len(g):
            share = safe_pct(g.iloc[0], g.sum())
            diagnostic.append(f"**{g.index[0]}** is the strongest product and contributes **{share:.1f}%** of total sales, making it a major driver of performance.")
        if len(g) >= 2:
            diagnostic.append(f"The gap between the top product (**{money(g.iloc[0])}**) and lowest-sales product (**{money(g.iloc[-1])}**) is **{money(g.iloc[0] - g.iloc[-1])}**.")
    if sales and category:
        g = ndf.groupby(category)[sales].sum().sort_values(ascending=False)
        if len(g):
            share = safe_pct(g.iloc[0], g.sum())
            if share >= 50:
                diagnostic.append(f"The leading category accounts for **{share:.1f}%** of sales, so category concentration is a major reason for the overall result.")
            else:
                diagnostic.append(f"The leading category contributes **{share:.1f}%** of sales; performance is spread across multiple categories.")
    if sales and cost and qty:
        profit = float((ndf[sales] - ndf[cost] * ndf[qty]).sum())
        margin = safe_pct(profit, ndf[sales].sum())
        diagnostic.append(f"Estimated gross profitability is **{margin:.1f}%**, based on sales minus unit cost × quantity.")
    if not diagnostic:
        diagnostic.append("The available file does not contain enough comparable dimensions to confidently explain the drivers of performance.")

    # Predictive: what next?
    predictive = []
    forecast = simple_forecast(df, sales, period, 3) if sales and period else None
    if forecast is not None:
        grouped, preds = forecast
        direction = "increase" if preds[-1] > grouped.iloc[-1] else "decrease" if preds[-1] < grouped.iloc[-1] else "remain broadly stable"
        change = safe_pct(preds[-1] - grouped.iloc[-1], abs(grouped.iloc[-1]))
        predictive.append(f"The simple trend model estimates that sales may **{direction} by about {abs(change):.1f}%** by the third forecast period compared with the latest observed period.")
        predictive.append(f"Estimated third forecast value: **{money(preds[-1])}**. This is a statistical trend estimate, not a guarantee.")
    else:
        predictive.append("A reliable trend forecast cannot be calculated yet. The file needs a recognizable time/period column and at least three comparable periods.")

    # Prescriptive: what should we do?
    prescriptive = []
    for item in recommendations[:4]:
        prescriptive.append(re.sub(r"\*\*", "", str(item)))
    if sales and category:
        g = ndf.groupby(category)[sales].sum().sort_values(ascending=False)
        if len(g) >= 2:
            prescriptive.append(f"Protect the leading category **{g.index[0]}**, while investigating why **{g.index[-1]}** trails it.")
    if sales and cost and qty:
        profit = float((ndf[sales] - ndf[cost] * ndf[qty]).sum())
        margin = safe_pct(profit, ndf[sales].sum())
        if margin < 20:
            prescriptive.append("Prioritize cost and pricing review because the estimated margin is below 20%.")
        else:
            prescriptive.append("Protect the current margin by monitoring unit costs on the highest-sales products.")
    if not prescriptive:
        prescriptive.append("Use the strongest detected drivers as the starting point, validate the figures, and investigate weak-performing segments before taking action.")

    return {
        "Descriptive — What happened?": descriptive[:5],
        "Diagnostic — Why?": diagnostic[:5],
        "Predictive — What next?": predictive[:4],
        "Prescriptive — What should we do?": prescriptive[:5],
    }


def simple_forecast(df, sales_col, period_col, periods=3):
    if not sales_col or not period_col:
        return None
    ndf = make_numeric(df)
    grouped = ndf.groupby(period_col)[sales_col].sum().dropna()
    if len(grouped) < 3:
        return None
    y = grouped.values.astype(float)
    x = np.arange(len(y), dtype=float)
    try:
        slope, intercept = np.polyfit(x, y, 1)
        future_x = np.arange(len(y), len(y) + periods, dtype=float)
        preds = slope * future_x + intercept
        return grouped, np.maximum(preds, 0)
    except Exception:
        return None


def create_excel_report(df, metrics, insights, warnings, recommendations, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame([metrics]).to_excel(writer, index=False, sheet_name="Summary")
        pd.DataFrame({"Insights": insights}).to_excel(writer, index=False, sheet_name="Insights")
        pd.DataFrame({"Warnings": warnings}).to_excel(writer, index=False, sheet_name="Warnings")
        pd.DataFrame({"Recommendations": recommendations}).to_excel(writer, index=False, sheet_name="Recommendations")
        df.to_excel(writer, index=False, sheet_name="Data")


def create_docx_report(title, metrics, insights, warnings, recommendations, path):
    if Document is None:
        return False
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph("Universal Document Analytics — V6")
    doc.add_heading("Executive Summary", level=1)
    for key, value in metrics.items():
        doc.add_paragraph(f"{key}: {value}")
    for heading, items in [("Key Insights", insights), ("Things to Watch", warnings), ("Recommendations", recommendations)]:
        doc.add_heading(heading, level=1)
        for item in items:
            doc.add_paragraph(re.sub(r"\*\*", "", str(item)), style="List Bullet")
    doc.save(path)
    return True


def create_pdf_report(title, metrics, insights, warnings, recommendations, path):
    if SimpleDocTemplate is None:
        return False
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12), Paragraph("Universal Document Analytics — V6", styles["Normal"]), Spacer(1, 12)]
    data = [[str(k), str(v)] for k, v in metrics.items()]
    if data:
        table = Table([["Metric", "Value"]] + data, colWidths=[180, 320])
        table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1f2937")), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), 0.4, colors.grey), ("VALIGN", (0,0), (-1,-1), "TOP")]))
        story += [table, Spacer(1, 14)]
    for heading, items in [("Key Insights", insights), ("Things to Watch", warnings), ("Recommendations", recommendations)]:
        story.append(Paragraph(heading, styles["Heading2"]))
        for item in items:
            story.append(Paragraph("• " + re.sub(r"\*\*", "", str(item)), styles["BodyText"]))
            story.append(Spacer(1, 5))
    doc.build(story)
    return True


def report_metrics_for_text(text, domain):
    years = sorted(set(re.findall(r"\b(?:19|20)\d{2}\b", text)))
    period = f"{years[0]}–{years[-1]}" if len(years) >= 2 else (years[0] if years else "Not detected")
    return {
        "Detected area": domain,
        "Words": f"{text_word_count(text):,}",
        "Numeric values": f"{len(text_numbers(text)):,}",
        "Money values": f"{len(text_money_values(text)):,}",
        "Period": period,
    }


# ============================================================
# SESSION STATE
# ============================================================
defaults = {
    "file_name": None,
    "file_kind": None,
    "df": None,
    "document_text": None,
    "domain": "General",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# HEADER
# ============================================================
st.title("📊 Universal Document Analytics")
st.markdown(
    '<div class="app-subtitle">'
    "Upload documents and analyze your data easily — without "
    "exposing technical code or internal Streamlit information."
    "</div>",
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("📂 Upload & Manage Files")

    uploaded = st.file_uploader(
        "Upload your document",
        type=[
            "xlsx",
            "xls",
            "xlsm",
            "csv",
            "json",
            "pdf",
            "docx",
            "txt",
        ],
        help=(
            "Supported: Excel (.xlsx/.xls/.xlsm), CSV, JSON, PDF, Word and TXT."
        ),
    )

    if (
        uploaded is not None
        and uploaded.name != st.session_state.file_name
    ):
        try:
            kind, data, text = read_uploaded_file(uploaded)

            st.session_state.file_name = uploaded.name
            st.session_state.file_kind = kind
            st.session_state.df = data
            st.session_state.document_text = text
            st.session_state.domain = detect_document_domain(data, text or "")

            st.success("File loaded successfully.")

        except Exception as exc:
            st.error(
                "We could not read this file. If it is an Excel file, open it in Excel "
                "and use Save As → Excel Workbook (.xlsx), then upload the saved copy."
            )
            # Keep technical details out of the main dashboard, but make them available
            # for troubleshooting when a deployment has a genuinely unusual file.
            with st.expander("Show troubleshooting details"):
                st.code(str(exc))

    if st.session_state.file_name:
        st.divider()

        st.write(
            f"**Current file:** "
            f"{st.session_state.file_name}"
        )

        st.write(f"**Detected area:** {st.session_state.domain}")
        if st.session_state.file_kind == "table":
            st.write("**Detected format:** Structured data")
        else:
            st.write("**Detected format:** Document")

        if st.button(
            "🗑️ Clear current file",
            use_container_width=True,
        ):
            for key in defaults:
                st.session_state[key] = None
            st.session_state.domain = "General"

            st.rerun()


# ============================================================
# NO FILE STATE
# ============================================================
if not st.session_state.file_name:
    st.info("👈 Upload a file from the sidebar to begin.")

    st.markdown("### Supported file types")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            """
            **📊 Structured data**

            - Excel `.xlsx`
            - Excel `.xls`
            - Excel `.xlsm`
            - CSV `.csv`
            - JSON `.json`
            """
        )

    with c2:
        st.markdown(
            """
            **📄 Documents**

            - PDF `.pdf`
            - Word `.docx`
            - TXT `.txt`
            """
        )

    with c3:
        st.markdown(
            """
            **🔎 Analysis**

            - Data preview
            - KPIs
            - Charts
            - Data quality
            - Questions
            """
        )

    st.stop()


# ============================================================
# STRUCTURED DATA
# ============================================================
if st.session_state.file_kind == "table":
    df = st.session_state.df
    metrics, ndf, cols = calculate_table_metrics(df)
    domain = st.session_state.domain
    insights, warnings, recommendations = generate_table_insights(df, cols)

    st.success(f"Loaded: **{st.session_state.file_name}**")

    overview_tab, insights_tab, analytics_tab, quality_tab, ask_tab, export_tab, raw_tab = st.tabs([
        "📊 Overview", "🔥 Insights", "📈 Analytics", "🔎 Data Quality", "💬 Ask Your File", "📥 Export", "📋 Raw Data"
    ])

    with overview_tab:
        st.subheader("✨ Executive Overview")
        st.caption(f"Detected area: **{domain}** • {len(df):,} records • {len(df.columns):,} fields")
        c1, c2, c3, c4 = st.columns(4)
        responsive_metric("Rows", f"{len(df):,}", c1)
        responsive_metric("Columns", f"{len(df.columns):,}", c2)
        responsive_metric("Missing Cells", f"{metrics['missing']:,}", c3)
        responsive_metric("Quality Score", f"{quality_score(df)}%", c4)

        st.markdown("### 💼 Key Business Metrics")
        k1, k2, k3, k4 = st.columns(4)
        responsive_metric("Total Sales / Revenue", money(metrics.get("sales")) if "sales" in metrics else "Not detected", k1)
        responsive_metric("Total Quantity", f"{metrics['quantity']:,.0f}" if "quantity" in metrics else "Not detected", k2)
        responsive_metric("Average Price", money(metrics.get("avg_price")) if "avg_price" in metrics else "Not detected", k3)
        responsive_metric("Calculated Profit", money(metrics.get("profit")) if "profit" in metrics else "Not enough data", k4)

        st.markdown("### 🧾 Executive Summary")
        if insights:
            st.info(" ".join(re.sub(r"\*\*", "", x) for x in insights[:3]))
        else:
            st.info("Not enough structured information was detected for an automatic summary.")

        if warnings:
            st.markdown("### ⚠️ Things to Watch")
            for item in warnings[:4]: st.warning(re.sub(r"\*\*", "", item))

        st.markdown("### 👀 Data at a Glance")
        st.dataframe(df.head(12), width="stretch")

    with insights_tab:
        st.subheader("🔥 Business Insights")
        if insights:
            for item in insights:
                st.markdown(f'<div class="insight-card"><h4>💡 Insight</h4><p>{item}</p></div>', unsafe_allow_html=True)
        if warnings:
            st.markdown("### ⚠️ Things to Watch")
            for item in warnings: st.warning(re.sub(r"\*\*", "", item))
        st.markdown("### 💡 Recommended Actions")
        for item in recommendations: st.markdown(f"- {item}")

        st.markdown("### 🧭 Analytics Type — Answers from Your File")
        analytics_answers = analytics_type_answers(df, ndf, cols, metrics, recommendations)
        a1, a2, a3, a4 = st.columns(4)
        answer_cards = [
            (a1, "📊 Descriptive", "What happened?", analytics_answers["Descriptive — What happened?"]),
            (a2, "🔎 Diagnostic", "Why?", analytics_answers["Diagnostic — Why?"]),
            (a3, "🔮 Predictive", "What next?", analytics_answers["Predictive — What next?"]),
            (a4, "🎯 Prescriptive", "What should we do?", analytics_answers["Prescriptive — What should we do?"]),
        ]
        for card, title, question, answers in answer_cards:
            with card:
                st.markdown(f"### {title}")
                st.markdown(f"**{question}**")
                for answer in answers:
                    st.markdown(f"- {answer}")

    with analytics_tab:
        st.subheader("📈 Automatic Analytics")
        chart_created = False
        sales_col, qty_col = cols["sales"], cols["quantity"]
        product_col, category_col, period_col = cols["product"], cols["category"], cols["period"]
        if sales_col and product_col:
            st.markdown("### Sales by Product")
            st.bar_chart(ndf.groupby(product_col)[sales_col].sum().sort_values(ascending=False).head(15)); chart_created = True
        if sales_col and category_col:
            st.markdown("### Sales by Category")
            st.bar_chart(ndf.groupby(category_col)[sales_col].sum().sort_values(ascending=False)); chart_created = True
        if sales_col and period_col:
            st.markdown("### Sales by Period")
            st.line_chart(ndf.groupby(period_col)[sales_col].sum()); chart_created = True
        if qty_col and product_col:
            st.markdown("### Quantity by Product")
            st.bar_chart(ndf.groupby(product_col)[qty_col].sum().sort_values(ascending=False).head(15)); chart_created = True

        forecast = simple_forecast(df, sales_col, period_col, 3)
        if forecast is not None:
            grouped, preds = forecast
            st.markdown("### 🔮 Simple Trend Forecast")
            st.caption("This is a statistical trend estimate, not a guarantee. It requires at least three periods.")
            future = pd.DataFrame({"Forecast": preds}, index=[f"Future {i+1}" for i in range(len(preds))])
            st.line_chart(pd.concat([grouped.rename("Actual"), future], axis=0))
        if not chart_created: st.info("No standard sales/product/category/time columns were detected for automatic charts.")

    with quality_tab:
        st.subheader("🔎 Data Quality")
        c1, c2, c3 = st.columns(3)
        c1.metric("Missing Cells", f"{metrics['missing']:,}")
        c2.metric("Duplicate Rows", f"{metrics['duplicates']:,}")
        c3.metric("Quality Score", f"{quality_score(df)}%")
        quality_table = pd.DataFrame({"Column": df.columns, "Data Type": [str(x) for x in df.dtypes], "Missing Values": [int(x) for x in df.isna().sum()], "Missing %": [round(x*100,2) for x in df.isna().mean()], "Unique Values": [int(df[c].nunique(dropna=True)) for c in df.columns]})
        st.dataframe(quality_table, width="stretch")
        if metrics["duplicates"] == 0: st.success("No duplicate rows detected.")
        else: st.warning(f"{metrics['duplicates']:,} duplicate rows detected.")
        if metrics["missing"] == 0: st.success("No missing values detected.")
        else: st.warning(f"{metrics['missing']:,} missing cells detected.")

    with ask_tab:
        st.subheader("💬 Ask Your File")
        st.caption("Questions are answered locally from the uploaded data. No AI or external service is required.")
        examples = ["What is the total sales?", "Which product has the highest sales?", "Which category has the highest sales?", "What is the average unit price?", "Which month had the highest sales?", "How many rows are there?", "What are the key insights?"]
        st.write("**Try:** " + " • ".join(examples[:5]))
        question = st.text_input("Your question", placeholder="Example: Which product has the highest sales?", key="table_question_v6")
        if st.button("🔎 Get Answer", type="primary", use_container_width=True, key="table_answer_v6"):
            if question.strip():
                if "key insight" in question.lower() or "business insight" in question.lower():
                    answer = "\n\n".join(f"- {x}" for x in insights[:6])
                elif "recommend" in question.lower() or "what should" in question.lower():
                    answer = "\n\n".join(f"- {x}" for x in recommendations[:6])
                else:
                    answer = answer_table_question(df, question)
                st.markdown("### Answer")
                st.markdown(answer)
            else: st.warning("Please enter a question first.")

    with export_tab:
        st.subheader("📥 Export Analysis Report")
        st.write("Create a clean report containing the main metrics, insights, warnings, recommendations, and data.")
        base = re.sub(r"[^A-Za-z0-9_-]+", "_", st.session_state.file_name.rsplit('.',1)[0])
        import tempfile, os
        with tempfile.TemporaryDirectory() as td:
            xlsx_path = os.path.join(td, f"{base}_analysis.xlsx")
            docx_path = os.path.join(td, f"{base}_analysis.docx")
            pdf_path = os.path.join(td, f"{base}_analysis.pdf")
            create_excel_report(df, metrics, insights, warnings, recommendations, xlsx_path)
            create_docx_report(f"Analysis Report — {st.session_state.file_name}", metrics, insights, warnings, recommendations, docx_path)
            pdf_ok = create_pdf_report(f"Analysis Report — {st.session_state.file_name}", metrics, insights, warnings, recommendations, pdf_path)
            with open(xlsx_path, "rb") as f: st.download_button("📊 Download Excel Report", f.read(), os.path.basename(xlsx_path), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with open(docx_path, "rb") as f: st.download_button("📝 Download Word Report", f.read(), os.path.basename(docx_path), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            if pdf_ok:
                with open(pdf_path, "rb") as f: st.download_button("📄 Download PDF Report", f.read(), os.path.basename(pdf_path), "application/pdf", use_container_width=True)
            else: st.info("PDF export requires reportlab. Install the V6 requirements file to enable it.")

    with raw_tab:
        st.subheader("📋 Raw Data")
        st.dataframe(df, width="stretch", height=600)
        st.download_button("⬇️ Download CSV", df.to_csv(index=False).encode("utf-8"), "analyzed_data.csv", "text/csv", use_container_width=True)


# ============================================================
# TEXT / UNSTRUCTURED DOCUMENTS
# ============================================================
else:
    text = st.session_state.document_text or ""
    domain = st.session_state.domain
    word_count = text_word_count(text)
    numeric_values = text_numbers(text)
    money_values = text_money_values(text)
    highlights = text_highlights(text, 8)
    topics = text_topics(text, 10)
    title = first_title(text, st.session_state.file_name)
    text_metrics = report_metrics_for_text(text, domain)

    st.markdown(f"## 📄 {title}")
    st.caption(f"Detected area: **{domain}** • {word_count:,} words • {len(numeric_values):,} numeric values")
    overview_tab, insights_tab, ask_tab, export_tab, raw_tab = st.tabs(["✨ Key Insights", "🧠 Important Content", "💬 Ask Document", "📥 Export", "📄 Original Content"])

    with overview_tab:
        st.markdown("### ⭐ Key Indicators")
        c1, c2, c3, c4 = st.columns(4)
        responsive_metric("📝 Words", f"{word_count:,}", c1)
        responsive_metric("🔢 Numbers", f"{len(numeric_values):,}", c2)
        responsive_metric("💰 Money Values", f"{len(money_values):,}", c3)
        years = sorted(set(re.findall(r"\b(?:19|20)\d{2}\b", text)))
        period = f"{years[0]}–{years[-1]}" if len(years) >= 2 else (years[0] if years else "Not detected")
        responsive_metric("📅 Period", period, c4)

        st.markdown("### 🧾 Executive Summary")
        if highlights: st.info(" ".join(highlights[:3]))
        else: st.info("No reliable summary could be generated from the available text.")
        st.markdown("### 🔥 Key Points")
        for item in highlights[:6]: st.markdown(f"- {item}")
        if topics:
            st.markdown("### 🏷️ Main Topics")
            st.write(" • ".join(topics))
        if money_values:
            st.markdown("### 💰 Important Values")
            money_cols = st.columns(min(4, len(money_values)))
            for card, value in zip(money_cols, money_values[:4]): responsive_metric("Detected value", value, card)

    with insights_tab:
        st.markdown("### 🧠 Important Content")
        for i, item in enumerate(highlights, 1):
            st.markdown(f'<div class="insight-card"><h4>💡 Key Point {i}</h4><p>{item}</p></div>', unsafe_allow_html=True)
        st.markdown("### 🏷️ Main Topics")
        if topics:
            cards = st.columns(min(4, len(topics)))
            for card, topic in zip(cards, topics): card.markdown(f'<div class="insight-card"><h4>🔹 {topic}</h4><p>Recurring topic detected.</p></div>', unsafe_allow_html=True)
        st.markdown("### 🔢 Document Statistics")
        st.write(f"**{len(numeric_values):,}** numeric values and **{len(money_values):,}** money values were detected.")

    with ask_tab:
        st.subheader("💬 Ask About This Document")
        st.caption("Built-in document search works locally without an external AI service.")
        q = st.text_input("Your question", placeholder="Example: Give me a summary of this document", key="text_question_v6")
        if st.button("🔎 Find Answer", type="primary", use_container_width=True, key="text_answer_v6"):
            if q.strip(): st.markdown("### Answer\n\n" + text_answer(text, q))
            else: st.warning("Please enter a question.")

    with export_tab:
        st.subheader("📥 Export Analysis Report")
        import tempfile, os
        base = re.sub(r"[^A-Za-z0-9_-]+", "_", st.session_state.file_name.rsplit('.',1)[0])
        with tempfile.TemporaryDirectory() as td:
            xlsx_path = os.path.join(td, f"{base}_analysis.xlsx")
            docx_path = os.path.join(td, f"{base}_analysis.docx")
            pdf_path = os.path.join(td, f"{base}_analysis.pdf")
            insights = highlights[:8]
            warnings = []
            recommendations = ["Review the highlighted sections and verify important figures against the original document."]
            create_excel_report(pd.DataFrame({"Extracted Text": [text]}), text_metrics, insights, warnings, recommendations, xlsx_path)
            create_docx_report(f"Document Analysis — {st.session_state.file_name}", text_metrics, insights, warnings, recommendations, docx_path)
            pdf_ok = create_pdf_report(f"Document Analysis — {st.session_state.file_name}", text_metrics, insights, warnings, recommendations, pdf_path)
            with open(xlsx_path, "rb") as f: st.download_button("📊 Download Excel Report", f.read(), os.path.basename(xlsx_path), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with open(docx_path, "rb") as f: st.download_button("📝 Download Word Report", f.read(), os.path.basename(docx_path), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            if pdf_ok:
                with open(pdf_path, "rb") as f: st.download_button("📄 Download PDF Report", f.read(), os.path.basename(pdf_path), "application/pdf", use_container_width=True)

    with raw_tab:
        st.markdown("### 📄 Original Extracted Content")
        st.caption("The original content is kept separate so the main screen stays focused on useful information.")
        with st.expander("Open full extracted content"):
            st.text_area("Extracted content", text, height=600, label_visibility="collapsed")

# ============================================================
# FOOTER
# ============================================================
st.divider()

st.markdown(
    '<div class="small-note">'
    "Universal Document Intelligence • V6 • AI-free • "
    "User-facing interface contains no internal debug output."
    "</div>",
    unsafe_allow_html=True,
)
