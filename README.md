import re
from io import BytesIO
from datetime import datetime

import streamlit as st
from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="HLD → NC Liste", layout="wide")


def norm(text: str) -> str:
    text = text or ""
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    return re.sub(r"[ \t]+", " ", text)


def read_pdf_text(uploaded_pdf) -> str:
    reader = PdfReader(uploaded_pdf)
    pages = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            pages.append(f"\n--- PAGE {i} ---\n" + (page.extract_text() or ""))
        except Exception:
            pages.append(f"\n--- PAGE {i} ---\n")
    return norm("\n".join(pages))


def search(pattern, text, default=""):
    m = re.search(pattern, text, flags=re.I | re.S)
    return m.group(1).strip() if m else default


def clean_line(value: str) -> str:
    value = (value or "").strip()
    value = re.sub(r"\s+", " ", value)
    return value.strip(" :-")


def short_segment(port: str) -> str:
    # 625200/7/73/C1 -> 7.73/C1 ; 625200/7/73/E/X -> 7.73/E/X
    parts = (port or "").split("/")
    if len(parts) >= 4:
        return f"{parts[1]}.{parts[2]}/" + "/".join(parts[3:])
    if len(parts) >= 3:
        return f"{parts[1]}.{parts[2]}"
    return port or ""


def extract_hld_data(text: str) -> dict:
    data = {}
    data["jira"] = clean_line(search(r"JIRA:\s*([A-Z]+-\s*\d+|[A-Z]+-\d+)", text)).replace(" ", "")
    data["project_name"] = clean_line(search(r"Projekt(?:name)?\s*:?\s*([0-9A-Z_\-]+_[A-Za-z]+_[A-Za-z]+_[A-Z]{2})", text))
    data["fibernode"] = clean_line(search(r"Fibernode:\s*([0-9A-Z]+-\d+)", text)) or clean_line(search(r"LNS\s+([0-9A-Z]+-\d+)_", text))
    data["hub"] = clean_line(search(r"Hub:\s*([A-Za-zÄÖÜäöüß\- ]+)", text)) or clean_line(search(r"HUB\s+([A-Za-zÄÖÜäöüß\- ]+)\s+de-", text))
    data["onkz"] = clean_line(search(r"ONKz:\s*(\d+)", text)) or clean_line(search(r"ONKz\s+(\d+)", text))
    data["asb"] = clean_line(search(r"Asb:\s*(\d+)", text))
    data["ortsteil"] = clean_line(search(r"Ortsteil:\s*([^\n]+)", text))
    data["ort"] = clean_line(search(r"Ort\s+([A-Za-zÄÖÜäöüß\- ]+)\s+Anzahl Kunden", text)) or data.get("ortsteil", "")
    data["cmts"] = clean_line(search(r"CMTS\s+([A-Za-z0-9\- ]+)", text))
    data["gf_length_m"] = clean_line(search(r"GF-Länge\s+(\d+)\s*m", text)) or clean_line(search(r"Länge:\s*~?\s*(\d+)m", text))
    data["datum"] = clean_line(search(r"(\d{2}\.\d{2}\.\d{4})", text))
    data["vrp"] = clean_line(search(r"(?:Standort DFN\s+[0-9A-Z\-/]+\s*)?(\d{4}-\d{3}-\d{3})", text))
    data["plz"] = clean_line(search(r"\n\s*(\d{5})\s+Heppenheim", text)) or clean_line(search(r"\n\s*(\d{5})\s+([A-Za-zÄÖÜäöüß\- ]+)", text))
    data["street"] = clean_line(search(r"\d{4}-\d{3}-\d{3}\s*\n\s*([^\n]+?)\s*\n\s*\d{5}\s+", text))
    data["hub_hostname"] = clean_line(search(r"HUB\s+[A-Za-zÄÖÜäöüß\- ]+\s+([a-z]{2}-[a-z]{3}\d+)", text))
    data["hub_address"] = clean_line(search(r"HUB\s+[A-Za-zÄÖÜäöüß\- ]+\s+[a-z]{2}-[a-z]{3}\d+\s*\n\s*([^\n]+)\s*\n\s*\d{5}", text))

    # Parse FN old/new table if present
    old_fn = clean_line(search(r"Bezeichnung FN alt\s+([0-9A-Z]+-\d+)", text))
    if old_fn and not data["fibernode"]:
        data["fibernode"] = old_fn

    domains = {}

    # Table: Bezeichnung FN neu 7113F-24 7113F-25 / HP / Anzahl Kunden Ki / Auslastung
    table_m = re.search(
        r"Bezeichnung FN neu\s+([0-9A-Z]+-\d+)\s+([0-9A-Z]+-\d+).*?HP\s+(\d+)\s+(\d+).*?Anzahl Kunden Ki\s+(\d+)\s+(\d+)",
        text, flags=re.I | re.S,
    )
    if table_m:
        n1, n2, hp1, hp2, hsi1, hsi2 = table_m.groups()
        domains[n1] = {"node": n1, "hp": int(hp1), "hsi": int(hsi1)}
        domains[n2] = {"node": n2, "hp": int(hp2), "hsi": int(hsi2)}

    # DS blocks from Netzschema page
    for m in re.finditer(r"DS Domäne:\s*([0-9A-Z]+-\d+)(.*?)(?=DS Domäne:|Standort DFN|--- PAGE|$)", text, flags=re.I | re.S):
        node, block = m.group(1), m.group(2)
        d = domains.setdefault(node, {"node": node})
        hp = search(r"HP\s*:\s*(\d+)", block)
        hsi = search(r"HSI\s*:\s*(\d+)", block)
        if hp: d["hp"] = int(hp)
        if hsi: d["hsi"] = int(hsi)
        # US descriptions can appear in order US2 then US1
        for um in re.finditer(r"US\s*([12])\s*Beschreibung:\s*([0-9/]+(?:/[A-Z0-9]+)*)\s*HSI\s*[–-]\s*US\s*:\s*(\d+)", block, flags=re.I):
            us_no, port, hsi_us = um.groups()
            d[f"us{us_no}_desc"] = port.strip()
            d[f"us{us_no}_hsi"] = int(hsi_us)

    # Fallback from Linientechnisches Konzept bullets
    for node, us_no, lines, port in re.findall(
        r"-\s*([0-9A-Z]+-\d+)\s*-\s*US\s*([12]).*?Linien:\s*([^\n]+)\s*US\d\s*Portbezeichnung:\s*([0-9/]+(?:/[A-Z0-9]+)*)",
        text, flags=re.I | re.S,
    ):
        d = domains.setdefault(node, {"node": node})
        d[f"us{us_no}_lines"] = clean_line(lines)
        d.setdefault(f"us{us_no}_desc", clean_line(port))

    # Known defaults from text/project if only one node found
    if not domains and data.get("fibernode"):
        domains[data["fibernode"]] = {"node": data["fibernode"]}

    data["domains"] = domains
    return data


def find_nc_sheet(wb):
    for name in wb.sheetnames:
        if "zuordnung" in name.lower() or "nc" in name.lower():
            return wb[name]
    return wb.active


def find_or_create_row(ws, node: str, preferred_rows=None):
    preferred_rows = preferred_rows or []
    # First search existing node in column C or D
    for row in range(1, ws.max_row + 1):
        c = str(ws.cell(row, 3).value or "").strip()
        d = str(ws.cell(row, 4).value or "").strip()
        if c == node or d == node:
            return row
    # Then use a preferred empty/placeholder row
    for row in preferred_rows:
        if not str(ws.cell(row, 3).value or "").strip() or "DaisyChain" not in str(ws.cell(row, 3).value or ""):
            return row
    return ws.max_row + 1


def setv(ws, cell, value):
    if value not in (None, ""):
        ws[cell] = value


def fill_nc_workbook(xlsx_file, data: dict) -> BytesIO:
    wb = load_workbook(xlsx_file)
    ws = find_nc_sheet(wb)

    # Header fields
    setv(ws, "C3", data.get("jira"))
    setv(ws, "C4", data.get("hub"))
    setv(ws, "C5", data.get("fibernode"))
    # Keep date as string to avoid locale confusion
    setv(ws, "C9", data.get("datum"))

    # Main table columns in NC-Zuordnungsliste
    col = {
        "lfd": 2, "old_node": 3, "dca_name": 4, "nettyp": 5, "access": 6, "dca": 7,
        "fttx": 8, "standort": 9, "plz": 10, "ort": 11, "strasse": 12,
        "us1_segment": 13, "us2_segment": 14, "asb": 15, "ot": 16, "hp": 17,
        "hsi": 18, "us1_hsi": 19, "us2_hsi": 20, "opt_tx1": 21, "rack": 22,
        "opt_tx2": 23, "gf_km": 24, "anzahl_gf": 25, "pegel": 26,
        "us1_beschreibung": 27, "us1_bezeichnung": 28, "serving1": 29,
        "us2_beschreibung": 30, "us2_bezeichnung": 31, "serving2": 32, "bemerkung": 33,
    }

    domains = list(data.get("domains", {}).values())
    # Sort nodes by suffix number for stable rows: 7113F-24 before 7113F-25
    def node_sort(d):
        m = re.search(r"-(\d+)$", d.get("node", ""))
        return int(m.group(1)) if m else 9999
    domains = sorted(domains, key=node_sort)

    preferred = [23, 25, 27, 29, 31, 33]
    for i, d in enumerate(domains):
        node = d.get("node")
        row = find_or_create_row(ws, node, preferred_rows=[preferred[i] if i < len(preferred) else ws.max_row + 1])
        ws.cell(row, col["lfd"]).value = i + 1
        ws.cell(row, col["old_node"]).value = node
        ws.cell(row, col["dca_name"]).value = node
        ws.cell(row, col["dca"]).value = ws.cell(row, col["dca"]).value or "DCA"
        ws.cell(row, col["standort"]).value = "GF-Node"
        ws.cell(row, col["plz"]).value = data.get("plz")
        ws.cell(row, col["ort"]).value = data.get("ort") or data.get("ortsteil")
        ws.cell(row, col["strasse"]).value = data.get("street")
        ws.cell(row, col["asb"]).value = int(data["asb"]) if str(data.get("asb", "")).isdigit() else data.get("asb")
        ws.cell(row, col["ot"]).value = data.get("ortsteil") or data.get("hub")
        ws.cell(row, col["hp"]).value = d.get("hp")
        ws.cell(row, col["hsi"]).value = d.get("hsi")
        ws.cell(row, col["us1_hsi"]).value = d.get("us1_hsi")
        ws.cell(row, col["us2_hsi"]).value = d.get("us2_hsi")
        if data.get("gf_length_m"):
            try:
                ws.cell(row, col["gf_km"]).value = round(float(data["gf_length_m"]) / 1000, 2)
            except Exception:
                ws.cell(row, col["gf_km"]).value = data.get("gf_length_m")
        ws.cell(row, col["anzahl_gf"]).value = ws.cell(row, col["anzahl_gf"]).value or 1
        ws.cell(row, col["us1_beschreibung"]).value = d.get("us1_desc")
        ws.cell(row, col["us2_beschreibung"]).value = d.get("us2_desc")
        ws.cell(row, col["us1_segment"]).value = short_segment(d.get("us1_desc", ""))
        ws.cell(row, col["us2_segment"]).value = short_segment(d.get("us2_desc", ""))

    # Add/update control sheet
    if "AUTO_Extrakt" in wb.sheetnames:
        del wb["AUTO_Extrakt"]
    sh = wb.create_sheet("AUTO_Extrakt")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    sh.append(["Feld", "Wert"])
    for cell in sh[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    rows = [
        ("JIRA", data.get("jira")), ("Projekt", data.get("project_name")), ("Fibernode/Ziel", data.get("fibernode")),
        ("Hub", data.get("hub")), ("ONKz", data.get("onkz")), ("ASB", data.get("asb")),
        ("Ortsteil", data.get("ortsteil")), ("Ort", data.get("ort")), ("PLZ", data.get("plz")),
        ("Straße", data.get("street")), ("VrP", data.get("vrp")), ("GF Länge m", data.get("gf_length_m")),
        ("Datum", data.get("datum")), ("CMTS", data.get("cmts")),
    ]
    for r in rows:
        sh.append(list(r))
    sh.append([])
    sh.append(["Node", "HP", "HSI", "US1 Beschreibung", "US1 HSI", "US2 Beschreibung", "US2 HSI"])
    for cell in sh[sh.max_row]:
        cell.fill = header_fill
        cell.font = header_font
    for d in domains:
        sh.append([d.get("node"), d.get("hp"), d.get("hsi"), d.get("us1_desc"), d.get("us1_hsi"), d.get("us2_desc"), d.get("us2_hsi")])
    for row in sh.iter_rows():
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col_letter, width in {"A": 24, "B": 48, "C": 12, "D": 24, "E": 12, "F": 24, "G": 12}.items():
        sh.column_dimensions[col_letter].width = width

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def data_to_rows(data):
    rows = [
        ("JIRA", data.get("jira")), ("Projekt", data.get("project_name")), ("Fibernode/Ziel", data.get("fibernode")),
        ("Hub", data.get("hub")), ("ONKz", data.get("onkz")), ("ASB", data.get("asb")),
        ("Ortsteil", data.get("ortsteil")), ("PLZ", data.get("plz")), ("Straße", data.get("street")),
        ("VrP", data.get("vrp")), ("GF Länge", data.get("gf_length_m")),
    ]
    return rows


st.title("HLD → NC Liste Generator")
st.write("Upload HLD PDF i NC Excel listu. Aplikacija izvlači podatke iz HLD-a i upisuje ih u NC listu + dodaje kontrolni sheet `AUTO_Extrakt`.")

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("1) HLD PDF", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("2) NC Excel lista", type=["xlsx"])

if pdf_file:
    with st.spinner("Čitam HLD PDF..."):
        text = read_pdf_text(pdf_file)
        data = extract_hld_data(text)
    st.subheader("Pronađeni podaci")
    st.table(data_to_rows(data))
    if data.get("domains"):
        st.subheader("Pronađeni Node / US podaci")
        st.dataframe(list(data["domains"].values()), use_container_width=True)
    else:
        st.warning("Nisam pronašao Node/US tablicu u PDF-u. Excel će ipak dobiti AUTO_Extrakt sheet za kontrolu.")

if pdf_file and xlsx_file:
    if st.button("Generiraj popunjenu NC listu", type="primary"):
        try:
            # Reset file pointers
            pdf_file.seek(0)
            xlsx_file.seek(0)
            text = read_pdf_text(pdf_file)
            data = extract_hld_data(text)
            out = fill_nc_workbook(xlsx_file, data)
            default_name = f"NC-Liste_{data.get('fibernode') or 'AUTO'}_AUTO.xlsx"
            st.success("Gotovo. Preuzmi popunjenu NC listu ispod.")
            st.download_button(
                "Download Excel",
                data=out.getvalue(),
                file_name=default_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error("Došlo je do greške pri generiranju.")
            st.exception(e)
