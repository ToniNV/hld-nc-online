#!/usr/bin/env python3
"""
HLD -> NC-Liste automation (first working version)

Usage CLI:
  python hld_nc_automation.py --pdf HLD.pdf --template NC-Liste.xlsx --output NC-Liste_AUTO.xlsx

Usage GUI:
  python hld_nc_automation.py

Dependencies:
  pip install pymupdf openpyxl
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except Exception as exc:
    raise SystemExit("PyMuPDF fehlt. Bitte installieren: pip install pymupdf") from exc

try:
    from openpyxl import load_workbook
except Exception as exc:
    raise SystemExit("openpyxl fehlt. Bitte installieren: pip install openpyxl") from exc


@dataclass
class NodeData:
    name: str = ""
    hp: Optional[int] = None
    hsi: Optional[int] = None
    us1_desc: str = ""
    us1_hsi: Optional[int] = None
    us2_desc: str = ""
    us2_hsi: Optional[int] = None
    us1_segment: str = ""
    us2_segment: str = ""


@dataclass
class ExtractedHLD:
    jira: str = ""
    project_name: str = ""
    hub: str = ""
    target_fibernode: str = ""
    fn_alt: str = ""
    fn_new: List[NodeData] = None
    grussz: str = ""
    onkz: str = ""
    asb: str = ""
    plz: str = ""
    ort: str = ""
    street: str = ""
    vrp: str = ""
    gf_length_m: Optional[int] = None
    cmts: str = ""
    fn_alt_hp: Optional[int] = None
    fn_alt_hsi: Optional[int] = None
    raw_warnings: List[str] = None

    def __post_init__(self):
        if self.fn_new is None:
            self.fn_new = []
        if self.raw_warnings is None:
            self.raw_warnings = []


def pdf_text_by_page(pdf_path: Path) -> List[str]:
    doc = fitz.open(str(pdf_path))
    pages = []
    for page in doc:
        text = page.get_text("text") or ""
        pages.append(text)
    return pages


def first_match(text: str, patterns: List[str], flags=re.I | re.M) -> str:
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return (m.group(1) if m.groups() else m.group(0)).strip()
    return ""


def parse_int(value: str) -> Optional[int]:
    if not value:
        return None
    value = value.replace(".", "").replace(",", "").strip()
    m = re.search(r"\d+", value)
    return int(m.group(0)) if m else None


def normalize_segment(desc: str) -> str:
    """Convert 625200/7/73/Y -> 7.73, or 625200/7/73/E/X -> 7.73 when needed for NC column."""
    m = re.search(r"\d{6}/(\d+)/(\d+)", desc)
    return f"{m.group(1)}.{m.group(2)}" if m else ""


def extract_hld(pdf_path: Path) -> ExtractedHLD:
    pages = pdf_text_by_page(pdf_path)
    text = "\n".join(pages)
    data = ExtractedHLD()

    data.jira = first_match(text, [r"JIRA:\s*([A-Z]+-\s*\d+)", r"JIRA:\s*([A-Z]+-\d+)"]).replace(" ", "")
    data.project_name = first_match(text, [r"Projektname:\s*([^\n]+)", r"Projekt:\s*([^\n]+)"])
    data.hub = first_match(text, [r"Hub:\s*([^\n]+)", r"HUB\s+([^\n]+)"])
    # Prefer explicit Fibernode from cover page, fallback project name token.
    data.target_fibernode = first_match(text, [r"Fibernode:\s*([0-9A-Z]+-[0-9]+)", r"LNS\s+([0-9A-Z]+-[0-9]+)_"])
    data.fn_alt = first_match(text, [r"Bezeichnung FN alt\s+([0-9A-Z]+-[0-9]+)"])
    data.grussz = first_match(text, [r"GRUSSZ:\s*(\d+)"])
    data.onkz = first_match(text, [r"ONKz:\s*(\d+)", r"ONKz\s+(\d+)"])
    data.asb = first_match(text, [r"Asb:\s*(\d+)", r"Asb\s*:\s*(\d+)"])
    data.cmts = first_match(text, [r"CMTS\s+([^\n]+)"])
    data.gf_length_m = parse_int(first_match(text, [r"GF-Länge\s+([0-9.]+)\s*m", r"Länge:\s*~?\s*([0-9.]+)m"]))
    data.fn_alt_hp = parse_int(first_match(text, [r"Anzahl HP\s+([0-9.]+)"]))
    data.fn_alt_hsi = parse_int(first_match(text, [r"Anzahl Kunden HSI \(Ki\)\s+([0-9.]+)"]))

    # Standort block: Standort DFN ... line, code, street, PLZ Ort
    standort = first_match(text, [r"Standort DFN[^\n]*\n([\s\S]{0,160}?\d{5}\s+[^\n]+)"])
    if standort:
        lines = [ln.strip() for ln in standort.splitlines() if ln.strip()]
        # Typical: 6252-007-073 / Darmstädter Str./Bürgermeister-Kunz-Str. / 64646 Heppenheim
        for ln in lines:
            if re.search(r"\d{4}-\d{3}-\d{3}", ln):
                data.vrp = re.search(r"\d{4}-\d{3}-\d{3}", ln).group(0)
            elif re.search(r"\d{5}\s+", ln):
                m = re.search(r"(\d{5})\s+(.+)", ln)
                data.plz, data.ort = m.group(1), m.group(2).strip()
            elif not data.street and not re.search(r"Standort|DCA|VrP|UM-DF", ln, re.I):
                data.street = ln

    # Fallback address fields from full text.
    if not data.vrp:
        data.vrp = first_match(text, [r"(\d{4}-\d{3}-\d{3})"])
    if not data.plz or not data.ort:
        m = re.search(r"(\d{5})\s+(Heppenheim|Bensheim|[A-ZÄÖÜ][A-Za-zÄÖÜäöüß\- ]+)", text)
        if m:
            data.plz, data.ort = m.group(1), m.group(2).strip()
    if not data.street:
        data.street = first_match(text, [r"([A-ZÄÖÜ][A-Za-zÄÖÜäöüß.\-]+\s+Str\./[A-ZÄÖÜ][A-Za-zÄÖÜäöüß\-]+\-Str\.)"])

    # Nodes after upgrade: parse table on page 3 first.
    nodes: Dict[str, NodeData] = {}
    page3 = pages[2] if len(pages) >= 3 else text
    m = re.search(
        r"Bezeichnung FN neu\s+([0-9A-Z]+-[0-9]+)\s+([0-9A-Z]+-[0-9]+).*?HP\s+(\d+)\s+(\d+).*?Anzahl Kunden Ki\s+(\d+)\s+(\d+)",
        page3,
        re.I | re.S,
    )
    if m:
        n1, n2 = m.group(1), m.group(2)
        nodes[n1] = NodeData(name=n1, hp=parse_int(m.group(3)), hsi=parse_int(m.group(5)))
        nodes[n2] = NodeData(name=n2, hp=parse_int(m.group(4)), hsi=parse_int(m.group(6)))

    # US details from Linientechnisches Konzept (page 6 preferred)
    page6 = pages[5] if len(pages) >= 6 else text
    for n in re.findall(r"([0-9A-Z]+-[0-9]+)\s+-\s+US\s*([12]).*?Portbezeichnung:\s*([0-9/]+(?:/[A-Z0-9]+)*)", page6, re.I | re.S):
        node_name, us_num, desc = n[0], n[1], n[2]
        node = nodes.setdefault(node_name, NodeData(name=node_name))
        if us_num == "1":
            node.us1_desc = desc
            node.us1_segment = normalize_segment(desc)
        else:
            node.us2_desc = desc
            node.us2_segment = normalize_segment(desc)

    # HSI-US counts from Netzschema; search around DS Domäne blocks.
    page4 = pages[3] if len(pages) >= 4 else text
    for node_name in list(nodes.keys()):
        block_match = re.search(rf"DS Domäne:\s*{re.escape(node_name)}([\s\S]{{0,240}}?)(?=DS Domäne:|Standort|$)", page4, re.I)
        if block_match:
            block = block_match.group(1)
            node = nodes[node_name]
            hp = parse_int(first_match(block, [r"HP\s*:\s*([0-9.]+)"]))
            hsi = parse_int(first_match(block, [r"HSI\s*:\s*([0-9.]+)"]))
            if hp is not None: node.hp = hp
            if hsi is not None: node.hsi = hsi
            for us_num, desc, count in re.findall(r"US\s*([12])\s*Beschreibung:\s*([0-9/]+(?:/[A-Z0-9]+)*)\s*HSI\s*[–-]\s*US\s*:\s*(\d+)", block, re.I):
                if us_num == "1":
                    node.us1_desc = desc
                    node.us1_hsi = parse_int(count)
                    node.us1_segment = normalize_segment(desc)
                else:
                    node.us2_desc = desc
                    node.us2_hsi = parse_int(count)
                    node.us2_segment = normalize_segment(desc)

    data.fn_new = sorted(nodes.values(), key=lambda x: x.name)

    # Warnings for missing key fields.
    for key in ["jira", "project_name", "hub", "target_fibernode", "plz", "ort", "street"]:
        if not getattr(data, key):
            data.raw_warnings.append(f"Nije pronađeno polje: {key}")
    if not data.fn_new:
        data.raw_warnings.append("Nisu pronađeni novi FN/Node podaci")

    return data


def fill_nc_list(template_path: Path, output_path: Path, data: ExtractedHLD) -> None:
    wb = load_workbook(template_path)
    ws = wb["NC- Zuordnungsliste"] if "NC- Zuordnungsliste" in wb.sheetnames else wb.active

    # Header fields
    if data.jira:
        ws["C3"] = data.jira
    if data.hub:
        ws["C4"] = data.hub
    if data.target_fibernode:
        ws["C5"] = data.target_fibernode

    # Data rows: keep the template's odd/even structure. Real DCA rows are 23, 25, 27...
    target_rows = [23, 25, 27, 29, 31, 33]
    for idx, node in enumerate(data.fn_new[:len(target_rows)]):
        row = target_rows[idx]
        ws.cell(row=row, column=2).value = idx + 1                       # Lfd. Nr.
        ws.cell(row=row, column=3).value = node.name                      # Alter GF-Node oder VrP-Name
        ws.cell(row=row, column=4).value = node.name                      # DCA-Name
        ws.cell(row=row, column=7).value = "DCA"                         # DCA / Daisychain / VNS
        ws.cell(row=row, column=9).value = "GF-Node"                     # bBk/VrP/GDR
        ws.cell(row=row, column=10).value = data.plz
        ws.cell(row=row, column=11).value = data.ort
        ws.cell(row=row, column=12).value = data.street
        ws.cell(row=row, column=13).value = node.us1_segment
        ws.cell(row=row, column=14).value = node.us2_segment
        ws.cell(row=row, column=15).value = parse_int(data.asb) if data.asb else None
        ws.cell(row=row, column=16).value = data.hub
        ws.cell(row=row, column=17).value = node.hp
        ws.cell(row=row, column=18).value = node.hsi
        ws.cell(row=row, column=19).value = node.us1_hsi
        ws.cell(row=row, column=20).value = node.us2_hsi
        ws.cell(row=row, column=24).value = round((data.gf_length_m or 0) / 1000) if data.gf_length_m else None
        ws.cell(row=row, column=25).value = 1
        ws.cell(row=row, column=27).value = node.us1_desc
        ws.cell(row=row, column=30).value = node.us2_desc

    # Add an audit sheet so the planner can verify what was found.
    if "AUTO_Extrakt" in wb.sheetnames:
        del wb["AUTO_Extrakt"]
    ex = wb.create_sheet("AUTO_Extrakt")
    ex.append(["Feld", "Wert"])
    for k, v in asdict(data).items():
        if k == "fn_new":
            continue
        ex.append([k, str(v)])
    ex.append([])
    ex.append(["Node", "HP", "HSI", "US1 Desc", "US1 HSI", "US2 Desc", "US2 HSI"])
    for n in data.fn_new:
        ex.append([n.name, n.hp, n.hsi, n.us1_desc, n.us1_hsi, n.us2_desc, n.us2_hsi])
    for col in range(1, 8):
        ex.column_dimensions[chr(64 + col)].width = 22

    wb.save(output_path)


def run(pdf: Path, template: Path, output: Path) -> ExtractedHLD:
    data = extract_hld(pdf)
    fill_nc_list(template, output, data)
    return data


def launch_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.title("HLD -> NC-Liste Automation")
    root.geometry("650x250")

    pdf_var = tk.StringVar()
    tpl_var = tk.StringVar()
    out_var = tk.StringVar()

    def pick_pdf():
        p = filedialog.askopenfilename(title="HLD PDF auswählen", filetypes=[("PDF", "*.pdf")])
        if p:
            pdf_var.set(p)
            if not out_var.get():
                out_var.set(str(Path(p).with_name(Path(p).stem + "_NC_AUTO.xlsx")))

    def pick_tpl():
        p = filedialog.askopenfilename(title="NC-Liste Template auswählen", filetypes=[("Excel", "*.xlsx")])
        if p:
            tpl_var.set(p)

    def pick_out():
        p = filedialog.asksaveasfilename(title="Ausgabe speichern", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            out_var.set(p)

    def start():
        try:
            if not pdf_var.get() or not tpl_var.get() or not out_var.get():
                messagebox.showwarning("Fehlt", "Bitte PDF, NC-Liste und Ausgabe auswählen.")
                return
            data = run(Path(pdf_var.get()), Path(tpl_var.get()), Path(out_var.get()))
            msg = f"Fertig!\n\nDatei: {out_var.get()}\n\nGefunden: {len(data.fn_new)} Node(s)"
            if data.raw_warnings:
                msg += "\n\nHinweise:\n- " + "\n- ".join(data.raw_warnings)
            messagebox.showinfo("Erfolg", msg)
        except Exception as exc:
            messagebox.showerror("Fehler", str(exc))

    def row(label, var, cmd, r):
        tk.Label(root, text=label, anchor="w", width=18).grid(row=r, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(root, textvariable=var, width=62).grid(row=r, column=1, padx=5, pady=10)
        tk.Button(root, text="...", command=cmd).grid(row=r, column=2, padx=5, pady=10)

    row("HLD PDF", pdf_var, pick_pdf, 0)
    row("NC-Liste Template", tpl_var, pick_tpl, 1)
    row("Ausgabe", out_var, pick_out, 2)
    tk.Button(root, text="NC-Liste generieren", command=start, height=2, width=25).grid(row=3, column=1, pady=20)
    root.mainloop()
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Extract HLD PDF data and fill Vodafone NC-Liste template.")
    parser.add_argument("--pdf", type=Path, help="HLD PDF")
    parser.add_argument("--template", type=Path, help="NC-Liste XLSX template")
    parser.add_argument("--output", type=Path, help="Output XLSX")
    args = parser.parse_args(argv)

    if not args.pdf and not args.template and not args.output:
        return launch_gui()
    if not args.pdf or not args.template or not args.output:
        parser.error("CLI mode needs --pdf, --template and --output")

    data = run(args.pdf, args.template, args.output)
    print(f"OK: {args.output}")
    if data.raw_warnings:
        print("Hinweise:")
        for warning in data.raw_warnings:
            print(" -", warning)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
