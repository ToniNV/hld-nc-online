from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from hld_nc_automation import extract_hld, fill_nc_list

st.set_page_config(page_title="HLD → NC Liste", page_icon="📄", layout="centered")

st.title("📄 HLD → NC Liste Generator")
st.caption("Upload HLD PDF + NC Excel template → download popunjene NC liste")

st.info(
    "Ova verzija je V1: čita standardni Vodafone/Telefonica LNS HLD PDF, "
    "popunjava osnovna polja u NC listi i dodaje sheet AUTO_Extrakt za kontrolu."
)

hld_pdf = st.file_uploader("1) Upload HLD PDF", type=["pdf"])
nc_template = st.file_uploader("2) Upload prazne / postojeće NC liste (.xlsx)", type=["xlsx"])

with st.expander("Šta se trenutno pokušava izvući iz PDF-a"):
    st.write(
        "JIRA, Projektname, Fibernode, HUB, VrP, PLZ, Ort, Straße, GF-Länge, "
        "FN neu, HP, HSI, US1/US2 Portbezeichnung i HSI-US vrijednosti."
    )

if hld_pdf and nc_template:
    if st.button("🔍 Analiziraj PDF i generiraj NC listu", type="primary"):
        with st.spinner("Obrađujem PDF i Excel..."):
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                pdf_path = tmp_path / hld_pdf.name
                template_path = tmp_path / nc_template.name
                output_path = tmp_path / f"{Path(nc_template.name).stem}_AUTO.xlsx"

                pdf_path.write_bytes(hld_pdf.getvalue())
                template_path.write_bytes(nc_template.getvalue())

                data = extract_hld(pdf_path)
                fill_nc_list(template_path, output_path, data)
                output_bytes = output_path.read_bytes()

        st.success("Gotovo — NC lista je generirana.")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Pronađeni Node/FN", len(data.fn_new))
        with col2:
            st.metric("JIRA", data.jira or "nije pronađeno")

        st.subheader("Kontrola pronađenih podataka")
        st.write({
            "Projekt": data.project_name,
            "Fibernode": data.target_fibernode,
            "HUB": data.hub,
            "VrP": data.vrp,
            "Adresa": f"{data.street}, {data.plz} {data.ort}".strip(", "),
            "GF Länge": data.gf_length_m,
        })

        if data.fn_new:
            st.write("Node podaci:")
            st.dataframe([
                {
                    "Node": n.name,
                    "HP": n.hp,
                    "HSI": n.hsi,
                    "US1": n.us1_desc,
                    "US1 HSI": n.us1_hsi,
                    "US2": n.us2_desc,
                    "US2 HSI": n.us2_hsi,
                }
                for n in data.fn_new
            ], use_container_width=True)

        if data.raw_warnings:
            st.warning("Provjeri ova polja:\n" + "\n".join(f"- {w}" for w in data.raw_warnings))

        st.download_button(
            "⬇️ Download popunjene NC liste",
            data=output_bytes,
            file_name=output_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.write("Uploadaj oba fajla pa klikni Generate.")

st.divider()
st.caption("V1 prototip. Prije slanja dalje uvijek provjeri sheet AUTO_Extrakt i popunjena polja u NC listi.")
