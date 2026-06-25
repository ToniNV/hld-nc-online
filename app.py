import streamlit as st

st.set_page_config(page_title="HLD → NC Liste", layout="wide")

st.title("HLD → NC Liste Generator")
st.write("Upload HLD PDF i NC Excel listu.")

pdf = st.file_uploader("HLD PDF", type=["pdf"])
xlsx = st.file_uploader("NC Excel lista", type=["xlsx"])

if pdf and xlsx:
    st.success("Datoteke su učitane. Sljedeći korak je obrada i generiranje Excel liste.")
