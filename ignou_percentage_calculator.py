import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
from fpdf import FPDF
import tempfile

# UI setup
st.set_page_config(page_title="IGNOU Grade Card", layout="centered")
st.title("🎓 IGNOU Grade Card Automation (Streamlit Cloud Safe)")

# Inputs
enrollment = st.text_input("Enrollment Number", max_chars=10)
gradecard_for = st.selectbox("Gradecard For", [
    ("1", "BCA/MCA/MP/PGDCA etc."),
    ("2", "BDP/BA/B.COM/B.Sc./ASSO Programmes"),
    ("3", "CBCS Programmes"),
    ("4", "Other Programmes")
], format_func=lambda x: x[1], index=0)

program_code = st.selectbox("Programme Code", [
    "BCA", "BCAOL", "BCA_NEW", "BCA_NEWOL", "MBF", "MCA", "MCAOL",
    "MCA_NEW", "MCA_NEWOL", "MP", "MPB", "PGDCA", "PGDCA_NEW",
    "PGDHRM", "PGDFM", "PGDOM", "PGDMM", "PGDFMP"
], index=5)

if st.button("🚀 Fetch Grade Card") and enrollment:
    try:
        # Build request
        url = "https://gradecard.ignou.ac.in/gradecardR.asp"
        payload = {
            "eno": enrollment,
            "prog": program_code,
            "Grade": gradecard_for[0],
            "submit": "Submit"
        }
        headers = {"User-Agent": "Mozilla/5.0"}

        # Send POST request
        response = requests.post(url, data=payload, headers=headers, timeout=20)

        # Parse HTML
        soup = BeautifulSoup(response.content, "html.parser")
        table = soup.find("table", id="ctl00_ContentPlaceHolder1_gvDetail") or soup.find("table")

        if not table:
            st.error("❌ Grade card table not found. Check if enrollment/program is valid.")
            st.stop()

        # Parse headers and rows
        headers = [th.get_text(strip=True) for th in table.find_all("th")]
        rows = []
        for tr in table.find_all("tr")[1:]:
            cols = [td.get_text(strip=True) for td in tr.find_all("td")]
            if len(cols) == len(headers):
                rows.append(cols)

        if not rows:
            st.warning("No valid rows found.")
            st.stop()

        df = pd.DataFrame(rows, columns=headers)

        # Process data
        df["Asgn1"] = pd.to_numeric(df["Asgn1"].replace(["-", "N/A", ""], 0), errors='coerce')
        df["TERM END THEORY"] = pd.to_numeric(df["TERM END THEORY"].replace(["-", "N/A", ""], 0), errors='coerce')
        df["TERM END PRACTICAL"] = pd.to_numeric(df["TERM END PRACTICAL"].replace(["-", "N/A", ""], 0), errors='coerce')
        df["COURSE"] = df["COURSE"].fillna("").astype(str)

        df_completed = df[df["STATUS"] == "COMPLETED"]
        df_non_lab = df_completed[~df_completed["COURSE"].str.startswith("MCSL")]

        df_calc = df_non_lab.copy()
        df_calc["30% Assignments"] = df_calc["Asgn1"] * 0.3
        df_calc["70% Theory"] = df_calc["TERM END THEORY"] * 0.7
        df_calc["Total (A+B)"] = df_calc["30% Assignments"] + df_calc["70% Theory"]

        total_subjects = len(df_calc)
        total_marks = total_subjects * 100
        obtained = df_calc["Total (A+B)"].sum()
        percentage = round((obtained / total_marks) * 100, 2) if total_marks else 0

        # Display table
        st.success("✅ Grade card fetched successfully!")
        st.dataframe(df_calc, use_container_width=True)
        st.metric("Final Percentage", f"{percentage}%")

        # PDF download
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, "IGNOU Grade Report", ln=True, align="C")
        for _, row in df_calc.iterrows():
            pdf.cell(200, 10, f"{row['COURSE']}: {row['Total (A+B)']:.2f}", ln=True)
        pdf.cell(200, 10, f"Final Percentage: {percentage}%", ln=True)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf.output(tmp.name)
        with open(tmp.name, "rb") as f:
            st.download_button("📄 Download PDF", f, file_name="grade_report.pdf")

    except Exception as e:
        st.error(f"❌ Failed: {str(e)}")
