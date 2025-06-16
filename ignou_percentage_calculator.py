import os
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from bs4 import BeautifulSoup
import pandas as pd
from fpdf import FPDF
import tempfile
from webdriver_manager.chrome import ChromeDriverManager
import logging
import time

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Streamlit setup
st.set_page_config(page_title="IGNOU Grade Card Automation", layout="centered")
st.title("🎓 IGNOU Grade Card % Calculator")

# Inputs
enrollment = st.text_input("Enrollment Number", max_chars=10)  # Removed default value
gradecard_for = st.selectbox("Gradecard For", [
    ("1", "BCA/MCA/MP/PGDCA etc."),
    ("2", "BDP/BA/B.COM/B.Sc./ASSO Programmes"),
    ("3", "CBCS Programmes"),
    ("4", "Other Programmes")
], format_func=lambda x: x[1], index=0)

valid_programs = [
    "BCA", "BCAOL", "BCA_NEW", "BCA_NEWOL", "MBF", "MCA", "MCAOL",
    "MCA_NEW", "MCA_NEWOL", "MP", "MPB", "PGDCA", "PGDCA_NEW",
    "PGDHRM", "PGDFM", "PGDOM", "PGDMM", "PGDFMP"
]
program_code = st.selectbox("Programme Code", valid_programs, index=valid_programs.index("MCAOL"))

# Initialize session state for button
if "processing" not in st.session_state:
    st.session_state.processing = False

if st.button("🚀 Fetch Grade Card", disabled=st.session_state.processing or not enrollment):
    st.session_state.processing = True
    driver = None
    max_retries = 2
    retry_count = 0

    while retry_count <= max_retries:
        try:
            # Validate enrollment number
            if not enrollment.isdigit() or len(enrollment) not in [9, 10]:
                st.error("❌ Enrollment number must be 9 or 10 digits.")
                st.session_state.processing = False
                st.stop()

            logging.info("Starting grade card fetch for enrollment: %s (Attempt %d/%d)", enrollment, retry_count + 1, max_retries + 1)

            # Setup Selenium
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            wait = WebDriverWait(driver, 30)

            # Log Chrome version
            chrome_version = driver.capabilities['browserVersion']
            logging.info("Using Chrome version: %s", chrome_version)

            # Navigate to IGNOU grade card page
            driver.get("https://gradecard.ignou.ac.in/gradecard/")
            logging.info("Navigated to IGNOU grade card page")

            # Select grade card type and program
            Select(wait.until(EC.presence_of_element_located((By.ID, "ddlGradecardfor")))).select_by_value(gradecard_for[0])
            Select(wait.until(EC.presence_of_element_located((By.ID, "ddlProgram")))).select_by_value(program_code)
            driver.find_element(By.ID, "txtEnrno").send_keys(enrollment)

            # Click login button
            btn = wait.until(EC.element_to_be_clickable((By.ID, "btnlogin")))
            driver.execute_script("arguments[0].click();", btn)
            logging.info("Submitted form")

            # Wait for results table
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvDetail")))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            logging.info("Parsed page source")

            # Check for error messages or CAPTCHAs
            error_message = soup.find("span", {"id": "ctl00_ContentPlaceHolder1_lblMsg"})
            if error_message and error_message.text.strip():
                st.error(f"❌ IGNOU website error: {error_message.text.strip()}")
                st.session_state.processing = False
                st.stop()

            # Extract table
            table = soup.find("table", {"id": "ctl00_ContentPlaceHolder1_gvDetail"})
            if not table:
                with open("page_source.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                logging.error("Grade card table not found. Page source saved to page_source.html")
                st.error("❌ Grade card table not found on the page. Page source saved to page_source.html for debugging.")
                st.session_state.processing = False
                st.stop()

            headers = [th.text.strip() for th in table.find_all("th")]
            rows = []
            for tr in table.find_all("tr")[1:]:
                cols = [td.text.strip() for td in tr.find_all("td")]
                if len(cols) == len(headers):
                    rows.append(cols)

            if not rows:
                st.error("❌ No valid data found in the grade card table.")
                st.session_state.processing = False
                st.stop()

            df = pd.DataFrame(rows, columns=headers)

            # Ensure COURSE column is string type and clean it
            if "COURSE" in df.columns:
                df["COURSE"] = df["COURSE"].astype(str).fillna("")
            else:
                st.error("❌ COURSE column missing in grade card table.")
                st.session_state.processing = False
                st.stop()

            # Convert columns to numeric
            for col in ["Asgn1", "TERM END THEORY", "TERM END PRACTICAL"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col].replace(["-", "N/A", ""], 0), errors='coerce').fillna(0)
                else:
                    st.warning(f"⚠️ Column {col} missing; assuming 0 for all rows.")
                    df[col] = 0

            # Filter completed courses and exclude non-MCSL lab courses
            df_calc = df[df["STATUS"] == "COMPLETED"].copy()
            df_calc = df_calc[df_calc["COURSE"].str.startswith("MCSL") | ~df_calc["COURSE"].str.contains("lab", case=False, na=False)]

            # Calculate scores
            df_calc["30% Assignments"] = df_calc["Asgn1"] * 0.3
            df_calc["70% Theory"] = df_calc.apply(
                lambda row: row["TERM END PRACTICAL"] * 0.7 if row["COURSE"].startswith("MCSL") else row["TERM END THEORY"] * 0.7, axis=1
            )
            df_calc["Total (A+B)"] = df_calc["30% Assignments"] + df_calc["70% Theory"]

            # Calculate totals
            totals = {
                "COURSE": "Total",
                "Asgn1": df_calc["Asgn1"].sum(),
                "TERM END THEORY": df_calc["TERM END THEORY"].sum(),
                "TERM END PRACTICAL": df_calc["TERM END PRACTICAL"].sum(),
                "30% Assignments": df_calc["30% Assignments"].sum(),
                "70% Theory": df_calc["70% Theory"].sum(),
                "Total (A+B)": df_calc["Total (A+B)"].sum()
            }

            # Calculate total possible marks and percentage
            num_subjects = len(df_calc)
            total_possible_marks = num_subjects * 100
            total_obtained_marks = df_calc["Total (A+B)"].sum()
            percentage = round((total_obtained_marks / total_possible_marks) * 100, 2) if total_possible_marks > 0 else 0

            # Prepare display DataFrame with totals
            df_calc_display = pd.concat([df_calc, pd.DataFrame([totals])], ignore_index=True)
            # Adjust index to start from 1
            df_calc_display.index = df_calc_display.index + 1

            # Display results with fixed column widths
            st.success("✅ Grade Card Parsed and Calculated!")
            st.subheader("✅ Completed Subjects (Used in Calculation)")
            column_config = {
                "COURSE": st.column_config.TextColumn(width=100),
                "Asgn1": st.column_config.NumberColumn(width=80, format="%.0f"),
                "TERM END THEORY": st.column_config.NumberColumn(width=100, format="%.0f"),
                "TERM END PRACTICAL": st.column_config.NumberColumn(width=100, format="%.0f"),
                "30% Assignments": st.column_config.NumberColumn(width=100, format="%.2f"),
                "70% Theory": st.column_config.NumberColumn(width=100, format="%.2f"),
                "Total (A+B)": st.column_config.NumberColumn(width=100, format="%.2f")
            }
            with st.container():
                st.dataframe(df_calc_display, use_container_width=True, column_config=column_config)

            # Display totals and percentage
            st.subheader("📊 Summary")
            st.metric("Final Percentage", f"{percentage}%")
            st.write(f"**Total Obtained Marks**: {total_obtained_marks:.2f} / {total_possible_marks:.0f}")
            st.write(f"**Total Assignment Marks**: {totals['Asgn1']:.0f}")
            st.write(f"**Total Theory Marks**: {totals['TERM END THEORY']:.0f}")
            st.write(f"**Total Practical Marks**: {totals['TERM END PRACTICAL']:.0f}")

            # Display incomplete courses
            df_incomplete = df[df["STATUS"] != "COMPLETED"]
            if not df_incomplete.empty:
                st.subheader("⚠️ Not Completed / Incomplete Subjects")
                st.dataframe(df_incomplete[["COURSE", "STATUS", "Asgn1", "TERM END THEORY", "TERM END PRACTICAL"]], use_container_width=True)

            # Generate PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(200, 10, "IGNOU Grade Report", ln=True, align="C")
            pdf.ln(10)
            pdf.set_font("Arial", size=12)
            for _, row in df_calc.iterrows():
                course = row["COURSE"] or "Unknown"
                total = row["Total (A+B)"] or 0
                asgn = row["Asgn1"] or 0
                theory = row["TERM END THEORY"] or 0
                practical = row["TERM END PRACTICAL"] or 0
                pdf.cell(200, 10, f"{course}: {total:.2f} (Asgn: {asgn}, TEE: {theory}, PRACT: {practical})", ln=True)
            pdf.ln(10)
            pdf.set_font("Arial", "B", 12)
            pdf.cell(200, 10, f"Final Percentage: {percentage}%", ln=True)
            pdf.cell(200, 10, f"Total Obtained Marks: {total_obtained_marks:.2f} / {total_possible_marks:.0f}", ln=True)
            pdf.cell(200, 10, f"Total Assignment Marks: {totals['Asgn1']:.0f}", ln=True)
            pdf.cell(200, 10, f"Total Theory Marks: {totals['TERM END THEORY']:.0f}", ln=True)
            pdf.cell(200, 10, f"Total Practical Marks: {totals['TERM END PRACTICAL']:.0f}", ln=True)

            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            pdf.output(tmp_file.name)
            with open(tmp_file.name, "rb") as f:
                st.download_button("📄 Download PDF", f, file_name="ignou_grade_report.pdf")

            st.session_state.processing = False
            break

        except (TimeoutException, NoSuchElementException, ElementNotInteractableException) as e:
            retry_count += 1
            logging.error("Attempt %d failed: %s", retry_count, str(e))
            if retry_count > max_retries:
                st.error(f"⏳ Failed to interact with the IGNOU website after {max_retries + 1} attempts: {str(e)}")
                st.session_state.processing = False
            else:
                st.warning(f"⚠️ Attempt {retry_count} failed. Retrying in 5 seconds...")
                time.sleep(5)
        except Exception as e:
            st.error(f"❌ Failed: {str(e)}")
            st.session_state.processing = False
            break
        finally:
            if driver:
                try:
                    driver.quit()
                    logging.info("WebDriver closed successfully")
                except Exception as e:
                    logging.error("Error closing WebDriver: %s", str(e))