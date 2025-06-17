import os
import shutil
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException, WebDriverException
from bs4 import BeautifulSoup
import pandas as pd
from fpdf import FPDF
import tempfile
import logging
import time
import subprocess

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Clear ChromeDriver cache
cache_path = "/home/appuser/.wdm"
if os.path.exists(cache_path):
    shutil.rmtree(cache_path)
    logging.info("Cleared ChromeDriver cache")

# Streamlit setup
st.set_page_config(
    page_title="IGNOU Grade Card Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        margin-top: 1rem;
    }
    .stDataFrame {
        width: 100%;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .stAlert {
        padding: 1rem;
        border-radius: 0.5rem;
    }
    .summary-box {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Title with custom styling
st.markdown("""
    <h1 style='text-align: center; color: #1E88E5; margin-bottom: 2rem;'>
        🎓 IGNOU Grade Card Calculator
    </h1>
""", unsafe_allow_html=True)

# Create two columns for input fields
col1, col2 = st.columns(2)

with col1:
    enrollment = st.text_input(
        "Enrollment Number",
        max_chars=10,
        placeholder="Enter 9 or 10-digit enrollment number",
        help="Enter your 9 or 10-digit IGNOU enrollment number"
    )
    gradecard_for = st.selectbox(
        "Gradecard For",
        [
            ("1", "BCA/MCA/MP/PGDCA etc."),
            ("2", "BDP/BA/B.COM/B.Sc./ASSO Programmes"),
            ("3", "CBCS Programmes"),
            ("4", "Other Programmes")
        ],
        format_func=lambda x: x[1],
        index=0,
        help="Select your program category"
    )

with col2:
    valid_programs = [
        "BCA", "BCAOL", "BCA_NEW", "BCA_NEWOL", "MBF", "MCA", "MCAOL",
        "MCA_NEW", "MCA_NEWOL", "MP", "MPB", "PGDCA", "PGDCA_NEW",
        "PGDHRM", "PGDFM", "PGDOM", "PGDMM", "PGDFMP"
    ]
    program_code = st.selectbox(
        "Programme Code",
        valid_programs,
        index=valid_programs.index("MCAOL"),
        help="Select your program code"
    )

# Initialize session state for button
if "processing" not in st.session_state:
    st.session_state.processing = False

# Function to find Chromium binary
def find_chromium_binary():
    possible_paths = [
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/lib/chromium-browser/chromium",
        "/usr/lib/chromium/chromium"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            logging.info("Found Chromium binary at: %s", path)
            return path
    logging.error("No Chromium binary found in paths: %s", possible_paths)
    return None

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

            # Find Chromium binary
            binary_path = find_chromium_binary()
            if not binary_path:
                raise WebDriverException("No Chromium binary found. Please ensure chromium is installed.")
            chrome_options.binary_location = binary_path

            # Use system ChromeDriver if available
            chromedriver_path = "/usr/bin/chromedriver"
            if os.path.exists(chromedriver_path):
                service = Service(chromedriver_path)
                logging.info("Using system ChromeDriver at: %s", chromedriver_path)
            else:
                from webdriver_manager.chrome import ChromeDriverManager
                service = Service(ChromeDriverManager().install())
                logging.info("Using webdriver-manager to install ChromeDriver")

            logging.info("Initializing ChromeDriver")
            driver = webdriver.Chrome(service=service, options=chrome_options)
            wait = WebDriverWait(driver, 60)  # Increased timeout to 60s

            # Log Chrome and ChromeDriver versions
            chrome_version = driver.capabilities['browserVersion']
            chromedriver_version = driver.capabilities['chrome']['chromedriverVersion'].split(' ')[0]
            logging.info("Using Chromium version: %s, ChromeDriver version: %s", chrome_version, chromedriver_version)

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

            # Wait for results table or error message
            wait.until(EC.any_of(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvDetail")),
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_lblMsg"))
            ))
            soup = BeautifulSoup(driver.page_source, "html.parser")
            logging.info("Parsed page source")

            # Check for CAPTCHA
            if soup.find("div", {"id": "captcha"}) or "captcha" in driver.page_source.lower():
                st.error("❌ CAPTCHA detected. Please try again later or access the website manually to verify.")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.info("Page source saved to: %s", tmp_file.name)
                    st.write(f"Page source saved to: {tmp_file.name}")
                st.session_state.processing = False
                st.stop()

            # Check for error messages
            error_message = soup.find("span", {"id": "ctl00_ContentPlaceHolder1_lblMsg"})
            if error_message and error_message.text.strip():
                st.error(f"❌ IGNOU website error: {error_message.text.strip()}")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.info("Page source saved to: %s", tmp_file.name)
                    st.write(f"Page source saved to: {tmp_file.name}")
                st.session_state.processing = False
                st.stop()

            # Extract table
            table = soup.find("table", {"id": "ctl00_ContentPlaceHolder1_gvDetail"})
            if not table:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.error("Grade card table not found. Page source saved to: %s", tmp_file.name)
                    st.error(f"❌ Grade card table not found. Check if enrollment ({enrollment}) and program ({program_code}) are valid. Page source saved to: {tmp_file.name}")
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
            df_calc_display.index = df_calc_display.index + 1  # Start serial number from 1

            # Display results with improved layout
            st.success("✅ Grade Card Parsed and Calculated!")
            
            # Summary section in a nice box
            st.markdown('<div class="summary-box">', unsafe_allow_html=True)
            st.subheader("📊 Summary")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Final Percentage", f"{percentage}%")
                st.write(f"**Total Obtained Marks**: {total_obtained_marks:.2f} / {total_possible_marks:.0f}")
            with col2:
                st.write(f"**Total Assignment Marks**: {totals['Asgn1']:.0f}")
                st.write(f"**Total Theory Marks**: {totals['TERM END THEORY']:.0f}")
                st.write(f"**Total Practical Marks**: {totals['TERM END PRACTICAL']:.0f}")
            st.markdown('</div>', unsafe_allow_html=True)

            # Completed subjects table
            st.subheader("✅ Completed Subjects")
            column_config = {
                "COURSE": st.column_config.TextColumn(
                    "Course",
                    width="medium",
                    help="Course Code"
                ),
                "Asgn1": st.column_config.NumberColumn(
                    "Assignment",
                    width="small",
                    format="%.0f",
                    help="Assignment Marks"
                ),
                "TERM END THEORY": st.column_config.NumberColumn(
                    "Theory",
                    width="small",
                    format="%.0f",
                    help="Theory Marks"
                ),
                "TERM END PRACTICAL": st.column_config.NumberColumn(
                    "Practical",
                    width="small",
                    format="%.0f",
                    help="Practical Marks"
                ),
                "30% Assignments": st.column_config.NumberColumn(
                    "30% Assignment",
                    width="small",
                    format="%.2f",
                    help="30% of Assignment Marks"
                ),
                "70% Theory": st.column_config.NumberColumn(
                    "70% Theory/Practical",
                    width="small",
                    format="%.2f",
                    help="70% of Theory/Practical Marks"
                ),
                "Total (A+B)": st.column_config.NumberColumn(
                    "Total",
                    width="small",
                    format="%.2f",
                    help="Total Marks"
                )
            }

            st.dataframe(
                df_calc_display,
                use_container_width=True,
                column_config=column_config,
                hide_index=False
            )

            # Incomplete subjects table
            df_incomplete = df[df["STATUS"] != "COMPLETED"]
            if not df_incomplete.empty:
                st.subheader("⚠️ Not Completed / Incomplete Subjects")
                incomplete_config = {
                    "COURSE": st.column_config.TextColumn(
                        "Course",
                        width="medium",
                        help="Course Code"
                    ),
                    "STATUS": st.column_config.TextColumn(
                        "Status",
                        width="small",
                        help="Course Status"
                    ),
                    "Asgn1": st.column_config.NumberColumn(
                        "Assignment",
                        width="small",
                        format="%.0f",
                        help="Assignment Marks"
                    ),
                    "TERM END THEORY": st.column_config.NumberColumn(
                        "Theory",
                        width="small",
                        format="%.0f",
                        help="Theory Marks"
                    ),
                    "TERM END PRACTICAL": st.column_config.NumberColumn(
                        "Practical",
                        width="small",
                        format="%.0f",
                        help="Practical Marks"
                    )
                }
                st.dataframe(
                    df_incomplete,
                    use_container_width=True,
                    column_config=incomplete_config,
                    hide_index=True
                )

            # PDF download button in a centered container
            st.markdown('<div style="text-align: center;">', unsafe_allow_html=True)
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
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
            pdf.output(tmp_file.name)
            with open(tmp_file.name, "rb") as f:
                st.download_button(
                    "📄 Download PDF Report",
                    f,
                    file_name="ignou_grade_report.pdf",
                    use_container_width=False
                )
            st.markdown('</div>', unsafe_allow_html=True)

            st.session_state.processing = False
            break

        except WebDriverException as e:
            retry_count += 1
            logging.error("Attempt %d failed due to WebDriver issue: %s", retry_count, str(e))
            if retry_count > max_retries:
                st.error(f"⏳ Failed to initialize ChromeDriver after {max_retries + 1} attempts: {str(e)}. Please ensure chromium is installed correctly.")
                try:
                    result = subprocess.run(["chromium", "--version"], capture_output=True, text=True)
                    logging.info("Chromium version check: %s", result.stdout or result.stderr)
                except Exception as debug_e:
                    logging.error("Failed to check chromium version: %s", str(debug_e))
                st.session_state.processing = False
            else:
                st.warning(f"⚠️ Attempt {retry_count} failed. Retrying in 5 seconds...")
                time.sleep(5)
        except (TimeoutException, NoSuchElementException, ElementNotInteractableException) as e:
            retry_count += 1
            logging.error("Attempt %d failed: %s", retry_count, str(e))
            if retry_count > max_retries:
                st.error(f"⏳ Failed to interact with the IGNOU website after {max_retries + 1} attempts: {str(e)}")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.info("Page source saved to: %s", tmp_file.name)
                    st.write(f"Page source saved to: {tmp_file.name}")
                st.session_state.processing = False
            else:
                st.warning(f"⚠️ Attempt {retry_count} failed. Retrying in 5 seconds...")
                time.sleep(5)
        except Exception as e:
            st.error(f"❌ Failed: {str(e)}")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
                tmp_file.write(driver.page_source.encode("utf-8"))
                logging.info("Page source saved to: %s", tmp_file.name)
                st.write(f"Page source saved to: {tmp_file.name}")
            st.session_state.processing = False
            break
        finally:
            if driver:
                try:
                    driver.quit()
                    logging.info("WebDriver closed successfully")
                except Exception as e:
                    logging.error("Error closing WebDriver: %s", str(e))