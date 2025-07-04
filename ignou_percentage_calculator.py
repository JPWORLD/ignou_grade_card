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
import uuid
import atexit
from datetime import datetime, timedelta
import threading
from queue import Queue

# Function to create temporary file with session ID and better error handling
def create_temp_file(suffix):
    temp_dir = tempfile.gettempdir()
    max_retries = 3
    for attempt in range(max_retries):
        try:
            file_name = f"ignou_grade_{st.session_state.session_id}_{uuid.uuid4()}{suffix}"
            file_path = os.path.join(temp_dir, file_name)
            # Create empty file to ensure we have write permissions
            with open(file_path, 'w') as f:
                pass
            st.session_state.temp_files.append(file_path)
            return file_path
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(0.1)

# Enhanced resource management with better concurrency support
class ResourceManager:
    def __init__(self):
        self.active_drivers = {}
        self.lock = threading.Lock()
        self.driver_semaphore = threading.Semaphore(5)  # Limit concurrent drivers
        self.last_cleanup = time.time()
        self.cleanup_interval = 300  # 5 minutes
    
    def add_driver(self, session_id, driver):
        with self.lock:
            # Cleanup old drivers periodically
            current_time = time.time()
            if current_time - self.last_cleanup > self.cleanup_interval:
                self._cleanup_old_drivers()
                self.last_cleanup = current_time
            
            # Acquire semaphore before adding new driver
            self.driver_semaphore.acquire()
            self.active_drivers[session_id] = {
                'driver': driver,
                'created_at': current_time,
                'last_used': current_time
            }
    
    def remove_driver(self, session_id):
        with self.lock:
            if session_id in self.active_drivers:
                try:
                    self.active_drivers[session_id]['driver'].quit()
                except Exception as e:
                    logging.error(f"Error closing driver for session {session_id}: {e}")
                del self.active_drivers[session_id]
                self.driver_semaphore.release()
    
    def _cleanup_old_drivers(self):
        current_time = time.time()
        expired_sessions = []
        for session_id, driver_info in self.active_drivers.items():
            # Cleanup drivers older than 10 minutes or inactive for 5 minutes
            if (current_time - driver_info['created_at'] > 600 or 
                current_time - driver_info['last_used'] > 300):
                expired_sessions.append(session_id)
        
        for session_id in expired_sessions:
            self.remove_driver(session_id)
    
    def update_last_used(self, session_id):
        with self.lock:
            if session_id in self.active_drivers:
                self.active_drivers[session_id]['last_used'] = time.time()
    
    def cleanup_all(self):
        with self.lock:
            for session_id in list(self.active_drivers.keys()):
                self.remove_driver(session_id)

# Enhanced rate limiting with better concurrency support
class RateLimiter:
    def __init__(self, max_requests=10, time_window=60):
        self.max_requests = max_requests
        self.time_window = time_window
        self.requests = []
        self.lock = threading.Lock()
    
    def check_rate_limit(self):
        with self.lock:
            current_time = time.time()
            # Remove old requests
            self.requests = [req_time for req_time in self.requests 
                           if current_time - req_time < self.time_window]
            
            if len(self.requests) >= self.max_requests:
                return False
            
            self.requests.append(current_time)
            return True

# Initialize rate limiter
rate_limiter = RateLimiter(max_requests=10, time_window=60)

# Setup logging with session ID and rotation
def setup_logging():
    # Get client IP address using new query_params
    client_ip = st.query_params.get("client_ip", "unknown")
    if client_ip == "unknown":
        try:
            import socket
            client_ip = socket.gethostbyname(socket.gethostname())
        except:
            client_ip = "unknown"
    
    # Create unique session ID using IP and timestamp
    session_id = f"{client_ip}_{int(time.time())}_{str(uuid.uuid4())[:8]}"
    
    # Setup minimal logging
    logging.basicConfig(level=logging.INFO)
    return session_id

# Initialize Streamlit page config first
st.set_page_config(
    page_title="IGNOU Grade Card Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state variables
if "session_id" not in st.session_state:
    st.session_state.session_id = setup_logging()
if "temp_files" not in st.session_state:
    st.session_state.temp_files = []
if "processing" not in st.session_state:
    st.session_state.processing = False
if "last_request_time" not in st.session_state:
    st.session_state.last_request_time = None
if "retry_count" not in st.session_state:
    st.session_state.retry_count = 0

resource_manager = ResourceManager()

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
    .stMetric [data-testid="stMetricValue"] {
        color: #1E88E5;
        font-size: 2rem;
        font-weight: bold;
    }
    .stMetric [data-testid="stMetricLabel"] {
        color: #262730;
        font-size: 1.2rem;
        font-weight: 500;
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

# Initialize input variables
enrollment = ""
gradecard_for = None
program_code = None

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

# Function to log enrollment number
def log_enrollment(enrollment_number):
    logging.info(f"Processing enrollment number: {enrollment_number}")

# Extract student details from the page
def extract_student_details(soup):
    try:
        # Find the student details table
        details_table = soup.find("table", {"id": "ctl00_ContentPlaceHolder1_gvDetail"})
        if details_table:
            # Get the first row which contains student details
            first_row = details_table.find("tr")
            if first_row:
                cells = first_row.find_all("td")
                if len(cells) >= 3:
                    return {
                        "enrollment": cells[0].text.strip(),
                        "name": cells[1].text.strip(),
                        "program": cells[2].text.strip()
                    }
    except Exception as e:
        logging.error(f"Error extracting student details: {str(e)}")
    return None

# Add new helper functions for robust element interaction
def wait_for_page_load(driver, timeout=30):
    """Wait for the page to be fully loaded and interactive"""
    try:
        # Wait for document ready state
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        # Additional wait for jQuery if present
        try:
            WebDriverWait(driver, 5).until(
                lambda d: d.execute_script('return jQuery.active == 0')
            )
        except:
            pass  # jQuery not present, continue
        # Small delay to ensure dynamic content is loaded
        time.sleep(1)
    except TimeoutException:
        logging.warning("Page load timeout - continuing anyway")

def wait_and_find_element(driver, by, value, timeout=10, clickable=False):
    """Wait for an element to be present and optionally clickable"""
    try:
        # First check if element exists in DOM
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        # Then check if it's visible
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, value))
        )
        # Finally check if it's clickable if requested
        if clickable:
            WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
        return element
    except TimeoutException:
        logging.error(f"Element not found or not ready: {by}={value}")
        return None

def safe_click(driver, element):
    """Safely click an element using JavaScript if regular click fails"""
    try:
        element.click()
    except ElementNotInteractableException:
        driver.execute_script("arguments[0].click();", element)

def ensure_page_loaded(driver):
    """Ensure the page is properly loaded before proceeding"""
    try:
        # Check if we're on the correct page
        if "gradecard.ignou.ac.in" not in driver.current_url:
            driver.get("https://gradecard.ignou.ac.in/gradecard/")
            wait_for_page_load(driver)
        
        # Wait for any loading indicators to disappear
        try:
            WebDriverWait(driver, 5).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "loading"))
            )
        except:
            pass  # No loading indicator found, continue
        
        # Verify page title or some key element
        if not driver.find_elements(By.ID, "ddlGradecardfor"):
            raise WebDriverException("Grade card page not properly loaded")
            
    except Exception as e:
        logging.error(f"Page load verification failed: {str(e)}")
        raise WebDriverException("Failed to load grade card page properly")

# Update the main processing block
if st.button("🚀 Fetch Grade Card", disabled=st.session_state.processing or not enrollment):
    if not rate_limiter.check_rate_limit():
        st.error("⚠️ Too many requests. Please wait a minute before trying again.")
        st.stop()
    
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

            logging.info(f"Session {st.session_state.session_id} - Starting grade card fetch for enrollment: {enrollment} (Attempt {retry_count + 1}/{max_retries + 1})")

            # Setup Selenium with optimized options
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-notifications")
            chrome_options.add_argument("--disable-popup-blocking")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Add unique user data directory for each session
            user_data_dir = f"/tmp/chrome_profile_{st.session_state.session_id}"
            chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
            chrome_options.add_argument("--profile-directory=Default")

            # Find Chromium binary
            binary_path = find_chromium_binary()
            if not binary_path:
                raise WebDriverException("No Chromium binary found. Please ensure chromium is installed.")
            chrome_options.binary_location = binary_path

            # Use system ChromeDriver if available
            chromedriver_path = "/usr/bin/chromedriver"
            if os.path.exists(chromedriver_path):
                service = Service(chromedriver_path)
                logging.info(f"Session {st.session_state.session_id} - Using system ChromeDriver at: {chromedriver_path}")
            else:
                from webdriver_manager.chrome import ChromeDriverManager
                service = Service(ChromeDriverManager().install())
                logging.info(f"Session {st.session_state.session_id} - Using webdriver-manager to install ChromeDriver")

            logging.info(f"Session {st.session_state.session_id} - Initializing ChromeDriver")
            driver = webdriver.Chrome(service=service, options=chrome_options)
            resource_manager.add_driver(st.session_state.session_id, driver)
            wait = WebDriverWait(driver, 60)  # Increased timeout to 60s

            # Log Chrome and ChromeDriver versions
            chrome_version = driver.capabilities['browserVersion']
            chromedriver_version = driver.capabilities['chrome']['chromedriverVersion'].split(' ')[0]
            logging.info(f"Session {st.session_state.session_id} - Using Chromium version: {chrome_version}, ChromeDriver version: {chromedriver_version}")

            # Ensure page is properly loaded with session-specific handling
            ensure_page_loaded(driver)
            resource_manager.update_last_used(st.session_state.session_id)
            
            # Wait for form elements with better error handling
            gradecard_select = wait_and_find_element(driver, By.ID, "ddlGradecardfor", timeout=15)
            if not gradecard_select:
                raise WebDriverException("Grade card type dropdown not found or not ready")
            Select(gradecard_select).select_by_value(gradecard_for[0])

            time.sleep(1)  # Increased delay for dropdown to update

            program_select = wait_and_find_element(driver, By.ID, "ddlProgram", timeout=15)
            if not program_select:
                raise WebDriverException("Program dropdown not found or not ready")
            Select(program_select).select_by_value(program_code)

            time.sleep(1)  # Increased delay for dropdown to update

            enrollment_input = wait_and_find_element(driver, By.ID, "txtEnrno", timeout=15)
            if not enrollment_input:
                raise WebDriverException("Enrollment number input not found or not ready")
            enrollment_input.clear()
            enrollment_input.send_keys(enrollment)

            time.sleep(1)  # Increased delay before clicking

            # Optimized login button click
            login_button = wait_and_find_element(driver, By.ID, "btnlogin", timeout=15, clickable=True)
            if not login_button:
                raise WebDriverException("Login button not found or not ready")
            
            # Try multiple click methods
            try:
                login_button.click()
            except:
                try:
                    driver.execute_script("arguments[0].click();", login_button)
                except:
                    # Last resort: try JavaScript event
                    driver.execute_script("""
                        var evt = new MouseEvent('click', {
                            bubbles: true,
                            cancelable: true,
                            view: window
                        });
                        arguments[0].dispatchEvent(evt);
                    """, login_button)

            # Wait for results with better error handling
            try:
                WebDriverWait(driver, 30).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvDetail")),
                        EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_lblMsg"))
                    )
                )
            except TimeoutException:
                raise WebDriverException("Timeout waiting for results or error message")

            soup = BeautifulSoup(driver.page_source, "html.parser")
            logging.info(f"Session {st.session_state.session_id} - Parsed page source")

            # Check for CAPTCHA
            if soup.find("div", {"id": "captcha"}) or "captcha" in driver.page_source.lower():
                st.error("❌ CAPTCHA detected. Please try again later or access the website manually to verify.")
                with create_temp_file('.html') as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.info(f"Session {st.session_state.session_id} - Page source saved to: {tmp_file}")
                    st.write(f"Page source saved to: {tmp_file}")
                st.session_state.processing = False
                st.stop()

            # Check for error messages
            error_message = soup.find("span", {"id": "ctl00_ContentPlaceHolder1_lblMsg"})
            if error_message and error_message.text.strip():
                st.error(f"❌ IGNOU website error: {error_message.text.strip()}")
                with create_temp_file('.html') as tmp_file:
                    tmp_file.write(driver.page_source.encode("utf-8"))
                    logging.info(f"Session {st.session_state.session_id} - Page source saved to: {tmp_file}")
                    st.write(f"Page source saved to: {tmp_file}")
                st.session_state.processing = False
                st.stop()

            # Extract table and student details
            table = soup.find("table", {"id": "ctl00_ContentPlaceHolder1_gvDetail"})
            if not table:
                st.error("❌ Grade card table not found. Please check your enrollment number and program code.")
                st.session_state.processing = False
                st.stop()

            # Extract student details
            student_details = extract_student_details(soup)
            
            # Display student details in a nice box
            if student_details:
                st.markdown('<div class="summary-box">', unsafe_allow_html=True)
                st.subheader("👤 Student Details")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.write(f"**Enrollment No:** {student_details['enrollment']}")
                with col2:
                    st.write(f"**Name:** {student_details['name']}")
                with col3:
                    st.write(f"**Programme Code:** {student_details['program']}")
                st.markdown('</div>', unsafe_allow_html=True)

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

            # Filter incomplete subjects
            df_incomplete = df[df["STATUS"] != "COMPLETED"].copy()
            # Remove the total row if it exists in incomplete subjects
            df_incomplete = df_incomplete[df_incomplete["COURSE"] != "Total"]

            # Display results with improved layout
            st.success("✅ Grade Card Parsed and Calculated!")
            
            # Summary section in a nice box with dynamic height
            st.markdown('<div class="summary-box" style="min-height: fit-content;">', unsafe_allow_html=True)
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

            # Download buttons in a row
            col1, col2 = st.columns(2)
            
            with col1:
                # Excel download for completed subjects
                excel_file = create_temp_file('.xlsx')
                try:
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as excel_buffer:
                        df_calc_display.to_excel(excel_buffer, sheet_name='Completed Subjects', index=False)
                        if not df_calc.empty:
                            df_calc.to_excel(excel_buffer, sheet_name='Completed Subjects', index=False)
                    
                    with open(excel_file, 'rb') as f:
                        st.download_button(
                            "📊 Download Excel Report",
                            f,
                            file_name=f"ignou_grade_report_{st.session_state.session_id}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    logging.error(f"Session {st.session_state.session_id} - Error creating Excel file: {str(e)}")
                    st.error("Failed to create Excel report. Please try again.")

            with col2:
                # PDF download
                pdf_file = create_temp_file('.pdf')
                try:
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("helvetica", "B", 14)
                    pdf.cell(200, 10, "IGNOU Grade Report", new_x="LMARGIN", new_y="NEXT", align="C")
                    pdf.ln(10)
                    
                    # Add student details section
                    if student_details:
                        pdf.set_font("helvetica", "B", 12)
                        pdf.cell(200, 10, "Student Details", new_x="LMARGIN", new_y="NEXT")
                        pdf.set_font("helvetica", size=12)
                        pdf.cell(200, 10, f"Enrollment No: {student_details['enrollment']}", new_x="LMARGIN", new_y="NEXT")
                        pdf.cell(200, 10, f"Name: {student_details['name']}", new_x="LMARGIN", new_y="NEXT")
                        pdf.cell(200, 10, f"Programme Code: {student_details['program']}", new_x="LMARGIN", new_y="NEXT")
                        pdf.ln(10)

                    # Add summary section
                    pdf.set_font("helvetica", "B", 12)
                    pdf.cell(200, 10, "Summary", new_x="LMARGIN", new_y="NEXT")
                    pdf.set_font("helvetica", size=12)
                    pdf.cell(200, 10, f"Final Percentage: {percentage}%", new_x="LMARGIN", new_y="NEXT")
                    pdf.cell(200, 10, f"Total Obtained Marks: {total_obtained_marks:.2f} / {total_possible_marks:.0f}", new_x="LMARGIN", new_y="NEXT")
                    pdf.cell(200, 10, f"Total Assignment Marks: {totals['Asgn1']:.0f}", new_x="LMARGIN", new_y="NEXT")
                    pdf.cell(200, 10, f"Total Theory Marks: {totals['TERM END THEORY']:.0f}", new_x="LMARGIN", new_y="NEXT")
                    pdf.cell(200, 10, f"Total Practical Marks: {totals['TERM END PRACTICAL']:.0f}", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(10)
                    
                    # Add completed subjects table
                    pdf.set_font("helvetica", "B", 12)
                    pdf.cell(200, 10, "Completed Subjects", new_x="LMARGIN", new_y="NEXT")
                    pdf.set_font("helvetica", size=10)
                    
                    # Table headers with serial number
                    headers = ["S.No.", "Course", "Assignment", "Theory", "Practical", "30% Assignment", "70% Theory", "Total"]
                    col_widths = [10, 35, 20, 20, 20, 25, 25, 20]
                    
                    # Add headers
                    for i, header in enumerate(headers):
                        pdf.cell(col_widths[i], 10, header, 1)
                    pdf.ln()
                    
                    # Add data rows with serial numbers
                    for idx, (_, row) in enumerate(df_calc_display.iterrows(), 1):
                        pdf.cell(col_widths[0], 10, str(idx), 1)
                        pdf.cell(col_widths[1], 10, str(row["COURSE"]), 1)
                        pdf.cell(col_widths[2], 10, f"{row['Asgn1']:.0f}", 1)
                        pdf.cell(col_widths[3], 10, f"{row['TERM END THEORY']:.0f}", 1)
                        pdf.cell(col_widths[4], 10, f"{row['TERM END PRACTICAL']:.0f}", 1)
                        pdf.cell(col_widths[5], 10, f"{row['30% Assignments']:.2f}", 1)
                        pdf.cell(col_widths[6], 10, f"{row['70% Theory']:.2f}", 1)
                        pdf.cell(col_widths[7], 10, f"{row['Total (A+B)']:.2f}", 1)
                        pdf.ln()
                    
                    # Add incomplete subjects if any
                    if not df_incomplete.empty:
                        pdf.ln(10)
                        pdf.set_font("helvetica", "B", 12)
                        pdf.cell(200, 10, "Incomplete Subjects", new_x="LMARGIN", new_y="NEXT")
                        pdf.set_font("helvetica", size=10)
                        
                        # Table headers for incomplete subjects with serial number
                        headers = ["S.No.", "Course", "Status", "Assignment", "Theory", "Practical"]
                        col_widths = [10, 45, 25, 25, 25, 25]
                        
                        # Add headers
                        for i, header in enumerate(headers):
                            pdf.cell(col_widths[i], 10, header, 1)
                        pdf.ln()
                        
                        # Add data rows with serial numbers
                        for idx, (_, row) in enumerate(df_incomplete.iterrows(), 1):
                            pdf.cell(col_widths[0], 10, str(idx), 1)
                            pdf.cell(col_widths[1], 10, str(row["COURSE"]), 1)
                            pdf.cell(col_widths[2], 10, str(row["STATUS"]), 1)
                            pdf.cell(col_widths[3], 10, f"{row['Asgn1']:.0f}", 1)
                            pdf.cell(col_widths[4], 10, f"{row['TERM END THEORY']:.0f}", 1)
                            pdf.cell(col_widths[5], 10, f"{row['TERM END PRACTICAL']:.0f}", 1)
                            pdf.ln()
                    
                    pdf.output(pdf_file)
                    
                    with open(pdf_file, 'rb') as f:
                        st.download_button(
                            "📄 Download PDF Report",
                            f,
                            file_name=f"ignou_grade_report_{st.session_state.session_id}.pdf",
                            use_container_width=True
                        )
                except Exception as e:
                    logging.error(f"Session {st.session_state.session_id} - Error creating PDF file: {str(e)}")
                    st.error("Failed to create PDF report. Please try again.")

            # Completed subjects table with dynamic height
            st.markdown('<div style="min-height: fit-content;">', unsafe_allow_html=True)
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
            st.markdown('</div>', unsafe_allow_html=True)

            # Incomplete subjects table with dynamic height
            df_incomplete = df[df["STATUS"] != "COMPLETED"]
            if not df_incomplete.empty:
                st.markdown('<div style="min-height: fit-content;">', unsafe_allow_html=True)
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
                st.markdown('</div>', unsafe_allow_html=True)

            # Clean up temporary files
            try:
                os.remove(excel_file)
                os.remove(pdf_file)
            except Exception as e:
                logging.error(f"Session {st.session_state.session_id} - Error cleaning up temporary files: {str(e)}")

            st.session_state.processing = False
            break  # Successfully completed, exit the retry loop

        except WebDriverException as e:
            retry_count += 1
            logging.error(f"Session {st.session_state.session_id} - Attempt {retry_count} failed due to WebDriver issue: {str(e)}")
            if retry_count > max_retries:
                st.error(f"⏳ Failed to initialize ChromeDriver after {max_retries + 1} attempts. Please try again later.")
                try:
                    result = subprocess.run(["chromium", "--version"], capture_output=True, text=True)
                    logging.info(f"Session {st.session_state.session_id} - Chromium version check: {result.stdout or result.stderr}")
                except Exception as debug_e:
                    logging.error(f"Session {st.session_state.session_id} - Failed to check chromium version: {str(debug_e)}")
                st.session_state.processing = False
            else:
                st.warning(f"⚠️ Attempt {retry_count} failed. Retrying in 3 seconds...")
                time.sleep(3)
        except Exception as e:
            logging.error(f"Session {st.session_state.session_id} - Error: {str(e)}")
            st.error("❌ An error occurred. Please try again later.")
            st.session_state.processing = False
            break
        finally:
            if driver:
                resource_manager.remove_driver(st.session_state.session_id)
            st.session_state.processing = False

# Update cleanup function to handle user data directories
def cleanup_temp_files():
    for file_path in st.session_state.temp_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logging.info(f"Cleaned up temporary file: {file_path}")
        except Exception as e:
            logging.error(f"Error cleaning up file {file_path}: {str(e)}")
            try:
                time.sleep(1)
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as retry_e:
                logging.error(f"Failed to remove file {file_path} after retry: {str(retry_e)}")
    
    # Cleanup Chrome user data directories
    try:
        for item in os.listdir("/tmp"):
            if item.startswith("chrome_profile_"):
                profile_path = os.path.join("/tmp", item)
                if os.path.isdir(profile_path):
                    shutil.rmtree(profile_path, ignore_errors=True)
    except Exception as e:
        logging.error(f"Error cleaning up Chrome profiles: {str(e)}")

# Register cleanup functions
atexit.register(cleanup_temp_files)
atexit.register(resource_manager.cleanup_all)