import os
import json
import time
import re
import logging
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# -------------------- Logging -------------------- #
logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')

# -------------------- Environment Variables from Secrets -------------------- #
USERNAME = os.environ.get('AROFLO_USERNAME')
PASSWORD = os.environ.get('AROFLO_PASSWORD')

# Validate that credentials are available
if not USERNAME or not PASSWORD:
    logging.error("Missing AROFLO credentials. Please set AROFLO_USERNAME and AROFLO_PASSWORD environment variables.")
    exit(1)

logging.info("Successfully loaded AROFLO credentials from environment variables")

# -------------------- Chrome options -------------------- #
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1280,720")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.add_experimental_option('useAutomationExtension', False)

# -------------------- WebDriver -------------------- #
try:
    # Use webdriver-manager to automatically handle ChromeDriver compatibility
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    logging.info("Chrome WebDriver initialized with webdriver-manager")
except Exception as e:
    logging.error(f"Failed to initialize Chrome WebDriver: {e}")
    raise

# -------------------- Job tracker -------------------- #
CONFIG_FILE = "job_tracker.json"
DEFAULT_START_JOB_ID = 3411

def load_job_tracker():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                last_index = data.get('last_index')
                if last_index is None:
                    last_index = DEFAULT_START_JOB_ID
                logging.info(f"Loaded last processed job index: {last_index}")
                return int(last_index)
        except Exception as e:
            logging.error(f"Error loading job tracker: {e}")
            return DEFAULT_START_JOB_ID
    else:
        logging.info(f"No existing job tracker found, starting from default job ID: {DEFAULT_START_JOB_ID}")
        return DEFAULT_START_JOB_ID

def save_job_tracker(index):
    try:
        data = {"last_index": index}
        with open(CONFIG_FILE, 'w') as f:
            json.dump(data, f, indent=2)
        logging.info(f"Saved last processed job index: {index}")
    except Exception as e:
        logging.error(f"Error saving job tracker: {e}")

# -------------------- Helper functions -------------------- #
def safe_find(by, value, timeout=15):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def safe_find_visible(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, value)))

def safe_div_click(selector, timeout=10):
    elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    elem.click()
    time.sleep(1)

def is_valid_email(text):
    """Check if text is a valid email address"""
    if not text:
        return False
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, text.strip()))

def check_email_field_content():
    """Check if email field contains a valid email address and return it"""
    try:
        # Try both -15 and -16 fields
        selectors = ["textarea[id$='-15']", "input[id$='-15']", 
                    "textarea[id$='-16']", "input[id$='-16']"]
        
        for selector in selectors:
            try:
                email_field = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                # Get email value from field
                email_value = email_field.get_attribute('value') or email_field.text or email_field.get_attribute('innerText')
                email_value = email_value.strip() if email_value else ""
                
                logging.info(f"Email field {selector} content: '{email_value}'")
                
                # Check if it's a valid email (ignore values like 'on')
                if is_valid_email(email_value) and email_value not in ['on', 'off', 'true', 'false']:
                    logging.info(f"Valid email found in field: {email_value}")
                    return email_value
                    
            except Exception:
                continue
        
        logging.info("No valid email found in any field")
        return None
            
    except Exception as e:
        logging.error(f"Error checking email field content: {e}")
        return None

def get_email_field():
    """Get the appropriate email field based on conditions"""
    try:
        if driver.find_elements(By.XPATH, "//th[contains(., 'Delivery Only')]"):
            selector = "textarea[id$='-15'], input[id$='-15']"
        else:
            selector = "textarea[id$='-16'], input[id$='-16']"

        try:
            email_field = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
            )
            logging.info(f"Found email field using selector: {selector}")
            return email_field
        except Exception:
            # Fallback to label-based approach
            logging.warning("Fallback: trying label-based XPath for Task Email Address")
            email_field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//td[contains(text(),'Task Email Address')]/following-sibling::td//textarea | "
                    "//td[contains(text(),'Task Email Address')]/following-sibling::td//input"
                ))
            )
            return email_field
    except Exception as e:
        logging.error(f"Error finding email field: {e}")
        return None

def paste_email(email):
    """Paste email into Task Email Address field"""
    try:
        email_field = get_email_field()
        if not email_field:
            logging.error("Could not find email field to paste into")
            return False

        # Clear and paste email
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", email_field)
        email_field.clear()
        time.sleep(0.5)
        
        # Try direct input first
        email_field.send_keys(email)
        time.sleep(0.5)

        # Verify email was entered correctly
        current_value = email_field.get_attribute("value") or email_field.text
        if current_value.strip() == email:
            logging.info(f"Email pasted successfully: {email}")
            return True

        # If send_keys fails, try clipboard paste
        logging.warning("Direct input failed, trying clipboard paste")
        pyperclip.copy(email)
        email_field.clear()
        time.sleep(0.5)
        email_field.send_keys(Keys.CONTROL, 'v')
        time.sleep(0.5)

        # Verify again
        current_value = email_field.get_attribute("value") or email_field.text
        if current_value.strip() == email:
            logging.info(f"Email pasted successfully via clipboard: {email}")
            return True

        logging.warning("Failed to paste email value after all attempts")
        return False

    except Exception as e:
        logging.error(f"Paste error: {e}")
        return False

def search_aroflo_email_in_page():
    """Search for an existing .aroflo.com email anywhere on the page"""
    try:
        page_source = driver.page_source
        matches = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]*\.aroflo\.com\b', page_source)
        if matches:
            logging.info(f"Found .aroflo.com email via regex: {matches[0]}")
            return matches[0]
        return None
    except Exception as e:
        logging.error(f"Error searching email in page: {e}")
        return None

def is_installer_checkin_task():
    """Check if the current task is an 'Installer Checkin' type"""
    try:
        # Look for 'Installer Checkin' in the task details
        installer_checkin_elements = driver.find_elements(By.XPATH, "//th[contains(., 'Installer Checkin')]")
        if installer_checkin_elements:
            logging.info("Found 'Installer Checkin' task type - skipping this task")
            return True
        return False
    except Exception as e:
        logging.error(f"Error checking for Installer Checkin task type: {e}")
        return False

# -------------------- Main workflow -------------------- #
try:
    last_index = load_job_tracker()

    logging.info("Navigating to Aroflo login...")
    driver.get("https://office.aroflo.com/?redirect=%2Fims%2FSite%2FHome%2Findex.cfm%3Fview%3D1%26tid%3DIMS.HME")
    
    # Enter credentials
    username_field = safe_find(By.CSS_SELECTOR, 'input[name="username"]')
    username_field.send_keys(USERNAME)
    
    password_field = safe_find(By.CSS_SELECTOR, 'input[name="password"]')
    password_field.send_keys(PASSWORD)

    submit_btn = safe_find(By.CSS_SELECTOR, 'button[type="submit"]')
    submit_btn.click()
    time.sleep(3)
    
    # Handle potential second login screen
    try:
        submit_btn = safe_find(By.CSS_SELECTOR, 'button[type="submit"]', timeout=5)
        submit_btn.click()
    except:
        pass

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    logging.info("Login successful!")

    logging.info("Navigating to tasks page...")
    safe_div_click(".afMegaMenu > button:nth-of-type(2)")
    safe_div_click("div:nth-of-type(1) .afMegaMenu__item .item-content .item-submenu a:nth-of-type(2)")

    # -------------------- Extract all job IDs -------------------- #
    job_id_cells = driver.find_elements(By.CSS_SELECTOR, "td.page-content-task-jobnumber")
    if not job_id_cells:
        logging.warning("No job ID cells found!")
    else:
        first_row_job_id = int(job_id_cells[0].text.strip())
        logging.info(f"Max job ID (first row): {first_row_job_id}")

        # Loop from last_index up to first_row_job_id (inclusive)
        for job_id in range(last_index, first_row_job_id + 1):
            logging.info(f"Processing Job ID: {job_id}")

            try:
                rows = driver.find_elements(By.XPATH, f"//td[contains(@class,'page-content-task-jobnumber') and text()='{job_id}']/..")
                if not rows:
                    logging.warning(f"Task {job_id} not found on the page. Skipping...")
                    continue

                task_link = rows[0].find_element(By.CSS_SELECTOR, "td.page-content-task-name a")
                task_link.click()
                time.sleep(3)

                # Check if this is an 'Installer Checkin' task type
                if is_installer_checkin_task():
                    logging.info(f"Skipping task {job_id} because it's an 'Installer Checkin' type")
                    safe_div_click(".afContainer__titlebar-actions-left > a:nth-of-type(1)")
                    save_job_tracker(job_id)
                    continue

                # Check if email field already contains a valid email in -15 or -16
                existing_email = check_email_field_content()
                if existing_email:
                    logging.info(f"Task already has email in field: {existing_email}. Skipping completely.")
                    safe_div_click(".afContainer__titlebar-actions-left > a:nth-of-type(1)")
                    save_job_tracker(job_id)
                    continue

                # If we get here, no email exists - perform Create workflow
                logging.info("No existing email found in fields. Performing Create workflow...")
                
                # FIRST: Try to click workflow button
                workflow_clicked = False
                try:
                    safe_div_click("div:nth-of-type(2) > .afBtn > .afIcon")
                    workflow_clicked = True
                    logging.info("Workflow button clicked successfully")
                except Exception as e:
                    logging.error(f"Failed to click workflow div for task {job_id}: {e}")

                # SECOND: Check if email was already created (even if workflow button failed)
                created_email = search_aroflo_email_in_page()
                if created_email:
                    logging.info(f"Email already created: {created_email}. Pasting into field.")
                    if paste_email(created_email):
                        logging.info("Successfully pasted created email into field")
                    else:
                        logging.error("Failed to paste created email")
                else:
                    # THIRD: If no email created yet, try Create button
                    if workflow_clicked:
                        try:
                            create_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[@role='button' and normalize-space()='Create']"))
                            )
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", create_btn)
                            create_btn.click()
                            logging.info("Create button clicked")
                            time.sleep(3)
                            
                            # Check again for created email after Create button
                            created_email = search_aroflo_email_in_page()
                            if created_email:
                                logging.info(f"Email created after Create button: {created_email}")
                                if not paste_email(created_email):
                                    logging.error("Failed to paste created email")
                            else:
                                logging.warning("No email created after Create button")
                                # Paste default email as fallback
                                default_email = "default@aroflo.com"
                                if paste_email(default_email):
                                    logging.info(f"Pasted default email: {default_email}")
                        except Exception as e:
                            logging.error(f"Create button not found for task {job_id}: {e}")
                            # Even if Create button fails, try default email
                            default_email = "default@aroflo.com"
                            if paste_email(default_email):
                                logging.info(f"Pasted default email after Create button failure: {default_email}")
                    else:
                        # If workflow button also failed, just use default email
                        logging.warning("Workflow button failed, using default email")
                        default_email = "default@aroflo.com"
                        if paste_email(default_email):
                            logging.info(f"Pasted default email: {default_email}")

                # Save the task
                try:
                    save_btn = safe_find(By.ID, "update_btn", timeout=5)
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", save_btn)
                    save_btn.click()
                    logging.info(f"Task {job_id} saved successfully")
                    time.sleep(2)
                except Exception as e:
                    logging.warning(f"Save button not found for task {job_id}: {e}")

                save_job_tracker(job_id)

                try:
                    safe_div_click(".afContainer__titlebar-actions-left > a:nth-of-type(1)")
                    time.sleep(1)
                except Exception as e:
                    logging.warning(f"Failed to return to main page after task {job_id}: {e}")

            except Exception as e:
                logging.error(f"Unexpected error while processing task {job_id}: {e}")
                continue

    logging.info("All workflows completed!")

except Exception as e:
    logging.error(f"Critical error in main workflow: {e}")
    raise

finally:
    driver.quit()
    logging.info("Chrome WebDriver closed")
