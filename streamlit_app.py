import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime, timedelta
import re

# --- إعدادات الصفحة ---
st.set_page_config(page_title="Exam Booking Bot", layout="wide")

# --- دوال مساعدة ---
def parse_name(full_name):
    full_name = str(full_name).strip()
    match = re.search(r'(.*?)\s*\((.*?)\)', full_name)
    if match: return match.group(1).strip(), match.group(2).strip()
    parts = full_name.split()
    if len(parts) > 1: return parts[0], parts[-1]
    return full_name, "."

def format_date_standard(date_str):
    try:
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%d-%m-%y'):
            try: return datetime.strptime(str(date_str), fmt).strftime('%d-%m-%Y')
            except: continue
        return str(date_str)
    except: return str(date_str)

# --- إعداد المتصفح للسحابة (Headless) ---
def get_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # تشغيل بدون واجهة
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # محاولة التثبيت والتشغيل
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        st.error(f"Error setting up Chrome Driver: {e}")
        return None

# --- دوال التفاعل مع الموقع ---
def select_mui_option(driver, wait, input_id, search_text, match_condition_func=None):
    try:
        input_elem = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
        driver.execute_script("arguments[0].click();", input_elem)
        input_elem.send_keys(Keys.CONTROL + "a", Keys.BACK_SPACE)
        if search_text:
            input_elem.send_keys(search_text)
            time.sleep(1.5)
        
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.MuiAutocomplete-option")))
            options = driver.find_elements(By.CSS_SELECTOR, "li.MuiAutocomplete-option")
        except:
            driver.find_element(By.TAG_NAME, 'body').click()
            return False, "No list"

        for opt in options:
            text = opt.text
            if match_condition_func:
                if match_condition_func(text):
                    driver.execute_script("arguments[0].click();", opt)
                    return True, text
            else:
                if search_text.lower() in text.lower():
                    driver.execute_script("arguments[0].click();", opt)
                    return True, text
        
        driver.find_element(By.TAG_NAME, 'body').click()
        return False, "Not found"
    except: return False, "Error"

def handle_time_selection(driver, wait, target_time_str):
    input_id = "timeList"
    target_clean = str(target_time_str).strip()[:5]
    try: target_dt = datetime.strptime(target_clean, "%H:%M")
    except: return False, "Format error"

    try:
        input_elem = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
        driver.execute_script("arguments[0].click();", input_elem)
        input_elem.send_keys(Keys.CONTROL + "a", Keys.BACK_SPACE)
        time.sleep(0.5)
        input_elem.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)

        options = driver.find_elements(By.CSS_SELECTOR, "li.MuiAutocomplete-option")
        valid_opts = []
        for opt in options:
            match = re.search(r'(\d{1,2}:\d{2})', opt.text)
            if match:
                try: valid_opts.append((datetime.strptime(match.group(1), "%H:%M"), opt))
                except: continue
        
        best_elem = None
        for dt, elem in valid_opts:
            if dt == target_dt:
                best_elem = elem; break
        
        if not best_elem:
            prevs = [x for x in valid_opts if x[0] < target_dt]
            if prevs:
                prevs.sort(key=lambda x: x[0], reverse=True)
                best_elem = prevs[0][1]

        if best_elem:
            driver.execute_script("arguments[0].click();", best_elem)
            return True, None
    except: pass
    
    # محاولة الكتابة المباشرة
    try:
        input_elem.send_keys(Keys.CONTROL + "a", Keys.BACK_SPACE)
        input_elem.send_keys(target_clean)
        input_elem.send_keys(Keys.ENTER)
        return True, None
    except: pass
    
    return False, "Failed"

def force_submit(driver, wait):
    try:
        btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit'] | //button[normalize-space()='Toevoegen']")))
        driver.execute_script("arguments[0].removeAttribute('disabled');", btn)
        driver.execute_script("arguments[0].disabled = false;", btn)
        driver.execute_script("arguments[0].click();", btn)
        return True
    except: return False

def set_react_date(driver, element, value):
    driver.execute_script("""
        let input = arguments[0];
        let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
        nativeInputValueSetter.call(input, arguments[1]);
        input.dispatchEvent(new Event('input', { bubbles: true}));
        input.dispatchEvent(new Event('change', { bubbles: true}));
        input.dispatchEvent(new Event('blur', { bubbles: true}));
    """, element, value)

# --- الواجهة ومنطق التشغيل ---
st.title("🤖 Exam Booking Automation Bot")
st.markdown("Automated booking system for MasterTolken (Headless Cloud Version)")

uploaded_file = st.file_uploader("Upload Excel/CSV File", type=['xlsx', 'csv'])

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date")
with col2:
    end_date = st.date_input("End Date")

if st.button("Start Automation"):
    if not uploaded_file:
        st.error("Please upload a file first.")
    else:
        # قراءة الملف
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';', dtype=str)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
            
            # الفلترة
            TARGET_CODES = ["BTH-I-T", "VM3-C-I-T", "RVM1-C-I-T", "ATH-I-T", "VM2-C-I-T", "VM3-D-IT", "VM2-D-IT"]
            df['Examen.datum_dt'] = pd.to_datetime(df['Examen.datum'], dayfirst=True, errors='coerce')
            mask_date = (df['Examen.datum_dt'] >= pd.to_datetime(start_date)) & (df['Examen.datum_dt'] <= pd.to_datetime(end_date))
            mask_code = df['Algemeen.product_code'].astype(str).str.strip().isin(TARGET_CODES)
            filtered_df = df[mask_date & mask_code]
            
            st.info(f"Found {len(filtered_df)} students to process.")
            
            if len(filtered_df) > 0:
                driver = get_driver()
                if driver:
                    wait = WebDriverWait(driver, 20)
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    failed_list = []
                    success_count = 0
                    
                    try:
                        # Login
                        status_text.text("Logging in...")
                        driver.get("https://rijschool.mastertolken.nl/")
                        try:
                            try: email = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='email']")))
                            except: email = driver.find_element(By.XPATH, "//input[contains(@name, 'mail')]")
                            email.send_keys("alrawirijschool@gmail.com")
                            pwd = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
                            pwd.send_keys("As@12345678")
                            pwd.submit()
                            time.sleep(3)
                        except Exception as e:
                            st.error(f"Login failed: {e}")

                        # Enter Form
                        try:
                            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Afronden') or contains(text(), 'afronden')]")))
                            btn.click()
                            wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'dd-mm-yyyy')]")))
                        except: pass

                        # Loop
                        for i, (index, row) in enumerate(filtered_df.iterrows()):
                            full_name = row['Kandidaat.samengestelde_naam']
                            status_text.text(f"Processing ({i+1}/{len(filtered_df)}): {full_name}")
                            progress_bar.progress((i + 1) / len(filtered_df))
                            
                            try:
                                # Date
                                try: date_input = driver.find_element(By.XPATH, "//input[contains(@placeholder, 'dd-mm-yyyy')]")
                                except: date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'dd-mm-yyyy')]")))
                                target_date = format_date_standard(row['Examen.datum'])
                                set_react_date(driver, date_input, target_date)

                                # Time
                                target_time = str(row['Examen.tijd']).split()[0][:5]
                                success, _ = handle_time_selection(driver, wait, target_time)
                                
                                # Code
                                prod = str(row['Algemeen.product_code']).strip()
                                select_mui_option(driver, wait, "examType", prod, lambda t: f"({prod})" in t)
                                
                                # Lang
                                select_mui_option(driver, wait, "languageList", "Arabisch", lambda t: "Syrisch" in t)
                                
                                # Location
                                raw_loc = str(row['Algemeen.locatie_naam']).split('(')[0]
                                search_keyword = raw_loc.split('-')[0].strip()
                                found_loc, msg_loc = select_mui_option(driver, wait, "examCenterList", search_keyword)
                                if not found_loc:
                                    failed_list.append(f"{full_name} (Location not found: {search_keyword})")
                                    driver.refresh()
                                    time.sleep(2)
                                    continue

                                # Gender
                                try: driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//input[@type='radio' and @value='M']"))
                                except: pass

                                # Data
                                fn, ln = parse_name(full_name)
                                data_map = [
                                    ("//label[contains(text(), 'Voornaam')]/following::input[1]", fn),
                                    ("//label[contains(text(), 'Achternaam')]/following::input[1]", ln),
                                    ("//label[contains(text(), 'Kandidaatsnummer')]/following::input[1]", row['Kandidaat.nummer']),
                                    ("//label[contains(text(), 'Emailadres')]/following::input[1]", "alrawirijschool@gmail.com"),
                                    ("//label[contains(text(), 'Telefoonnummer')]/following::input[1]", "0685338583")
                                ]
                                for xpath, val in data_map:
                                    try:
                                        inp = driver.find_element(By.XPATH, xpath)
                                        inp.send_keys(Keys.CONTROL + "a", Keys.BACK_SPACE)
                                        time.sleep(0.1)
                                        inp.send_keys(str(val))
                                    except: pass
                                
                                driver.find_element(By.TAG_NAME, 'body').click()
                                
                                # Submit
                                if force_submit(driver, wait):
                                    try:
                                        wait.until(lambda d: d.find_element(By.XPATH, "//label[contains(text(), 'Voornaam')]/following::input[1]").get_attribute('value') == "")
                                        success_count += 1
                                    except: failed_list.append(f"{full_name} (Verify manually)")
                                else:
                                    failed_list.append(f"{full_name} (Button failed)")

                            except Exception as e:
                                failed_list.append(f"{full_name} (Error: {e})")
                                driver.refresh()
                                time.sleep(2)
                        
                        driver.quit()
                        status_text.text("Done!")
                        
                        # Report
                        st.success(f"Completed! Success: {success_count}, Failed: {len(failed_list)}")
                        if failed_list:
                            st.warning("Students that failed:")
                            st.write(failed_list)

                    except Exception as e:
                        st.error(f"Fatal error: {e}")
        except Exception as e:
            st.error(f"Error reading file: {e}")