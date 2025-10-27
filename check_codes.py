import os
import time
import re
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

EXCEL_PATH = "example_codes.xlsx"
SITE_URL = "https://online.olimpiada.ru"

LOG_LEVEL = "DEFAULT"  # DEFAULT или DEBUG
SEARCH_MODE = "REPORT"  # FIRST_MATCH или REPORT

HEADLESS = False
WAIT_TIMEOUT = 10
SLEEP_BETWEEN = 1.0

TARGET_CLASS = "5А"
TARGET_NAME = "Иван Иванов"


CODE_PATTERN = re.compile(r"^[a-z0-9]+/[a-z0-9]+/\d+/[a-z0-9]+$", re.IGNORECASE)

def normalize_subject_name(s):
    return str(s).replace(" ", "").lower() if s else ""

def log(msg, level="INFO"):
    if level == "DEBUG" and LOG_LEVEL != "DEBUG":
        return
    print(f"[{level}] {msg}")

def setup_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_window_size(1200, 900)
    return driver

def read_codes_with_position_for_class(path, column_name, class_column_index=4, target_class=TARGET_CLASS):
    df = pd.read_excel(path, engine="openpyxl", header=0)
    df.columns = [str(c) for c in df.columns]
    real_columns = [c for c in df.columns if not c.lower().startswith("unnamed") and not c.lower().startswith("9 класс")]
    norm_map = {normalize_subject_name(c): c for c in real_columns}
    column_norm = normalize_subject_name(column_name)
    if column_norm not in norm_map:
        raise KeyError(f"Колонка '{column_name}' не найдена.")
    column_real = norm_map[column_norm]

    results = []
    for idx, row in df.iterrows():
        if len(row) > class_column_index and str(row.iloc[class_column_index]).strip() == target_class:
            val = row[column_real]
            if pd.isna(val):
                continue
            code = str(val).strip()
            if code and CODE_PATTERN.match(code):
                results.append((idx + 2, column_real, code))
    return results, norm_map.keys()

def find_input_and_submit(driver, code):
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    try:
        input_el = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.ui-textinput__input"))
        )
        input_el.clear()
        input_el.send_keys(code)
        log(f"Введён код: {code}", "DEBUG")
        ActionChains(driver).move_to_element(input_el).click().perform()
        input_el.send_keys(Keys.TAB)
        time.sleep(0.3)
        btn_el = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "button.smt-login-user-form__register-btn"))
        wait.until(lambda d: btn_el.is_enabled())
        log("Кнопка 'Войти' активна, клик", "DEBUG")
        btn_el.click()
        return True
    except Exception as e:
        log(f"Ошибка при вводе кода и клике: {e}", "DEBUG")
        return False

def check_result_and_logout(driver):
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    try:
        page_text = driver.find_element(By.TAG_NAME, "body").text
        if "Добро пожаловать!" in page_text:
            return False, "код не использован", 0.0, 0.0
        ui = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.user_info")))
        full_text = ui.text
        name = full_text.replace("Выйти", "").strip()
        name_match = (normalize_subject_name(name) == normalize_subject_name(TARGET_NAME)) if TARGET_NAME else True
        try:
            scores_el = driver.find_element(By.CSS_SELECTOR, "div.header__button_text")
            scores_text = scores_el.text
            match = re.search(r"Ваш результат:\s*([\d.,]+)\s*из\s*([\d.,]+)", scores_text)
            if match:
                score = float(match.group(1).replace(",", "."))
                max_score = float(match.group(2).replace(",", "."))
            else:
                score = 0.0
                max_score = 0.0
        except Exception:
            score = 0.0
            max_score = 0.0
        try:
            logout_btn = ui.find_element(By.CSS_SELECTOR, "div.user_info__logout")
            logout_btn.click()
        except Exception:
            pass
        return name_match, name, score, max_score
    except Exception:
        return False, "", 0.0, 0.0

def print_progress(current, total, success_count, bar_length=40):
    percent = current / total
    filled = int(bar_length * percent)
    empty = bar_length - filled
    bar = "█" * filled + "-" * empty
    sys.stdout.write(f"\rПрогресс: |{bar}| {current}/{total} ({percent*100:.1f}%)")
    sys.stdout.flush()

def generate_unique_filename(base_name):
    if not os.path.exists(base_name):
        return base_name
    i = 2
    while True:
        new_name = f"{os.path.splitext(base_name)[0]}_{i}.xlsx"
        if not os.path.exists(new_name):
            return new_name
        i += 1

def main():
    driver = setup_driver()
    try:
        df_preview = pd.read_excel(EXCEL_PATH, nrows=0, engine="openpyxl")
        real_columns = [c for c in df_preview.columns if not c.lower().startswith("unnamed") and not c.lower().startswith("9 класс")]
        print("Доступные предметы:")
        for c in real_columns:
            print(" -", c)
        subject_input = input("\nВведите точное название предмета: ")
        subject_norm = normalize_subject_name(subject_input)
        norm_map = {normalize_subject_name(c): c for c in real_columns}
        if subject_norm not in norm_map:
            print(f"Ошибка: предмет '{subject_input}' не найден")
            return
        subject = norm_map[subject_norm]

        codes, _ = read_codes_with_position_for_class(EXCEL_PATH, subject, class_column_index=4, target_class=TARGET_CLASS)
        total_codes = len(codes)
        print(f"Найдено {total_codes} кодов. Проверка началась...")

        results_report = []
        success_count = 0

        for idx, (row_idx, column_name, code) in enumerate(codes, 1):
            if LOG_LEVEL == "DEBUG":
                print(f"\nСтрока {row_idx}, столбец '{column_name}' — проверка кода: {code}")
            try:
                driver.get(SITE_URL)
                time.sleep(0.5)
                submitted = find_input_and_submit(driver, code)
                if not submitted:
                    driver.get(SITE_URL.rstrip("/") + "/" + code.lstrip("/"))
                time.sleep(1.0)
                ok, user_info_text, score, max_score = check_result_and_logout(driver)
                
                if SEARCH_MODE == "FIRST_MATCH":
                    if ok and (user_info_text != "код не использован"):
                        results_report.append((user_info_text, code, score, max_score))
                        success_count += 1
                        break
                        
                elif SEARCH_MODE == "REPORT":
                    if user_info_text != "код не использован":
                        results_report.append((user_info_text, code, score, max_score))
                        success_count += 1
                        
            except Exception as e:
                log(f"Ошибка при проверке '{code}': {repr(e)}", "DEBUG")
            finally:
                time.sleep(SLEEP_BETWEEN)
                
            if LOG_LEVEL != "DEBUG" and SEARCH_MODE != "FIRST_MATCH":
                print_progress(idx, total_codes, success_count)

        if LOG_LEVEL != "DEBUG":
            print()

        if SEARCH_MODE != "FIRST_MATCH" and results_report:
            df_report = pd.DataFrame(results_report, columns=["ИФ", "Код", "Баллы", "Макс. баллы"])
            df_report = df_report.sort_values(by="Баллы", ascending=False)
            report_file = generate_unique_filename(f"Отчет_{subject.replace(' ', '_')}.xlsx")
            df_report.to_excel(report_file, index=False)
            print(f"✅ Отчет сформирован: {report_file}")

        elif SEARCH_MODE == "FIRST_MATCH" and results_report:
            print(f"✅ Найдено совпадение: {results_report[0][0]} ({results_report[0][2]}/{results_report[0][3]} баллов)")


        print("\n✅ Проверка завершена.")
        for fi, code, score, max_score in results_report:
            print(f"{fi} | {code} | {score}/{max_score} баллов")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
