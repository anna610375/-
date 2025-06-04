import time
import re
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ====== НАСТРОЙКИ ======
USERNAME = "ansa174"
PASSWORD = "Yrx7d)73!"
STATUS_IDS = ["17541", "26873"]  # Завершено, Готов, ожидает оплаты
TOTAL_PAGES = 30
OUTPUT_FILE = "Результат_сделок.xlsx"
# ========================


def parse_number(text):
    """Оставляет цифры, точку и запятую, заменяет запятую на точку, возвращает float."""
    cleaned = re.sub(r'[^\d\.,-]', '', text).replace(',', '.')
    try:
        return float(cleaned)
    except:
        return 0.0


def extract_date(date_str):
    """Извлекает дату в формате дд.мм.гггг из строки и возвращает datetime или NaT"""
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(date_str))
    if m:
        try:
            return datetime.strptime(m.group(1), '%d.%m.%Y')
        except:
            return pd.NaT
    return pd.NaT

# Инициализация Selenium
def init_driver():
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )

try:
    driver = init_driver()

    # 1) Авторизация
    driver.get("https://printoffice24.com/login")
    time.sleep(2)
    driver.find_element(By.NAME, "username").send_keys(USERNAME)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD, Keys.RETURN)
    time.sleep(5)

    # 2) Фильтр статусов
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@data-id="deal_status_select"]'))
    ).click()
    time.sleep(1)
    for opt in driver.find_elements(By.CSS_SELECTOR, '.deal_list_select option[selected]'):
        driver.execute_script("arguments[0].selected = false;", opt)
    for opt in driver.find_elements(By.XPATH, '//select[@id="deal_status_select"]/option'):
        if opt.get_attribute("value") in STATUS_IDS:
            driver.execute_script("arguments[0].selected = true;", opt)
    driver.execute_script("$('#deal_status_select').selectpicker('refresh');")
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "h3.panel-title").click()
    time.sleep(3)

    # 3) Сбор ссылок + даты выдачи из списка
    deal_end_dates = {}
    deal_ids = []
    for page in range(1, TOTAL_PAGES + 1):
        print(f"📄 Страница {page}")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "deal_list_table"))
        )
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
        time.sleep(2)
        rows = driver.find_elements(By.CSS_SELECTOR, "#deal_list_table tbody tr")
        for row in rows:
            try:
                did = row.find_element(By.CSS_SELECTOR, "td.clmn_num").get_attribute("data-id")
                ed = row.find_element(By.CSS_SELECTOR, "td.clmn_end_date").text.strip().replace("\n", " ")
                if did and did not in deal_ids:
                    deal_ids.append(did)
                    deal_end_dates[did] = ed
            except:
                pass
        if page < TOTAL_PAGES:
            btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//a[@class="page_pagin" and normalize-space()="{page+1}"]'))
            )
            first = rows[0]
            btn.click()
            WebDriverWait(driver, 10).until(EC.staleness_of(first))
            time.sleep(2)

    links = [f"https://printoffice24.com/editDeal/{did}" for did in deal_ids]

    # 4) Парсинг каждой сделки
    all_rows = []
    for url in links:
        print("⏳ Обрабатываю", url)
        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.deal-element-row"))
        )
        time.sleep(1)

        deal_id = url.rsplit("/", 1)[-1]
        try:
            deal_number = driver.find_element(By.ID, "page_title").get_attribute("value").split("№")[-1].strip()
        except:
            deal_number = ""
        try:
            cb = driver.find_element(By.ID, "div_client_info")
            client_name = cb.find_element(
                By.XPATH,
                './/strong[contains(text(),"Название компании")]/following-sibling::br/following-sibling::span'
            ).text.strip()
        except:
            client_name = ""

        temp = []
        deal_ready = None
        els = driver.find_elements(By.CSS_SELECTOR, "tr.deal-element-row")
        for el in els:
            full = el.find_element(By.CLASS_NAME, "deal-element-name").text.strip()
            if " / " in full:
                name, ptype = full.split(" / ", 1)
            else:
                name, ptype = full, ""
            qty = el.find_elements(By.XPATH, './/td[@class="text-right"]')[1].text.replace('\xa0','').strip()
            rev = parse_number(el.find_elements(By.XPATH, './/td[@class="text-right"]')[5].text)
            try:
                cost = parse_number(el.find_element(By.XPATH, './/span[contains(@class,"hidden-values")]').get_attribute("data-val"))
            except:
                cost = 0.0
            profit = rev - cost
            margin = round(profit / rev * 100, 2) if rev else 0.0

            try:
                desc = el.find_element(By.CLASS_NAME, "deal-elm-description").text
                rd = next((ln.split("Дата готовности:")[1].strip() for ln in desc.split("\n") if "Дата готовности:" in ln), "")
            except:
                rd = ""
            if rd:
                deal_ready = rd

            temp.append({
                "Номер сделки": deal_number,
                "Клиент": client_name,
                "Наименование": name,
                "Вид печати": ptype,
                "Выручка": rev,
                "Себестоимость": cost,
                "Прибыль": profit,
                "Маржинальность (%)": margin,
                "Дата готовности": rd or deal_ready or deal_end_dates.get(deal_id, "")
            })
        all_rows.extend(temp)

    df = pd.DataFrame(all_rows)

    # 5) Сводка по типу печати
    def classify(pt):
        if pt == "УФ печать": return "УФ печать"
        if pt == "Цифровая печать": return "Цифровая печать"
        return "Перезаказ"

    df["Класс печати"] = df["Вид печати"].apply(classify)

    summary = df.groupby("Класс печати").agg({
        "Выручка": "sum",
        "Себестоимость": "sum"
    }).reset_index()
    summary["Прибыль"] = summary["Выручка"] - summary["Себестоимость"]
    summary["Маржинальность (%)"] = summary.apply(
        lambda x: round(x["Прибыль"] / x["Выручка"] * 100, 2) if x["Выручка"] else 0,
        axis=1
    )

    # 6) Сводка по месяцам и категориям печати
    df["ParsedDate"] = df["Дата готовности"].apply(extract_date)
    df["Месяц"] = df["ParsedDate"].dt.to_period('M').dt.to_timestamp()

    monthly_summary = df.groupby(["Месяц", "Класс печати"]).agg(
        Выручка=('Выручка', 'sum'),
        Себестоимость=('Себестоимость', 'sum')
    ).reset_index()
    monthly_summary['Прибыль'] = monthly_summary['Выручка'] - monthly_summary['Себестоимость']
    monthly_summary['Маржинальность (%)'] = (
        monthly_summary['Прибыль'] / monthly_summary['Выручка'] * 100
    ).round(2)

    # 7) Запись в файл с тремя листами
    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Данные по сделкам", index=False)
        summary.to_excel(writer, sheet_name="По типу печати", index=False)
        monthly_summary.to_excel(writer, sheet_name="По месяцам", index=False)

    print(f"✅ Всё готово — «{OUTPUT_FILE}»")

finally:
    driver.quit()
