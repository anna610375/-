# -*- coding: utf-8 -*-
# Скрипт автоматизации: собирает вчерашние поступления из FinTablo и вносит оплаты в PrintOffice24
# Запуск:
#     python C:/Users/anna6/Downloads/parse_fin_tablo_and_apply_payments.py

import os
import re
import time
import json
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ====== НАСТРОЙКИ ======
PROFILE_PATH = r"C:\Users\anna6\AppData\Local\Google\Chrome\User Data"
TARGET_PROFILE_NAME = "Пользователь 1"
FIN_TABLO_URL = (
    "https://my.fintablo.ru/dds/index?"
    "date-nav-type=%2B0+days&"
    "RealTransactionSearch%5Btype%5D=in&"
    "RealTransactionSearch%5BfromDate%5D={date}&"
    "RealTransactionSearch%5BtoDate%5D={date}"
)
PO24_BASE = "https://printoffice24.com"
# Селекторы PrintOffice24
SEARCH_INPUT_ID = 'deals_search_input'
DEAL_LINK_XPATH = "//div[@style='margin-right:18px']/b[text()='{deal}']/ancestor::tr//a[contains(@class,'move_to')]"
PLUS_SELECTOR = 'i.fa.fa-plus-square'
# ================================

def find_profile_dir():
    state = json.load(open(os.path.join(PROFILE_PATH, 'Local State'), encoding='utf-8'))
    for dir_name, info in state.get('profile', {}).get('info_cache', {}).items():
        if info.get('name', '').replace('\xa0',' ') == TARGET_PROFILE_NAME:
            return dir_name
    raise RuntimeError(f"Профиль '{TARGET_PROFILE_NAME}' не найден")


def init_driver():
    profile_dir = find_profile_dir()
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={PROFILE_PATH}")
    opts.add_argument(f"--profile-directory={profile_dir}")
    opts.add_argument("--start-maximized")
    opts.add_experimental_option('excludeSwitches',['enable-automation'])
    opts.add_experimental_option('useAutomationExtension', False)
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)


def parse_amount(text):
    val = re.sub(r"[^\d\.,]","", text).replace(',','.')
    try:
        return float(val)
    except:
        return 0.0


def extract_deal_id(text):
    m = re.search(r"(?:№)?(\d{4})", text)
    return m.group(1) if m else ''


def main():
    # 1) Запуск драйвера
    driver = init_driver()
    try:
        # 2) Сбор вчерашних поступлений из FinTablo
        date_dot = (datetime.now() - timedelta(days=1)).strftime('%d.%m.%Y')
        date_slash = date_dot.replace('.', '/')
        url = FIN_TABLO_URL.format(date=date_dot)
        print(f"Открываем FinTablo: {url}")
        driver.get(url)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'table tbody tr')))
        time.sleep(1)

        payments = []
        rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, 'td')
            if len(cols) < 3:
                continue
            ab = cols[1].text.splitlines()
            amount = parse_amount(ab[0] if ab else '')
            bank = ab[1] if len(ab)>1 else ''
            cl = cols[2].text.splitlines()
            client = cl[0] if cl else ''
            deal = extract_deal_id(cl[1] if len(cl)>1 else '')
            if deal:
                payments.append({'Дата': date_slash, 'Номер сделки':deal, 'Сумма':amount, 'Сбербанк':bank, 'Клиент':client})
        if not payments:
            print("Нет поступлений за вчера.")
            return

        # Сохранение для проверки
        df = pd.DataFrame(payments)
        path = os.path.join(os.getcwd(), 'check_payments.xlsx')
        df.to_excel(path, index=False)
        print(f"Сохранили список: {path}")

        # 3) Разнесение оплат в PrintOffice24
        print(f"Начинаем разнесение {len(payments)} платежей...")
        # 3.1 авторизация
        driver.get(f"{PO24_BASE}/login")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))).click()
        time.sleep(2)

        for p in payments:
            print(f"Сделка {p['Номер сделки']}")
            # 3.2 поиск сделки: переходим на страницу списка сделок
            driver.get(f"{PO24_BASE}/dealsList")
            # ждём появления поля поиска
            inp = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, SEARCH_INPUT_ID))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", inp)
            inp.clear()
            inp.send_keys(p['Номер сделки'])
            inp.send_keys(Keys.ENTER)
            # ждём и кликаем по сделке
            xpath = DEAL_LINK_XPATH.format(deal=p['Номер сделки'])
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            ).click()

            # 3.3 добавление платежа
            # прокрутка до кнопки + и клик
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, PLUS_SELECTOR))
            ).click()

            # ввод суммы
            amt = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'payment_amount'))
            )
            amt.clear()
            amt.send_keys(str(p['Сумма']))

            # ввод даты (формат DD/MM/YYYY)
            date_el = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'datepicker_new_payment'))
            )
            # удаляем текущее значение
            date_el.click()
            date_el.send_keys(Keys.CONTROL, 'a')
            date_el.send_keys(Keys.DELETE)
            # вводим дату из таблицы
            date_el.send_keys(p['Дата'])

            # выбор метода оплаты
            sel = Select(WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.NAME, 'pay_method'))
            ))
            if 'Сбербанк' in p['Сбербанк']:
                sel.select_by_visible_text('Сбербанк - перевод на карту / с карты')
            else:
                sel.select_by_visible_text('Безнал')

                        # создать платёж
            create_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'deal_paid_create'))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", create_btn)
            create_btn.click()

            # подтвердить через текст кнопки 'Подтвердить'
            confirm_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Подтвердить']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", confirm_btn)
            confirm_btn.click()
            time.sleep(1)

        print('✅ Разнесение завершено!')

    finally:
        driver.quit()

if __name__=='__main__':
    main()

