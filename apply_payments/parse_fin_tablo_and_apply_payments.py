# -*- coding: utf-8 -*-
# Скрипт автоматизации: собирает вчерашние поступления из FinTablo и вносит оплаты в PrintOffice24
# Запуск:
#     python C:/Users/anna6/Downloads/parse_fin_tablo_and_apply_payments.py

import os
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
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from shared_utils.utils import parse_amount, extract_deal_id
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
    path = os.path.join(PROFILE_PATH, 'Local State')
    with open(path, encoding='utf-8') as f:
        state = json.load(f)
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





def safe_click(driver, by, value, timeout=10, retries=3):
    """Click an element safely, retrying if it becomes stale or intercepted."""
    for attempt in range(retries):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            element.click()
            return
        except (StaleElementReferenceException, ElementClickInterceptedException):
            if attempt == retries - 1:
                raise
            time.sleep(0.5)


def safe_send_keys(driver, by, value, keys, timeout=10, retries=3, clear=False):
    """Send keys to an element with retries handling stale/intercepted issues."""
    for attempt in range(retries):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            if clear:
                element.clear()
            element.click()
            # allow passing a list/tuple of keys
            if isinstance(keys, (list, tuple)):
                element.send_keys(*keys)
            else:
                element.send_keys(keys)
            return
        except (StaleElementReferenceException, ElementClickInterceptedException):
            if attempt == retries - 1:
                raise
            time.sleep(0.5)


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
        path = os.path.abspath(os.path.join(os.getcwd(), 'check_payments.xlsx'))
        with pd.ExcelWriter(path) as writer:
            df.to_excel(writer, index=False, sheet_name='FinTablo')
        print(f"Сохранили список: {path}")

        # 3) Разнесение оплат в PrintOffice24
        print(f"Начинаем разнесение {len(payments)} платежей...")
        # 3.1 авторизация (если требуется)
        driver.get(f"{PO24_BASE}/login")
        try:
            safe_click(driver, By.CSS_SELECTOR, 'button[type="submit"]', timeout=5)
            time.sleep(2)
        except Exception:
            # возможно, профиль уже авторизован
            pass

        results = []

        for p in payments:
            print(f"Сделка {p['Номер сделки']}")
            try:
                # 3.2 поиск сделки: переходим на страницу списка сделок
                driver.get(f"{PO24_BASE}/dealsList")
                # ждём появления поля поиска
                safe_send_keys(
                    driver,
                    By.ID,
                    SEARCH_INPUT_ID,
                    [p['Номер сделки'], Keys.ENTER],
                    clear=True,
                )
                # ждём и кликаем по сделке
                xpath = DEAL_LINK_XPATH.format(deal=p['Номер сделки'])
                safe_click(driver, By.XPATH, xpath)

                # 3.3 добавление платежа
                # прокрутка до кнопки + и клик
                safe_click(driver, By.CSS_SELECTOR, PLUS_SELECTOR)

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
                safe_click(driver, By.ID, 'deal_paid_create')

                # подтвердить через текст кнопки 'Подтвердить'
                safe_click(driver, By.XPATH, "//button[normalize-space()='Подтвердить']")
                time.sleep(1)

                results.append({**p, 'Статус': 'OK'})
            except Exception as e:
                results.append({**p, 'Статус': f'Ошибка: {e}'})

        print('✅ Разнесение завершено!')

        if results:
            with pd.ExcelWriter(path, mode='a', if_sheet_exists='replace') as writer:
                pd.DataFrame(results).to_excel(writer, index=False, sheet_name='PrintOffice')

    finally:
        driver.quit()

if __name__=='__main__':
    main()

