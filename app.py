import streamlit as st
import csv
import io
import time
import re
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# Инициализируем настройки Streamlit
st.set_page_config(page_title="Парсер Закупок 44-ФЗ", page_icon="📊", layout="wide")
st.title("📊 Парсер поставщиков Zakupki.gov.ru")
st.write("Вставьте ссылки из поиска госзакупок, нажмите кнопку и скачайте готовый отчет с контактами и Telegram-ссылками.")

# Обновленный список полей
CSV_FIELDS = [
    "Наименование поставщика", "Страна, код", "Адрес места нахождения",
    "Почтовый адрес", "Телефон", "Ссылка на Telegram", "E-mail поставщика", "Статус поставщика",
    "Наименование контракта", "Вид работ (классификация)",
    "Место выполнения работ", "Номер реестровой записи", "Ссылка на карточку контракта",
]

WORK_TYPE_KEYWORDS = [
    ("Пожарная сигнализация / СОУЭ", ["пожарн", "соуэ", "оповещен"]),
    ("Слаботочные системы", ["слаботочн", "видеонаблюден", "скс", "структурированн"]),
    ("ЭОМ / электромонтаж", ["электромонтаж", "эом", "электроснабжен", "электрическ"]),
    ("Вентиляция / ОВиК", ["вентиляц", "кондиционир", "овик", "отоплен"]),
    ("Капитальный ремонт", ["капитальный ремонт", "капремонт"]),
    ("Реконструкция", ["реконструкц"]),
    ("Отделочные работы", ["отделочн", "отделк"]),
    ("Ремонт помещений", ["ремонт помещен"]),
    ("СМР / строительство", ["строительно-монтажн", "смр", "строительств"]),
    ("Ремонт (прочее)", ["ремонт"]),
]

def classify_work_type(contract_title: str) -> str:
    title_lower = (contract_title or "").lower()
    for label, keywords in WORK_TYPE_KEYWORDS:
        if any(kw in title_lower for kw in keywords):
            return label
    return "Не классифицировано"

def make_telegram_link(phone_text: str) -> str:
    """Извлекает цифры из строки телефона и переводит в ссылку вида https://t.me"""
    if not phone_text:
        return ""
    digits = re.sub(r"\D", "", phone_text)
    if not digits:
        return ""
    if digits.startswith("8") and len(digits) == 11:
        digits = "7" + digits[1:]
    elif len(digits) == 10 and digits.startswith("9"):
        digits = "7" + digits
    return f"https://t.me+{digits}/"

def parse_contract_card(page, card_url, contract_title):
    try:
        page.goto(card_url, timeout=30000)
        page.wait_for_selector(".contractRegNum", timeout=15000)
    except PWTimeout:
        return None

    try:
        reg_num = page.locator(".contractRegNum").inner_text().strip()
    except: reg_num = ""

    try:
        delivery_place = page.locator("//span[contains(text(), 'Место поставки товара') or contains(text(), 'Место выполнения работ')]/following-sibling::span").inner_text().strip()
    except: delivery_place = ""

    try:
        supplier_name = page.locator("//td[contains(text(), 'Полное наименование')]/following-sibling::td").first.inner_text().strip()
    except:
        supplier_name = "Требуется проверка селекторов"

    # Парсинг телефона
    try:
        phone = page.locator("//td[contains(text(), 'Номер телефона') or contains(text(), 'Телефон') and not(contains(text(), 'факс'))]/following-sibling::td").first.inner_text().strip()
    except:
        phone = ""

    # Парсинг E-mail
    try:
        email = page.locator("//td[contains(text(), 'Адрес электронной почты') or contains(text(), 'E-mail') or contains(text(), 'e-mail')]/following-sibling::td").first.inner_text().strip()
    except:
        email = ""

    tg_link = make_telegram_link(phone)

    return {
        "Наименование поставщика": supplier_name,
        "Страна, код": "Российская Федерация",
        "Адрес места нахождения": "", "Почтовый адрес": "", 
        "Телефон": phone,
        "Ссылка на Telegram": tg_link,
        "E-mail поставщика": email,  # Новое поле
        "Статус поставщика": "Активен",
        "Наименование контракта": contract_title,
        "Вид работ (классификация)": classify_work_type(contract_title),
        "Место выполнения работ": delivery_place,
        "Номер реестровой записи": reg_num,
        "Ссылка на карточку контракта": card_url,
    }

# --- ИНТЕРФЕЙС СТРАНИЦЫ ---
urls_input = st.text_area(
    "Вставьте URL ссылок результатов поиска (каждая ссылка с новой строки):",
    height=200,
    placeholder="https://zakupki.gov.ru?...\nhttps://..."
)

if st.button("🚀 Начать сбор данных", type="primary"):
    search_urls = [url.strip() for url in urls_input.split("\n") if url.strip()]
    
    if not search_urls:
        st.warning("⚠️ Пожалуйста, добавьте хотя бы одну ссылку.")
    else:
        output = io.StringIO()
        output.write('\ufeff') # Записываем BOM для корректного отображения кириллицы в Excel
        writer = csv.DictWriter(output, fieldnames=CSV_FIELDS, delimiter=";")
        writer.writeheader()

        status_text = st.empty()
        progress_bar = st.progress(0)
        log_box = st.empty()
        
        results_count = 0
        seen_contracts = set()

        with sync_playwright() as p:
            status_text.text("Запуск системного браузера на сервере...")
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            page = context.new_page()

            for idx, url in enumerate(search_urls):
                status_text.text(f"Обработка поисковой ленты {idx+1} из {len(search_urls)}...")
                
                try:
                    page.goto(url, timeout=45000)
                    page.wait_for_selector(".search-registry-entry-block", timeout=20000)
                    entries = page.locator(".search-registry-entry-block").all()
                except Exception as e:
                    st.error(f"Ошибка загрузки страницы поиска: {e}")
                    continue

                for entry_idx, entry in enumerate(entries):
                    try:
                        link_element = entry.locator(".registry-entry__header-mid__number a").first
                        href = link_element.get_attribute("href")
                        card_url = urljoin("https://zakupki.gov.ru", href)
                        contract_id = link_element.inner_text().strip()
                        contract_title = entry.locator(".registry-entry__body-value").first.inner_text().strip()
                    except:
                        continue

                    if contract_id in seen_contracts:
                        continue

                    log_box.info(f"🔎 Проверяем контракт №{contract_id} ({entry_idx+1}/{len(entries)})")
                    time.sleep(1.5)

                    new_page = context.new_page()
                    try:
                        row_data = parse_contract_card(new_page, card_url, contract_title)
                        if row_data:
                            row_data["Номер реестровой записи"] = contract_id
                            writer.writerow(row_data)
                            seen_contracts.add(contract_id)
                            results_count += 1
                    except Exception as e:
                        print(f"Ошибка карточки: {e}")
                    finally:
                        new_page.close()
                
                progress_bar.progress((idx + 1) / len(search_urls))

            browser.close()

        status_text.success(f"✨ Сбор завершен! Успешно обработано контрактов: {results_count}")
        log_box.empty()
        
        st.download_button(
            label="📥 Скачать результаты в Excel (CSV)",
            data=output.getvalue(),
            file_name="zakupki_contacts_report.csv",
            mime="text/csv"
        )
