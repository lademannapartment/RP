import asyncio
import datetime as dt
import json
import os
import socket
import sys
import time
import re
from typing import Tuple, Dict
import gc as py_gc
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError
from playwright.async_api import async_playwright, Page, TimeoutError

# --- КОНСТАНТЫ ---
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1t8jZFnJ5PxW9ry8SPN8WAGhbsgICn5F9BZCR9d4vB5U/edit"
SHEET_NAME_GDYNIA_APART = "RankGdyniaApart"
SHEET_NAME_GDANSK_APART = "RankGdanskApart"  # Новый лист
SHEET_NAME_RANK_ABN = "Rank ABN"
CREDENTIALS_FILE = "silken-glyph-443313-i4-86edc01b6c81.json"
STAY_NIGHTS = 1
TIMEOUT_SEC = 60
DAYS_FORWARD = 60
MAX_RANK = 400

# Гдыня (Колонки 6, 7, 8, 9 -> F, G, H, I)
LISTING_COLUMNS_GDYNIA_APARTMENTS = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 6,
    "Sailor Apartment- 10 minutes to Seaside & City Center": 7,
}

# Гданьск (Начиная с колонки 9 -> I)
LISTING_COLUMNS_GDANSK_APARTMENTS = {
    "Swish Apartment -10 Min to Old Town Gdańsk,PKP & Shopping Mall": 9,
    "Milky Apartment with Balcony - 3 km to Baltic Sea": 10,
    "Homely Apartment with Balcony - 3 km to Baltic Sea": 11,
    "Wine Apartment - 5 min to Old Town Gdańsk": 12,
    "Wave Apartment with Amazing View- 2 km to Baltic Sea": 13,
    "Anchor Apartment Your Docking Point in the Heart of Gdańsk": 14,
    "Boho Apartment - 10 Min to Old Town Gdańsk,PKP & Shopping Mall": 15,
    "Luna Premium Apartment at Old Town Gdansk": 16,
    "Libertas Premium Apartment at Old Town Gdansk": 17,
    "Postcard Apartment Scenic Escape 10 Minutes to Gdańsk Old Town and Seaside": 18,
    "Panorama Apartment with Amazing View- 2 km to Baltic Sea, 15 Minutes to Gdańsk Old Town": 19,
    "Old Town Heaven Apartment": 20,
    "Seredino Blue Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 21,
    "Seredino Green Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 22,
    "Seredino Relax Apartment- Romantic Escape with a Bathtub in the Bedroom, 5 Minutes to Gdańsk Old Town & Seaside": 23,
    "Seredino Navy Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 24,
    "Laura Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 25,
    "Coast Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 26,
    "Golden Apartment - Premium Apartment at Old Town Gdansk": 27,
    "Amber Apartment - Premium Apartment at Old Town Gdansk": 28,
    "Bloom Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 29,
    "Legit Apartment - Old Town Gdansk": 30,
    "Elite Apartment 15 min to Old Town Gdansk": 31,
    "Seashell Apartment - at Gdańsk Stogi, 5 min by car to Baltic Sea": 32,
    "Rout Studio 5 min to Old Town Gdansk": 33,
    "Queem Studio 10 min to Old Town Gdansk": 34,
    "Exclusive Apartment at Old Town Gdansk": 35,
    "Riverview Apartment at Old Town Gdańsk": 36,
    "Coconut Apartment - 5 min to Baltic Sea": 37
}

# Global gspread client
gc = None


# ---------------------------------------------
# Вспомогательные функции
# ---------------------------------------------
def col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def authorize_gspread():
    global gc
    # GitHub Actions передает секрет через переменную окружения
    creds_json_str = os.environ.get("GOOGLE_CREDS")

    if creds_json_str:
        try:
            creds_json = json.loads(creds_json_str)
            creds = Credentials.from_service_account_info(
                creds_json, scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
        except Exception as e:
            print(f"Ошибка чтения GOOGLE_CREDS: {e}")
            sys.exit(1)
    else:
        # Для локального запуска
        creds = Credentials.from_service_account_file(
            "silken-glyph-443313-i4-86edc01b6c81.json", scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
    gc = gspread.authorize(creds)


def initialize_gspread(sheet_name: str):
    if gc is None: authorize_gspread()
    try:
        sh = gc.open_by_url(SPREADSHEET_URL)
        return sh.worksheet(sheet_name)
    except Exception as e:
        print(f"❌ Ошибка листа {sheet_name}: {e}")
        return None


def find_row_by_date(ws, target_date: dt.date, retries: int = 3):
    for attempt in range(1, retries + 1):
        try:
            values = ws.col_values(1)
            for idx, cell in enumerate(values, start=1):
                if not cell.strip(): continue
                for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                    try:
                        if dt.datetime.strptime(cell.strip(), fmt).date() == target_date:
                            return idx
                    except ValueError:
                        continue
            return None
        except Exception:
            time.sleep(2)
    return None


# Словарь соответствия колонок для листа Rank ABN (Гдыня)
# AL=38, AM=39, AN=40, AO=41
ABN_GDYNIA_COLUMNS = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 38,
    "Sailor Apartment- 10 minutes to Seaside & City Center": 39,
    "Azure Apartment - Seaside & City Center Gdynia": 40,
    "Portlight Apartment -at Marina Yacht Park, Old Town Gdynia": 41
}


def update_spreadsheet_data(ws: gspread.Worksheet, row_index: int, ranks_map: dict, listings_map: dict, city_name: str,
                            mmrent_count: int = None):
    SLEEP_TIME = 1.2
    print(f"\n[GSheets] Обновление данных для {city_name}...")

    # 1. Запись MMRent (только в основной лист Гданьска)
    if mmrent_count is not None and city_name == "Gdansk":
        try:
            ws.update(range_name=f"C{row_index}", values=[[mmrent_count]])
            time.sleep(SLEEP_TIME)
        except:
            pass

    # 2. Инициализация листа Rank ABN
    ws_abn = initialize_gspread(SHEET_NAME_RANK_ABN)
    our_found_ranks = [r for r in ranks_map.values() if r is not None]

    # 3. Основной цикл записи
    for title, col_idx in listings_map.items():
        raw_rank = ranks_map.get(title)

        # Определяем координаты ячеек
        cell_main = f"{col_to_letter(col_idx)}{row_index}"

        # Определяем колонку для листа ABN
        if city_name == "Gdynia":
            target_col_abn = ABN_GDYNIA_COLUMNS.get(title, col_idx)
        else:
            target_col_abn = col_idx
        cell_abn = f"{col_to_letter(target_col_abn)}{row_index}"

        # --- ЛОГИКА ЗАПИСИ ИЛИ ОЧИСТКИ ---

        if raw_rank is not None:
            # Если нашли — записываем новые ранги
            higher_than_us = len([r for r in our_found_ranks if r < raw_rank])
            abn_rank = raw_rank - higher_than_us

            try:
                ws.update(range_name=cell_main, values=[[raw_rank]], value_input_option='USER_ENTERED')
                if ws_abn:
                    ws_abn.update(range_name=cell_abn, values=[[abn_rank]], value_input_option='USER_ENTERED')
                print(f"   ✅ {title[:15]}: Найдено #{raw_rank} -> Записано")
            except:
                pass
        else:
            # ЕСЛИ НЕ НАШЛИ — ОЧИЩАЕМ СТАРЫЕ ДАННЫЕ
            try:
                ws.update(range_name=cell_main, values=[[""]])  # Пустые кавычки удалят старое число
                if ws_abn:
                    ws_abn.update(range_name=cell_abn, values=[[""]])
                print(f"   🗑️ {title[:15]}: НЕ НАЙДЕНО -> Ячейка очищена")
            except:
                pass

        time.sleep(SLEEP_TIME)

# ---------------------------------------------
# Логика скрейпинга
# ---------------------------------------------
def build_gdynia_apartments_url(checkin: dt.date, checkout: dt.date) -> str:
    checkin_str, checkout_str = checkin.strftime("%Y-%m-%d"), checkout.strftime("%Y-%m-%d")
    base = "label=gdynia-riAzz8lSi05Ov0n2ROkbRwS752305354988%3Apl%3Ata%3Ap1%3Ap2%3Aac%3Aap%3Aneg%3Afi%3Atiaud-2382347442848%3Akwd-1011677789197%3Alp9196285%3Ali%3Adec%3Adm%3Appccp%3DUmFuZG9tSVYkc2RlIyh9YavywThF4buZtMEeOgSC-o4&gclid=CjwKCAiA95fLBhBPEiwATXUsxDJb9ebWiR5iLK74z9D1n-Yum-k89bwZbXj7Tpm7Kn8rHej3iAo4WhoC04gQAvD_BwE&aid=311097&dest_id=-501414&dest_type=city&order=price&nflt=ht_id%3D201"
    return f"https://www.booking.com/searchresults.pl.html?{base}&checkin={checkin_str}&checkout={checkout_str}"


def build_gdansk_apartments_url(checkin: dt.date, checkout: dt.date) -> str:
    checkin_str, checkout_str = checkin.strftime("%Y-%m-%d"), checkout.strftime("%Y-%m-%d")
    base = "label=gdynia-riAzz8lSi05Ov0n2ROkbRwS752305354988%3Apl%3Ata%3Ap1%3Ap2%3Aac%3Aap%3Aneg%3Afi%3Atiaud-2382347442848%3Akwd-1011677789197%3Alp9196285%3Ali%3Adec%3Adm%3Appccp%3DUmFuZG9tSVYkc2RlIyh9YavywThF4buZtMEeOgSC-o4&aid=311097&ss=Gdańsk&dest_id=-501400&dest_type=city&order=price&nflt=ht_id%3D201"
    return f"https://www.booking.com/searchresults.pl.html?{base}&checkin={checkin_str}&checkout={checkout_str}"


async def scrape_cards_and_get_ranks(page: Page, listings_map: dict) -> Tuple[Dict[str, int], int]:
    PROPERTY_CARD_SELECTOR = 'div[data-testid="property-card"]'

    def clean_text(text: str) -> str:
        if not text: return ""
        return re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', text).lower()

    found_ranks = {title: None for title in listings_map.keys()}
    remaining = set(listings_map.keys())
    mmrent_count = 0
    seen_cards_count = 0

    print(f"   📈 Поиск объектов (цель: до {MAX_RANK})...")

    for i in range(40):  # Итерации скроллинга
        cards = await page.query_selector_all(PROPERTY_CARD_SELECTOR)
        current_count = len(cards)

        # --- НОВАЯ СТРОКА ЛОГА ---
        if current_count > seen_cards_count:
            print(f"      [Страница] Загружено карточек: {current_count}")

        # Проверяем новые карточки
        for rank, card in enumerate(cards[seen_cards_count:], start=seen_cards_count + 1):
            if rank > MAX_RANK: break

            try:
                element = await card.query_selector('div[data-testid="title"]')
                if not element: continue
                title_clean = clean_text(await element.inner_text())

                if "mmrent" in title_clean:
                    mmrent_count += 1

                for listing_name in list(remaining):
                    if clean_text(listing_name) in title_clean or title_clean in clean_text(listing_name):
                        found_ranks[listing_name] = rank
                        remaining.remove(listing_name)
                        print(f"   ✅ НАЙДЕНО: #{rank} {listing_name[:20]}...")
            except:
                continue

        seen_cards_count = current_count

        if not remaining:
            print(f"   🎯 Все объекты найдены на позиции {rank}. Останавливаю скролл.")
            break

        if seen_cards_count >= MAX_RANK:
            print(f"   ⏹ Достигнут лимит в {MAX_RANK} карточек.")
            break

        # Скроллим дальше
        try:
            show_more_button = page.locator('button:has-text("Załaduj"), button:has-text("Show more")').first
            if await show_more_button.is_visible(timeout=2000):
                await show_more_button.click(force=True)
                await asyncio.sleep(4)
            else:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                await asyncio.sleep(2)
        except:
            break

    return found_ranks, mmrent_count


async def process_date_for_city(page: Page, city_name: str, ws: gspread.Worksheet, url_builder: callable,
                                listings_map: dict, date_obj: dt.date):
    if not ws: return
    print(f"\n>>>> ДАТА: {date_obj.strftime('%d.%m.%Y')} — {city_name} <<<<")
    row_index = find_row_by_date(ws, date_obj)
    if not row_index: return

    # Функция блокировки ресурсов (картинки, шрифты)
    async def block_aggressively(route):
        if route.request.resource_type in ["image", "media", "font"]:
            await route.abort()
        else:
            await route.continue_()

    url = url_builder(date_obj, date_obj + dt.timedelta(days=STAY_NIGHTS))
    try:
        # 1. Включаем блокировку ресурсов перед переходом
        await page.route("**/*", block_aggressively)

        # 2. Переходим по ссылке
        await page.goto(url, timeout=90000, wait_until="domcontentloaded")
        await asyncio.sleep(2)

        # 3. Закрываем куки, если появились
        try:
            cookie_btn = page.locator('button#onetrust-accept-btn-handler')
            if await cookie_btn.is_visible(timeout=3000):
                await cookie_btn.click()
        except:
            pass

        # 4. Собираем данные
        ranks_map, mmrent_count = await scrape_cards_and_get_ranks(page, listings_map)

        # 5. Отправляем в Google Таблицы
        update_spreadsheet_data(ws, row_index, ranks_map, listings_map, city_name, mmrent_count)

        # 6. ВАЖНО: Отключаем блокировку для этого запроса, чтобы не накапливать обработчики
        await page.unroute("**/*")

    except Exception as e:
        print(f"❌ Ошибка {city_name} на {date_obj}: {e}")

import subprocess
async def main_async():
    # --- ДОБАВЬТЕ ЭТИ СТРОКИ ---
    print("Checking/Installing Playwright browser...")
    subprocess.run(["python", "-m", "playwright", "install", "chromium"])
    # ---------------------------
    authorize_gspread()
    ws_gdynia = initialize_gspread(SHEET_NAME_GDYNIA_APART)
    ws_gdansk = initialize_gspread(SHEET_NAME_GDANSK_APART)

    now_str = dt.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    if ws_gdynia: ws_gdynia.update_cell(1, 1, f"📅 Skan: {now_str}")
    if ws_gdansk: ws_gdansk.update_cell(1, 1, f"📅 Skan: {now_str}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--js-flags='--max-old-space-size=256'"  # Лимит памяти для скриптов
            ]
        )
        # Создаем контекст один раз
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
            locale="pl-PL",
            service_workers="block"  # Блокируем фоновые процессы
        )

        today = dt.date.today()
        #today = dt.date(2026, 1, 15)
        for i in range(DAYS_FORWARD):
            date_obj = today + dt.timedelta(days=i)

            # СОЗДАЕМ ЧИСТУЮ СТРАНИЦУ
            page = await context.new_page()

            try:
                await process_date_for_city(page, "Gdynia", ws_gdynia, build_gdynia_apartments_url,
                                            LISTING_COLUMNS_GDYNIA_APARTMENTS, date_obj)

                await process_date_for_city(page, "Gdansk", ws_gdansk, build_gdansk_apartments_url,
                                            LISTING_COLUMNS_GDANSK_APARTMENTS, date_obj)
            finally:
                await page.close()
                py_gc.collect()  # Теперь конфликта не будет
                print(f"🧹 ОЗУ очищена после {date_obj}")

        await browser.close()
    print("\n🏁 Все города и даты обработаны.")

if __name__ == "__main__":
    asyncio.run(main_async())
