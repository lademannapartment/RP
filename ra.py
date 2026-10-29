import asyncio
import datetime as dt
import json
import os
import socket
import sys
import time
import re
import random
import subprocess
from typing import Tuple, Dict
import gc as py_gc
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError
from playwright.async_api import async_playwright, Page

# --- КОНСТАНТЫ ---
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1t8jZFnJ5PxW9ry8SPN8WAGhbsgICn5F9BZCR9d4vB5U/edit"
SHEET_NAME_GDYNIA_APART = "RankGdyniaApart"
SHEET_NAME_GDANSK_APART = "RankGdanskApart"  # Новый лист
SHEET_NAME_RANK_ABN = "Rank ABN"
CREDENTIALS_FILE = "silken-glyph-443313-i4-cdf14b36288a.json"
STAY_NIGHTS = 1
TIMEOUT_SEC = 60
DAYS_FORWARD = 60
MAX_RANK = 400
MAX_PARALLEL_PAGES = 2   # сколько дат обрабатывать одновременно
DATE_RETRIES = 2         # сколько раз повторить дату при ошибке

# Гдыня (Колонки 6, 7, 8, 9 -> F, G, H, I)
LISTING_COLUMNS_GDYNIA_APARTMENTS = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 6,
    "Sailor Apartment- 10 minutes to Seaside & City Center": 7,
}

# Гданьск (Начиная с колонки 9 -> I)
LISTING_COLUMNS_GDANSK_APARTMENTS = {
    "Swish Apartment -10 Min to Old Town Gdańsk,PKP & Shopping Mall": 9,
    "Wave Apartment with Amazing View- 2 km to Baltic Sea": 10,
    "Anchor Apartment Your Docking Point in the Heart of Gdańsk": 11,
    "Boho Apartment - 10 Min to Old Town Gdańsk,PKP & Shopping Mall": 12,
    "Luna Premium Apartment at Old Town Gdansk": 13,
    "Libertas Premium Apartment at Old Town Gdansk": 14,
    "Panorama Apartment with Amazing View- 2 km to Baltic Sea, 15 Minutes to Gdańsk Old Town": 15,
    "Old Town Heaven Apartment": 16,
    "Seredino Blue Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 17,
    "Seredino Green Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 18,
    "Seredino Relax Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 19,
    "Seredino Navy Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 20,
    "Laura Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 21,
    "Bloom Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 22,
    "Seashell Apartment - at Gdańsk Stogi, 5 min by car to Baltic Sea": 23,
    "Coconut Apartment - 5 min to Baltic Sea": 24,
    "Granaria Apartment - Heart of Motlawa Promenade at Old Town Gdansk": 25,
    "Casa Verde Apartment at Old Town Gdańsk": 26,
    "Casa Pluma Apartment at Old Town Gdańsk": 27,
    "Rout Studio 5 min to Old Town Gdansk": 28,
    "Queem Studio 10 min to Old Town Gdansk": 29,
    "Avenue 77 Apartment - 15 Minutes to Seaside": 30,
    "Honey Apartment- Cozy Retreat, 10 Minutes to Gdańsk Old Town": 31,
    "Birch Studio - Peaceful Retreat, 20 Minutes to Gdańsk Old Town & Seaside": 32
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
            CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
    gc = gspread.authorize(creds)


_spreadsheet = None
_ws_cache = {}


def initialize_gspread(sheet_name: str):
    global _spreadsheet
    if sheet_name in _ws_cache:
        return _ws_cache[sheet_name]
    if gc is None: authorize_gspread()
    try:
        if _spreadsheet is None:
            _spreadsheet = gc.open_by_url(SPREADSHEET_URL)
        ws = _spreadsheet.worksheet(sheet_name)
        _ws_cache[sheet_name] = ws
        return ws
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
ABN_GDYNIA_COLUMNS = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 41,
    "Sailor Apartment- 10 minutes to Seaside & City Center": 42,
}


def update_spreadsheet_data(ws: gspread.Worksheet, row_index: int, ranks_map: dict, listings_map: dict, city_name: str,
                            mmrent_count: int = None):
    print(f"\n[GSheets] Обновление данных для {city_name}...")

    main_updates = []
    if mmrent_count is not None and city_name == "Gdansk":
        main_updates.append({"range": f"C{row_index}", "values": [[mmrent_count]]})

    # Лист Rank ABN (берётся из кэша)
    ws_abn = initialize_gspread(SHEET_NAME_RANK_ABN)
    abn_updates = []
    our_found_ranks = [r for r in ranks_map.values() if r is not None]

    for title, col_idx in listings_map.items():
        raw_rank = ranks_map.get(title)
        cell_main = f"{col_to_letter(col_idx)}{row_index}"

        if city_name == "Gdynia":
            target_col_abn = ABN_GDYNIA_COLUMNS.get(title, col_idx)
        else:
            target_col_abn = col_idx
        cell_abn = f"{col_to_letter(target_col_abn)}{row_index}"

        if raw_rank is not None:
            higher_than_us = len([r for r in our_found_ranks if r < raw_rank])
            abn_rank = raw_rank - higher_than_us
            main_val, abn_val = raw_rank, abn_rank
            print(f"   ✅ {title[:15]}: Найдено #{raw_rank} -> Записано")
        else:
            # Не нашли — очищаем старые данные
            main_val, abn_val = "", ""
            print(f"   🗑️ {title[:15]}: НЕ НАЙДЕНО -> Ячейка очищена")

        main_updates.append({"range": cell_main, "values": [[main_val]]})
        abn_updates.append({"range": cell_abn, "values": [[abn_val]]})

    # Одна пакетная запись на каждый лист вместо записи по ячейкам
    try:
        if main_updates:
            ws.batch_update(main_updates, value_input_option='USER_ENTERED')
    except Exception as e:
        print(f"   ❌ Ошибка записи в лист {city_name}: {e}")

    if ws_abn:
        try:
            ws_abn.batch_update(abn_updates, value_input_option='USER_ENTERED')
        except Exception as e:
            print(f"   ❌ Ошибка записи в Rank ABN: {e}")

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


async def wait_for_more_cards(page: Page, before: int, timeout_ms: int = 8000) -> bool:
    try:
        await page.wait_for_function(
            "n => document.querySelectorAll('div[data-testid=\"property-card\"]').length > n",
            arg=before, timeout=timeout_ms)
        return True
    except Exception:
        return False


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
                if not title_clean: continue

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
            print(f"   🎯 Все объекты найдены (загружено карточек: {seen_cards_count}). Останавливаю скролл.")
            break

        if seen_cards_count >= MAX_RANK:
            print(f"   ⏹ Достигнут лимит в {MAX_RANK} карточек.")
            break

        # Скроллим дальше: ждём появления новых карточек, а не фиксированную паузу
        try:
            show_more_button = page.locator('button:has-text("Załaduj"), button:has-text("Show more")').first
            if await show_more_button.is_visible(timeout=1000):
                await show_more_button.click(force=True)
            else:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            if not await wait_for_more_cards(page, current_count):
                break  # новых карточек нет — выдача закончилась
        except:
            break

    return found_ranks, mmrent_count


async def block_aggressively(route):
    if route.request.resource_type in ["image", "media", "font"]:
        await route.abort()
    else:
        await route.continue_()


async def process_date_for_city(page: Page, city_name: str, ws, url_builder: callable,
                                listings_map: dict, date_obj: dt.date):
    if not ws: return
    print(f"\n>>>> ДАТА: {date_obj.strftime('%d.%m.%Y')} — {city_name} <<<<")
    row_index = await asyncio.to_thread(find_row_by_date, ws, date_obj)
    if not row_index: return

    url = url_builder(date_obj, date_obj + dt.timedelta(days=STAY_NIGHTS))
    try:
        await page.goto(url, timeout=90000, wait_until="domcontentloaded")
        # Ждём появления карточек вместо фиксированной паузы
        try:
            await page.wait_for_selector('div[data-testid="property-card"]', timeout=15000)
        except Exception:
            pass

        # Закрываем куки, если появились
        try:
            cookie_btn = page.locator('button#onetrust-accept-btn-handler')
            if await cookie_btn.is_visible(timeout=1000):
                await cookie_btn.click()
        except:
            pass

        ranks_map, mmrent_count = await scrape_cards_and_get_ranks(page, listings_map)
        await asyncio.to_thread(update_spreadsheet_data, ws, row_index, ranks_map, listings_map, city_name, mmrent_count)

    except Exception as e:
        print(f"❌ Ошибка {city_name} на {date_obj}: {e}")
        raise  # пусть верхний уровень повторит дату


import subprocess
async def process_one_date(context, date_obj: dt.date, ws_gdynia, ws_gdansk, sem: asyncio.Semaphore):
    async with sem:
        # Небольшая случайная пауза, чтобы страницы не открывались одновременно
        await asyncio.sleep(random.uniform(1, 3))
        for attempt in range(1, DATE_RETRIES + 1):
            page = await context.new_page()
            try:
                await process_date_for_city(page, "Gdynia", ws_gdynia, build_gdynia_apartments_url,
                                            LISTING_COLUMNS_GDYNIA_APARTMENTS, date_obj)
                await process_date_for_city(page, "Gdansk", ws_gdansk, build_gdansk_apartments_url,
                                            LISTING_COLUMNS_GDANSK_APARTMENTS, date_obj)
                print(f"🧹 Дата {date_obj} обработана")
                return
            except Exception as e:
                print(f"⚠️ Попытка {attempt}/{DATE_RETRIES} для {date_obj} не удалась: {e}")
                await asyncio.sleep(5)
            finally:
                await page.close()
        print(f"❌ Дата {date_obj} пропущена после {DATE_RETRIES} попыток")


async def main_async():
    # Установка браузера Chromium для Playwright (если ещё не установлен)
    print("Checking/Installing Playwright browser...")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
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
                "--js-flags=--max-old-space-size=256"  # Лимит памяти для скриптов
            ]
        )
        # Контекст создаём один раз; блокировку картинок/шрифтов задаём на весь контекст один раз
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
            locale="pl-PL",
            service_workers="block"  # Блокируем фоновые процессы
        )
        await context.route("**/*", block_aggressively)

        today = dt.date.today()
        sem = asyncio.Semaphore(MAX_PARALLEL_PAGES)
        tasks = [
            process_one_date(context, today + dt.timedelta(days=i), ws_gdynia, ws_gdansk, sem)
            for i in range(DAYS_FORWARD)
        ]
        await asyncio.gather(*tasks)
        py_gc.collect()

        await browser.close()
    print("\n🏁 Все города и даты обработаны.")

if __name__ == "__main__":
    asyncio.run(main_async())
