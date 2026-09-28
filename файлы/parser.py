import os
import sys
import time
import logging
import urllib.parse
from typing import List, Dict, Optional, Tuple

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Настройка системного логирования для отображения хода работы в консоли
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)


def load_env_file(filepath: str = ".env") -> None:
    """
    Загружает переменные окружения из файла .env.
    Реализован собственный разбор файла на случай, если библиотека python-dotenv
    не установлена в системе — это гарантирует запуск кода у любого заказчика.
    """
    try:
        from dotenv import load_dotenv
        load_dotenv(filepath)
        return
    except ImportError:
        # Падение внешней библиотеки не блокирует работу приложения
        pass

    if not os.path.exists(filepath):
        return

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip().strip("'\"")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception as e:
        # Логируем ошибку чтения конфигурации, не прерывая выполнение
        logging.warning(f"Не удалось прочитать конфигурационный файл {filepath}: {e}")


class CatalogParser:
    """
    Класс для сборки каталога товаров.
    Изоляция логики в классе позволяет повторно использовать сессию запросов
    и легко масштабировать парсер под другие разделы сайта.
    """

    def __init__(self):
        # Конфигурация вынесена в переменные окружения для защиты от жесткого хардкода
        # и быстрой настройки под задачи клиента без вмешательства в исходный код.
        self.start_url = os.environ.get(
            "START_URL", 
            "http://books.toscrape.com/catalogue/page-1.html"
        )
        # Ограничение по страницам (по умолчанию 5) позволяет провести быстрый тест
        self.max_pages = int(os.environ.get("MAX_PAGES", "5"))
        self.delay = float(os.environ.get("REQUEST_DELAY", "1.0"))
        self.timeout = int(os.environ.get("REQUEST_TIMEOUT", "10"))
        self.output_file = os.environ.get("OUTPUT_FILE", "books_catalog.xlsx")

        # Настоящий User-Agent снижает риск блокировки со стороны защитных систем сайта
        user_agent = os.environ.get(
            "USER_AGENT",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )

        # Сессия requests поддерживает переиспользование TCP-соединений (keep-alive),
        # что существенно повышает скорость работы при многостраничном обходе.
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"
        })

    def _fetch_page(self, url: str, retries: int = 3) -> Optional[str]:
        """
        Выполняет HTTP-запрос с повторными попытками и экспоненциальной задержкой.
        Это защищает процесс от случайных обрывов сети или временной недоступности сайта.
        """
        for attempt in range(1, retries + 1):
            try:
                response = self.session.get(url, timeout=self.timeout)
                
                # Явная проверка статус-кода заменяет попытки распарсить 404/500 страницы
                if response.status_code == 200:
                    # Корректная установка кодировки предотвращает "крякозябры" в тексте
                    response.encoding = "utf-8"
                    return response.text

                logging.warning(
                    f"Сервер вернул код ответа {response.status_code} для {url}. "
                    f"Попытка {attempt} из {retries}."
                )
            except requests.RequestException as e:
                # Перехват сетевых исключений (таймаут, DNS, разрыв соединения)
                logging.warning(
                    f"Сетевая ошибка при доступе к {url}: {e}. "
                    f"Попытка {attempt} из {retries}."
                )

            if attempt < retries:
                # Экспоненциальное увеличение паузы (exponential backoff) снижает нагрузку на сеть
                sleep_time = self.delay * (1.5 ** (attempt - 1))
                time.sleep(sleep_time)

        return None

    def _parse_product_pod(self, pod: BeautifulSoup, current_url: str) -> Optional[Dict[str, str]]:
        """
        Извлекает атрибуты одной карточки товара.
        Использование безопасных проверок для каждого поля изолирует сбои:
        если у одного товара изменится верстка, остальные товары все равно соберутся.
        """
        try:
            # 1. Название товара
            # На данном ресурсе полные названия находятся в атрибуте title тега <a>,
            # тогда как видимый текст тега обрезается двоеточием на странице.
            title_elem = pod.select_one("h3 a")
            if not title_elem:
                return None

            title = title_elem.get("title") or title_elem.get_text(strip=True)

            # 2. Цена товара
            price_elem = pod.select_one(".price_color")
            price = price_elem.get_text(strip=True) if price_elem else "Н/Д"

            # 3. Статус наличия
            # Удаление лишних пробелов и табуляций приводит статус к опрятному виду
            avail_elem = pod.select_one(".availability")
            availability = " ".join(avail_elem.get_text().split()) if avail_elem else "Н/Д"

            # 4. Абсолютная ссылка на карточку
            # urljoin корректно восстанавливает полный адрес из относительных путей типа '../page.html'
            href = title_elem.get("href", "")
            link = urllib.parse.urljoin(current_url, href) if href else "Н/Д"

            return {
                "title": title,
                "price": price,
                "availability": availability,
                "link": link
            }
        except Exception as e:
            # Логируем аномалию конкретной карточки без остановки всей программы
            logging.warning(f"Ошибка обработки карточки товара: {e}")
            return None

    def _get_fallback_data(self) -> List[Dict[str, str]]:
        """
        Встроенный демонстрационный набор данных.
        Если сайт недоступен или отсутствует интернет, программа выполнит показательный
        прогон, чтобы заказчик сразу увидел итоговый Excel-файл без ошибок и настроек.
        """
        return [
            {
                "title": "A Light in the Attic (Демо-образец)",
                "price": "£51.77",
                "availability": "In stock",
                "link": "http://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"
            },
            {
                "title": "Tipping the Velvet (Демо-образец)",
                "price": "£53.74",
                "availability": "In stock",
                "link": "http://books.toscrape.com/catalogue/tipping-the-velvet_999/index.html"
            },
            {
                "title": "Soumission (Демо-образец)",
                "price": "£50.10",
                "availability": "In stock",
                "link": "http://books.toscrape.com/catalogue/soumission_998/index.html"
            },
            {
                "title": "Behind Closed Doors (Демо-образец)",
                "price": "£52.15",
                "availability": "In stock",
                "link": "http://books.toscrape.com/catalogue/behind-closed-doors_997/index.html"
            }
        ]

    def parse_catalog(self) -> Tuple[List[Dict[str, str]], bool]:
        """
        Выполняет итеративный постраничный обход каталога.
        Возвращает собранную коллекцию товаров и признак показательного прогона.
        """
        all_books: List[Dict[str, str]] = []
        current_url = self.start_url
        page_count = 0
        is_fallback = False

        logging.info(f"Запуск парсинга каталога. Старт: {current_url}")

        while current_url and page_count < self.max_pages:
            page_count += 1
            logging.info(f"Загрузка страницы {page_count} из {self.max_pages}: {current_url}")

            html_content = self._fetch_page(current_url)

            # Если первая же страница не ответила, переключаемся на автономный показательный прогон
            if not html_content:
                if page_count == 1:
                    logging.error(
                        "ВНИМАНИЕ: Нет подключения кbooks.toscrape.com (сеть недоступна). "
                        "Включен показательный прогон на встроенных образцах данных. "
                        "Результат все равно сформирован в виде готовой таблицы Excel!"
                    )
                    return self._get_fallback_data(), True
                else:
                    logging.warning(
                        f"Не удалось загрузить {current_url}. Завершение сбора на странице {page_count - 1}."
                    )
                    break

            soup = BeautifulSoup(html_content, "html.parser")
            pods = soup.select("article.product_pod")

            if not pods:
                logging.warning(f"На странице {current_url} не найдены карточки товаров.")
                break

            parsed_on_page = 0
            for pod in pods:
                item = self._parse_product_pod(pod, current_url)
                if item:
                    all_books.append(item)
                    parsed_on_page += 1

            logging.info(f"Успешно обработано товаров на странице: {parsed_on_page}")

            # Постраничный переход через поиски кнопки Next в пагинации
            next_btn = soup.select_one("li.next a")
            if next_btn and next_btn.get("href"):
                next_href = next_btn["href"]
                current_url = urllib.parse.urljoin(current_url, next_href)
            else:
                logging.info("Достигнута последняя страница каталога.")
                current_url = None

            # Задержка между запросами регулирует интенсивность обращения к серверу
            if current_url and page_count < self.max_pages:
                time.sleep(self.delay)

        return all_books, is_fallback


def export_to_excel(data: List[Dict[str, str]], filepath: str) -> str:
    """
    Генерирует стилизованный Excel-файл (.xlsx) с форматированием заголовков,
    кликабельными гиперссылками и автоматической подгонкой ширины колонок.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Каталог товаров"

    # Явное включение сетки повышает удобство чтения таблицы пользователем
    ws.views.sheetView[0].showGridLines = True

    headers = ["№", "Название товара", "Цена", "Наличие", "Ссылка на товар"]
    ws.append(headers)

    # Оформление шапки таблицы: темно-синяя заливка, контрастный белый текст, центрирование
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    ws.row_dimensions[1].height = 28

    # Стили для строк с данными
    data_font = Font(name="Calibri", size=10)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    for idx, item in enumerate(data, start=1):
        row_idx = idx + 1

        ws.cell(row=row_idx, column=1, value=idx).alignment = align_center
        ws.cell(row=row_idx, column=2, value=item.get("title", "")).alignment = align_left
        ws.cell(row=row_idx, column=3, value=item.get("price", "")).alignment = align_right
        ws.cell(row=row_idx, column=4, value=item.get("availability", "")).alignment = align_center

        link_url = item.get("link", "")
        link_cell = ws.cell(row=row_idx, column=5, value=link_url)
        link_cell.alignment = align_left

        if link_url.startswith("http"):
            # Создаем полноценную гиперссылку для прямого перехода из Excel
            link_cell.hyperlink = link_url
            link_cell.font = link_font
        else:
            link_cell.font = data_font

        # Наложение границ ячеек
        for col_idx in range(1, 6):
            c = ws.cell(row=row_idx, column=col_idx)
            if col_idx != 5 or not link_url.startswith("http"):
                c.font = data_font
            c.border = thin_border

        ws.row_dimensions[row_idx].height = 20

    # Вычисление ширины колонок по длиннейшей ячейке с разумным ограничением сверху (max 65),
    # чтобы сверхдлинные ссылки не растягивали столбцы на несколько экранов.
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)

        calculated_width = max(max_len + 4, 12)
        ws.column_dimensions[col_letter].width = min(calculated_width, 65)

    # Обработка ошибки доступа: если целевой файл уже открыт у пользователя в Excel,
    # программа сохраняет результат под новым именем, не аварийно завершаясь.
    final_filepath = filepath
    try:
        wb.save(final_filepath)
    except PermissionError:
        base_name, ext = os.path.splitext(filepath)
        final_filepath = f"{base_name}_result{ext}"
        logging.warning(
            f"Файл '{filepath}' заблокирован другой программой. "
            f"Таблица выгружена в файл: '{final_filepath}'"
        )
        wb.save(final_filepath)

    return final_filepath


def main():
    # Автоматическая загрузка локальных параметров
    load_env_file()

    print("=" * 65)
    print("  АВТОМАТИЧЕСКИЙ СБОР КАТАЛОГА ТОВАРОВ И ВЫГРУЗКА В EXCEL")
    print("=" * 65)

    parser = CatalogParser()
    items, is_fallback = parser.parse_catalog()

    if not items:
        logging.error("Не удалось сформировать набор данных для выгрузки.")
        sys.exit(1)

    saved_path = export_to_excel(items, parser.output_file)

    print("-" * 65)
    if is_fallback:
        print(" [ПОКАЗАТЕЛЬНЫЙ ПРОГОН ВЫПОЛНЕН]")
        print(" Обработан демонстрационный набор данных из-за недоступности сети.")
    else:
        print(" [СБОР ДАННЫХ УСПЕШНО ЗАВЕРШЕН]")

    print(f" Собрано позиций:      {len(items)}")
    print(f" Сформирован файл Excel: {os.path.abspath(saved_path)}")
    print("=" * 65)


if __name__ == "__main__":
    main()
