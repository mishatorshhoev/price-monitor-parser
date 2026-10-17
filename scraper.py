"""Сбор товаров с сайта: обход категорий и страниц каталога."""
from __future__ import annotations

import logging
import time
from dataclasses import asdict, dataclass
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

BASE_URL = "https://books.toscrape.com/"
HEADERS = {"User-Agent": "Mozilla/5.0 (price-monitor demo; +https://github.com/)"}
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# Перевод категорий учебного сайта на русский
CATEGORIES_RU = {
    "Travel": "Путешествия", "Mystery": "Детективы", "Historical Fiction": "Исторические романы",
    "Sequential Art": "Комиксы", "Classics": "Классика", "Philosophy": "Философия", "Romance": "Любовные романы",
    "Womens Fiction": "Женская проза", "Fiction": "Художественная литература", "Childrens": "Детские книги",
    "Religion": "Религия", "Nonfiction": "Нон-фикшн", "Music": "Музыка", "Default": "Разное",
    "Science Fiction": "Научная фантастика", "Sports and Games": "Спорт и игры", "Add a comment": "Без категории",
    "Fantasy": "Фэнтези", "New Adult": "Для молодёжи 18+", "Young Adult": "Подростковая литература",
    "Science": "Наука", "Poetry": "Поэзия", "Paranormal": "Мистика", "Art": "Искусство", "Psychology": "Психология",
    "Autobiography": "Автобиографии", "Parenting": "Воспитание детей", "Adult Fiction": "Проза для взрослых",
    "Humor": "Юмор", "Horror": "Ужасы", "History": "История", "Food and Drink": "Кулинария",
    "Christian Fiction": "Христианская проза", "Business": "Бизнес", "Biography": "Биографии",
    "Thriller": "Триллеры", "Contemporary": "Современная проза", "Spirituality": "Духовность",
    "Academic": "Учебная литература", "Self Help": "Саморазвитие", "Historical": "Историческая проза",
    "Christian": "Христианство", "Suspense": "Остросюжетные", "Short Stories": "Рассказы", "Novels": "Романы",
    "Health": "Здоровье", "Politics": "Политика", "Cultural": "Культура", "Erotica": "Эротика", "Crime": "Криминал",
}


@dataclass
class Product:
    title: str
    category: str
    price: float  # в валюте сайта (фунты); в рубли переводится при выводе
    rating: int
    in_stock: bool
    url: str

    def to_dict(self) -> dict:
        return asdict(self)


class Fetcher:
    """Скачивает страницы вежливо: с паузой между запросами и повторами при сбоях."""

    def __init__(self, delay: float = 0.3, retries: int = 3, timeout: int = 15):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.delay = delay
        self.retries = retries
        self.timeout = timeout

    def get(self, url: str) -> str:
        for attempt in range(1, self.retries + 1):
            try:
                resp = self.session.get(url, timeout=self.timeout)
                resp.raise_for_status()
                resp.encoding = "utf-8"
                time.sleep(self.delay)
                return resp.text
            except requests.RequestException as e:
                log.warning("Ошибка загрузки %s (попытка %d/%d): %s", url, attempt, self.retries, e)
                time.sleep(2 * attempt)
        raise RuntimeError(f"Не удалось загрузить {url}")


# ---------- разбор HTML (отдельно от загрузки, чтобы легко тестировать) ----------

def parse_categories(html: str, page_url: str) -> list[tuple[str, str]]:
    """Список (название категории, ссылка) из бокового меню."""
    soup = BeautifulSoup(html, "html.parser")
    links = soup.select("div.side_categories ul li ul li a")
    result = []
    for a in links:
        name = a.get_text(strip=True)
        result.append((CATEGORIES_RU.get(name, name), urljoin(page_url, a["href"])))
    return result


def parse_price(text: str) -> float:
    digits = "".join(ch for ch in text if ch.isdigit() or ch == ".")
    return float(digits) if digits else 0.0


def parse_listing(html: str, page_url: str, category: str) -> tuple[list[Product], str | None]:
    """Товары со страницы каталога и ссылка на следующую страницу (или None)."""
    soup = BeautifulSoup(html, "html.parser")
    products = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        rating_el = card.select_one("p.star-rating")
        rating_word = next((c for c in rating_el.get("class", []) if c in RATINGS), "") if rating_el else ""
        stock_el = card.select_one("p.instock.availability, p.availability")
        products.append(Product(
            title=link.get("title") or link.get_text(strip=True),
            category=category,
            price=parse_price(card.select_one("p.price_color").get_text()),
            rating=RATINGS.get(rating_word, 0),
            in_stock=bool(stock_el and "in stock" in stock_el.get_text().lower()),
            url=urljoin(page_url, link["href"]),
        ))
    next_el = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_el["href"]) if next_el else None
    return products, next_url


# ---------- обход сайта ----------

def scrape(fetcher: Fetcher, max_categories: int | None = None) -> list[Product]:
    """Обходит все категории и все страницы внутри каждой."""
    categories = parse_categories(fetcher.get(BASE_URL), BASE_URL)
    if max_categories:
        categories = categories[:max_categories]
    log.info("Найдено категорий: %d", len(categories))

    all_products: list[Product] = []
    seen_urls: set[str] = set()
    for i, (name, url) in enumerate(categories, 1):
        page_url: str | None = url
        count = 0
        while page_url:
            items, page_url = parse_listing(fetcher.get(page_url), page_url, name)
            for p in items:
                if p.url not in seen_urls:  # защита от дублей
                    seen_urls.add(p.url)
                    all_products.append(p)
                    count += 1
        log.info("[%d/%d] %s: %d товаров", i, len(categories), name, count)
    return all_products
