"""Мониторинг цен: собирает каталог, сравнивает с прошлым запуском, сохраняет Excel-отчёт.

Запуск:
    python main.py                  # весь каталог (~1000 товаров, 2–4 минуты)
    python main.py --categories 5   # только первые 5 категорий, для быстрой проверки
    python main.py --simulate       # демо: показать, как выглядит отчёт с изменениями цен
"""
from __future__ import annotations

import argparse
import logging
import random
from datetime import datetime
from pathlib import Path

import currency
import report
import scraper

SNAPSHOT = Path("data/last_run.json")
REPORTS_DIR = Path("reports")


def simulate_changes(old: dict[str, dict], seed: int = 42) -> dict[str, dict]:
    """Демо-режим: меняет часть «прошлых» цен и наличия, чтобы отчёт показал изменения."""
    rnd = random.Random(seed)
    old = {url: dict(p) for url, p in old.items()}
    urls = list(old)
    rnd.shuffle(urls)
    for url in urls[:25]:  # 25 товаров «изменили цену»
        old[url]["price"] = round(old[url]["price"] * rnd.uniform(0.8, 1.25), 2)
    for url in urls[25:30]:  # 5 товаров «раньше не было в наличии»
        old[url]["in_stock"] = False
    for url in urls[30:35]:  # 5 товаров «новые»: убираем их из прошлого снимка
        del old[url]
    old["https://books.toscrape.com/catalogue/removed-demo-book/index.html"] = {
        "title": "Книга, которую убрали с сайта (демо)", "category": "Художественная литература", "price": 19.99,
        "rating": 3, "in_stock": True, "url": "https://books.toscrape.com/catalogue/removed-demo-book/index.html",
    }
    return old


def main() -> None:
    ap = argparse.ArgumentParser(description="Мониторинг цен с выгрузкой в Excel")
    ap.add_argument("--categories", type=int, default=None, help="сколько категорий обойти (по умолчанию все)")
    ap.add_argument("--simulate", action="store_true", help="демо-режим: искусственные изменения цен")
    ap.add_argument("--rate", type=float, default=None, help="курс фунта в рублях вручную (если ЦБ недоступен)")
    ap.add_argument("--delay", type=float, default=0.3, help="пауза между запросами, сек")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
    log = logging.getLogger("main")

    started = datetime.now()
    if args.rate:
        rate, rate_date = args.rate, ""
    else:
        rate, rate_date = currency.get_rate("GBP")
    products = [p.to_dict() for p in scraper.scrape(scraper.Fetcher(delay=args.delay), args.categories)]
    log.info("Собрано товаров: %d", len(products))

    old = report.load_snapshot(SNAPSHOT)
    if args.simulate:
        old = simulate_changes(old or {p["url"]: p for p in products})
    changes = report.compare(old, products)

    stamp = started.strftime("%Y-%m-%d_%H-%M-%S")
    out = REPORTS_DIR / f"prices_{stamp}.xlsx"
    meta = {
        "site": scraper.BASE_URL, "date": started.strftime("%d.%m.%Y %H:%M"), "simulated": args.simulate,
        "rate": rate, "rate_date": rate_date,
    }
    report.build_report(out, products, changes, meta, first_run=old is None)
    if not args.simulate:  # демо-прогон не портит настоящую историю цен
        report.save_snapshot(SNAPSHOT, products)

    log.info("Изменений с прошлого запуска: %d", len(changes))
    log.info("Готово за %d сек. Отчёт: %s", (datetime.now() - started).seconds, out.resolve())


if __name__ == "__main__":
    main()
