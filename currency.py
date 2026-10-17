"""Курс валют ЦБ РФ на сегодня."""
from __future__ import annotations

import logging
import xml.etree.ElementTree as ET

import requests

log = logging.getLogger(__name__)

CBR_XML = "https://www.cbr.ru/scripts/XML_daily.asp"
CBR_JSON_MIRROR = "https://www.cbr-xml-daily.ru/daily_json.js"


def _from_cbr_xml(code: str) -> tuple[float, str]:
    resp = requests.get(CBR_XML, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    root = ET.fromstring(resp.content)
    for v in root.findall("Valute"):
        if v.findtext("CharCode") == code:
            value = float(v.findtext("Value").replace(",", "."))
            nominal = int(v.findtext("Nominal"))
            return value / nominal, root.get("Date", "")
    raise ValueError(f"Валюта {code} не найдена в ответе ЦБ")


def _from_mirror(code: str) -> tuple[float, str]:
    resp = requests.get(CBR_JSON_MIRROR, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    v = data["Valute"][code]
    date = data.get("Date", "")[:10]
    if len(date) == 10:  # 2026-10-17 -> 17.10.2026
        date = f"{date[8:10]}.{date[5:7]}.{date[:4]}"
    return v["Value"] / v["Nominal"], date


def get_rate(code: str = "GBP") -> tuple[float, str]:
    """Возвращает (рублей за 1 единицу валюты, дата курса). Сначала сайт ЦБ, потом зеркало."""
    for source in (_from_cbr_xml, _from_mirror):
        try:
            rate, date = source(code)
            log.info("Курс ЦБ на %s: 1 %s = %.2f ₽", date, code, rate)
            return rate, date
        except Exception as e:  # сеть, формат ответа и т.п.
            log.warning("Не удалось получить курс (%s): %s", source.__name__, e)
    raise RuntimeError("Курс ЦБ недоступен. Укажите его вручную: python main.py --rate 105.5")
