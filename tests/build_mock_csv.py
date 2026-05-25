"""
Мок-CSV для тестирования валютного отчёта.

Читает config/bonds.yaml, для каждого ISIN генерит правдоподобные
рыночные данные. Используется только для отладки на машине без
доступа к реальному MOEX ISS CSV.

Usage:
    python tests/build_mock_csv.py /tmp/bonds_usd_mock.csv
"""
import csv
import random
import sys
from pathlib import Path

import yaml

COLUMNS = [
    "TRADEDATE", "ISIN", "SECID", "SHORTNAME", "NAME", "BOARDID", "TYPE",
    "MATDATE", "OFFERDATE", "FACEVALUE", "FACEUNIT", "ISSUESIZE", "LOTVALUE",
    "COUPONPERCENT", "COUPONVALUE", "COUPONFREQUENCY", "COUPONPERIOD",
    "OPEN", "LOW", "HIGH", "CLOSE", "WAPRICE", "VOLUME", "VALUE", "NUMTRADES",
    "ACCINT", "YIELDCLOSE", "YIELDATWAP", "DURATION_DAYS", "DURATION_YEARS",
    "STATUS",
]

TYPICAL_FACEVALUES = {
    "USD": [100, 1000, 200000],
    "EUR": [1000, 100, 100000],
    "CHF": [5000],
    "CNY": [1000, 10000],
}


def build_row(isin, currency, shortname, matdate_iso, is_perpetual=False):
    facevalue = random.choice(TYPICAL_FACEVALUES.get(currency, [1000]))
    coupon = round(random.uniform(2.5, 12.0), 2)
    price = round(random.uniform(70, 110), 2)
    ytm = round(random.uniform(2.5, 10), 2)
    dur_years = None if is_perpetual else round(random.uniform(0.1, 10.0), 1)

    return {
        "TRADEDATE": "2026-04-16",
        "ISIN": isin,
        "SECID": "",
        "SHORTNAME": shortname,
        "NAME": shortname,
        "BOARDID": "TQOD" if currency == "USD" else "TQOE",
        "TYPE": "exchange_bond",
        "MATDATE": "" if is_perpetual else matdate_iso,
        "OFFERDATE": "",
        "FACEVALUE": facevalue,
        "FACEUNIT": currency,
        "ISSUESIZE": 500_000,
        "LOTVALUE": facevalue,
        "COUPONPERCENT": coupon,
        "COUPONVALUE": "",
        "COUPONFREQUENCY": random.choice([1, 2, 4, 12]),
        "COUPONPERIOD": "",
        "OPEN": "", "LOW": "", "HIGH": "",
        "CLOSE": price,
        "WAPRICE": price,
        "VOLUME": "", "VALUE": "", "NUMTRADES": "",
        "ACCINT": "",
        "YIELDCLOSE": ytm,
        "YIELDATWAP": ytm,
        "DURATION_DAYS": "" if is_perpetual else int((dur_years or 0) * 365),
        "DURATION_YEARS": "" if is_perpetual else dur_years,
        "STATUS": "OK",
    }


def main():
    out_path = sys.argv[1] if len(sys.argv) > 1 else "/tmp/bonds_usd_mock.csv"
    random.seed(42)

    yaml_path = Path(__file__).resolve().parent.parent / "config" / "bonds.yaml"
    with open(yaml_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    n_total = 0
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()

        for sec_id, sec in cfg.get("sections", {}).items():
            currency = sec.get("currency", "USD")
            for entry in sec.get("bonds", []):
                isin = entry["isin"] if isinstance(entry, dict) else entry
                is_perpetual = isin in ("RU000A105QW3", "RU000A105QX1")
                year = random.randint(2026, 2035)
                month = random.randint(1, 12)
                day = random.randint(1, 28)
                matdate = f"{year}-{month:02d}-{day:02d}"
                shortname = f"Bond {isin[-5:]}"
                w.writerow(build_row(isin, currency, shortname, matdate, is_perpetual))
                n_total += 1

    print(f"Wrote {n_total} rows to {out_path}")


if __name__ == "__main__":
    main()
