from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
import yfinance as yf


EXCEL_BESTAND = Path("prijzen.xlsx")
TIJDZONE = ZoneInfo("Europe/Amsterdam")
START_DATUM = datetime(2025, 1, 1)


def haal_crypto_prijs_op(coin, datum):
    timestamp = int(datum.timestamp())

    url = f"https://api.coingecko.com/api/v3/coins/{coin}/history"
    params = {
        "date": datum.strftime("%d-%m-%Y"),
        "localization": "false"
    }

    r = requests.get(url, params=params)
    data = r.json()

    return data["market_data"]["current_price"]["eur"]


def haal_aandeel_prijs(ticker, datum):
    t = yf.Ticker(ticker)
    hist = t.history(start=datum.strftime("%Y-%m-%d"),
                     end=(datum + timedelta(days=1)).strftime("%Y-%m-%d"))

    if hist.empty:
        return None

    return float(hist["Close"].iloc[0])


def maak_excel_als_nodig():
    if not EXCEL_BESTAND.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Prijzen"

        ws.append([
            "Datum",
            "Bitcoin EUR",
            "Ethereum EUR",
            "ASML EUR",
            "NVIDIA USD"
        ])

        wb.save(EXCEL_BESTAND)


def laatste_datum(ws):
    if ws.max_row <= 1:
        return None

    return datetime.strptime(ws.cell(row=ws.max_row, column=1).value, "%Y-%m-%d")


def main():

    maak_excel_als_nodig()

    wb = load_workbook(EXCEL_BESTAND)
    ws = wb["Prijzen"]

    last_date = laatste_datum(ws)

    if last_date is None:
        current = START_DATUM
    else:
        current = last_date + timedelta(days=1)

    vandaag = datetime.now(TIJDZONE)

    while current.date() <= vandaag.date():

        datum_str = current.strftime("%Y-%m-%d")
        print("Fetching", datum_str)

        btc = haal_crypto_prijs_op("bitcoin", current)
        eth = haal_crypto_prijs_op("ethereum", current)

        asml = haal_aandeel_prijs("ASML.AS", current)
        nvda = haal_aandeel_prijs("NVDA", current)

        ws.append([
            datum_str,
            btc,
            eth,
            asml,
            nvda
        ])

        current += timedelta(days=1)

    wb.save(EXCEL_BESTAND)

    print("Data update compleet")


if __name__ == "__main__":
    main()
