from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
import yfinance as yf


EXCEL_BESTAND = Path("prijzen.xlsx")
TIJDZONE = ZoneInfo("Europe/Amsterdam")


def haal_crypto_prijzen_op():
    url = "https://api.coingecko.com/api/v3/simple/price"
    params = {
        "ids": "bitcoin,ethereum",
        "vs_currencies": "eur"
    }
    response = requests.get(url, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()

    return {
        "bitcoin_eur": data["bitcoin"]["eur"],
        "ethereum_eur": data["ethereum"]["eur"],
    }


def haal_aandelen_op():
    asml = yf.Ticker("ASML.AS")
    nvda = yf.Ticker("NVDA")

    asml_hist = asml.history(period="5d", interval="1d")
    nvda_hist = nvda.history(period="5d", interval="1d")

    if asml_hist.empty:
        raise ValueError("Geen koersdata gevonden voor ASML.AS")
    if nvda_hist.empty:
        raise ValueError("Geen koersdata gevonden voor NVDA")

    asml_prijs = float(asml_hist["Close"].dropna().iloc[-1])
    nvda_prijs = float(nvda_hist["Close"].dropna().iloc[-1])

    return {
        "asml_eur": round(asml_prijs, 2),
        "nvda_usd": round(nvda_prijs, 2),
    }


def maak_excel_als_nodig():
    if not EXCEL_BESTAND.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Prijzen"
        ws.append([
            "Datum",
            "Tijd",
            "Bitcoin EUR",
            "Ethereum EUR",
            "ASML EUR",
            "NVIDIA USD",
        ])
        wb.save(EXCEL_BESTAND)


def laatste_datum(ws):
    if ws.max_row <= 1:
        return None
    return ws.cell(row=ws.max_row, column=1).value


def main():
    nu = datetime.now(TIJDZONE)
    datum = nu.strftime("%Y-%m-%d")
    tijd = nu.strftime("%H:%M:%S")

    print(f"Testmodus: lokale NL-tijd is {tijd}, we gaan door.")

    maak_excel_als_nodig()

    wb = load_workbook(EXCEL_BESTAND)
    ws = wb["Prijzen"]

    if laatste_datum(ws) == datum:
        print(f"Bestand was al bijgewerkt voor {datum}")
        return

    crypto = haal_crypto_prijzen_op()
    aandelen = haal_aandelen_op()

    ws.append([
        datum,
        tijd,
        crypto["bitcoin_eur"],
        crypto["ethereum_eur"],
        aandelen["asml_eur"],
        aandelen["nvda_usd"],
    ])

    wb.save(EXCEL_BESTAND)
    print(f"Toegevoegd voor {datum} {tijd}")


if __name__ == "__main__":
    main()
