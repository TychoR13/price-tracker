from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import os

import requests
from openpyxl import Workbook, load_workbook
import yfinance as yf


EXCEL_BESTAND = Path("prijzen.xlsx")
TIJDZONE = ZoneInfo("Europe/Amsterdam")


# ---------- CONFIG ----------
TICKERS_EUR = {
    "ASML EUR": "ASML.AS",
    "Pharming EUR": "PHARM.AS",
    "TDIV EUR": "TDIV.AS",
    "EUNL EUR": "IWDA.AS",   # FIX
    "VUAA EUR": "VUAA.DE",   # FIX
    "Magnum EUR": "7RM.DU",
    "DFNS EUR": "DFNS.PA",
}


# USD → EUR
TICKERS_USD_TO_EUR = {
    "AGNC EUR": "AGNC",
    "NVIDIA EUR": "NVDA",
}


# ---------- CRYPTO ----------
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
        "Bitcoin EUR": round(data["bitcoin"]["eur"], 2),
        "Ethereum EUR": round(data["ethereum"]["eur"], 2),
    }


# ---------- FX ----------
def haal_eurusd_op():
    fx = yf.Ticker("EURUSD=X")
    hist = fx.history(period="5d")

    if hist.empty:
        raise ValueError("Geen EUR/USD data")

    return float(hist["Close"].dropna().iloc[-1])


# ---------- METALEN ----------
def haal_metalen_op(eurusd):
    goud = yf.Ticker("GC=F").history(period="5d")["Close"].dropna().iloc[-1]
    zilver = yf.Ticker("SI=F").history(period="5d")["Close"].dropna().iloc[-1]
    platina = yf.Ticker("PL=F").history(period="5d")["Close"].dropna().iloc[-1]

    factor = 32.1507466  # ounce → kg

    return {
        "Goud EUR/kg": round((goud * factor) / eurusd, 2),
        "Zilver EUR/kg": round((zilver * factor) / eurusd, 2),
        "Platina EUR/kg": round((platina * factor) / eurusd, 2),
    }


# ---------- AANDELEN ----------
def haal_aandelen_op(eurusd):
    resultaten = {}

    # EUR assets
    for naam, ticker in TICKERS_EUR.items():
        hist = yf.Ticker(ticker).history(period="5d")

        if hist.empty:
            print(f"Geen data voor {naam}")
            resultaten[naam] = None
            continue

        prijs = float(hist["Close"].dropna().iloc[-1])
        resultaten[naam] = round(prijs, 2)

    # USD → EUR assets
    for naam, ticker in TICKERS_USD_TO_EUR.items():
        hist = yf.Ticker(ticker).history(period="5d")

        if hist.empty:
            print(f"Geen data voor {naam}")
            resultaten[naam] = None
            continue

        prijs_usd = float(hist["Close"].dropna().iloc[-1])
        resultaten[naam] = round(prijs_usd / eurusd, 2)

    return resultaten


# ---------- EXCEL ----------
def maak_excel_als_nodig(kolommen):
    if not EXCEL_BESTAND.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Prijzen"
        ws.append(["Datum", "Tijd"] + kolommen)
        wb.save(EXCEL_BESTAND)


def laatste_datum(ws):
    if ws.max_row <= 1:
        return None
    return ws.cell(row=ws.max_row, column=1).value


# ---------- MAIN ----------
def main():
    nu = datetime.now(TIJDZONE)
    datum = nu.strftime("%Y-%m-%d")
    tijd = nu.strftime("%H:%M:%S")

    github_event_name = os.getenv("GITHUB_EVENT_NAME", "")

    if github_event_name != "workflow_dispatch" and nu.hour != 0:
        print("Niet middernacht → skip")
        return

    eurusd = haal_eurusd_op()

    crypto = haal_crypto_prijzen_op()
    aandelen = haal_aandelen_op(eurusd)
    metalen = haal_metalen_op(eurusd)

    alle_data = {**crypto, **aandelen, **metalen}

    kolommen = list(alle_data.keys())

    maak_excel_als_nodig(kolommen)

    wb = load_workbook(EXCEL_BESTAND)
    ws = wb["Prijzen"]

    if laatste_datum(ws) == datum:
        print("Vandaag al gedaan")
        return

    ws.append([datum, tijd] + [alle_data[k] for k in kolommen])

    wb.save(EXCEL_BESTAND)
    print("Toegevoegd!")


if __name__ == "__main__":
    main()
