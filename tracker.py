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
TICKERS = {
    "ASML EUR": "ASML.AS",
    "Pharming EUR": "PHARM.AS",
    "TDIV EUR": "TDIV.AS",
    "EUNL EUR": "EUNL.AS",
    "VUAA EUR": "VUAA.AS",
    "Magnum EUR": "7RM.DE",
    "DFNS EUR": "DFNS",
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
        "Bitcoin EUR": data["bitcoin"]["eur"],
        "Ethereum EUR": data["ethereum"]["eur"],
    }


# ---------- FX RATE ----------
def haal_eurusd_op():
    fx = yf.Ticker("EURUSD=X")
    hist = fx.history(period="1d")

    if hist.empty:
        raise ValueError("Geen EUR/USD data")

    return float(hist["Close"].iloc[-1])


# ---------- METALEN ----------
def haal_metalen_op(eurusd):
    # USD per ounce
    goud = yf.Ticker("GC=F").history(period="1d")["Close"].iloc[-1]
    zilver = yf.Ticker("SI=F").history(period="1d")["Close"].iloc[-1]
    platina = yf.Ticker("PL=F").history(period="1d")["Close"].iloc[-1]

    # ounce → kg = 32.1507
    factor = 32.1507

    return {
        "Goud EUR/kg": round((goud * factor) / eurusd, 2),
        "Zilver EUR/kg": round((zilver * factor) / eurusd, 2),
        "Platina EUR/kg": round((platina * factor) / eurusd, 2),
    }


# ---------- AANDELEN ----------
def haal_aandelen_op(eurusd):
    resultaten = {}

    for naam, ticker in TICKERS.items():
        t = yf.Ticker(ticker)
        hist = t.history(period="5d")

        if hist.empty:
            continue

        prijs = float(hist["Close"].dropna().iloc[-1])

        # Amerikaanse aandelen omrekenen
        if ticker.endswith(".AS") or ticker.endswith(".DE"):
            resultaten[naam] = round(prijs, 2)
        else:
            # USD → EUR
            resultaten[naam] = round(prijs / eurusd, 2)

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
        print(f"Niet uitgevoerd: {tijd}")
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

    rij = [datum, tijd] + [alle_data[k] for k in kolommen]

    ws.append(rij)
    wb.save(EXCEL_BESTAND)

    print(f"Toegevoegd: {datum} {tijd}")


if __name__ == "__main__":
    main()
