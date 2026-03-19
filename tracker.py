from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
import yfinance as yf
from openpyxl import Workbook, load_workbook


EXCEL_BESTAND = Path("prijzen.xlsx")
TIJDZONE = ZoneInfo("Europe/Amsterdam")
START_DATUM = "2026-01-01"


TICKERS = {
    "Bitcoin EUR": "BTC-EUR",
    "Ethereum EUR": "ETH-EUR",
    "ASML EUR": "ASML.AS",
    "Pharming EUR": "PHARM.AS",
    "TDIV EUR": "TDIV.AS",
    "EUNL EUR": "IWDA.AS",
    "VUAA EUR": "VUAA.DE",
    "Magnum EUR": "7RM.DU",
    "DFNS EUR": "DFNS.PA",
    "AGNC EUR": "AGNC",
    "NVIDIA EUR": "NVDA",
    "Goud EUR/kg": "GC=F",
    "Zilver EUR/kg": "SI=F",
    "Platina EUR/kg": "PL=F",
}


def download_data():
    data = {}

    for naam, ticker in TICKERS.items():
        df = yf.download(
            ticker,
            start=START_DATUM,
            interval="1d",
            progress=False,
        )

        if df.empty:
            continue

        serie = df["Close"]
        data[naam] = serie

    combined = pd.DataFrame(data)

    # USD → EUR voor bepaalde assets
    eurusd = yf.download("EURUSD=X", start=START_DATUM)["Close"]

    for kolom in ["AGNC EUR", "NVIDIA EUR"]:
        if kolom in combined.columns:
            combined[kolom] = combined[kolom] / eurusd

    # metalen ounce → kg
    factor = 32.1507466
    for kolom in ["Goud EUR/kg", "Zilver EUR/kg", "Platina EUR/kg"]:
        if kolom in combined.columns:
            combined[kolom] = combined[kolom] * factor / eurusd

    combined = combined.round(2)

    return combined


def maak_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Prijzen"

    headers = ["Datum"] + list(df.columns)
    ws.append(headers)

    for datum, row in df.iterrows():
        ws.append([datum.strftime("%Y-%m-%d")] + list(row))

    wb.save(EXCEL_BESTAND)


def main():
    df = download_data()
    maak_excel(df)
    print("Historische data correct opgeslagen!")


if __name__ == "__main__":
    main()
