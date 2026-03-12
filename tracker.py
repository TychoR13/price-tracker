from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

import pandas as pd
import yfinance as yf
from openpyxl import Workbook, load_workbook


EXCEL_BESTAND = Path("prijzen.xlsx")
TIJDZONE = ZoneInfo("Europe/Amsterdam")
START_DATUM = "2025-01-01"


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


def lees_bestaande_datums(ws):
    bestaande_datums = set()

    if ws.max_row <= 1:
        return bestaande_datums

    for row in ws.iter_rows(min_row=2, values_only=True):
        datum_waarde = row[0]
        if datum_waarde:
            bestaande_datums.add(str(datum_waarde))

    return bestaande_datums


def download_slotkoersen(start_datum: str, eind_datum: str) -> pd.DataFrame:
    tickers = {
        "Bitcoin EUR": "BTC-EUR",
        "Ethereum EUR": "ETH-EUR",
        "ASML EUR": "ASML.AS",
        "NVIDIA USD": "NVDA",
    }

    frames = []

    for kolomnaam, ticker in tickers.items():
        df = yf.download(
            ticker,
            start=start_datum,
            end=eind_datum,
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=False,
        )

        if df.empty:
            continue

        if "Close" not in df.columns:
            continue

        serie = df["Close"].copy()
        serie.name = kolomnaam
        frames.append(serie)

    if not frames:
        raise ValueError("Geen koersdata ontvangen van Yahoo Finance.")

    gecombineerd = pd.concat(frames, axis=1)
    gecombineerd.index = pd.to_datetime(gecombineerd.index).tz_localize(None)
    gecombineerd.sort_index(inplace=True)

    return gecombineerd


def voeg_ontbrekende_rijen_toe(ws, data: pd.DataFrame, bestaande_datums: set[str]):
    nu = datetime.now(TIJDZONE)
    vandaag_str = nu.strftime("%Y-%m-%d")
    nu_tijd_str = nu.strftime("%H:%M:%S")

    toegevoegde_rijen = 0

    for datum, row in data.iterrows():
        datum_str = datum.strftime("%Y-%m-%d")

        if datum_str in bestaande_datums:
            continue

        bitcoin = row.get("Bitcoin EUR")
        ethereum = row.get("Ethereum EUR")
        asml = row.get("ASML EUR")
        nvidia = row.get("NVIDIA USD")

        # Sla alleen een rij op als er minstens 1 waarde is
        if pd.isna(bitcoin) and pd.isna(ethereum) and pd.isna(asml) and pd.isna(nvidia):
            continue

        # Voor vandaag zetten we de echte tijd, voor historische rijen 00:00:00
        tijd_str = nu_tijd_str if datum_str == vandaag_str else "00:00:00"

        ws.append([
            datum_str,
            tijd_str,
            None if pd.isna(bitcoin) else round(float(bitcoin), 2),
            None if pd.isna(ethereum) else round(float(ethereum), 2),
            None if pd.isna(asml) else round(float(asml), 2),
            None if pd.isna(nvidia) else round(float(nvidia), 2),
        ])

        toegevoegde_rijen += 1

    return toegevoegde_rijen


def main():
    maak_excel_als_nodig()

    wb = load_workbook(EXCEL_BESTAND)
    ws = wb["Prijzen"]

    bestaande_datums = lees_bestaande_datums(ws)

    vandaag = datetime.now(TIJDZONE).date()
    eind_datum = (vandaag + timedelta(days=1)).strftime("%Y-%m-%d")

    print(f"Download historische + actuele data vanaf {START_DATUM} t/m {vandaag}")

    data = download_slotkoersen(START_DATUM, eind_datum)
    aantal = voeg_ontbrekende_rijen_toe(ws, data, bestaande_datums)

    wb.save(EXCEL_BESTAND)

    print(f"Klaar. {aantal} nieuwe rijen toegevoegd.")


if __name__ == "__main__":
    main()
