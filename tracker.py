def download_data():
    data = {}

    for naam, ticker in TICKERS.items():
        print(f"Download: {naam} ({ticker})")

        try:
            df = yf.download(
                ticker,
                start=START_DATUM,
                interval="1d",
                progress=False,
            )

            if df.empty:
                print(f"❌ Geen data voor {naam}")
                continue

            serie = df["Close"]
            data[naam] = serie

        except Exception as e:
            print(f"❌ Fout bij {naam}: {e}")

    # 🔥 BELANGRIJK: check of we data hebben
    if not data:
        raise ValueError("Geen enkele ticker gaf data terug")

    combined = pd.DataFrame(data)

    # FX ophalen
    try:
        eurusd = yf.download("EURUSD=X", start=START_DATUM, progress=False)["Close"]
        combined["EURUSD"] = eurusd
    except:
        print("⚠️ FX fallback gebruikt")
        combined["EURUSD"] = 1

    # USD → EUR
    for kolom in ["AGNC EUR", "NVIDIA EUR"]:
        if kolom in combined.columns:
            combined[kolom] = combined[kolom] / combined["EURUSD"]

    # metalen → kg
    factor = 32.1507466
    for kolom in ["Goud EUR/kg", "Zilver EUR/kg", "Platina EUR/kg"]:
        if kolom in combined.columns:
            combined[kolom] = combined[kolom] * factor / combined["EURUSD"]

    combined = combined.drop(columns=["EURUSD"], errors="ignore")

    combined = combined.round(2)

    return combined
