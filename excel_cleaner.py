"""
excel_cleaner.py — Excel dosyası temizleyici ve özet tablo üretici.
Kullanım: python excel_cleaner.py dosya.xlsx [--sheet Sayfa1] [--out cikti.xlsx]
"""

import argparse
import sys
import pandas as pd


def load(path: str, sheet) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    # Tamamen boş satır/sütun sil
    df = df.dropna(how="all").dropna(axis=1, how="all")

    # Sütun adlarını düzelt: küçük harf, boşluk → alt çizgi
    df.columns = (
        df.columns.str.strip().str.lower().str.replace(r"\s+", "_", regex=True)
    )

    # Yinelenen satırları sil
    df = df.drop_duplicates()

    # String sütunlarda baş/son boşluk temizle
    str_cols = df.select_dtypes(include="object").columns
    df[str_cols] = df[str_cols].apply(lambda c: c.str.strip())

    return df.reset_index(drop=True)


def summarize(df: pd.DataFrame) -> pd.DataFrame:
    numeric = df.select_dtypes(include="number")
    if numeric.empty:
        print("Uyarı: Sayısal sütun bulunamadı; özet boş.")
        return pd.DataFrame()

    summary = numeric.agg(["count", "mean", "median", "std", "min", "max"])
    summary.loc["eksik"] = numeric.isna().sum()
    return summary.round(2)


def save(df: pd.DataFrame, summary: pd.DataFrame, out_path: str) -> None:
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Temiz_Veri", index=False)
        if not summary.empty:
            summary.to_excel(writer, sheet_name="Ozet")
    print(f"Kaydedildi → {out_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel temizleyici")
    parser.add_argument("dosya", help="Girdi .xlsx dosyası")
    parser.add_argument("--sheet", default=0, help="Sayfa adı veya indeksi (varsayılan: 0)")
    parser.add_argument("--out", default="cikti.xlsx", help="Çıktı dosyası")
    args = parser.parse_args()

    print(f"Yükleniyor: {args.dosya}")
    df_raw = load(args.dosya, args.sheet)
    print(f"  Ham: {df_raw.shape[0]} satır × {df_raw.shape[1]} sütun")

    df_clean = clean(df_raw)
    print(f"  Temiz: {df_clean.shape[0]} satır × {df_clean.shape[1]} sütun")

    summary = summarize(df_clean)
    if not summary.empty:
        print("\n── Özet ──")
        print(summary.to_string())

    save(df_clean, summary, args.out)


if __name__ == "__main__":
    main()
