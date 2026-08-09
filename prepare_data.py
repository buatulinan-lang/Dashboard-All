"""
Konversi file Excel data (format: satu sheet per cabang, sama seperti
Gabungan_Semua_Cabang.xlsx) menjadi data/latest_data.csv.gz — file bawaan
yang dipakai dashboard secara otomatis tanpa perlu upload manual.

Kenapa perlu dikonversi? File Excel aslinya biasanya cukup besar (puluhan
MB), sedangkan GitHub lewat browser cuma bisa upload file sampai 25MB.
Setelah dikonversi (dedup + kompresi), ukurannya jadi jauh lebih kecil.

Cara pakai (di folder dashboard_app ini):
    python prepare_data.py path/ke/file_excel_terbaru.xlsx

Setelah selesai, file data/latest_data.csv.gz akan dibuat/ditimpa.
Langkah selanjutnya: upload/timpa file itu ke folder data/ di repo GitHub
Anda, lalu reboot app di Streamlit Cloud (atau tunggu redeploy otomatis).
"""
import sys
from pathlib import Path

import pandas as pd


def main():
    if len(sys.argv) != 2:
        print("Cara pakai: python prepare_data.py path/ke/file_excel_terbaru.xlsx")
        sys.exit(1)

    src = Path(sys.argv[1])
    if not src.exists():
        print(f"File tidak ditemukan: {src}")
        sys.exit(1)

    print(f"Membaca {src} ...")
    xls = pd.ExcelFile(src, engine='openpyxl')
    frames = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if df.empty:
            continue
        df['CABANG'] = sheet
        frames.append(df)

    if not frames:
        print("File tidak berisi data yang bisa dibaca.")
        sys.exit(1)

    full = pd.concat(frames, ignore_index=True, sort=False)

    original_cols = [c for c in full.columns if c != 'CABANG']
    before = len(full)
    full = full.drop_duplicates(subset=original_cols + ['CABANG'], keep='first').reset_index(drop=True)
    after = len(full)

    out_dir = Path(__file__).parent / "data"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / "latest_data.csv.gz"
    full.to_csv(out_path, index=False, compression='gzip')

    size_mb = out_path.stat().st_size / (1024 * 1024)
    print(f"Selesai. {before:,} baris mentah -> {after:,} baris unik (duplikat dihapus).")
    print(f"File tersimpan: {out_path} ({size_mb:.1f} MB)")
    if size_mb > 25:
        print("PERINGATAN: file masih di atas 25MB, upload lewat browser GitHub mungkin gagal. "
              "Pertimbangkan Git LFS atau kompres data lebih lanjut.")
    print("\nLangkah selanjutnya: upload/timpa file 'data/latest_data.csv.gz' ini ke repo GitHub Anda, "
          "lalu reboot app di Streamlit Cloud.")


if __name__ == '__main__':
    main()
