import json
import os
import tempfile
from datetime import datetime
from collections import Counter
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
from flask import Flask, request, send_file
import pandas as pd

app = Flask(__name__)

replacement_nama_fasilitas = {
    "AEON Credit Services Indonesia": "AEON Credit",
    "Adira Dinamika Multi Finance": "Adira",
    "Akulaku Finance Indonesia": "Akulaku",
    "Atome Finance Indonesia": "Atome Finance",
    "Astra Multi Finance": "Astra MF",
    "BFI Finance Indonesia": "BFI",
    "BIMA Multi Finance": "Bima MF",
    "BPD Jawa Barat dan Banten": "BJB",
    "BPD Jawa Barat dan Banten Syariah": "BJB Syariah",
    "BPD Jawa Timur": "Bank Jatim",
    "BPD Sumatera Utara": "Bank Sumut",
    "Bank BCA Syariah": "BCA Syariah",
    "Bank CIMB Niaga": "CIMB Niaga",
    "Bank Central Asia": "BCA",
    "Bank DBS Indonesia": "Bank DBS",
    "Bank Danamon": "Danamon",
    "Bank Danamon Indonesia": "Danamon",
    "Bank Danamon Syariah": "Danamon Syariah",
    "Bank Hibank Indonesia": "Hibank",
    "Bank HSBC Indonesia": "HSBC",
    "Bank KEB Hana Indonesia": "Bank KEB Hana",
    "Bank Mandiri": "Bank Mandiri",
    "Bank Mandiri Taspen": "Bank Mantap",
    "Bank Mayapada Internasional": "Bank Mayapada",
    "Bank Maybank Indonesia": "Maybank",
    "Bank Mega Syariah": "Bank Mega Syariah",
    "Bank Muamalat Indonesia": "Bank Muamalat",
    "Bank Negara Indonesia": "BNI",
    "Bank Neo Commerce": "Akulaku",
    "Bank OCBC NISP": "OCBC NISP",
    "Bank Panin Indonesia": "Panin Bank",
    "Bank Permata": "Bank Permata",
    "Bank Rakyat Indonesia": "BRI",
    "Bank Sahabat Sampoerna": "Bank Sampoerna",
    "Bank Saqu Indonesia ( ": "Bank Saqu",
    "Bank Seabank Indonesia": "Seabank",
    "Bank SMBC Indonesia": "Bank SMBC",
    "Bank Syariah Indonesia": "BSI",
    "Bank Tabungan Negara": "BTN",
    "Bank UOB Indonesia": "Bank UOB",
    "Bank Woori Saudara": "BWS",
    "Bank Woori Saudara Indonesia 1906":"BWS",
    "Bussan Auto Finance": "BAF",
    "Cakrawala Citra Mega Multifinance":"CCM Finance",
    "Commerce Finance": "Seabank",
    "Dana Mandiri Sejahtera": "Dana Mandiri",
    "Esta Dana Ventura": "Esta Dana",
    "Federal International Finance": "FIF",
    "Home Credit Indonesia": "Home Credit",
    "Indodana Multi Finance": "Indodana MF",
    "Indomobil Finance Indonesia": "IMFI",
    "Indonesia Airawata Finance (": "Indonesia Airawata Finance",
    "JACCS Mitra Pinasthika Mustika Finance Indonesia": "JACCS",
    "KB Finansia Multi Finance": "Kreditplus",
    "Kredivo Finance Indonesia": "Kredivo",
    "Krom Bank Indonesia": "Krom Bank",
    "Mandala Multifinance": "Mandala MF",
    "Mandiri Utama Finance": "MUF",
    "Maybank Syariah": "Maybank Syariah",
    "Mega Auto Finance": "MAF",
    "Mega Central Finance": "MCF",
    "Mitra Bisnis Keluarga Ventura": "MBK",
    "Multifinance Anak Bangsa": "MF Anak Bangsa",
    "Panin Bank": "Panin Bank",
    "Permodalan Nasional Madani": "PNM",
    "Pratama Interdana Finance": "Pratama Finance",
    "Standard Chartered Bank": "Standard Chartered",
    "Summit Oto Finance": "Summit Oto",
    "Super Bank Indonesia": "Superbank",
    "Wahana Ottomitra Multiartha": "WOM",
    "Bank Jago": "Bank Jago",
    "Bank BTPN Syariah,": "BTPNS",
    "Bina Artha Ventura": "BAV"
}


def bersihkan_nama_fasilitas(nama):
    if not nama:
        return ""
    nama = nama.strip().replace("PT ", "").replace("PT.", "")
    return replacement_nama_fasilitas.get(nama, nama)

@app.route("/api/proses", methods=["POST"])
def proses():
    files = request.files.getlist("files")
    hasil_semua = []

    for file in files:
        try:
            content = file.read().decode("latin-1")
            data = json.loads(content)
        except:
            continue

        fasilitas = data.get('individual', {}).get('fasilitas', {}).get('kreditPembiayan', [])
        nama_debitur = data.get('individual', {}).get('dataPokokDebitur', [{}])[0].get('namaDebitur', '')
        filename = file.filename or "file.txt"
        nik = os.path.splitext(filename)[0].replace("NIK_", "")

        total_baki = sum(int(f.get('bakiDebet', 0)) for f in fasilitas)
        jumlah_fasilitas = len(fasilitas)
        nama_fasilitas = [bersihkan_nama_fasilitas(f.get('ljkKet', '')) for f in fasilitas]
        fasilitas_joined = "; ".join(nama_fasilitas)

        hasil_semua.append({
            "NIK": "'" + nik,
            "Nama Debitur": nama_debitur,
            "Jumlah Fasilitas": jumlah_fasilitas,
            "Total Baki Debet": total_baki,
            "Daftar Fasilitas": fasilitas_joined
        })

    if not hasil_semua:
        return {"error": "Tidak ada file valid"}, 400

    df = pd.DataFrame(hasil_semua)

    tempdir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_path = os.path.join(tempdir, f"Hasil_SLIK_{timestamp}.xlsx")
    df.to_excel(output_path, index=False)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    wb.save(output_path)

    return send_file(output_path, as_attachment=True, download_name="hasil_slik.xlsx")
