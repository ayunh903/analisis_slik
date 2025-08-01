import json
from http.server import BaseHTTPRequestHandler
import pandas as pd
import os
from datetime import datetime
from collections import Counter
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side

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


def bersihkan_nama_fasilitas(nama_fasilitas: str) -> str:
    if not nama_fasilitas:
        return ""
    lower_fasilitas = nama_fasilitas.lower()
    if "d/h" in lower_fasilitas:
        nama_bersih = nama_fasilitas[:lower_fasilitas.find("d/h")].strip()
    elif "d.h" in lower_fasilitas:
        nama_bersih = nama_fasilitas[:lower_fasilitas.find("d.h")].strip()
    else:
        nama_bersih = nama_fasilitas.strip()
    for pattern in ["PT ", "PT.", "PD.", "(Persero)", "(Perseroda)", "(UUS)", " Tbk"]:
        nama_bersih = nama_bersih.replace(pattern, "")
        nama_bersih = nama_bersih.replace("Bank Perekonomian Rakyat Syariah", "BPRS")
    nama_bersih = nama_bersih.replace("Bank Perekonomian Rakyat", "BPR")
    nama_bersih = nama_bersih.replace("Koperasi Simpan Pinjam", "KSP")
    nama_bersih = nama_bersih.strip()
    for nama_asli, alias in replacement_nama_fasilitas.items():
        if nama_asli.lower() == nama_bersih.lower():
            return alias
    return nama_bersih

def handler_process(data_json):
    files_data = json.loads(data_json)
    hasil_semua = []

    for file_data in files_data:
        filename = file_data.get("filename", "")
        nik = os.path.splitext(filename)[0]
        if nik.upper().startswith("NIK_"):
            nik = nik[4:]

        fasilitas = file_data.get("fasilitas", [])
        nama_debitur = file_data.get("nama", "")

        kol_1_list, kol_25_list, wo_list, lovi_list = [], [], [], []
        jumlah_fasilitas_aktif = 0
        total_plafon, total_baki_debet, baki_debet_kol25wo = 0, 0, 0
        excluded_fasilitas = {"BTPNS", "Bank Jago", "BAV"}

        for item in fasilitas:
            nama_fasilitas = item.get("ljkKet", "")
            kondisi_ket = (item.get("kondisiKet") or '').lower()
            kualitas = item.get('kualitas', '')
            jumlah_hari_tunggakan = int(item.get('jumlahHariTunggakan', 0))
            tanggal_kondisi = item.get('tanggalKondisi', '')
            baki_debet = int(item.get('bakiDebet', 0))
            plafon_awal = int(item.get('plafonAwal', 0))

            nama_fasilitas_bersih = bersihkan_nama_fasilitas(nama_fasilitas)
            baki_debet_format = "{:,.0f}".format(baki_debet).replace(",", ".")

            if kondisi_ket == 'fasilitas aktif' and kualitas == '1' and jumlah_hari_tunggakan <= 30:
                fasilitas_teks = nama_fasilitas_bersih
            elif kondisi_ket == 'fasilitas aktif':
                fasilitas_teks = f"{nama_fasilitas_bersih} Kol {kualitas}/{jumlah_hari_tunggakan} {baki_debet_format}"
            elif kondisi_ket in ['dihapusbukukan', 'hapus tagih']:
                try:
                    tahun_wo = int(str(tanggal_kondisi)[:4])
                except:
                    tahun_wo = ""
                fasilitas_teks = f"{nama_fasilitas_bersih} WO {tahun_wo} {baki_debet_format}"
            else:
                fasilitas_teks = nama_fasilitas_bersih

            if kondisi_ket == 'fasilitas aktif':
                jumlah_fasilitas_aktif += 1
                total_plafon += plafon_awal
                total_baki_debet += baki_debet
                if kualitas == '1' and jumlah_hari_tunggakan == 0:
                    kol_1_list.append(nama_fasilitas_bersih)
                else:
                    kol_25_list.append(fasilitas_teks)
                    if nama_fasilitas_bersih not in excluded_fasilitas:
                        baki_debet_kol25wo += baki_debet
            elif kondisi_ket in ['dihapusbukukan', 'hapus tagih']:
                wo_list.append(fasilitas_teks)
                if nama_fasilitas_bersih not in excluded_fasilitas:
                    baki_debet_kol25wo += baki_debet

        if jumlah_fasilitas_aktif > 0 and not kol_25_list and not wo_list:
            rekomendasi = "OK"
        elif baki_debet_kol25wo <= 250_000:
            rekomendasi = "OK"
        else:
            rekomendasi = "NOT OK"

        hasil_semua.append({
            'NIK': "'" + nik,
            'Nama Debitur': nama_debitur,
            'Rekomendasi': rekomendasi,
            'Jumlah Fasilitas': jumlah_fasilitas_aktif,
            'Total Plafon Awal': total_plafon,
            'Total Baki Debet': total_baki_debet,
            'Kol 1': "; ".join(kol_1_list),
            'Kol 2-5': "; ".join(kol_25_list),
            'WO/dihapusbukukan': "; ".join(wo_list),
            'LOVI': ""
        })

    return json.dumps(hasil_semua)

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers['Content-Length'])
        body = self.rfile.read(content_length).decode()
        result = handler_process(body)
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(result.encode())
