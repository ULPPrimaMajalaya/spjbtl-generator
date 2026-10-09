import csv

# Daftar header lengkap disesuaikan dengan placeholder template Word
headers = [
    "NAMA_PELANGGAN_2",
    "NO_PIHAK_1_COVER",
    "TGL_AGUSTUS_COVER",
    "DAYA_LAMA",
    "DAYA_BARU",
    "NO_PIHAK_1",
    "HARI_SURAT",
    "TGL_LENGKAP",
    "BLN_LENGKAP",
    "THN_LENGKAP",
    "TGL_ANGKA",
    "NAMA_PERUSAHAAN_2",
    "BIDANG_INDUSTRI",
    "NOMOR_KBLI",
    "ALAMAT_PERUSAHAAN_2",
    "NO_AKTA",
    "TGL_AKTA",
    "NOTARIS_AKTA",
    "NAMA_PENANGGUNG_JAWAB",
    "JABATAN_PIHAK_2",
    "IDPEL",
    "UNITAP_UNITUP",
    "NO_HP_2",
    "TARIF_KODE",
    "TEGANGAN_KODE",
    "TEGANGAN_TERBILANG",
    "FREKUENSI",
    "FREKUENSI_TERBILANG",
    "KOORDINAT_2",
    "EMAIL_2",
    "NIK_2",
    "NPWP_2",
    "NAMA_NPWP_2",
    "ALAMAT_NPWP_2",
    "NO_SURAT_PERMOHONAN",
    "TGL_SURAT_PERMOHONAN",
    "DAYA_BARU_TERBILANG",
    "TARIF_BP_PER_VA",
    "BIAYA_BP_ANGKA",
    "BIAYA_BP_TERBILANG",
    "TARIF_UJL_PER_VA",
    "DAYA_UJL",
    "UJL_ANGKA",
    "UJL_TERBILANG",
    "NO_SURAT_KUASA",
    "TGL_SURAT_KUASA",
]

# Data Ekstraksi dari Dokumen Pasang Baru (PB) - GLENN SUTISNA
row_pb = [
    "GLENN SUTISNA",  # NAMA_PELANGGAN_2
    "0013.Pj/AGA.04.01/F02110400/2025",  # NO_PIHAK_1_COVER
    "12 NOVEMBER 2025",  # TGL_AGUSTUS_COVER
    "-",  # DAYA_LAMA (Strip untuk Pasang Baru)
    "865.000",  # DAYA_BARU
    "0013.Pj/AGA.04.01/F02110400/2025",  # NO_PIHAK_1
    "Rabu",  # HARI_SURAT
    "Dua Belas",  # TGL_LENGKAP
    "November",  # BLN_LENGKAP
    "Dua Ribu Dua Puluh Lima",  # THN_LENGKAP
    "12-11-2025",  # TGL_ANGKA
    "GLENN SUTISNA",  # NAMA_PERUSAHAAN_2
    "Industri Plastik",  # BIDANG_INDUSTRI
    "22299",  # NOMOR_KBLI
    (
        "Jl. Raya Sapan Km 01 No. 8, Desa Tegalluar, Kec. Gedebage, Kab."
        " Bandung, Jawa Barat"
    ),  # ALAMAT_PERUSAHAAN_2
    "-",  # NO_AKTA
    "-",  # TGL_AKTA
    "-",  # NOTARIS_AKTA
    "Glenn Nathan Sutisna",  # NAMA_PENANGGUNG_JAWAB
    "Direktur",  # JABATAN_PIHAK_2
    "535312613214",  # IDPEL
    "53MJA / 53531",  # UNITAP_UNITUP
    "08121473312 / 087888877226",  # NO_HP_2
    "I3 / 865.000 VA / 20 kV",  # TARIF_KODE
    "20 kV",  # TEGANGAN_KODE
    "dua puluh kilo Volt",  # TEGANGAN_TERBILANG
    "50",  # FREKUENSI
    "lima puluh Hertz",  # FREKUENSI_TERBILANG
    "X= -6.976356; Y= 107.68725",  # KOORDINAT_2
    "glennathansutisna@gmail.com",  # EMAIL_2
    "3273151205900001",  # NIK_2
    "01.325.681.3-429.000",  # NPWP_2
    "GLENN SUTISNA",  # NAMA_NPWP_2
    (
        "Jl. Raya Sapan Km 01 No. 8, Desa Tegalluar, Kec. Gedebage, Kab."
        " Bandung, Jawa Barat"
    ),  # ALAMAT_NPWP_2
    "-",  # NO_SURAT_PERMOHONAN
    "10 Oktober 2025",  # TGL_SURAT_PERMOHONAN
    "Delapan Ratus Enam Puluh Lima Ribu Volt Ampere",  # DAYA_BARU_TERBILANG
    "Rp 631,-",  # TARIF_BP_PER_VA
    "Rp 545.815.000,-",  # BIAYA_BP_ANGKA
    (
        "Lima Ratus Empat Puluh Lima Juta Delapan Ratus Lima Belas Ribu Rupiah"
    ),  # BIAYA_BP_TERBILANG
    "Rp. 225,-",  # TARIF_UJL_PER_VA
    "865.000 VA",  # DAYA_UJL
    "Rp 194.625.000,-",  # UJL_ANGKA
    (
        "Seratus Sembilan Puluh Empat Juta Enam Ratus Dua Puluh Lima Ribu"
        " Rupiah"
    ),  # UJL_TERBILANG
    "MPF/IX/0012",  # NO_SURAT_KUASA
    "24 November 2025",  # TGL_SURAT_KUASA
]

# Simpan ke file CSV
with open(
    "Database_Mail_Merge_SPJBTL_PB.csv", mode="w", newline="", encoding="utf-8"
) as file:
    writer = csv.writer(file)
    writer.writerow(headers)
    writer.writerow(row_pb)

print("File Database_Mail_Merge_SPJBTL_PB.csv berhasil dibuat!")