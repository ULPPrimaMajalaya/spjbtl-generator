import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def load_css(file_name="style.css"):
    """Fungsi untuk membaca dan menerapkan file CSS external"""
    try:
        with open(file_name, "r") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"File {file_name} tidak ditemukan. Menggunakan tampilan standar.")

def get_template_excel(kode_transaksi):
    """Fungsi untuk membuat file Excel template ber-border dan rapi"""
    if kode_transaksi == "PD":
        data = {
            'NAMA_PELANGGAN_2': ['PT PANFILA INDOSARI II'],
            'TGL_COVER': ['26 AGUSTUS 2026'],
            'DAYA_LAMA': ['233.000'],
            'DAYA_BARU': ['865.000'],
            'NO_PIHAK_1': ['0012.Pj/AGA.04.01/F02110400/2026'],
            'HARI_SURAT': ['Kamis'],
            'TGL_LENGKAP': ['Sepuluh'],
            'BLN_LENGKAP': ['September'],
            'THN_LENGKAP': ['Dua Ribu Dua Puluh Enam'],
            'TGL_ANGKA': ['10-09-2026'],
            'NAMA_PERUSAHAAN_2': ['PT PANFILA INDOSARI'],
            'BIDANG_INDUSTRI': ['Industri Air Kemasan'],
            'ALAMAT_PERUSAHAAN_2': ['Jl. Raya Bandung - Garut Nyalindung, Nagrog, Kec. Cicalengka, Kab. Bandung, Jawa Barat'],
            'NO_AKTA': ['06'],
            'TGL_AKTA': ['26 Maret 2026'],
            'NOTARIS_AKTA': ['Karla Priscilla Suryawinata, S.H., M.Kn.'],
            'NAMA_PENANGGUNG_JAWAB': ['FERRY LABORA'],
            'JABATAN_PIHAK_2': ['DIREKTUR'],
            'IDPEL': ['535330290652'],
            'UNITAP_UNITUP': ['53MJA/53536'],
            'NO_HP_2': ['0811929110'],
            'TARIF_KODE': ['I3 / 865.000 VA / 20 KV'],
            'TEGANGAN_KODE': ['20 KV'],
            'KOORDINAT_2': ['X= -6.9249; Y= 107.6425'],
            'EMAIL_2': ['bulndewi@gmail.com;b.ganjar@yahoo.co.id'],
            'NIK_2': ['3204081811820002'],
            'NPWP_2': ['01.910.164.1-424000'],
            'NAMA_NPWP_2': ['PT PANFILA INDOSARI'],
            'ALAMAT_NPWP_2': ['Jl. Jenderal Ahmad Yani, 514 D, RT 000/RW 000, Babakan Surabaya, Kiaracondong, Kota Bandung, Jawa Barat, 40281'],
            'NO_SURAT_PERMOHONAN': ['003/PIC/VI/2026'],
            'TGL_SURAT_PERMOHONAN': ['8 Juni 2026'],
            'DAYA_BARU_TERBILANG': ['Delapan Ratus Enam Puluh Lima Ribu Volt Ampere'],
            'TARIF_BP_PER_VA': ['Rp 631,-'],
            'BIAYA_BP_ANGKA': ['Rp 398.792.000,-'],
            'BIAYA_BP_TERBILANG': ['Tiga Ratus Sembilan Puluh Delapan Juta Tujuh Ratus Sembilan Puluh Dua Ribu Rupiah'],
            'DAYA_UJL': ['865.000 VA'],
            'UJL_ANGKA': ['Rp. 147.559.000,-'],
            'UJL_TERBILANG': ['Seratus Empat Puluh Tujuh Juta Lima Ratus Lima Puluh Sembilan Ribu Rupiah']
        }

    elif kode_transaksi == "PB_Murni":
        data = {
            'NAMA_PELANGGAN_2': ['GLENN SUTISNA'],
            'TGL_COVER': ['12 NOVEMBER 2025'],
            'DAYA_LAMA': ['-'],
            'DAYA_BARU': ['865.000'],
            'NO_PIHAK_1': ['0013.Pj/AGA.04.01/F02110400/2025'],
            'HARI_SURAT': ['Rabu'],
            'TGL_LENGKAP': ['Dua Belas'],
            'BLN_LENGKAP': ['November'],
            'THN_LENGKAP': ['Dua Ribu Dua Puluh Lima'],
            'TGL_ANGKA': ['12-11-2025'],
            'NAMA_PERUSAHAAN_2': ['GLENN SUTISNA'],
            'BIDANG_INDUSTRI': ['Industri Plastik'],
            'NOMOR_KBLI': ['22299'],
            'ALAMAT_PERUSAHAAN_2': ['Jl. Raya Sapan Km 01 No. 8, Desa Tegalluar, Kec. Gedebage, Kab. Bandung, Jawa Barat'],
            'NO_AKTA': ['-'],
            'TGL_AKTA': ['-'],
            'NOTARIS_AKTA': ['-'],
            'NAMA_PENANGGUNG_JAWAB': ['Glenn Nathan Sutisna'],
            'JABATAN_PIHAK_2': ['Direktur'],
            'IDPEL': ['535312613214'],
            'UNITAP_UNITUP': ['53MJA / 53531'],
            'NO_HP_2': ['08121473312 / 087888877226'],
            'TARIF_KODE': ['I3 / 865.000 VA / 20 kV'],
            'TEGANGAN_KODE': ['20 kV'],
            'TEGANGAN_TERBILANG': ['dua puluh kilo Volt'],
            'FREKUENSI': ['50'],
            'FREKUENSI_TERBILANG': ['lima puluh Hertz'],
            'KOORDINAT_2': ['X= -6.976356; Y= 107.68725'],
            'EMAIL_2': ['glennathansutisna@gmail.com'],
            'NIK_2': ['3273151205900001'],
            'NPWP_2': ['01.325.681.3-429.000'],
            'NAMA_NPWP_2': ['GLENN SUTISNA'],
            'ALAMAT_NPWP_2': ['Jl. Raya Sapan Km 01 No. 8, Desa Tegalluar, Kec. Gedebage, Kab. Bandung, Jawa Barat'],
            'NO_SURAT_PERMOHONAN': ['-'],
            'TGL_SURAT_PERMOHONAN': ['10 Oktober 2025'],
            'DAYA_BARU_TERBILANG': ['Delapan Ratus Enam Puluh Lima Ribu Volt Ampere'],
            'TARIF_BP_PER_VA': ['Rp 631,-'],
            'BIAYA_BP_ANGKA': ['Rp 545.815.000,-'],
            'BIAYA_BP_TERBILANG': ['Lima Ratus Empat Puluh Lima Juta Delapan Ratus Lima Belas Ribu Rupiah'],
            'TARIF_UJL_PER_VA': ['Rp. 225,-'],
            'DAYA_UJL': ['865.000 VA'],
            'UJL_ANGKA': ['Rp 194.625.000,-'],
            'UJL_TERBILANG': ['Seratus Sembilan Puluh Empat Juta Enam Ratus Dua Puluh Lima Ribu Rupiah'],
            'NO_SURAT_KUASA': ['MPF/IX/0012'],
            'TGL_SURAT_KUASA': ['24 November 2025']
        }

    elif kode_transaksi == "PB_TR_TM":
        data = {
        'NAMA_PELANGGAN_2': ['SEKOLAH RAKYAT SOREANG'],
        'TGL_COVER': ['26 AGUSTUS 2026'],
        'DAYA_LAMA': ['41.5'],
        'DAYA_BARU': ['555'],
        'NO_PIHAK_1': ['0002.Pj/AGA.04.01/F02110400/2026'], # Dipakai untuk cover & isi
        'HARI_SURAT': ['Rabu'],
        'TGL_LENGKAP': ['Dua Puluh Enam'],
        'BLN_LENGKAP': ['Agustus'],
        'THN_LENGKAP': ['Dua Ribu Dua Puluh Enam'],
        'TGL_ANGKA': ['26 - 08 - 2026'],
        'NAMA_PERUSAHAAN_2': ['SEKOLAH RAKYAT SOREANG'],
        'DESKRIPSI_PROFIL_PELANGGAN': ['suatu fasilitas pendidikan berasrama yang digagas pemerintah di bawah koordinasi Kementerian Sosial sebagai program afirmatif inklusif untuk menjangkau anak-anak dari keluarga tidak mampu'],
        'ALAMAT_PERUSAHAAN_2': ['Jl. Raya Soreang - Banjaran, Soreang, Kec. Soreang, Kabupaten Bandung, Jawa Barat 40911'],
        'NO_SURAT_KETERANGAN': ['095/BA-PRIMA/SRJB2/QC/VIII/2026'],
        'NAMA_PENANGGUNG_JAWAB': ['WASKITO ADY'],
        'JABATAN_PIHAK_2': ['PROJECT MANAGER PT ABIPRAYA – PRIMA, KSO'],
        'NAMA_KONSORSIUM_KONTRAKTOR': ['PT ABIPRAYA – PRIMA, KSO'],
        'NO_SURAT_PERMOHONAN': ['No. 095/BA-PRIMA/SRJB2/QC/VIII/2026'],
        'JABATAN_PENANDATANGAN_PELANGGAN': ['Kepala Sekolah Rakyat Soreang'],
        'NAMA_PENANDATANGAN_PELANGGAN': ['TITIN SRI SUPRIHATIN'],
        'NIK_PENANDATANGAN_PELANGGAN': ['3273055412800003'],
        'IDPEL': ['535322488735'],
        'UNITAP_UNITUP': ['53MJA/53536'],
        'NO_HP_2': ['083821499772'],
        'TARIF_KODE': ['I3 / 555.000 VA / 20 KV'],
        'KOORDINAT_2': ['X= -7.03250; Y= 107.53550'],
        'PERUNTUKAN_DAYA': ['Pendidikan (Sosial)'],
        'EMAIL_2': ['srt4kab.bandung@gmail.com'],
        'NPWP_2': ['24.603.986.1-428.000'],
        'NAMA_NPWP_2': ['TINTIN SRI SUPRIHATIN'],
        'ALAMAT_NPWP_2': ['JL. ANDIR NO. 129/78 CIROYOM ANDIR KOTA BANDUNG, JAWA BARAT'],
        'TGL_SURAT_PERMOHONAN': ['09 April 2026'],
        'NAMA_KONTRAKTOR_PELAKSANA': ['PT Gapura Abadi'],
        'SURAT_KEPUTUSAN_DASAR': ['Daftar Nama Kepala Sekolah pada 182 (Seratus Delapan Puluh Dua) Titik Lokasi Sekolah Rakyat'],
        'BIDANG_INDUSTRI': ['Pendidikan'],
        'KEMENTERIAN_PEMBINA': ['Kementerian Sosial'],
        'DAYA_BARU_TERBILANG': ['Lima Ratus Lima Puluh Lima Ribu Volt Ampere'],
        'TARIF_BP_PER_VA': ['Rp 631,-'],
        'BIAYA_BP_ANGKA': ['Rp 631.491.914,-'],
        'BIAYA_BP_TERBILANG': ['Enam Ratus Tiga Puluh Satu Juta Empat Ratus Sembilan Puluh Satu Ribu Sembilan Ratus Empat Belas Rupiah'],
        'TARIF_UJL_PER_VA': ['148,-'],
        'DAYA_BARU_VA': ['555.000'],
        'BIAYA_UJL_ANGKA': ['Rp. 75.139.131,-'],
        'BIAYA_UJL_TERBILANG': ['Tujuh Puluh Lima Juta Seratus Tiga Puluh Sembilan Ribu Seratus Tiga Puluh Satu Rupiah']
    }

    else:
        # Fallback default (untuk TD, BN, dll)
        data = {
            'NO_SPJBTL': ['001/SPJBTL/2026'],
            'TANGGAL_SPJBTL': ['05 Oktober 2026'],
            'IDPEL': ['531012345678'],
            'NAMA_PELANGGAN_1': ['Nama Pelanggan Contoh'],
            'ALAMAT': ['Alamat Lengkap Pelanggan']
        }

    df_template = pd.DataFrame(data)
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Data_Pelanggan')
        worksheet = writer.sheets['Data_Pelanggan']

        # Styling OpenPyXL
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        header_fill = PatternFill(start_color='1B263B', end_color='1B263B', fill_type='solid')
        header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        data_font = Font(name='Segoe UI', size=10)
        align_left = Alignment(horizontal='left', vertical='center')

        # Terapkan Border dan Styling
        for row in worksheet.iter_rows(min_row=1, max_row=len(df_template) + 1, min_col=1, max_col=len(df_template.columns)):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.font = data_font
                    cell.alignment = align_left

        # AutoFit Lebar Kolom
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if len(val) > max_len:
                    max_len = len(val)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    return buffer.getvalue()