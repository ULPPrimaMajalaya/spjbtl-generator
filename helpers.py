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
            'NAMA_PELANGGAN_2': ['PT MAJU SEJAHTERA ABADI II'],
            'TGL_COVER': ['15 OKTOBER 2026'],
            'DAYA_LAMA': ['233.000'],
            'DAYA_BARU': ['865.000'],
            'NO_PIHAK_1': ['0001.Pj/AGA.04.01/F02110400/2026'],
            'HARI_SURAT': ['Senin'],
            'TGL_LENGKAP': ['Dua Puluh'],
            'BLN_LENGKAP': ['Oktober'],
            'THN_LENGKAP': ['Dua Ribu Dua Puluh Enam'],
            'TGL_ANGKA': ['20-10-2026'],
            'NAMA_PERUSAHAAN_2': ['PT MAJU SEJAHTERA ABADI'],
            'BIDANG_INDUSTRI': ['Industri Manufaktur Tekstil'],
            'ALAMAT_PERUSAHAAN_2': ['Jl. Raya Majalaya - Rancaekek No. 123, Kec. Majalaya, Kab. Bandung, Jawa Barat'],
            'NO_AKTA': ['01'],
            'TGL_AKTA': ['12 Januari 2026'],
            'NOTARIS_AKTA': ['Budi Santoso, S.H., M.Kn.'],
            'NAMA_PENANGGUNG_JAWAB': ['ANDI PRATAMA'],
            'JABATAN_PIHAK_2': ['DIREKTUR'],
            'IDPEL': ['535330000001'],
            'UNITAP_UNITUP': ['53MJA/53536'],
            'NO_HP_2': ['081234567890'],
            'TARIF_KODE': ['I3 / 865.000 VA / 20 KV'],
            'TEGANGAN_KODE': ['20 KV'],
            'KOORDINAT_2': ['X= -7.0485; Y= 107.7562'],
            'EMAIL_2': ['admin.dummy@contohperusahaan.co.id'],
            'NIK_2': ['3204000000000001'],
            'NPWP_2': ['01.000.000.0-424000'],
            'NAMA_NPWP_2': ['PT MAJU SEJAHTERA ABADI'],
            'ALAMAT_NPWP_2': ['Jl. Raya Majalaya - Rancaekek No. 123, RT 001/RW 002, Majalaya, Kab. Bandung, Jawa Barat, 40382'],
            'NO_SURAT_PERMOHONAN': ['001/DUMMY/X/2026'],
            'TGL_SURAT_PERMOHONAN': ['10 Oktober 2026'],
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
            'NAMA_PELANGGAN_2': ['CV CAHAYA MANDIRI'],
            'TGL_COVER': ['12 DESEMBER 2026'],
            'DAYA_LAMA': ['-'],
            'DAYA_BARU': ['865.000'],
            'NO_PIHAK_1': ['0002.Pj/AGA.04.01/F02110400/2026'],
            'HARI_SURAT': ['Jumat'],
            'TGL_LENGKAP': ['Dua Belas'],
            'BLN_LENGKAP': ['Desember'],
            'THN_LENGKAP': ['Dua Ribu Dua Puluh Enam'],
            'TGL_ANGKA': ['12-12-2026'],
            'NAMA_PERUSAHAAN_2': ['CV CAHAYA MANDIRI'],
            'BIDANG_INDUSTRI': ['Industri Logam Dasar'],
            'NOMOR_KBLI': ['24101'],
            'ALAMAT_PERUSAHAAN_2': ['Jl. Industri Raya No. 45, Desa Ciparay, Kec. Ciparay, Kab. Bandung, Jawa Barat'],
            'NO_AKTA': ['05'],
            'TGL_AKTA': ['15 Februari 2026'],
            'NOTARIS_AKTA': ['Siti Rahmawati, S.H., M.Kn.'],
            'NAMA_PENANGGUNG_JAWAB': ['HENDRA SETIAWAN'],
            'JABATAN_PIHAK_2': ['Direktur Utama'],
            'IDPEL': ['535310000002'],
            'UNITAP_UNITUP': ['53MJA / 53531'],
            'NO_HP_2': ['081298765432'],
            'TARIF_KODE': ['I3 / 865.000 VA / 20 kV'],
            'TEGANGAN_KODE': ['20 kV'],
            'TEGANGAN_TERBILANG': ['dua puluh kilo Volt'],
            'FREKUENSI': ['50'],
            'FREKUENSI_TERBILANG': ['lima puluh Hertz'],
            'KOORDINAT_2': ['X= -7.050012; Y= 107.750123'],
            'EMAIL_2': ['cahayamandiri.dummy@email.com'],
            'NIK_2': ['3204051203850002'],
            'NPWP_2': ['02.456.789.1-429.000'],
            'NAMA_NPWP_2': ['CV CAHAYA MANDIRI'],
            'ALAMAT_NPWP_2': ['Jl. Industri Raya No. 45, Desa Ciparay, Kec. Ciparay, Kab. Bandung, Jawa Barat'],
            'NO_SURAT_PERMOHONAN': ['010/CM/XII/2026'],
            'TGL_SURAT_PERMOHONAN': ['1 Desember 2026'],
            'DAYA_BARU_TERBILANG': ['Delapan Ratus Enam Puluh Lima Ribu Volt Ampere'],
            'TARIF_BP_PER_VA': ['Rp 631,-'],
            'BIAYA_BP_ANGKA': ['Rp 545.815.000,-'],
            'BIAYA_BP_TERBILANG': ['Lima Ratus Empat Puluh Lima Juta Delapan Ratus Lima Belas Ribu Rupiah'],
            'TARIF_UJL_PER_VA': ['Rp. 225,-'],
            'DAYA_UJL': ['865.000 VA'],
            'UJL_ANGKA': ['Rp 194.625.000,-'],
            'UJL_TERBILANG': ['Seratus Sembilan Puluh Empat Juta Enam Ratus Dua Puluh Lima Ribu Rupiah'],
            'NO_SURAT_KUASA': ['SK/XII/0002'],
            'TGL_SURAT_KUASA': ['5 Desember 2026']
        }

    elif kode_transaksi == "PB_TR_TM":
        data = {
        'NAMA_PELANGGAN_2': ['YAYASAN BAKTI NUSANTARA UTAMA'],
        'TGL_COVER': ['18 OKTOBER 2026'],
        'DAYA_LAMA': ['41.5'],
        'DAYA_BARU': ['555'],
        'NO_PIHAK_1': ['0003.Pj/AGA.04.01/F02110400/2026'],
        'HARI_SURAT': ['Minggu'],
        'TGL_LENGKAP': ['Delapan Belas'],
        'BLN_LENGKAP': ['Oktober'],
        'THN_LENGKAP': ['Dua Ribu Dua Puluh Enam'],
        'TGL_ANGKA': ['18 - 10 - 2026'],
        'NAMA_PERUSAHAAN_2': ['YAYASAN BAKTI NUSANTARA UTAMA'],
        'DESKRIPSI_PROFIL_PELANGGAN': ['suatu lembaga sosial non-profit yang bergerak di bidang layanan pendidikan terpadu dan pelatihan keterampilan kerja bagi masyarakat umum'],
        'ALAMAT_PERUSAHAAN_2': ['Jl. Raya Katapang No. 88, Katapang, Kec. Katapang, Kabupaten Bandung, Jawa Barat 40921'],
        'NO_SURAT_KETERANGAN': ['042/BA-DUMMY/BNU/X/2026'],
        'NAMA_PENANGGUNG_JAWAB': ['BAGAS SAPUTRA'],
        'JABATAN_PIHAK_2': ['PROJECT MANAGER PT CIPTA KARYA – NUSANTARA, KSO'],
        'NAMA_KONSORSIUM_KONTRAKTOR': ['PT CIPTA KARYA – NUSANTARA, KSO'],
        'NO_SURAT_PERMOHONAN': ['No. 042/BA-DUMMY/BNU/X/2026'],
        'JABATAN_PENANDATANGAN_PELANGGAN': ['Ketua Yayasan Bakti Nusantara Utama'],
        'NAMA_PENANDATANGAN_PELANGGAN': ['DEWI LESTARI'],
        'NIK_PENANDATANGAN_PELANGGAN': ['3204094508920004'],
        'IDPEL': ['535322000999'],
        'UNITAP_UNITUP': ['53MJA/53536'],
        'NO_HP_2': ['082122334455'],
        'TARIF_KODE': ['I3 / 555.000 VA / 20 KV'],
        'KOORDINAT_2': ['X= -7.01234; Y= 107.54321'],
        'PERUNTUKAN_DAYA': ['Pendidikan (Sosial)'],
        'EMAIL_2': ['baktinusantara.dummy@gmail.com'],
        'NPWP_2': ['03.111.222.3-428.000'],
        'NAMA_NPWP_2': ['DEWI LESTARI'],
        'ALAMAT_NPWP_2': ['JL. KOPO PERMAI II BLOK C NO. 10, KABUPATEN BANDUNG, JAWA BARAT'],
        'TGL_SURAT_PERMOHONAN': ['10 September 2026'],
        'NAMA_KONTRAKTOR_PELAKSANA': ['PT Cipta Karya Mandiri'],
        'SURAT_KEPUTUSAN_DASAR': ['Daftar Keputusan Pengurus Yayasan Pendidikan Mitra Pembangunan'],
        'BIDANG_INDUSTRI': ['Pendidikan & Sosial'],
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