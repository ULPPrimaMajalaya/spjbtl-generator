import io
import os
import zipfile
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st

# Import modul dari helpers.py
from helpers import load_css, get_template_excel

# 1. Konfigurasi Halaman Web
st.set_page_config(
    page_title="SPJBTL Generator Prima Majalaya", page_icon="⚡", layout="wide"
)

# 2. Muat Custom CSS dari File External
load_css("style.css")

# 3. Sidebar Navigasi Sebelah Kiri
with st.sidebar:
    if os.path.exists("logo_pln.png"):
        st.image("logo_pln.png", width=120)
    else:
        st.image(
            "https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_PT_PLN.svg",
            width=120,
        )

    st.title("SPJBTL Lainnya")
    st.markdown("Pilih jenis pembuatan dokumen SPJBTL:")

    menu_terpilih = st.radio(
        "Jenis Permohonan:",
        [
            "Perubahan Daya (PD)",
            "Penyambungan Baru (PB)",
            "Turun Daya (TD)",
            "Balik Nama (BN)",
        ],
        index=0,
    )

    sub_pb = None
    if menu_terpilih == "Penyambungan Baru (PB)":
        st.markdown("---")
        sub_pb = st.radio(
            "Kategori PB:",
            [
                "PB Murni (Pelanggan Baru)",
                "PB TR ke TM (Pelanggan Prima)",
            ],
            index=0,
        )

    st.markdown("---")
    st.caption("PLN Prima Majalaya")


# Function Reusable untuk Memproses Pembuatan Dokumen Surat
def render_generator_page(
    kode_transaksi, judul_halaman, deskripsi_halaman, default_prefix_file
):
    st.markdown(
        f'<div class="main-title"> {judul_halaman}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div class="sub-title">{deskripsi_halaman}</div>',
        unsafe_allow_html=True,
    )

    # Download Template Excel Rapi Ber-border
    excel_bytes = get_template_excel(kode_transaksi)
    st.download_button(
        label=f"📥 Download Template Excel Standar ({kode_transaksi})",
        data=excel_bytes,
        file_name=f"Template_Excel_{kode_transaksi}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"dl_tpl_{kode_transaksi}"
    )

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            f"""
            <div class="css-card">
                <h4 style="color:#ffb703; margin-top:0;">Unggah Data Excel / CSV ({kode_transaksi})</h4>
                <p style="font-size:0.85rem; color:#f8fafc; font-weight: 500;">File .xlsx atau .csv berisi data pelanggan {kode_transaksi}</p>
            </div>
        """,
            unsafe_allow_html=True,
        )
        uploaded_excel = st.file_uploader(
            f"Pilih file Excel/CSV untuk {kode_transaksi}",
            type=["xlsx", "csv"],
            key=f"excel_{kode_transaksi}",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown(
            f"""
            <div class="css-card">
                <h4 style="color:#ffb703; margin-top:0;">Unggah Template Word ({kode_transaksi})</h4>
                <p style="font-size:0.85rem; color:#f8fafc; font-weight: 500;">File .docx master SPJBTL khusus {kode_transaksi}</p>
            </div>
        """,
            unsafe_allow_html=True,
        )
        uploaded_template = st.file_uploader(
            f"Pilih file Word untuk {kode_transaksi}",
            type=["docx"],
            key=f"docx_{kode_transaksi}",
            label_visibility="collapsed",
        )

    if uploaded_excel and uploaded_template:
        st.markdown("---")

        # 1. Baca file Excel/CSV (Kode Bawaan Kamu)
        if uploaded_excel.name.endswith(".csv"):
            df = pd.read_csv(uploaded_excel, dtype=str)
        else:
            df = pd.read_excel(uploaded_excel, dtype=str)

        # ----------------------------------------------------
        # 2. FITUR MAPPING KOLOM OTOMATIS (BARU DITAMBAHKAN)
        # ----------------------------------------------------
        df.columns = df.columns.str.strip()

        mapping_kolom = {
            # === INFORMASI UTAMA PELANGGAN ===
            'nm plg': 'NAMA_PELANGGAN_2',
            'nm_plg': 'NAMA_PELANGGAN_2',
            'nm pelanggan': 'NAMA_PELANGGAN_2',
            'nama plg': 'NAMA_PELANGGAN_2',
            'nama pelanggan': 'NAMA_PELANGGAN_2',
            'nama_pelanggan': 'NAMA_PELANGGAN_2',
            'nama_pelanggan_2': 'NAMA_PELANGGAN_2',
            'nama': 'NAMA_PELANGGAN_2',
            'pelanggan': 'NAMA_PELANGGAN_2',
            'customer': 'NAMA_PELANGGAN_2',
            'customer name': 'NAMA_PELANGGAN_2',
            'customer_name': 'NAMA_PELANGGAN_2',
            'nama pemohon': 'NAMA_PELANGGAN_2',
            'nama_pemohon': 'NAMA_PELANGGAN_2',
            'pemohon': 'NAMA_PELANGGAN_2',

            # IDPEL & Nomor Registrasi
            'idpel': 'IDPEL',
            'id_pel': 'IDPEL',
            'id pel': 'IDPEL',
            'no id': 'IDPEL',
            'no_id': 'IDPEL',
            'id_pelanggan': 'IDPEL',
            'id pelanggan': 'IDPEL',
            'nopel': 'IDPEL',
            'no_pelanggan': 'IDPEL',
            'no pelanggan': 'IDPEL',
            'no agenda': 'IDPEL',
            'no_agenda': 'IDPEL',
            'nomor agenda': 'IDPEL',
            'nomor_agenda': 'IDPEL',
            'no_registrasi': 'IDPEL',
            'no registrasi': 'IDPEL',

            # === DOKUMEN COVER & SURAT ===
            'tgl_agustus_cover': 'TGL_COVER',
            'tgl_spjbtl': 'TGL_COVER',
            'tgl_cover': 'TGL_COVER',
            'tanggal_cover': 'TGL_COVER',
            'tgl cover': 'TGL_COVER',
            'no_pihak_1': 'NO_PIHAK_1',
            'no pihak 1': 'NO_PIHAK_1',
            'no_surat': 'NO_PIHAK_1',
            'no surat': 'NO_PIHAK_1',
            'nomor surat': 'NO_PIHAK_1',
            'nomor_surat': 'NO_PIHAK_1',
            'no_spjbtl': 'NO_PIHAK_1',
            'no spjbtl': 'NO_PIHAK_1',
            'no_perjanjian': 'NO_PIHAK_1',
            'no perjanjian': 'NO_PIHAK_1',
            'hari_surat': 'HARI_SURAT',
            'hari': 'HARI_SURAT',
            'hari_buat': 'HARI_SURAT',
            'tgl_lengkap': 'TGL_LENGKAP',
            'tanggal_lengkap': 'TGL_LENGKAP',
            'bln_lengkap': 'BLN_LENGKAP',
            'bulan_lengkap': 'BLN_LENGKAP',
            'thn_lengkap': 'THN_LENGKAP',
            'tahun_lengkap': 'THN_LENGKAP',
            'tgl_angka': 'TGL_ANGKA',
            'tanggal_angka': 'TGL_ANGKA',

            # === DATA PERUSAHAAN & PERIZINAN ===
            'nm_perusahaan': 'NAMA_PERUSAHAAN_2',
            'nama_perusahaan': 'NAMA_PERUSAHAAN_2',
            'nama perusahaan': 'NAMA_PERUSAHAAN_2',
            'nama_perusahaan_2': 'NAMA_PERUSAHAAN_2',
            'perusahaan': 'NAMA_PERUSAHAAN_2',
            'pt/cv': 'NAMA_PERUSAHAAN_2',
            'nama pt': 'NAMA_PERUSAHAAN_2',
            'nama_pt': 'NAMA_PERUSAHAAN_2',
            'bidang_industri': 'BIDANG_INDUSTRI',
            'bidang industri': 'BIDANG_INDUSTRI',
            'jenis_usaha': 'BIDANG_INDUSTRI',
            'kbli': 'NOMOR_KBLI',
            'nomor_kbli': 'NOMOR_KBLI',
            'no_kbli': 'NOMOR_KBLI',
            'almt': 'ALAMAT_PERUSAHAAN_2',
            'almt_perusahaan': 'ALAMAT_PERUSAHAAN_2',
            'alamat_perusahaan': 'ALAMAT_PERUSAHAAN_2',
            'alamat_perusahaan_2': 'ALAMAT_PERUSAHAAN_2',
            'alamat perusahaan': 'ALAMAT_PERUSAHAAN_2',
            'alamat_kantor': 'ALAMAT_PERUSAHAAN_2',
            'alamat kantor': 'ALAMAT_PERUSAHAAN_2',
            'alamat_lokasi': 'ALAMAT_PERUSAHAAN_2',
            'alamat lokasi': 'ALAMAT_PERUSAHAAN_2',
            
            # Akta & Legalitas
            'no_akta': 'NO_AKTA',
            'no akta': 'NO_AKTA',
            'nomor akta': 'NO_AKTA',
            'nomor_akta': 'NO_AKTA',
            'tgl_akta': 'TGL_AKTA',
            'tanggal_akta': 'TGL_AKTA',
            'tgl akta': 'TGL_AKTA',
            'notaris_akta': 'NOTARIS_AKTA',
            'notaris': 'NOTARIS_AKTA',
            'nama_notaris': 'NOTARIS_AKTA',

            # Penanggung Jawab & Direksi
            'nm_penanggung_jawab': 'NAMA_PENANGGUNG_JAWAB',
            'nama_penanggung_jawab': 'NAMA_PENANGGUNG_JAWAB',
            'nama penanggung jawab': 'NAMA_PENANGGUNG_JAWAB',
            'penanggung_jawab': 'NAMA_PENANGGUNG_JAWAB',
            'pic': 'NAMA_PENANGGUNG_JAWAB',
            'nama_pic': 'NAMA_PENANGGUNG_JAWAB',
            'nama pic': 'NAMA_PENANGGUNG_JAWAB',
            'nama_direktur': 'NAMA_PENANGGUNG_JAWAB',
            'nama direktur': 'NAMA_PENANGGUNG_JAWAB',
            'direktur': 'NAMA_PENANGGUNG_JAWAB',
            'nama_direktur_2': 'NAMA_DIREKTUR_2',
            'NAMA_DIREKTUR_2': 'NAMA_PENANGGUNG_JAWAB',
            'jabatan_pihak_2': 'JABATAN_PIHAK_2',
            'jabatan': 'JABATAN_PIHAK_2',
            'jabatan_pic': 'JABATAN_PIHAK_2',
            'jabatan pic': 'JABATAN_PIHAK_2',

            # === KONTAK, LOKASI & IDENTITAS ===
            'unitap_unitup': 'UNITAP_UNITUP',
            'unitup': 'UNITAP_UNITUP',
            'unit_up': 'UNITAP_UNITUP',
            'unit_layanan': 'UNITAP_UNITUP',
            'ulp': 'UNITAP_UNITUP',
            'up3': 'UNITAP_UNITUP',
            'no_hp': 'NO_HP_2',
            'no_hp_2': 'NO_HP_2',
            'no hp': 'NO_HP_2',
            'nohp': 'NO_HP_2',
            'no_telp': 'NO_HP_2',
            'no telp': 'NO_HP_2',
            'telepon': 'NO_HP_2',
            'wa': 'NO_HP_2',
            'no_wa': 'NO_HP_2',
            'koordinat': 'KOORDINAT_2',
            'koordinat_2': 'KOORDINAT_2',
            'titik_koordinat': 'KOORDINAT_2',
            'email': 'EMAIL_2',
            'email_2': 'EMAIL_2',
            'alamat_email': 'EMAIL_2',
            'nik': 'NIK_2',
            'nik_2': 'NIK_2',
            'nik_pic': 'NIK_2',
            'no_ktp': 'NIK_2',
            'ktp': 'NIK_2',
            'npwp': 'NPWP_2',
            'npwp_2': 'NPWP_2',
            'no_npwp': 'NPWP_2',
            'nama_npwp': 'NAMA_NPWP_2',
            'nama_npwp_2': 'NAMA_NPWP_2',
            'nama npwp': 'NAMA_NPWP_2',
            'alamat_npwp': 'ALAMAT_NPWP_2',
            'alamat_npwp_2': 'ALAMAT_NPWP_2',
            'alamat npwp': 'ALAMAT_NPWP_2',

            # === TARIF, DAYA & KELISTRIKAN ===
            'daya_lama': 'DAYA_LAMA',
            'daya lama': 'DAYA_LAMA',
            'daya_awal': 'DAYA_LAMA',
            'daya_exist': 'DAYA_LAMA',
            'daya_eksisting': 'DAYA_LAMA',
            'daya_baru': 'DAYA_BARU',
            'daya baru': 'DAYA_BARU',
            'daya_mohon': 'DAYA_BARU',
            'daya_permohonan': 'DAYA_BARU',
            'daya': 'DAYA_BARU',
            'tarif_kode': 'TARIF_KODE',
            'tarif': 'TARIF_KODE',
            'gol_tarif': 'TARIF_KODE',
            'golongan_tarif': 'TARIF_KODE',
            'tarif/daya': 'TARIF_KODE',
            'tarif_daya': 'TARIF_KODE',
            'tegangan_kode': 'TEGANGAN_KODE',
            'tegangan': 'TEGANGAN_KODE',
            'tegangan_terbilang': 'TEGANGAN_TERBILANG',
            'frekuensi': 'FREKUENSI',
            'frekuensi_terbilang': 'FREKUENSI_TERBILANG',
            'daya_baru_terbilang': 'DAYA_BARU_TERBILANG',
            'daya_terbilang': 'DAYA_BARU_TERBILANG',

            # === BIAYA & PERMOHONAN ===
            'no_surat_permohonan': 'NO_SURAT_PERMOHONAN',
            'no surat permohonan': 'NO_SURAT_PERMOHONAN',
            'no_permohonan': 'NO_SURAT_PERMOHONAN',
            'surat_mohon_no': 'NO_SURAT_PERMOHONAN',
            'tgl_surat_permohonan': 'TGL_SURAT_PERMOHONAN',
            'tgl surat permohonan': 'TGL_SURAT_PERMOHONAN',
            'tanggal_permohonan': 'TGL_SURAT_PERMOHONAN',
            'tarif_bp_per_va': 'TARIF_BP_PER_VA',
            'biaya_bp_angka': 'BIAYA_BP_ANGKA',
            'biaya_bp': 'BIAYA_BP_ANGKA',
            'bp_angka': 'BIAYA_BP_ANGKA',
            'bp': 'BIAYA_BP_ANGKA',
            'biaya_penyambungan': 'BIAYA_BP_ANGKA',
            'biaya_bp_terbilang': 'BIAYA_BP_TERBILANG',
            'bp_terbilang': 'BIAYA_BP_TERBILANG',
            'tarif_ujl_per_va': 'TARIF_UJL_PER_VA',
            'daya_ujl': 'DAYA_UJL',
            'ujl_angka': 'UJL_ANGKA',
            'ujl': 'UJL_ANGKA',
            'biaya_ujl': 'UJL_ANGKA',
            'uang_jaminan': 'UJL_ANGKA',
            'ujl_terbilang': 'UJL_TERBILANG',
            'no_surat_kuasa': 'NO_SURAT_KUASA',
            'no surat kuasa': 'NO_SURAT_KUASA',
            'tgl_surat_kuasa': 'TGL_SURAT_KUASA',
            'tgl surat kuasa': 'TGL_SURAT_KUASA',
        }

        new_columns = []
        for col in df.columns:
            col_clean = str(col).lower().strip()
            if col_clean in mapping_kolom:
                new_columns.append(mapping_kolom[col_clean])
            else:
                new_columns.append(col)
        
        df.columns = new_columns
        # ----------------------------------------------------

        # 3. Pembersihan Otomatis (Kode Bawaan Kamu)
        df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
        df = df.fillna("")

        if "IDPEL" in df.columns and "NAMA_PELANGGAN_2" in df.columns:
            df["Pilihan_Pelanggan"] = (
                df["IDPEL"].astype(str) + " — " + df["NAMA_PELANGGAN_2"].astype(str)
            )
        elif "IDPEL" in df.columns and "NAMA_PELANGGAN_1" in df.columns:
            df["Pilihan_Pelanggan"] = (
                df["IDPEL"].astype(str) + " — " + df["NAMA_PELANGGAN_1"].astype(str)
            )
        elif "IDPEL" in df.columns and "Nama_Pelanggan" in df.columns:
            df["Pilihan_Pelanggan"] = (
                df["IDPEL"].astype(str) + " — " + df["Nama_Pelanggan"].astype(str)
            )
        else:
            df["Pilihan_Pelanggan"] = [
                f"Baris {i+1} — {row.iloc[0]}" for i, row in df.iterrows()
            ]

        st.subheader("Pilih Pelanggan")
        opsi_pelanggan = df["Pilihan_Pelanggan"].tolist()
        pilih_semua = st.checkbox(
            "Pilih Semua Pelanggan", key=f"chk_{kode_transaksi}"
        )

        if pilih_semua:
            pelanggan_terpilih = st.multiselect(
                "Daftar Pelanggan Terpilih:",
                options=opsi_pelanggan,
                default=opsi_pelanggan,
                key=f"multi_{kode_transaksi}",
            )
        else:
            pelanggan_terpilih = st.multiselect(
                "Cari dan pilih satu atau beberapa pelanggan:",
                options=opsi_pelanggan,
                key=f"multi_{kode_transaksi}",
            )

        df_filtered = df[df["Pilihan_Pelanggan"].isin(pelanggan_terpilih)]
        st.info(
            f"Total pelanggan terpilih: **{len(df_filtered)} pelanggan**"
        )

        if len(df_filtered) > 0:
            with st.expander("👁️Pratinjau Data Terpilih"):
                st.dataframe(
                    df_filtered.drop(columns=["Pilihan_Pelanggan"], errors="ignore"),
                    use_container_width=True,
                    hide_index=True,
                )

            if st.button(
                f"Generate Surat SPJBTL {kode_transaksi} Terpilih",
                use_container_width=True,
                key=f"btn_{kode_transaksi}",
            ):
                zip_buffer = io.BytesIO()

                with st.spinner("Memproses dokumen... Mohon tunggu sebentar."):
                    with zipfile.ZipFile(
                        zip_buffer, "a", zipfile.ZIP_DEFLATED, False
                    ) as zip_file:
                        for index, row in df_filtered.iterrows():
                            doc = DocxTemplate(uploaded_template)
                            context = row.to_dict()
                            doc.render(context)

                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)

                            idpel = str(row.get("IDPEL", f"Pelanggan_{index+1}"))
                            nama_pelanggan = str(
                                row.get(
                                    "NAMA_PELANGGAN_2",
                                    row.get(
                                        "NAMA_PELANGGAN_1",
                                        row.get(
                                            "Nama_Pelanggan",
                                            f"Dokumen_{index+1}",
                                        ),
                                    ),
                                )
                            ).replace("/", "_")

                            filename = f"{default_prefix_file}_{idpel}_{nama_pelanggan}.docx"
                            zip_file.writestr(filename, doc_io.getvalue())

                st.success("🎉 Seluruh dokumen terpilih berhasil dibuat!")

                st.download_button(
                    label="Unduh Hasil Dokumen (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Hasil_SPJBTL_{kode_transaksi}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key=f"dl_{kode_transaksi}",
                )
        else:
            st.warning("⚠️ Silakan pilih minimal 1 pelanggan untuk melanjutkan.")


# 4. Routing Halaman Berdasarkan Pilihan Sidebar
if menu_terpilih == "Perubahan Daya (PD)":
    render_generator_page(
        kode_transaksi="PD",
        judul_halaman="SPJBTL — Perubahan Daya (PD)",
        deskripsi_halaman="Sistem Automasi Pembuatan Surat Perjanjian Jual Beli Tenaga Listrik untuk Perubahan Daya",
        default_prefix_file="SPJBTL_PD",
    )

elif menu_terpilih == "Penyambungan Baru (PB)":
    if sub_pb == "PB Murni (Pelanggan Baru)":
        render_generator_page(
            kode_transaksi="PB_Murni",
            judul_halaman="SPJBTL — Penyambungan Baru Murni (PB)",
            deskripsi_halaman="Sistem Automasi Pembuatan SPJBTL untuk Pelanggan Baru Murni PLN",
            default_prefix_file="SPJBTL_PB_Murni",
        )
    else:
        render_generator_page(
            kode_transaksi="PB_TR_TM",
            judul_halaman="SPJBTL — Penyambungan Baru TR ke TM (Pelanggan Prima)",
            deskripsi_halaman="Sistem Automasi Pembuatan SPJBTL untuk Penambahan Layanan Prima (TR ke TM)",
            default_prefix_file="SPJBTL_PB_Prima",
        )

elif menu_terpilih == "Turun Daya (TD)":
    render_generator_page(
        kode_transaksi="TD",
        judul_halaman="SPJBTL — Turun Daya (TD)",
        deskripsi_halaman="Sistem Automasi Pembuatan Surat Perjanjian Jual Beli Tenaga Listrik untuk Turun Daya",
        default_prefix_file="SPJBTL_TD",
    )

elif menu_terpilih == "Balik Nama (BN)":
    render_generator_page(
        kode_transaksi="BN",
        judul_halaman="SPJBTL — Balik Nama (BN)",
        deskripsi_halaman="Sistem Automasi Pembuatan Surat Perjanjian Jual Beli Tenaga Listrik untuk Balik Nama",
        default_prefix_file="SPJBTL_BN",
    )