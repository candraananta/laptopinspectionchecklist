# Import library yang diperlukan
import streamlit as st  # Untuk membuat antarmuka web
import pandas as pd    # Untuk manipulasi dan analisis data
from datetime import datetime  # Untuk menangani format tanggal dan waktu
from pymongo import MongoClient  # Untuk koneksi ke database MongoDB
from pymongo.server_api import ServerApi  # Untuk menentukan versi API MongoDB
import urllib.parse  # Untuk encoding karakter khusus pada username dan password
from pandas import ExcelWriter  # Untuk membuat file Excel
import io  # Untuk menangani input/output bytes stream
from reportlab.lib import colors  # Untuk mengatur warna pada PDF
from reportlab.lib.pagesizes import A4, portrait  # Untuk mengatur ukuran dan orientasi halaman PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer  # Komponen untuk membuat PDF
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle  # Untuk mengatur gaya teks pada PDF
from reportlab.lib.units import inch  # Untuk mengatur satuan ukuran pada PDF


# Custom CSS untuk mengatur ukuran font
st.markdown("""
    <style>
    .main-title {
        font-size: 42px !important;
        font-weight: bold;
        text-align: left;
        margin-bottom: 20px;
    }
    .sub-title {
        font-size: 28px !important;
        text-align: left;
        color: #4A4A4A;
        margin-bottom: 30px;
    }
    .description {
        font-size: 20px !important;
        text-align: center;
        font-style: italic;
        color: #666666;
    }
    .logo-img {
        float: left;
        width: 110x;
        height: 110px;
        margin-right: 20px;
        margin-top:20px
    }
    </style>
""", unsafe_allow_html=True)

def get_mongo_client():
    """Fungsi untuk membuat koneksi ke MongoDB Atlas"""
    # Melakukan encode pada username dan password untuk menghindari karakter khusus
    username = urllib.parse.quote_plus(st.secrets["mongo"]["username"])
    password = urllib.parse.quote_plus(st.secrets["mongo"]["password"])
    
    # Membuat string koneksi MongoDB dengan kredensial yang sudah di-encode
    connection_string = f"mongodb+srv://{username}:{password}@{st.secrets['mongo']['cluster']}.am62x.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
    
    # Membuat instance MongoDB client dengan string koneksi
    client = MongoClient(connection_string, server_api=ServerApi('1'))
    
    return client

def save_to_mongodb(data):
    """Fungsi untuk menyimpan data checklist ke MongoDB"""
    try:
        # Mendapatkan koneksi MongoDB
        client = get_mongo_client()
        # Memilih database dan collection yang akan digunakan
        db = client[st.secrets["mongo"]["database_name"]]
        collection = db["hardware_checks"]
        
        # Melakukan tes koneksi ke database
        client.admin.command('ping')  # Akan raise exception jika gagal
        st.success("Terhubung ke MongoDB!")
        
        # Menyimpan data dan mengembalikan ID dokumen yang tersimpan
        result = collection.insert_one(data)
        return {
            'success': result.acknowledged,  # True jika operasi berhasil
            'document_id': str(result.inserted_id)  # Convert ObjectId ke string
        }
    
    
    except Exception as e:
        # Menangani error yang mungkin terjadi saat operasi database
        st.error(f"⛔ Gagal menyimpan ke database. Pastikan:")
        st.error("1. Internet aktif")
        st.error("2. IP Anda diizinkan di MongoDB Atlas")
        st.error(f"Detail error: {e}")
        return None
    finally:
        # Menutup koneksi database jika ada
        if 'client' in locals():
            client.close()

def main():
    """Fungsi utama untuk membuat antarmuka Streamlit"""
    # Mengatur judul halaman
    # Alternatif menggunakan HTML untuk layout yang lebih presisi:
    st.markdown("""
        <div class="container">
            <img class="logo-img" src="https://upload.wikimedia.org/wikipedia/commons/a/a7/Logo_AKR_Corporindo.png">
            <div>
                <p class="main-title">System Health Inspection Checklist</p>
                <p class="sub-title">PC & Laptop</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Membuat bagian header dengan input fields
    with st.expander("Checklist Information", expanded=True):
    # Membuat 3 kolom untuk input informasi
        col1, col2, col3 = st.columns(3)
    with col1:
        # Kosongkan field jika form berhasil disubmit
        location = st.text_input("Location")
    with col2:
        user_name = st.text_input("Nama User")
    with col3:
        checklist_date = st.date_input("Tanggal Checklist", datetime.now())
    
    # Mendefinisikan struktur data checklist
    data = {
        "No": list(range(1, 31)),  # Membuat list nomor 1-30
        "Category": ["Hardware"]*10 + ["Software"]*10 + ["Laptop"]*10,  # Kategori untuk setiap item
        "Test Item": [
            # Item Hardware (1-10)
            "Power Supply (PSU)", 
            "CPU Temperature",
            "Fan & Heatsink",
            "RAM",
            "HDD/SSD",
            "VGA/GPU",
            "Port USB/LAN/Audio",
            "Monitor",
            "Keyboard & Mouse",
            "Baterai CMOS",
            
            # Item Software (11-20)
            "OS & Update",
            "Antivirus",
            "Startup Program",
            "Disk Usage",
            "CPU & RAM Usage",
            "System Log Error",
            "Aplikasi Utama",
            "Koneksi Internet",
            "Backup & Restore",
            "Lisensi Software",
            
            # Item Laptop (21-30)
            "Baterai Laptop",
            "Adapter/Charger",
            "Layar Internal",
            "Touchpad",
            "Keyboard Internal",
            "Webcam",
            "Mikrofon",
            "Speaker Internal",
            "Engsel & Body",
            "Port HDMI/USB-C/SD"
        ],
        # Mendefinisikan standar pengujian untuk setiap item
        "Testing Standard/Normal": [
            # Standar Hardware
            "Output static 5V, 12V", 
            "Max 70°C saat full load",
            "Putaran lancar, tanpa suara aneh",
            "Tidak ada error saat tes",
            "Health > 90%, tanpa bad sector",
            "Suhu normal, performa lancar",
            "Semua port berfungsi",
            "Warna cerah, resolusi sesuai",
            "Semua tombol berfungsi",
            "Tegangan ±3V",
            
            # Standar Software
            "Versi terbaru, tidak ada error",
            "Real-time protection aktif",
            "Hanya aplikasi penting yang aktif",
            "<80% dari total kapasitas",
            "CPU <10%, RAM <50% saat idle",
            "Tidak ada error kritikal",
            "Berfungsi baik dan update",
            "Ping <10ms, koneksi stabil",
            "Backup rutin, restore bisa dilakukan",
            "Lisensi aktif dan legal",
            
            # Standar Laptop
            "Kapasitas > 80%",
            "Output tegangan stabil",
            "Tidak flicker, warna cerah",
            "Responsif, tidak delay",
            "Semua tombol berfungsi",
            "Gambar jernih",
            "Suara jelas, tidak noise",
            "Suara tidak pecah",
            "Kokoh, tidak longgar",
            "Semua port bisa digunakan"
        ],
        # Mendefinisikan metode pengujian untuk setiap item
        "Testing Method": [
            # Metode pengujian Hardware
            "Gunakan PSU Tester (manual) atau multimeter digital untuk mengukur. Periksa pada konektor 24-pin dan 8-pin CPU.", 
            "Install HWMonitor (free) atau CoreTemp (free). Jalankan saat idle & full load (gunakan stress test seperti Prime95).",
            "Gunakan SpeedFan (link) atau cek langsung di BIOS untuk melihat RPM fan. Perhatikan adanya suara kasar atau getaran.",
            "Jalankan MemTest86 (link), boot melalui USB, dan jalankan minimal 2 pass. Amati jika muncul error.",
            "Gunakan CrystalDiskInfo (link) untuk SMART check, atau HDDScan (link) untuk surface test.",
            "Gunakan GPU-Z (link) untuk monitoring, dan FurMark (link) untuk stress test suhu dan stabilitas.",
            "Tes dengan flashdisk, LAN cable, headset. Periksa Device Manager untuk status perangkat. Gunakan loopback plug untuk port LAN jika tersedia.",
            "Gunakan Dead Pixel Buddy (link) untuk tes dead pixel. Cek juga brightness dan warna lewat Display Settings atau software kalibrasi layar.",
            "Gunakan Keyboard Tester (link) dan tes klik mouse kiri/kanan + scroll. Bisa juga menggunakan aplikasi MouseTester untuk respons.",
            "Bongkar casing PC, lalu ukur langsung tegangan baterai menggunakan multimeter digital.",
            
            # Metode pengujian Software
            "Jalankan Windows Update, dan gunakan sfc /scannow di Command Prompt untuk cek integritas file sistem. Pastikan update berhasil dan tidak ada corrupt files.",
            "Jalankan full scan menggunakan Windows Defender atau Kaspersky Free. Pastikan database signature terbaru.",
            "Buka Task Manager > Startup, disable aplikasi tidak penting. Gunakan Autoruns dari Sysinternals untuk manajemen lebih detail.",
            "Gunakan Windows Explorer atau WinDirStat (link) untuk melihat pemakaian storage yang besar.",
            "Gunakan Task Manager > Performance atau Resource Monitor. Gunakan aplikasi seperti Chrome/Excel untuk simulasi beban ringan.",
            "Buka Event Viewer > Windows Logs > System, filter event dengan level Critical atau Error, periksa ID kejadian dan deskripsinya.",
            "Jalankan aplikasi satu per satu, tes fungsi utama. Lihat juga error log jika tersedia di dalam aplikasi (misal Excel > Options > Add-ins).",
            "Gunakan ping 8.8.8.8 -t di CMD untuk kestabilan, dan Speedtest untuk bandwidth. Gunakan juga Traceroute jika ping tidak stabil.",
            "Lakukan backup ke external drive atau cloud (OneDrive/Google Drive). Simulasikan restore 1 file sebagai tes. Gunakan juga Macrium Reflect untuk full image.",
            "Cek aktivasi dengan slmgr /xpr di CMD untuk Windows. Cek lisensi aplikasi lain lewat About section. Simpan bukti pembelian.",
            
            # Metode pengujian Laptop
            "Gunakan BatteryInfoView (link) untuk cek kapasitas desain vs real, cycle count, wear level.",
            "Gunakan multimeter digital untuk mengukur output DC charger (biasanya 19V). Pastikan konektor tidak longgar.",
            "Gunakan Dead Pixel Buddy (link), tes layar fullscreen. Cek juga brightness dan saturasi.",
            "Gerakkan kursor, klik kiri/kanan langsung di touchpad. Gunakan Control Panel > Mouse > Touchpad Settings untuk kalibrasi.",
            "Tes dengan KeyboardTester (link) atau Notepad sambil ketik semua huruf. Periksa delay atau tombol nyangkut.",
            "Buka Camera App bawaan Windows atau aplikasi meeting (Zoom/Google Meet). Tes pencahayaan, fokus, noise.",
            "Buka Sound Settings > Input > Test Microphone, atau tes rekam suara di aplikasi Voice Recorder. Gunakan headset untuk pembanding.",
            "Putar file musik/video, lalu periksa di Sound Settings untuk balance L/R. Gunakan Ear Test MP3 sebagai referensi suara.",
            "Lakukan buka/tutup layar 10–20 kali, periksa bunyi, goyangan, dan retakan fisik. Gunakan obeng kecil jika perlu membuka casing.",
            "Gunakan kabel HDMI, flashdisk USB-C, dan kartu SD asli. Cek di Device Manager jika tidak terbaca. Tes koneksi ke proyektor/monitor eksternal."
        ],
        # Mendefinisikan dampak jika kondisi tidak normal
        "Impact if Abnormal": [
            # Dampak Hardware
            "PC tidak menyala, tidak ada daya", 
            "Overheating, system freeze/restart",
            "Overheating CPU, performa menurun",
            "Crash, BSOD, performa lambat",
            "Lambat, gagal boot, data hilang",
            "Gambar rusak, layar hitam, performa rendah",
            "Tidak terhubung ke perangkat penting",
            "Tampilan terganggu, visual error",
            "Input error, lambat bekerja",
            "BIOS reset, konfigurasi hilang",
            
            # Dampak Software
            "Crash sistem, aplikasi gagal dijalankan",
            "Rentan virus, pencurian data",
            "Boot lambat, konsumsi resource tinggi",
            "Sistem lambat, tidak bisa simpan file",
            "Lambat bahkan saat idle",
            "Gangguan sistem mendadak",
            "Produktivitas terganggu",
            "Putus koneksi ke server/cloud",
            "Data hilang permanen",
            "Tidak bisa update, software dinonaktifkan",
            
            # Dampak Laptop
            "Laptop mati mendadak, tidak tahan lama",
            "Tidak mengisi, baterai rusak",
            "Gangguan penglihatan",
            "Navigasi sulit",
            "Salah input saat mengetik",
            "Tidak bisa meeting online",
            "Tidak terdengar lawan bicara",
            "Gangguan audio saat meeting",
            "Layar bisa rusak parah",
            "Tidak bisa output ke layar eksternal"
        ],
        # Mendefinisikan rekomendasi pemeliharaan
        "Maintenance Recommendation": [
            # Pemeliharaan Hardware
            "3-5 tahun atau jika ada gejala kegagalan daya", 
            "Cek kipas, ganti thermal paste setiap 1 tahun",
            "Ganti jika suara kasar/macet",
            "5 tahun atau jika mulai bermasalah",
            "HDD: 3–5 tahun, SSD: 5–7 tahun",
            "Ganti jika error visual atau overheating",
            "Ganti jika rusak/tidak responsif",
            "5–7 tahun atau jika flicker/warna tidak stabil",
            "Ganti jika tidak responsif",
            "2–3 tahun atau jika waktu sering reset",
            
            # Pemeliharaan Software
            "Reinstall jika tidak stabil",
            "Update harian, ganti jika tidak efektif",
            "Optimasi bulanan",
            "Tambah storage jika sering penuh",
            "Tambah RAM jika selalu penuh",
            "Monitoring bulanan",
            "Reinstall jika sering crash",
            "Upgrade jika lambat",
            "Cek restore setiap bulan",
            "Perpanjang lisensi tahunan",
            
            # Pemeliharaan Laptop
            "2–4 tahun tergantung pemakaian",
            "2–3 tahun atau bila rusak",
            "5–7 tahun atau bila rusak",
            "Ganti jika rusak",
            "Ganti jika tidak berfungsi",
            "Ganti internal atau pakai webcam USB",
            "Ganti atau pakai eksternal",
            "Ganti jika rusak",
            "Perbaiki jika longgar",
            "Servis jika tidak terdeteksi"
        ],
        # Kolom kosong untuk hasil tes dan catatan
        "Test Result": [""]*30,  # Membuat 30 baris kosong untuk hasil tes
        "Notes": [""]*30  # Membuat 30 baris kosong untuk catatan
    }
    
    # Membuat DataFrame dari dictionary data
    df = pd.DataFrame(data)
    
    # Menampilkan judul tabel
    st.write("### Testing Items")
    
    # Menambahkan CSS kustom untuk tampilan yang lebih baik
    st.markdown("""
        <style>
        .block-container {
            max-width: 1800px;
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        .element-container {
            width: 100%;
        }
        .stTextArea textarea {
            width: 100%;
        }
        .stRadio > label {
            min-width: 100px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Membuat form untuk checklist
    with st.form("hardware_test_form"):
        # Iterasi setiap baris dalam DataFrame
        for index, row in df.iterrows():
            # Membuat kolom untuk layout horizontal
            cols = st.columns([1, 2, 2, 4, 4, 1, 4])
            
            # Menampilkan nomor item
            with cols[0]:
                st.markdown(f"<div style='text-align: center'><strong>{row['No']}</strong></div>", unsafe_allow_html=True)
            
            # Menampilkan item tes dan kategori
            with cols[1]:
                st.markdown(f"<div style='min-width: 200px'><strong>Item:</strong> {row['Test Item']}<br><em>{row['Category']}</em></div>", unsafe_allow_html=True)
            
            # Menampilkan standar pengujian
            with cols[2]:
                st.markdown(f"<div style='min-width: 200px'><strong>Standard:</strong><br>{row['Testing Standard/Normal']}</div>", unsafe_allow_html=True)
            
            # Menampilkan metode pengujian
            with cols[3]:
                st.markdown(f"<div style='min-width: 200px'><strong>Method:</strong><br>{row['Testing Method']}</div>", unsafe_allow_html=True)
            
            # Menampilkan dampak dan pemeliharaan
            with cols[4]:
                st.markdown(f"<div style='min-width: 200px'><strong>Impact:</strong><br>{row['Impact if Abnormal']}<br><strong>Maintenance:</strong><br>{row['Maintenance Recommendation']}</div>", unsafe_allow_html=True)
            
            # Menambahkan radio button untuk hasil tes
            with cols[5]:
                test_result = st.radio(
                    "Result",
                    ["", "Yes", "No"],
                    key=f"result_{index}",
                    horizontal=False
                )
                df.at[index, "Test Result"] = test_result
            
            # Menambahkan text area untuk catatan
            with cols[6]:
                notes = st.text_area(
                    "Notes",
                    value="",
                    key=f"notes_{index}",
                    height=100,
                    max_chars=None,
                    help="Enter any additional notes here"
                )
                df.at[index, "Notes"] = notes
            
            # Menambahkan garis pemisah antar item
            st.markdown("---")
        
        # Menambahkan tombol submit
        
        submitted = st.form_submit_button("Save Results")

    
    # Inisialisasi session state untuk data form
    if 'form_submitted' not in st.session_state:
        st.session_state.form_submitted = False
        st.session_state.form_data = None
        st.session_state.excel_buffer = None
     
    # Menangani pengiriman form
    if submitted:
        # Validasi semua tes telah diisi
        if not all(df['Test Result'].isin(['Yes', 'No'])):
            st.warning("Harap isi **SEMUA** hasil pengecekan (Yes/No)!")
            st.stop()
        
        # Menyiapkan data untuk MongoDB
        document = {
            "metadata": {
                "location": location,
                "user_name": user_name,
                "checklist_date": checklist_date.strftime("%Y-%m-%d"),
                "submission_time": datetime.now().isoformat()
            },
            "tests": df.to_dict('records')
        }
        
        # Menyimpan ke MongoDB
        inserted_id = save_to_mongodb(document)
        
        # Update session state jika penyimpanan berhasil
        if inserted_id:
            st.session_state.form_submitted = True
            st.session_state.form_data = {
                'document': document,
                'df': df,
                'location': location,
                'user_name': user_name,
                'checklist_date': checklist_date,
                'inserted_id': inserted_id
            }

    # Menampilkan hasil dan opsi unduh
    if st.session_state.form_submitted and st.session_state.form_data:
        data = st.session_state.form_data
        
        # Menampilkan pesan sukses
        st.success(f"Test results saved successfully! Document ID: {data['inserted_id']}")
        st.write("### Saved Results")
        
        # Menampilkan metadata
        st.write(f"**Location:** {data['location']}")
        st.write(f"**User Name:** {data['user_name']}")
        st.write(f"**Checklist Date:** {data['checklist_date'].strftime('%Y-%m-%d')}")
        
        # Menampilkan tabel hasil
        st.dataframe(data['df'][["No", "Test Item", "Test Result", "Notes"]])
        
        # Membuat file Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            # Menulis metadata
            metadata_df = pd.DataFrame([
                ["Location", data['location']],
                ["User Name", data['user_name']],
                ["Checklist Date", data['checklist_date'].strftime("%Y-%m-%d")],
                ["Submission Time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                ["", ""]
            ], columns=["Field", "Value"])
            metadata_df.to_excel(writer, sheet_name='Hasil Checklist', index=False, startrow=0)
            
            # Menulis data checklist
            data['df'].to_excel(writer, sheet_name='Hasil Checklist', index=False, startrow=len(metadata_df) + 2)
            
            # Format file Excel
            workbook = writer.book
            worksheet = writer.sheets['Hasil Checklist']
            
            # Menambahkan format header
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1
            })
            
            # Menerapkan format header
            for col_num, value in enumerate(data['df'].columns.values):
                worksheet.write(len(metadata_df) + 2, col_num, value, header_format)
            
            # Mengatur lebar kolom
            worksheet.set_column('A:Z', 20)
        
        # Membuat file PDF
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=portrait(A4),
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30
        )
        
        # Menyiapkan elemen PDF
        elements = []
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=30
        )
        
        # Menambahkan konten PDF
        elements.append(Paragraph("Hardware Testing Checklist", title_style))
        elements.append(Paragraph(f"Location: {data['location']}", styles["Normal"]))
        elements.append(Paragraph(f"User Name: {data['user_name']}", styles["Normal"]))
        elements.append(Paragraph(f"Date: {data['checklist_date'].strftime('%Y-%m-%d')}", styles["Normal"]))
        elements.append(Spacer(1, 20))
        
        # Membuat tabel PDF
        table_data = [["No", "Test Item", "Result", "Notes"]]
        for _, row in data['df'].iterrows():
            table_data.append([
                str(row['No']),
                row['Test Item'],
                row['Test Result'],
                row['Notes']
            ])
        
        # Format tabel PDF
        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('WORDWRAP', (0, 0), (-1, -1), True),
        ]))
        
        elements.append(table)
        doc.build(elements)
        
        # Membuat tombol unduh
        col1, col2 = st.columns(2)
        
        # Tombol unduh Excel
        with col1:
            st.download_button(
                "📥 Download Excel",
                excel_buffer.getvalue(),
                "hardware_checklist.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='download-excel'
            )
        
        # Tombol unduh PDF
        with col2:
            st.download_button(
                "📄 Download PDF",
                pdf_buffer.getvalue(),
                "hardware_checklist.pdf",
                "application/pdf",
                key='download-pdf'
            )
            
if __name__ == "__main__":
    # Menjalankan fungsi utama
    main()
