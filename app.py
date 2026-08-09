import streamlit as st
import pandas as pd
import io

# Coba impor BeautifulSoup untuk simulasi web scraping
try:
    from bs4 import BeautifulSoup
    bs4_installed = True
except ImportError:
    bs4_installed = False

# Coba impor openpyxl untuk mengecek instalasi fitur Excel
try:
    import openpyxl
    openpyxl_installed = True
except ImportError:
    openpyxl_installed = False

# =====================================================================
# INISIALISASI SESSION STATE
# =====================================================================
if 'mulai_belajar' not in st.session_state:
    st.session_state.mulai_belajar = False
if 'nama_siswa' not in st.session_state:
    st.session_state.nama_siswa = ""
if 'kelas_siswa' not in st.session_state:
    st.session_state.kelas_siswa = ""
if 'absen_siswa' not in st.session_state:
    st.session_state.absen_siswa = ""

# Inisialisasi variabel jawaban agar nilainya abadi (tidak terhapus saat ganti menu)
if 'j_soal1' not in st.session_state:
    st.session_state.j_soal1 = "Pilih jawaban Anda"
if 'j_soal2' not in st.session_state:
    st.session_state.j_soal2 = "Pilih jawaban Anda"
if 'j_soal3' not in st.session_state:
    st.session_state.j_soal3 = "Pilih jawaban Anda"
if 'j_soal4' not in st.session_state:
    st.session_state.j_soal4 = "Pilih jawaban Anda"
if 'j_refleksi' not in st.session_state:
    st.session_state.j_refleksi = ""

# =====================================================================
# HALAMAN 1: FORM IDENTITAS
# =====================================================================
if not st.session_state.mulai_belajar:
    st.title("LKPD Interaktif BAB 1")
    st.subheader("Data, Informasi, dan Validasinya")
    st.write("LKPD ini digunakan untuk belajar tentang data, informasi, web scraping, dan validasi informasi.")

    st.divider()

    st.header("Form Identitas Siswa")
    nama = st.text_input("Nama Lengkap:")
    kelas = st.text_input("Kelas:")
    nomor_absen = st.text_input("Nomor Absen:")

    if st.button("Mulai Belajar"):
        if nama and kelas and nomor_absen:
            st.session_state.nama_siswa = nama
            st.session_state.kelas_siswa = kelas
            st.session_state.absen_siswa = nomor_absen
            st.session_state.mulai_belajar = True
            st.rerun()
        else:
            st.warning("Mohon isi semua data identitas terlebih dahulu sebelum mulai.")

# =====================================================================
# HALAMAN 2: MENU AKTIVITAS & EVALUASI
# =====================================================================
else:
    # --- SIDEBAR (NAVIGASI DI SAMPING) ---
    st.sidebar.title("Menu LKPD")
    st.sidebar.write(f"Siswa: **{st.session_state.nama_siswa}**")
    st.sidebar.write(f"Kelas: **{st.session_state.kelas_siswa}**")
    st.sidebar.write(f"Absen: **{st.session_state.absen_siswa}**")
    
    st.sidebar.divider()
    
    # Navigasi antar aktivitas
    menu_pilihan = st.sidebar.radio(
        "Pilih Aktivitas Belajar:",
        [
            "Aktivitas 1: Data vs Informasi", 
            "Aktivitas 2: Simulasi Web Scraping", 
            "Aktivitas 3: Validasi Informasi",
            "Rapor Hasil Belajar"
        ]
    )
    
    st.sidebar.divider()
    
    # Tombol keluar/reset
    if st.sidebar.button("Keluar / Reset Data"):
        st.session_state.mulai_belajar = False
        st.session_state.j_soal1 = "Pilih jawaban Anda"
        st.session_state.j_soal2 = "Pilih jawaban Anda"
        st.session_state.j_soal3 = "Pilih jawaban Anda"
        st.session_state.j_soal4 = "Pilih jawaban Anda"
        st.session_state.j_refleksi = ""
        st.rerun()

    # --- KONTEN AKTIVITAS 1: DATA VS INFORMASI ---
    if menu_pilihan == "Aktivitas 1: Data vs Informasi":
        st.title("Aktivitas 1: Perbedaan Data & Informasi")
        st.write("""
        Sebelum belajar lebih jauh, mari pahami perbedaan dasar antara data dan informasi:
        
        1. **Data** adalah fakta mentah, angka, atau catatan kejadian yang belum diolah dan belum memiliki arti yang jelas bagi orang lain.
           * *Contoh:* Angka `38`, kata `Merah`, angka-angka tinggi badan siswa.
        
        2. **Informasi** adalah data yang sudah diproses, diatur, atau diolah sedemikian rupa sehingga memiliki arti, konteks, dan berguna bagi pembacanya.
           * *Contoh:* Kalimat `"Suhu tubuh Andi saat ini 38°C (Demam)"` atau grafik rata-rata tinggi badan siswa kelas X.
        """)

        st.divider()

        # Visualisasi Data vs Informasi
        st.subheader("💡 Visualisasi: Mengubah Data Menjadi Informasi")
        st.write("Bayangkan kita memiliki data catatan suhu udara kota Jakarta selama 5 hari:")
        
        data_suhu = pd.DataFrame({
            "Hari": ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"],
            "Suhu (°C)": [31, 33, 29, 34, 30]
        })
        st.table(data_suhu)
        
        st.write("""
        Tabel angka di atas adalah **Data** mentah. 
        Jika kita mengubahnya menjadi **Grafik Batang** di bawah ini, kita bisa langsung menyimpulkan dengan mudah bahwa **Kamis** adalah hari yang paling panas. 
        Grafik dan kesimpulan ini disebut sebagai **Informasi**!
        """)
        
        chart_data = data_suhu.set_index("Hari")
        st.bar_chart(chart_data)

        st.divider()

        st.header("Uji Pemahaman Mandiri")
        st.write("Pilihlah jawaban yang paling tepat untuk pernyataan di bawah ini:")

        opsi_soal1 = ["Pilih jawaban Anda", "Data", "Informasi"]
        idx_soal1 = opsi_soal1.index(st.session_state.j_soal1)
        soal1 = st.radio(
            "1. Catatan daftar nilai ujian matematika murid kelas 10 (misal: 80, 75, 90, 85) yang belum dirata-rata disebut sebagai...",
            opsi_soal1,
            index=idx_soal1
        )
        st.session_state.j_soal1 = soal1

        opsi_soal2 = ["Pilih jawaban Anda", "Data", "Informasi"]
        idx_soal2 = opsi_soal2.index(st.session_state.j_soal2)
        soal2 = st.radio(
            "2. Laporan grafik peningkatan rata-rata nilai matematika kelas 10 dari bulan lalu ke bulan ini disebut sebagai...",
            opsi_soal2,
            index=idx_soal2
        )
        st.session_state.j_soal2 = soal2

        if st.button("Cek Jawaban"):
            if soal1 == "Pilih jawaban Anda" or soal2 == "Pilih jawaban Anda":
                st.warning("Silakan pilih jawaban untuk semua soal terlebih dahulu!")
            else:
                st.subheader("Hasil Koreksi:")
                if soal1 == "Data":
                    st.success("Soal 1: BENAR! Catatan nilai mentah yang belum diolah adalah **Data**.")
                else:
                    st.error("Soal 1: SALAH. Catatan nilai mentah belum diolah, jadi itu adalah **Data**.")

                if soal2 == "Informasi":
                    st.success("Soal 2: BENAR! Laporan grafik yang sudah diolah dan memiliki arti adalah **Informasi**.")
                else:
                    st.error("Soal 2: SALAH. Laporan grafik sudah diproses dan memiliki makna bagi pembacanya, jadi itu adalah **Informasi**.")

    # --- KONTEN AKTIVITAS 2: SIMULASI WEB SCRAPING ---
    elif menu_pilihan == "Aktivitas 2: Simulasi Web Scraping":
        st.title("Aktivitas 2: Mengenal Web Scraping")
        st.write("""
        **Web Scraping** adalah teknik untuk mengambil data secara otomatis dari halaman website menggunakan kode program. 
        Karena sebuah halaman web disusun menggunakan bahasa pemrograman **HTML** (HyperText Markup Language), 
        maka program Python kita bertugas mencari dan memisahkan data penting dari tumpukan kode HTML tersebut.
        """)
        
        st.divider()
        
        st.subheader("1. Contoh Kode HTML Mentah (Sumber Data)")
        st.write("Bayangkan kita ingin mengambil daftar berita hangat dari potongan HTML di bawah ini:")
        
        html_mentah = """<ul>
  <li class="berita">Kemajuan AI di Indonesia Semakin Pesat</li>
  <li class="berita">Belajar Python Sangat Menyenangkan bagi Pemula</li>
  <li class="berita">Tips dan Trik Membuat Web dengan Streamlit</li>
</ul>"""
        st.code(html_mentah, language="html")

        st.subheader("2. Kode Python untuk Scraping")
        st.write("Di Python, kita biasa menggunakan library bernama **BeautifulSoup** untuk mencari tag HTML secara otomatis:")
        
        kode_python_contoh = """from bs4 import BeautifulSoup

# Memasukkan HTML mentah ke BeautifulSoup
soup = BeautifulSoup(html_mentah, 'html.parser')

# Mengambil semua tag <li> yang memiliki class 'berita'
daftar = soup.find_all('li', class_='berita')

# Mengambil teks bersihnya saja
hasil = []
for item in daftar:
    hasil.append(item.text)"""
        st.code(kode_python_contoh, language="python")

        st.divider()

        st.subheader("3. Uji Coba Jalankan Scraping")
        st.write("Klik tombol di bawah untuk melihat bagaimana BeautifulSoup menyaring HTML mentah menjadi teks bersih:")

        if st.button("Jalankan Scraping"):
            if not bs4_installed:
                st.error("Pustaka `beautifulsoup4` belum terinstall di komputer Anda.")
                st.info("Silakan jalankan perintah berikut di terminal Anda untuk menginstalnya: \n`pip install beautifulsoup4`")
            else:
                soup = BeautifulSoup(html_mentah, 'html.parser')
                daftar = soup.find_all('li', class_='berita')
                
                st.success("Scraping Berhasil!")
                st.write("Berikut data bersih yang berhasil diekstrak oleh program Python:")
                for index, item in enumerate(daftar, 1):
                    st.info(f"**Data {index}:** {item.text}")

    # --- KONTEN AKTIVITAS 3: VALIDASI INFORMASI ---
    elif menu_pilihan == "Aktivitas 3: Validasi Informasi":
        st.title("Aktivitas 3: Validasi Informasi")
        st.write("""
        **Validasi Informasi** adalah proses menyaring, memeriksa, dan memastikan kebenaran dari data atau informasi yang kita terima.
        Di dunia maya, banyak beredar informasi palsu atau biasa disebut **Hoaks (Hoax)**. 
        Sebagai pelajar, kita harus kritis dengan memeriksa sumber berita sebelum memercayainya atau menyebarkannya.
        """)

        st.divider()

        st.subheader("Studi Kasus: Membaca Pesan Berantai")
        st.warning("""
        **Pesan Berantai WhatsApp:**
        "INFO PENTING! Mulai besok pagi, seluruh jaringan internet seluler di Indonesia akan dimatikan total selama 3 hari karena adanya badai matahari ekstrem yang membakar satelit bumi. Harap sebarkan pesan ini ke grup keluarga dan teman dekat Anda agar semua bersiap-siap membeli bahan makanan!"
        """)

        st.write("Mari kita analisis pesan di atas dengan menjawab kuis di bawah ini:")

        opsi_soal3 = ["Pilih jawaban Anda", "Ya, ada sumber resminya", "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi"]
        idx_soal3 = opsi_soal3.index(st.session_state.j_soal3)
        soal3 = st.radio(
            "1. Apakah pesan di atas menyebutkan tanggal kejadian spesifik dan mencantumkan sumber lembaga resmi (seperti BMKG atau NASA)?",
            opsi_soal3,
            index=idx_soal3
        )
        st.session_state.j_soal3 = soal3

        opsi_soal4 = ["Pilih jawaban Anda", "Informasi Valid (Fakta)", "Informasi Palsu (Hoaks)"]
        idx_soal4 = opsi_soal4.index(st.session_state.j_soal4)
        soal4 = st.radio(
            "2. Berdasarkan analisis Anda, pesan berantai tersebut tergolong sebagai...",
            opsi_soal4,
            index=idx_soal4
        )
        st.session_state.j_soal4 = soal4

        if st.button("Verifikasi Berita"):
            if soal3 == "Pilih jawaban Anda" or soal4 == "Pilih jawaban Anda":
                st.warning("Silakan pilih jawaban untuk semua soal terlebih dahulu!")
            else:
                st.subheader("Hasil Verifikasi:")
                if soal3 == "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi":
                    st.success("Soal 1: BENAR! Pesan hoaks biasanya tidak memiliki tanggal jelas dan tidak menyebutkan sumber resmi.")
                else:
                    st.error("Soal 1: SALAH. Perhatikan baik-baik, pesan tersebut tidak menyebutkan lembaga resmi mana pun sebagai sumbernya.")

                if soal4 == "Informasi Palsu (Hoaks)":
                    st.success("Soal 2: BENAR! Kabar tersebut adalah kabar bohong (hoaks) yang dibuat untuk menimbulkan kepanikan.")
                else:
                    st.error("Soal 2: SALAH. Kabar pemutusan internet nasional akibat badai matahari adalah hoaks.")

    # --- KONTEN RAPOR HASIL BELAJAR ---
    elif menu_pilihan == "Rapor Hasil Belajar":
        st.title("Rapor Hasil Belajar LKPD")
        st.write("Halaman ini menyajikan rangkuman hasil jawaban Anda dari seluruh aktivitas belajar di LKPD ini.")

        ans1 = st.session_state.j_soal1
        ans2 = st.session_state.j_soal2
        ans3 = st.session_state.j_soal3
        ans4 = st.session_state.j_soal4

        belum_selesai = (
            ans1 == "Pilih jawaban Anda" or 
            ans2 == "Pilih jawaban Anda" or 
            ans3 == "Pilih jawaban Anda" or 
            ans4 == "Pilih jawaban Anda"
        )

        if belum_selesai:
            st.warning("Anda belum menyelesaikan seluruh kuis di **Aktivitas 1** dan **Aktivitas 3**. Silakan lengkapi jawaban Anda terlebih dahulu untuk melihat Rapor!")
        else:
            # Hitung skor
            skor = 0
            eval1 = "SALAH"
            eval2 = "SALAH"
            eval3 = "SALAH"
            eval4 = "SALAH"

            if ans1 == "Data":
                skor += 25
                eval1 = "BENAR"
            if ans2 == "Informasi":
                skor += 25
                eval2 = "BENAR"
            if ans3 == "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi":
                skor += 25
                eval3 = "BENAR"
            if ans4 == "Informasi Palsu (Hoaks)":
                skor += 25
                eval4 = "BENAR"

            # Efek Balon Sukses
            st.balloons()
            
            st.subheader("Kartu Hasil Belajar")
            
            pesan_motivasi = ""
            if skor == 100:
                pesan_motivasi = "🎉 Luar biasa sempurna! Kamu memahami seluruh konsep data, informasi, dan validasi dengan sangat baik!"
            elif skor >= 75:
                pesan_motivasi = "👍 Kerja bagus! Pemahaman kamu sudah sangat baik. Hanya butuh sedikit ketelitian lagi!"
            else:
                pesan_motivasi = "📚 Tetap semangat! Coba baca kembali materi-materinya dan diskusikan dengan teman atau gurumu."

            st.info(f"""
            **IDENTITAS SISWA:**
            * Nama: {st.session_state.nama_siswa}
            * Kelas: {st.session_state.kelas_siswa}
            * Nomor Absen: {st.session_state.absen_siswa}
            
            **NILAI AKHIR ANDA:**
            ## {skor} / 100
            
            *{pesan_motivasi}*
            """)

            st.subheader("Rincian Evaluasi Jawaban")
            st.write(f"1. Soal Nilai Mentah (Aktivitas 1): **{ans1}** ({eval1})")
            st.write(f"2. Soal Grafik Nilai (Aktivitas 1): **{ans2}** ({eval2})")
            st.write(f"3. Soal Ciri Berita (Aktivitas 3): **{ans3}** ({eval3})")
            st.write(f"4. Soal Kategori Berita (Aktivitas 3): **{ans4}** ({eval4})")

            st.divider()
            
            # Form Refleksi
            st.subheader("📝 Refleksi Belajar")
            st.write("Tuliskan kesimpulan atau apa yang telah kamu pelajari setelah menyelesaikan LKPD ini:")
            refleksi_siswa = st.text_area(
                "Tanggapan/Refleksi Siswa:",
                value=st.session_state.j_refleksi,
                placeholder="Tulis di sini... (Misal: Hari ini saya belajar cara membedakan data mentah dengan informasi yang matang, serta bagaimana cara memverifikasi pesan hoaks di internet.)"
            )
            st.session_state.j_refleksi = refleksi_siswa

            st.divider()
            st.subheader("Ekspor Laporan Hasil Belajar")
            
            # --- OPSI 1: UNDUH SEBAGAI EXCEL (.xlsx) ---
            st.write("### Opsi A: Format Excel (.xlsx) - Sangat Direkomendasikan")
            
            if not openpyxl_installed:
                st.error("Pustaka `openpyxl` belum terinstall di lingkungan Python Anda.")
                st.info("Silakan jalankan perintah ini di terminal Anda untuk menginstalnya: \n`pip install openpyxl` \n\nSetelah selesai instalasi, silakan muat ulang (refresh) halaman browser Anda.")
            else:
                data_excel = {
                    "Kategori Data": [
                        "Identitas Siswa", "Identitas Siswa", "Identitas Siswa", 
                        "Soal Uji Pemahaman 1", "Soal Uji Pemahaman 2", 
                        "Soal Validasi 1", "Soal Validasi 2", "Skor Akhir",
                        "Refleksi Siswa"
                    ],
                    "Variabel / Pertanyaan": [
                        "Nama Siswa", "Kelas", "Nomor Absen",
                        "Catatan nilai mentah yang belum diolah",
                        "Laporan grafik peningkatan nilai",
                        "Pesan berantai memiliki sumber resmi?",
                        "Kategori pesan berantai",
                        "Total Nilai LKPD",
                        "Apa yang dipelajari"
                    ],
                    "Jawaban / Hasil": [
                        st.session_state.nama_siswa,
                        st.session_state.kelas_siswa,
                        st.session_state.absen_siswa,
                        ans1, ans2, ans3, ans4,
                        f"{skor} / 100",
                        refleksi_siswa if refleksi_siswa else "(Tidak diisi)"
                    ],
                    "Evaluasi": [
                        "-", "-", "-",
                        eval1, eval2, eval3, eval4,
                        "LULUS" if skor >= 75 else "BELUM LULUS",
                        "-"
                    ]
                }
                
                df = pd.DataFrame(data_excel)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Hasil_LKPD_Siswa')
                
                st.download_button(
                    label="📁 Unduh Laporan format Excel (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name=f"Laporan_LKPD_{st.session_state.nama_siswa.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # --- OPSI 2: UNDUH SEBAGAI TEKS (.txt) ---
            st.write("### Opsi B: Format Teks (.txt)")
            teks_laporan = f"""=== LAPORAN HASIL BELAJAR LKPD INTERAKTIF ===
BAB 1: Data, Informasi, dan Validasinya

IDENTITAS SISWA:
- Nama: {st.session_state.nama_siswa}
- Kelas: {st.session_state.kelas_siswa}
- Nomor Absen: {st.session_state.absen_siswa}

EVALUASI JAWABAN:
1. Soal Nilai Mentah: {ans1} -> {eval1}
2. Soal Grafik Nilai: {ans2} -> {eval2}
3. Soal Ciri Berita: {ans3} -> {eval3}
4. Soal Kategori Berita: {ans4} -> {eval4}

NILAI AKHIR: {skor} / 100

REFLEKSI SISWA:
{refleksi_siswa if refleksi_siswa else '(Tidak diisi)'}
=============================================
"""
            st.download_button(
                label="📄 Unduh Laporan format Teks (.txt)",
                data=teks_laporan,
                file_name=f"Laporan_LKPD_{st.session_state.nama_siswa.replace(' ', '_')}.txt",
                mime="text/plain"
            )
