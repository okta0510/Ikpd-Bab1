import streamlit as st
import pandas as pd
import io
from questions import QUIZ_QUESTIONS

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
def init_session_states():
    # Identitas
    if 'mulai_belajar' not in st.session_state:
        st.session_state.mulai_belajar = False
    if 'nama_siswa' not in st.session_state:
        st.session_state.nama_siswa = ""
    if 'kelas_siswa' not in st.session_state:
        st.session_state.kelas_siswa = ""
    if 'absen_siswa' not in st.session_state:
        st.session_state.absen_siswa = ""
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "1. 🏠 Beranda LKPD"
        
    # Read triggers
    if 'tujuan_read' not in st.session_state:
        st.session_state.tujuan_read = False
    if 'materi_read' not in st.session_state:
        st.session_state.materi_read = False

    # Halaman 5 (Aktivitas)
    if 'j_soal1' not in st.session_state:
        st.session_state.j_soal1 = "Pilih jawaban Anda"
    if 'j_soal2' not in st.session_state:
        st.session_state.j_soal2 = "Pilih jawaban Anda"
    if 'j_soal3' not in st.session_state:
        st.session_state.j_soal3 = "Pilih jawaban Anda"
    if 'j_soal4' not in st.session_state:
        st.session_state.j_soal4 = "Pilih jawaban Anda"
    if 'act1_checked' not in st.session_state:
        st.session_state.act1_checked = False
    if 'act3_checked' not in st.session_state:
        st.session_state.act3_checked = False

    # Game 1: Tebak Operator
    if 'g1_answers' not in st.session_state:
        st.session_state.g1_answers = [None] * 5
    if 'g1_checked' not in st.session_state:
        st.session_state.g1_checked = False

    # Game 2: Pasangkan Konsep
    if 'g2_answers' not in st.session_state:
        st.session_state.g2_answers = [None] * 5
    if 'g2_checked' not in st.session_state:
        st.session_state.g2_checked = False

    # Game 3: Detektif Informasi
    if 'g3_answers' not in st.session_state:
        st.session_state.g3_answers = [None] * 5
    if 'g3_checked' not in st.session_state:
        st.session_state.g3_checked = False

    # Game 4: CRAAP Test
    if 'g4_answers' not in st.session_state:
        st.session_state.g4_answers = [None] * 5
    if 'g4_checked' not in st.session_state:
        st.session_state.g4_checked = False

    # Game 5: Susun Langkah Scraping
    if 'g5_answers' not in st.session_state:
        st.session_state.g5_answers = [None] * 4
    if 'g5_checked' not in st.session_state:
        st.session_state.g5_checked = False

    # Game 6: Tebak Kode Python
    if 'g6_answers' not in st.session_state:
        st.session_state.g6_answers = [None] * 5
    if 'g6_checked' not in st.session_state:
        st.session_state.g6_checked = False

    # Game 7: TTS
    if 'g7_answers' not in st.session_state:
        st.session_state.g7_answers = [""] * 5
    if 'g7_checked' not in st.session_state:
        st.session_state.g7_checked = False

    # Game 8: Tantangan Kasus
    if 'g8_answers' not in st.session_state:
        st.session_state.g8_answers = [None] * 5
    if 'g8_checked' not in st.session_state:
        st.session_state.g8_checked = False

    # Kuis Akhir
    if 'quiz_answers' not in st.session_state:
        st.session_state.quiz_answers = [None] * 10
    if 'quiz_checked' not in st.session_state:
        st.session_state.quiz_checked = False

    # Refleksi
    if 'j_refleksi' not in st.session_state:
        st.session_state.j_refleksi = ""

init_session_states()

# =====================================================================
# CUSTOM CSS AESTHETICS (Google Font Outfit, premium look)
# =====================================================================
def inject_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"], .stText, .stMarkdown {
        font-family: 'Outfit', sans-serif;
    }
    
    .title-gradient {
        background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 2.3rem;
        margin-bottom: 0.2rem;
        text-align: center;
    }
    
    .subtitle-text {
        color: #4b5563;
        font-weight: 500;
        font-size: 1.1rem;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    
    .stButton>button {
        transition: all 0.3s ease;
        border-radius: 8px;
        font-weight: 600;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(139, 92, 246, 0.3);
    }
    
    .materi-header {
        color: #8b5cf6;
        border-bottom: 2px solid #8b5cf6;
        padding-bottom: 5px;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

inject_custom_css()

# =====================================================================
# RESET FUNCTION
# =====================================================================
def reset_all_data():
    st.session_state.mulai_belajar = False
    st.session_state.nama_siswa = ""
    st.session_state.kelas_siswa = ""
    st.session_state.absen_siswa = ""
    st.session_state.current_page = "1. 🏠 Beranda LKPD "
    
    st.session_state.tujuan_read = False
    st.session_state.materi_read = False
    
    st.session_state.j_soal1 = "Pilih jawaban Anda"
    st.session_state.j_soal2 = "Pilih jawaban Anda"
    st.session_state.j_soal3 = "Pilih jawaban Anda"
    st.session_state.j_soal4 = "Pilih jawaban Anda"
    st.session_state.act1_checked = False
    st.session_state.act3_checked = False
    
    st.session_state.g1_answers = [None] * 5
    st.session_state.g1_checked = False
    
    st.session_state.g2_answers = [None] * 5
    st.session_state.g2_checked = False
    
    st.session_state.g3_answers = [None] * 5
    st.session_state.g3_checked = False
    
    st.session_state.g4_answers = [None] * 5
    st.session_state.g4_checked = False
    
    st.session_state.g5_answers = [None] * 4
    st.session_state.g5_checked = False
    
    st.session_state.g6_answers = [None] * 5
    st.session_state.g6_checked = False
    
    st.session_state.g7_answers = [""] * 5
    st.session_state.g7_checked = False
    
    st.session_state.g8_answers = [None] * 5
    st.session_state.g8_checked = False
    
    st.session_state.quiz_answers = [None] * 10
    st.session_state.quiz_checked = False
    
    st.session_state.j_refleksi = ""

# =====================================================================
# HITUNG SKOR & PROGRESS
# =====================================================================
def get_score_aktivitas():
    score = 0
    if st.session_state.j_soal1 == "Data": score += 25
    if st.session_state.j_soal2 == "Informasi": score += 25
    if st.session_state.j_soal3 == "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi": score += 25
    if st.session_state.j_soal4 == "Informasi Palsu (Hoaks)": score += 25
    return score

def get_score_games():
    g1_correct = sum(1 for idx, ans in enumerate(st.session_state.g1_answers) if ans == [2, 1, 1, 1, 1][idx])
    g2_correct = sum(1 for idx, ans in enumerate(st.session_state.g2_answers) if ans == [1, 2, 3, 4, 5][idx])
    g3_correct = sum(1 for idx, ans in enumerate(st.session_state.g3_answers) if ans == [2, 2, 1, 2, 2][idx])
    g4_correct = sum(1 for idx, ans in enumerate(st.session_state.g4_answers) if ans == [1, 5, 2, 3, 4][idx])
    return (g1_correct + g2_correct + g3_correct + g4_correct) * 5 # Total 20 soal x 5 = 100

def get_score_puzzle():
    # Game 5: Susun Langkah (Maks 25)
    g5_correct = sum(1 for idx, ans in enumerate(st.session_state.g5_answers) if ans == [3, 1, 4, 2][idx])
    g5_score = g5_correct * 6.25
    
    # Game 6: Tebak Kode Python (Maks 25)
    g6_correct = sum(1 for idx, ans in enumerate(st.session_state.g6_answers) if ans == [2, 1, 2, 1, 1][idx])
    g6_score = g6_correct * 5
    
    # Game 7: TTS (Maks 25)
    g7_correct = sum(1 for idx, ans in enumerate(st.session_state.g7_answers) if ans.strip().upper() == ["SCRAPING", "DATA", "PYTHON", "HOAKS", "CRAAP"][idx])
    g7_score = g7_correct * 5
    
    # Game 8: Tantangan Kasus (Maks 25)
    g8_correct = sum(1 for idx, ans in enumerate(st.session_state.g8_answers) if ans == [2, 2, 2, 2, 2][idx])
    g8_score = g8_correct * 5
    
    return g5_score + g6_score + g7_score + g8_score

def get_score_kuis():
    kuis_correct = 0
    for idx, question in enumerate(QUIZ_QUESTIONS):
        if st.session_state.quiz_answers[idx] == question["answer"]:
            kuis_correct += 1
    return kuis_correct * 10 # 10 soal x 10 = 100

# =====================================================================
# TEKA-TEKI SILANG DATA & RENDERER
# =====================================================================
grid_letters = {
    (0, 2): 'S', (1, 2): 'C', (2, 2): 'R',
    (3, 1): 'D', (3, 2): 'A', (3, 3): 'T', (3, 4): 'A',
    (4, 2): 'P', (4, 3): 'Y', (4, 4): 'T', (4, 5): 'H', (4, 6): 'O', (4, 7): 'N',
    (5, 2): 'I', (5, 5): 'O',
    (6, 2): 'N', (6, 3): 'C', (6, 4): 'R', (6, 5): 'A', (6, 6): 'A', (6, 7): 'P',
    (7, 2): 'G', (7, 5): 'K',
    (8, 5): 'S'
}

grid_clue_numbers = {
    (0, 2): '1',
    (3, 1): '2',
    (4, 2): '3',
    (4, 5): '4',
    (6, 3): '5'
}

def get_revealed_letter(r, c):
    ans1_ok = st.session_state.g7_answers[0].strip().upper() == "SCRAPING"
    ans2_ok = st.session_state.g7_answers[1].strip().upper() == "DATA"
    ans3_ok = st.session_state.g7_answers[2].strip().upper() == "PYTHON"
    ans4_ok = st.session_state.g7_answers[3].strip().upper() == "HOAKS"
    ans5_ok = st.session_state.g7_answers[4].strip().upper() == "CRAAP"
    
    char = grid_letters[(r, c)]
    revealed = False
    
    if (r, c) in [(0,2), (1,2), (2,2), (5,2), (6,2), (7,2)]:
        revealed = ans1_ok
    elif (r, c) in [(3,1), (3,3), (3,4)]:
        revealed = ans2_ok
    elif (r, c) == (3,2):
        revealed = ans1_ok or ans2_ok
    elif (r, c) == (4,2):
        revealed = ans1_ok or ans3_ok
    elif (r, c) in [(4,3), (4,4), (4,6), (4,7)]:
        revealed = ans3_ok
    elif (r, c) == (4,5):
        revealed = ans3_ok or ans4_ok
    elif (r, c) in [(5,5), (7,5), (8,5)]:
        revealed = ans4_ok
    elif (r, c) in [(6,3), (6,4), (6,6), (6,7)]:
        revealed = ans5_ok
    elif (r, c) == (6,5):
        revealed = ans4_ok or ans5_ok
        
    return char if revealed else ""

def render_crossword_html():
    html = '<div style="display: flex; justify-content: center; margin: 20px 0;"><table style="border-collapse: collapse; border: 2px solid #444; background-color: #f3f4f6; font-family: sans-serif;">'
    for r in range(9):
        html += "<tr>"
        for c in range(8):
            if (r, c) in grid_letters:
                letter = get_revealed_letter(r, c)
                clue_num = grid_clue_numbers.get((r, c), "")
                
                html += '<td style="width: 40px; height: 40px; border: 1px solid #9ca3af; background-color: #ffffff; color: #1e1b4b; text-align: center; font-size: 20px; font-weight: bold; position: relative; padding: 0; box-shadow: inset 0 0 4px rgba(0,0,0,0.1);">'
                if clue_num:
                    html += f'<span style="font-size: 9px; position: absolute; top: 1px; left: 2px; color: #4b5563; font-weight: normal;">{clue_num}</span>'
                html += f'<span style="line-height: 40px;">{letter}</span>'
                html += "</td>"
            else:
                # Black cell
                html += '<td style="width: 40px; height: 40px; border: 1px solid #d1d5db; background-color: #1f2937; padding: 0;"></td>'
        html += "</tr>"
    html += "</table></div>"
    return html

# =====================================================================
# MENU UTAMA SIDEBAR
# =====================================================================
st.sidebar.title("Menu LKPD")

if st.session_state.mulai_belajar:
    st.sidebar.markdown(f"""
    <div style="background-color: rgba(99, 102, 241, 0.1); padding: 10px; border-radius: 8px; border-left: 4px solid #6366f1;">
      👤 Siswa: <b>{st.session_state.nama_siswa}</b><br>
      🏫 Kelas: <b>{st.session_state.kelas_siswa}</b><br>
      🔢 Absen: <b>{st.session_state.absen_siswa}</b>
    </div>
    """, unsafe_allow_html=True)
else:
    st.sidebar.info("💡 Isi identitas siswa di Halaman 2 untuk membuka misi belajar.")

st.sidebar.divider()

menu_options = [
    "1. 🏠 Beranda LKPD",
    "2. 📝 Identitas Siswa",
    "3. 🎯 Tujuan Pembelajaran",
    "4. 📖 Materi Pembelajaran",
    "5. ⚡ Aktivitas Interaktif",
    "6. 🎮 Game & Tantangan",
    "7. 🧩 Puzzle / TTS",
    "8. ✍️ Kuis Akhir",
    "9. 🏆 Hasil dan Skor",
    "10. 💬 Refleksi"
]

# Ambil index halaman aktif saat ini
try:
    active_idx = menu_options.index(st.session_state.current_page)
except ValueError:
    active_idx = 0

selected_page = st.sidebar.radio("Pilih Misi Belajar:", menu_options, index=active_idx)

# Update state halaman aktif
if selected_page != st.session_state.current_page:
    st.session_state.current_page = selected_page
    st.rerun()

st.sidebar.divider()
if st.sidebar.button("🔄 Reset Semua Data"):
    reset_all_data()
    st.success("Data berhasil direset!")
    st.rerun()

# =====================================================================
# GLOBAL PROGRESS BAR
# =====================================================================
if st.session_state.mulai_belajar:
    pts = 0
    if st.session_state.mulai_belajar: pts += 1
    if st.session_state.tujuan_read: pts += 1
    if st.session_state.materi_read: pts += 1
    if st.session_state.act1_checked: pts += 1
    if st.session_state.act3_checked: pts += 1
    if st.session_state.g1_checked and st.session_state.g2_checked: pts += 1
    if st.session_state.g3_checked and st.session_state.g4_checked: pts += 1
    if st.session_state.g5_checked and st.session_state.g6_checked: pts += 1
    if st.session_state.g7_checked and st.session_state.g8_checked: pts += 1
    if st.session_state.quiz_checked: pts += 1
    
    st.write(f"🎮 **Misi Belajar Selesai: {pts * 10}%** ({pts} dari 10 aktivitas selesai)")
    st.progress(pts / 10.0)
    st.divider()

# =====================================================================
# PROTEKSI PENGISIAN IDENTITAS
# =====================================================================
# Jika siswa mencoba membuka halaman 3-10 tanpa login
locked_pages = menu_options[2:]
if not st.session_state.mulai_belajar and st.session_state.current_page in locked_pages:
    st.markdown('<h1 class="title-gradient">🔒 Misi Belajar Terkunci</h1>', unsafe_allow_html=True)
    st.warning("⚠️ Anda harus mengisi Identitas Siswa terlebih dahulu sebelum memulai petualangan belajar!")
    if st.button("Isi Identitas Sekarang"):
        st.session_state.current_page = "2. 📝 Identitas Siswa"
        st.rerun()
    st.stop()

# =====================================================================
# KONTEN HALAMAN-HALAMAN
# =====================================================================

# --- HALAMAN 1: BERANDA LKPD ---
if st.session_state.current_page == "1. 🏠 Beranda LKPD":
    st.markdown('<h1 class="title-gradient">SMAN 10 BANDAR LAMPUNG<br>LKPD INFORMATIKA KELAS 10<br>DATA, INFORMASI, DAN VALIDASINYA</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle-text">Petualangan Belajar Interaktif untuk Kelas X SMA</p>', unsafe_allow_html=True)
    
    st.info("""
    Selamat datang di **Misi Belajar LKPD Interaktif**! 🚀
    Di sini Anda tidak hanya membaca materi atau menjawab kuis membosankan. Anda akan menjadi seorang **Detektif Informasi** 🕵️‍♂️ yang akan memecahkan teka-teki, menjalankan kode Python untuk web scraping, menguji data lewat CRAAP Test, dan mengisi teka-teki silang!
    """)
    
    st.markdown("### 🏆 Fitur Utama Petualangan Belajar Ini:")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        * 🎯 **Misi Bertahap**: Belajar terstruktur dari pengenalan materi hingga evaluasi kuis akhir.
        * 🎮 **8 Mini Game Interaktif**: Tebak operator pencarian, menjodohkan konsep, hingga detektif kebenaran hoaks.
        * 💻 **Simulasi Scraping Web**: Jalankan perintah BeautifulSoup Python secara langsung.
        """)
    with col2:
        st.markdown("""
        * 🧩 **Teka-Teki Silang (TTS)**: Grid interaktif yang mendeteksi jawaban benar secara langsung.
        * 📊 **Laporan Excel & Teks**: Hasil belajar Anda dapat diunduh dalam file Excel profesional untuk diserahkan ke Guru!
        * 📈 **Progress Bar**: Pantau persentase penyelesaian misi belajarmu secara real-time.
        """)
        
    st.divider()
    if st.button("Lanjut ke Identitas Siswa ➡️"):
        st.session_state.current_page = "2. 📝 Identitas Siswa"
        st.rerun()

# --- HALAMAN 2: IDENTITAS SISWA ---
elif st.session_state.current_page == "2. 📝 Identitas Siswa":
    st.markdown('<h1 class="title-gradient">📝 Identitas Siswa</h1>', unsafe_allow_html=True)
    st.write("Silakan isi data identitas Anda secara lengkap untuk memulai perjalanan belajar dan mencatat skor.")
    
    st.divider()
    
    nama_input = st.text_input("Nama Lengkap:", value=st.session_state.nama_siswa, placeholder="Contoh: Budi Santoso")
    kelas_input = st.text_input("Kelas:", value=st.session_state.kelas_siswa, placeholder="Contoh: X MIPA 1")
    absen_input = st.text_input("Nomor Absen:", value=st.session_state.absen_siswa, placeholder="Contoh: 07")
    
    st.divider()
    
    if st.button("🚀 Mulai Petualangan Belajar!"):
        if nama_input.strip() and kelas_input.strip() and absen_input.strip():
            st.session_state.nama_siswa = nama_input
            st.session_state.kelas_siswa = kelas_input
            st.session_state.absen_siswa = absen_input
            st.session_state.mulai_belajar = True
            st.session_state.current_page = "3. 🎯 Tujuan Pembelajaran"
            st.success("Identitas tersimpan! Selamat memulai misi belajar.")
            st.rerun()
        else:
            st.warning("⚠️ Mohon isi semua kolom identitas terlebih dahulu sebelum melanjutkan!")

# --- HALAMAN 3: TUJUAN PEMBELAJARAN ---
elif st.session_state.current_page == "3. 🎯 Tujuan Pembelajaran":
    st.session_state.tujuan_read = True
    st.markdown('<h1 class="title-gradient">🎯 Tujuan Pembelajaran</h1>', unsafe_allow_html=True)
    st.write("Sebelum memulai misi, ketahui terlebih dahulu kompetensi yang harus Anda kuasai:")
    
    st.markdown("""
    <div style="background-color: rgba(59, 130, 246, 0.05); padding: 25px; border-radius: 12px; border: 1px solid rgba(59, 130, 246, 0.2);">
      <h3>Setelah menyelesaikan LKPD ini, Anda diharapkan mampu:</h3>
      <ol>
        <li><b>Membedakan Data & Informasi</b>: Menjelaskan perbedaan mendasar antara data mentah dengan informasi yang sudah diolah.</li>
        <li><b>Menguasai Pencarian Informasi</b>: Menerapkan kata kunci pencarian yang efektif serta menggunakan operator pencarian (search operators) Google dengan benar.</li>
        <li><b>Memahami Web Scraping</b>: Memahami konsep pengumpulan data otomatis dari website menggunakan bahasa Python (BeautifulSoup) secara teoritis dan praktis di Google Colab.</li>
        <li><b>Memvalidasi Informasi</b>: Melakukan pengujian kredibilitas dan kebenaran suatu informasi internet menggunakan metode <b>CRAAP Test</b> dan teknik <b>Lateral Reading</b> untuk menangkal berita bohong (hoaks).</li>
      </ol>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    if st.button("Mulai Belajar Materi 📖"):
        st.session_state.current_page = "4. 📖 Materi Pembelajaran"
        st.rerun()

# --- HALAMAN 4: MATERI PEMBELAJARAN ---
elif st.session_state.current_page == "4. 📖 Materi Pembelajaran":
    st.session_state.materi_read = True
    st.markdown('<h1 class="title-gradient">📖 Ringkasan Materi Pembelajaran</h1>', unsafe_allow_html=True)
    st.write("Bacalah rangkuman materi berikut dengan saksama. Materi ini akan menjadi bekal untuk menjawab semua game dan kuis nanti.")
    
    # KATEGORI 1: Data dan Informasi
    with st.expander("📊 Bagian 1: Data vs Informasi"):
        st.markdown('<h4 class="materi-header">1. Data dan Informasi</h4>', unsafe_allow_html=True)
        st.write("""
        * **Data**: Catatan atas kumpulan fakta mentah, angka, tulisan, atau gambar yang belum diproses dan belum memiliki makna jelas. Contoh: Angka `38`, `37`, `39`.
        * **Informasi**: Data yang telah diolah, dikelompokkan, dan dikaitkan dengan konteks tertentu sehingga memiliki kegunaan dan arti bagi penerimanya. Contoh: `"Tinggi rata-rata siswa kelas X adalah 165 cm."`
        """)
    
    # KATEGORI 2: Pencarian Informasi & Operator
    with st.expander("🔎 Bagian 2: Teknik Pencarian Informasi Google"):
        st.markdown('<h4 class="materi-header">2. Pencarian Informasi & Kata Kunci</h4>', unsafe_allow_html=True)
        st.write("""
        Pencarian informasi yang sukses di internet sangat bergantung pada ketepatan penentuan **kata kunci (keywords)**. Hindari kalimat terlalu panjang dan fokus pada subjek inti.
        """)
        
        st.markdown('<h4 class="materi-header">3. Operator Pencarian (Search Operators)</h4>', unsafe_allow_html=True)
        st.write("""
        Gunakan operator khusus di mesin pencari Google untuk hasil yang spesifik:
        * **`site:`**: Membatasi pencarian hanya pada situs web atau domain tertentu saja. Contoh: `penerimaan siswa baru site:kemdikbud.go.id`
        * **`filetype:`**: Mencari hasil pencarian dalam bentuk format file tertentu (misal: PDF, DOC, XLS). Contoh: `materi informatika kelas x filetype:pdf`
        * **`intitle:`**: Menyaring hasil agar kata kunci pencarian wajib ada pada judul (title) artikel web. Contoh: `intitle:kecerdasan buatan`
        * **Tanda Kutip Double (`"..."`)**: Mencari frasa kata secara berurutan dan persis sama tanpa variasi. Contoh: `"belajar pemrograman python"`
        * **Minus (`-`)**: Mengecualikan kata tertentu agar tidak muncul dalam pencarian. Contoh: `jaguar -mobil` (mencari hewan jaguar, bukan merek mobil).
        """)
        
        st.markdown('<h4 class="materi-header">4. Pencarian Gambar & Suara</h4>', unsafe_allow_html=True)
        st.write("""
        * **Pencarian Gambar**: Kita bisa mengunggah file gambar untuk mencari gambar serupa atau sumber aslinya (Reverse Image Search).
        * **Pencarian Suara (Voice Search)**: Menggunakan input suara manusia sebagai masukan query ke search engine.
        """)
        
    # KATEGORI 3: Web Scraping
    with st.expander("💻 Bagian 3: Koleksi Data & Web Scraping"):
        st.markdown('<h4 class="materi-header">5. Pengumpulan Data: Manual vs Otomatis</h4>', unsafe_allow_html=True)
        st.write("""
        * **Pengumpulan Manual**: Menyalin teks/tabel satu per satu secara manual dari halaman web (lemah efisiensi jika data berjumlah besar).
        * **Web Scraping**: Teknik mengekstrak data dari dokumen halaman web secara otomatis menggunakan program komputer.
        """)
        
        st.markdown('<h4 class="materi-header">6. Python untuk Web Scraping</h4>', unsafe_allow_html=True)
        st.write("""
        Dua pustaka (library) utama Python yang sangat populer untuk web scraping dasar:
        * **`requests`**: Mengirim permintaan HTTP (seperti GET request) ke server web untuk mengambil dokumen HTML mentah dari alamat URL.
        * **`BeautifulSoup`** (dari paket `bs4`): Menganalisis dan mengurai (parsing) struktur tag HTML agar kita bisa mencari, mengambil, dan menyaring bagian teks tertentu berdasarkan nama tag (seperti `<h1>`, `<li>`, `<a>`) dan kelasnya (`class_`).
        """)
        
        st.markdown('<h4 class="materi-header">7. Google Colab</h4>', unsafe_allow_html=True)
        st.write("""
        Platform berbasis cloud gratis milik Google untuk menulis dan menjalankan kode program Python secara interaktif di browser internet tanpa perlu menginstal aplikasi Python di komputer lokal.
        """)
        
    # KATEGORI 4: Validasi Informasi
    with st.expander("🛡️ Bagian 4: Validitas Informasi & Periksa Fakta"):
        st.markdown('<h4 class="materi-header">8. Validitas Informasi & Kredibilitas Sumber</h4>', unsafe_allow_html=True)
        st.write("""
        Tidak semua tulisan di internet bernilai benar. Banyak berita bohong (hoaks). Validasi informasi bertujuan menyaring data yang kita dapatkan agar bebas dari kesalahan atau bias.
        """)
        
        st.markdown('<h4 class="materi-header">9. CRAAP Test</h4>', unsafe_allow_html=True)
        st.write("""
        Metode evaluasi sumber informasi untuk menguji keandalan artikel dengan aspek:
        * **C (Currency)**: Kebaruan aspek waktu informasi (kapan diterbitkan/diperbarui?).
        * **R (Relevance)**: Kesesuaian konten dengan kebutuhan pencarian informasi Anda.
        * **A (Authority)**: Kejelasan siapa pembuat/penulisnya dan apa keahlian mereka.
        * **A (Accuracy)**: Kebenaran data, didukung referensi ilmiah, serta bebas kesalahan tata bahasa.
        * **P (Purpose)**: Tujuan pembuatan informasi (edukasi, opini, jualan, atau propaganda?).
        """)
        
        st.markdown('<h4 class="materi-header">10. Lateral Reading (Membaca Lateral)</h4>', unsafe_allow_html=True)
        st.write("""
        Teknik memverifikasi kredibilitas situs web dengan cara keluar dari halaman tersebut (membuka tab baru di browser) lalu menelusuri pendapat situs/berita independen tepercaya tentang situs yang sedang kita periksa.
        """)
        
    st.divider()
    if st.button("Mulai Aktivitas Interaktif ⚡"):
        st.session_state.current_page = "5. ⚡ Aktivitas Interaktif"
        st.rerun()

# --- HALAMAN 5: AKTIVITAS INTERAKTIF ---
elif st.session_state.current_page == "5. ⚡ Aktivitas Interaktif":
    st.markdown('<h1 class="title-gradient">⚡ Aktivitas Interaktif</h1>', unsafe_allow_html=True)
    st.write("Selesaikan 3 aktivitas berikut yang mempertahankan fitur orisinal proyek LKPD Anda:")
    
    tab_act1, tab_act2, tab_act3 = st.tabs([
        "📊 Aktivitas 1: Data vs Informasi",
        "💻 Aktivitas 2: Simulasi Web Scraping",
        "🛡️ Aktivitas 3: Validasi Informasi"
    ])
    
    # --- AKTIVITAS 1 ---
    with tab_act1:
        st.subheader("💡 Mengubah Data Menjadi Informasi")
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
        st.write("**Uji Pemahaman Mandiri:**")
        
        opsi_soal1 = ["Pilih jawaban Anda", "Data", "Informasi"]
        idx_soal1 = opsi_soal1.index(st.session_state.j_soal1) if st.session_state.j_soal1 in opsi_soal1 else 0
        soal1 = st.radio(
            "1. Catatan daftar nilai ujian matematika murid kelas 10 (misal: 80, 75, 90, 85) yang belum dirata-rata disebut sebagai...",
            opsi_soal1,
            index=idx_soal1,
            key="act1_q1"
        )
        st.session_state.j_soal1 = soal1
        
        opsi_soal2 = ["Pilih jawaban Anda", "Data", "Informasi"]
        idx_soal2 = opsi_soal2.index(st.session_state.j_soal2) if st.session_state.j_soal2 in opsi_soal2 else 0
        soal2 = st.radio(
            "2. Laporan grafik peningkatan rata-rata nilai matematika kelas 10 dari bulan lalu ke bulan ini disebut sebagai...",
            opsi_soal2,
            index=idx_soal2,
            key="act1_q2"
        )
        st.session_state.j_soal2 = soal2
        
        if st.button("Cek Jawaban Aktivitas 1", key="btn_check_act1"):
            if soal1 == "Pilih jawaban Anda" or soal2 == "Pilih jawaban Anda":
                st.warning("⚠️ Silakan pilih jawaban untuk semua soal terlebih dahulu!")
            else:
                st.session_state.act1_checked = True
                st.success("Jawaban Aktivitas 1 tersimpan!")
                st.rerun()
                
        if st.session_state.act1_checked:
            st.markdown("#### **Hasil Koreksi:**")
            if soal1 == "Data":
                st.success("🎉 Soal 1: BENAR! Catatan nilai mentah yang belum diolah adalah **Data**.")
            else:
                st.error("❌ Soal 1: BELUM TEPAT. Catatan nilai mentah belum diolah, jadi itu adalah **Data**.")
                
            if soal2 == "Informasi":
                st.success("🎉 Soal 2: BENAR! Laporan grafik yang sudah diolah dan memiliki arti adalah **Informasi**.")
            else:
                st.error("❌ Soal 2: BELUM TEPAT. Laporan grafik sudah diproses dan memiliki makna bagi pembacanya, jadi itu adalah **Informasi**.")
                
    # --- AKTIVITAS 2 ---
    with tab_act2:
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

        if st.button("Jalankan Scraping", key="btn_run_scraping"):
            if not bs4_installed:
                st.error("❌ Pustaka `beautifulsoup4` belum terinstall di komputer Anda.")
                st.info("Silakan jalankan perintah berikut di terminal Anda untuk menginstalnya: \n`pip install beautifulsoup4`")
            else:
                soup = BeautifulSoup(html_mentah, 'html.parser')
                daftar = soup.find_all('li', class_='berita')
                
                st.success("🎉 Scraping Berhasil!")
                st.write("Berikut data bersih yang berhasil diekstrak oleh program Python:")
                for index, item in enumerate(daftar, 1):
                    st.info(f"**Data {index}:** {item.text}")

    # --- AKTIVITAS 3 ---
    with tab_act3:
        st.subheader("Studi Kasus: Membaca Pesan Berantai")
        st.warning("""
        **Pesan Berantai WhatsApp:**
        "INFO PENTING! Mulai besok pagi, seluruh jaringan internet seluler di Indonesia akan dimatikan total selama 3 hari karena adanya badai matahari ekstrem yang membakar satelit bumi. Harap sebarkan pesan ini ke grup keluarga dan teman dekat Anda agar semua bersiap-siap membeli bahan makanan!"
        """)

        st.write("Mari kita analisis pesan di atas dengan answering kuis di bawah ini:")

        opsi_soal3 = ["Pilih jawaban Anda", "Ya, ada sumber resminya", "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi"]
        idx_soal3 = opsi_soal3.index(st.session_state.j_soal3) if st.session_state.j_soal3 in opsi_soal3 else 0
        soal3 = st.radio(
            "1. Apakah pesan di atas menyebutkan tanggal kejadian spesifik dan mencantumkan sumber lembaga resmi (seperti BMKG atau NASA)?",
            opsi_soal3,
            index=idx_soal3,
            key="act3_q1"
        )
        st.session_state.j_soal3 = soal3

        opsi_soal4 = ["Pilih jawaban Anda", "Informasi Valid (Fakta)", "Informasi Palsu (Hoaks)"]
        idx_soal4 = opsi_soal4.index(st.session_state.j_soal4) if st.session_state.j_soal4 in opsi_soal4 else 0
        soal4 = st.radio(
            "2. Berdasarkan analisis Anda, pesan berantai tersebut tergolong sebagai...",
            opsi_soal4,
            index=idx_soal4,
            key="act3_q2"
        )
        st.session_state.j_soal4 = soal4

        if st.button("Verifikasi Berita", key="btn_check_act3"):
            if soal3 == "Pilih jawaban Anda" or soal4 == "Pilih jawaban Anda":
                st.warning("⚠️ Silakan pilih jawaban untuk semua soal terlebih dahulu!")
            else:
                st.session_state.act3_checked = True
                st.success("Jawaban Aktivitas 3 tersimpan!")
                st.rerun()

        if st.session_state.act3_checked:
            st.markdown("#### **Hasil Verifikasi:**")
            if soal3 == "Tidak, hanya menyuruh menyebarkan tanpa sumber resmi":
                st.success("🎉 Soal 1: BENAR! Pesan hoaks biasanya tidak memiliki tanggal jelas dan tidak menyebutkan sumber resmi.")
            else:
                st.error("❌ Soal 1: BELUM TEPAT. Perhatikan baik-baik, pesan tersebut tidak menyebutkan lembaga resmi mana pun sebagai sumbernya.")

            if soal4 == "Informasi Palsu (Hoaks)":
                st.success("🎉 Soal 2: BENAR! Kabar tersebut adalah kabar bohong (hoaks) yang dibuat untuk menimbulkan kepanikan.")
            else:
                st.error("❌ Soal 2: BELUM TEPAT. Kabar pemutusan internet nasional akibat badai matahari adalah hoaks.")

    st.divider()
    if st.button("Lanjut ke Game & Tantangan Belajar 🎮"):
        st.session_state.current_page = "6. 🎮 Game & Tantangan"
        st.rerun()

# --- HALAMAN 6: GAME & TANTANGAN (Game 1, 2, 3, 4) ---
elif st.session_state.current_page == "6. 🎮 Game & Tantangan":
    st.markdown('<h1 class="title-gradient">🎮 Game & Tantangan</h1>', unsafe_allow_html=True)
    st.write("Selesaikan 4 game seru di bawah ini dengan membuka masing-masing tab:")
    
    t1, t2, t3, t4 = st.tabs([
        "🎮 Game 1: Tebak Operator",
        "🤝 Game 2: Pasangkan Konsep",
        "🔎 Game 3: Detektif Informasi",
        "🛡️ Game 4: CRAAP Test"
    ])
    
    # --- GAME 1: TEBAK OPERATOR ---
    with t1:
        st.subheader("🎮 Game 1: Tebak Operator Pencarian")
        st.write("Pilihlah operator pencarian Google yang tepat sesuai deskripsi fungsinya!")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Pikirkan tentang operator mana yang membatasi pencarian pada domain tertentu (seperti .gov, .edu), judul artikel, pencarian frasa persis, atau pengecualian kata.")
            
        g1_q1 = st.radio(
            "1. Saya digunakan untuk mencari informasi terbatas dari website atau domain tertentu saja.",
            ["Pilih...", "intitle:", "site:", "inurl:", "intext:"],
            index=st.session_state.g1_answers[0] if st.session_state.g1_answers[0] is not None else 0,
            key="g1_q1"
        )
        g1_q2 = st.radio(
            "2. Saya digunakan untuk mencari dokumen dengan format file tertentu seperti PDF, XLS, atau PPT.",
            ["Pilih...", "filetype:", "intitle:", "site:", "inurl:"],
            index=st.session_state.g1_answers[1] if st.session_state.g1_answers[1] is not None else 0,
            key="g1_q2"
        )
        g1_q3 = st.radio(
            "3. Saya digunakan untuk mencari frasa kata yang sama persis dan berurutan.",
            ["Pilih...", "tanda kutip (\"\")", "OR", "AND", "minus (-)"],
            index=st.session_state.g1_answers[2] if st.session_state.g1_answers[2] is not None else 0,
            key="g1_q3"
        )
        g1_q4 = st.radio(
            "4. Saya digunakan untuk mengecualikan kata tertentu agar tidak ikut tampil pada hasil pencarian.",
            ["Pilih...", "minus (-)", "OR", "site:", "inurl:"],
            index=st.session_state.g1_answers[3] if st.session_state.g1_answers[3] is not None else 0,
            key="g1_q4"
        )
        g1_q5 = st.radio(
            "5. Saya digunakan untuk mencari halaman web yang mengandung kata tertentu di dalam judulnya.",
            ["Pilih...", "intitle:", "site:", "inurl:", "filetype:"],
            index=st.session_state.g1_answers[4] if st.session_state.g1_answers[4] is not None else 0,
            key="g1_q5"
        )
        
        if st.button("Periksa Jawaban Game 1"):
            choices = [g1_q1, g1_q2, g1_q3, g1_q4, g1_q5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua pertanyaan terlebih dahulu!")
            else:
                st.session_state.g1_answers = [
                    ["Pilih...", "intitle:", "site:", "inurl:", "intext:"].index(g1_q1),
                    ["Pilih...", "filetype:", "intitle:", "site:", "inurl:"].index(g1_q2),
                    ["Pilih...", "tanda kutip (\"\")", "OR", "AND", "minus (-)"].index(g1_q3),
                    ["Pilih...", "minus (-)", "OR", "site:", "inurl:"].index(g1_q4),
                    ["Pilih...", "intitle:", "site:", "inurl:", "filetype:"].index(g1_q5)
                ]
                st.session_state.g1_checked = True
                st.success("Jawaban Game 1 disimpan!")
                st.rerun()
                
        if st.session_state.g1_checked:
            st.markdown("#### **Evaluasi Jawaban Game 1:**")
            feedback_g1 = [
                ("site:", "site: membatasi pencarian hanya pada domain atau situs tertentu."),
                ("filetype:", "filetype: menyaring dokumen berdasarkan ekstensi file."),
                ("tanda kutip (\"\")", "Tanda kutip ganda mencari kecocokan kata persis."),
                ("minus (-)", "Operator minus mengecualikan hasil kata tertentu."),
                ("intitle:", "intitle: mencari kata tertentu dalam judul artikel.")
            ]
            ans_states = [g1_q1, g1_q2, g1_q3, g1_q4, g1_q5]
            for i, val in enumerate(ans_states):
                correct_text = feedback_g1[i][0]
                expl = feedback_g1[i][1]
                if val == correct_text:
                    st.success(f"✔️ Soal {i+1}: BENAR! {expl}")
                else:
                    st.error(f"❌ Soal {i+1}: BELUM TEPAT. Seharusnya **{correct_text}**. ({expl})")

    # --- GAME 2: PASANGKAN KONSEP ---
    with t2:
        st.subheader("🤝 Game 2: Pasangkan Konsep Utama")
        st.write("Pasangkan konsep di kolom kiri dengan fungsinya yang benar di kolom kanan!")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Pikirkan perbedaan bahan mentah (Data) vs hasil olahan (Informasi). Ingat juga fungsi web scraping, CRAAP test untuk menguji kebenaran, serta membaca lateral.")
            
        opsi_def = [
            "Pilih...",
            "Fakta mentah yang belum diolah dan belum memiliki arti bagi orang lain.",
            "Hasil pengolahan data yang sudah memiliki arti, konteks, dan kegunaan.",
            "Proses pengumpulan/pengambilan data secara otomatis dari halaman web menggunakan kode.",
            "Metode evaluasi untuk menguji kredibilitas informasi internet (Currency, Relevance, Authority, Accuracy, Purpose).",
            "Memverifikasi informasi dengan membuka tab browser baru untuk mencari tahu apa kata sumber lain."
        ]
        
        g2_c1 = st.selectbox("1. Konsep: DATA dipasangkan dengan...", opsi_def, index=st.session_state.g2_answers[0] if st.session_state.g2_answers[0] is not None else 0, key="g2_c1")
        g2_c2 = st.selectbox("2. Konsep: INFORMASI dipasangkan dengan...", opsi_def, index=st.session_state.g2_answers[1] if st.session_state.g2_answers[1] is not None else 0, key="g2_c2")
        g2_c3 = st.selectbox("3. Konsep: WEB SCRAPING dipasangkan dengan...", opsi_def, index=st.session_state.g2_answers[2] if st.session_state.g2_answers[2] is not None else 0, key="g2_c3")
        g2_c4 = st.selectbox("4. Konsep: CRAAP TEST dipasangkan dengan...", opsi_def, index=st.session_state.g2_answers[3] if st.session_state.g2_answers[3] is not None else 0, key="g2_c4")
        g2_c5 = st.selectbox("5. Konsep: LATERAL READING dipasangkan dengan...", opsi_def, index=st.session_state.g2_answers[4] if st.session_state.g2_answers[4] is not None else 0, key="g2_c5")
        
        if st.button("Periksa Jawaban Game 2"):
            choices = [g2_c1, g2_c2, g2_c3, g2_c4, g2_c5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua pertanyaan terlebih dahulu!")
            else:
                st.session_state.g2_answers = [opsi_def.index(c) for c in choices]
                st.session_state.g2_checked = True
                st.success("Jawaban Game 2 disimpan!")
                st.rerun()
                
        if st.session_state.g2_checked:
            st.markdown("#### **Evaluasi Jawaban Game 2:**")
            correct_keys = [1, 2, 3, 4, 5]
            concepts = ["DATA", "INFORMASI", "WEB SCRAPING", "CRAAP TEST", "LATERAL READING"]
            for idx, ans_idx in enumerate(st.session_state.g2_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ {concepts[idx]}: BENAR! Pasangannya tepat.")
                else:
                    st.error(f"❌ {concepts[idx]}: BELUM TEPAT. Anda memilih: '{opsi_def[ans_idx]}'. Seharusnya dipasangkan dengan: '{opsi_def[correct_keys[idx]]}'")

    # --- GAME 3: DETEKTIF INFORMASI ---
    with t3:
        st.subheader("🔎 Game 3: Detektif Informasi (Pemeriksaan Hoaks)")
        st.write("Sebagai Detektif Informasi, pilihlah tindakan terbaik untuk menangani kasus penyebaran informasi berikut!")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Ingat bahwa pesan hoaks biasanya tidak mencantumkan lembaga resmi, manipulatif, bias, serta menggunakan bahasa kepanikan. Utamakan verifikasi lewat situs tepercaya atau melakukan lateral reading.")
            
        g3_q1 = st.radio(
            "1. Kasus IMEI HP disadap: Menerima pesan berantai 'Ketik *#06# untuk cek penyadapan. Angka ganjil berarti disadap, segera ke polisi.'",
            ["Pilih...", 
             "Menyebarkannya ke semua grup chat agar teman sekelas waspada.", 
             "Mengabaikan karena merupakan hoaks lama tentang manipulasi kode IMEI.", 
             "Panik dan membawa HP ke konter servis terdekat."],
            index=st.session_state.g3_answers[0] if st.session_state.g3_answers[0] is not None else 0,
            key="g3_q1"
        )
        g3_q2 = st.radio(
            "2. Kasus Singkong Beracun: Menonton video viral di TikTok yang mengklaim ada singkong beracun mematikan di sebuah desa tanpa nama lokasi yang jelas.",
            ["Pilih...", 
             "Mengunggah ulang video di status media sosial Anda.", 
             "Mencari klarifikasi resmi dengan kata kunci 'hoaks video singkong beracun' di mesin pencari atau situs TurnBackHoax.id.", 
             "Berhenti makan singkong seumur hidup."],
            index=st.session_state.g3_answers[1] if st.session_state.g3_answers[1] is not None else 0,
            key="g3_q2"
        )
        g3_q3 = st.radio(
            "3. Kasus Daun Pepaya Kanker: Sebuah blog gratisan menulis artikel 'Daun pepaya dapat menyembuhkan penyakit kanker kronis secara instan dalam 1 jam.'",
            ["Pilih...", 
             "Membuka tab baru (lateral reading) untuk memeriksa kebenaran artikel di situs medis terpercaya.", 
             "Percaya karena tulisan terlihat meyakinkan.", 
             "Menimbun daun pepaya untuk dijual."],
            index=st.session_state.g3_answers[2] if st.session_state.g3_answers[2] is not None else 0,
            key="g3_q3"
        )
        g3_q4 = st.radio(
            "4. Kasus Phishing Rekening: Anda mendapat email dari admin-bank-rakyat@gmail.com yang mengatakan rekening Anda dibekukan dan Anda harus mengeklik link yang tersedia.",
            ["Pilih...", 
             "Klik link dan mengisi data kartu ATM.", 
             "Memeriksa pengirim (domain gratisan adalah phishing) dan menghubungi call center resmi bank.", 
             "Mengabaikan saja tanpa melaporkan."],
            index=st.session_state.g3_answers[3] if st.session_state.g3_answers[3] is not None else 0,
            key="g3_q4"
        )
        g3_q5 = st.radio(
            "5. Kasus Referensi Tugas: Anda menemukan artikel sains menarik di web, namun tidak ada nama penulis, tanggal terbit, maupun daftar referensi di dalamnya.",
            ["Pilih...", 
             "Menjadikannya referensi tugas sekolah.", 
             "Meragukan keakuratannya karena tidak memenuhi aspek Authority dan Accuracy.", 
             "Menyalin artikel tersebut."],
            index=st.session_state.g3_answers[4] if st.session_state.g3_answers[4] is not None else 0,
            key="g3_q5"
        )
        
        if st.button("Periksa Jawaban Game 3"):
            choices = [g3_q1, g3_q2, g3_q3, g3_q4, g3_q5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua pertanyaan terlebih dahulu!")
            else:
                st.session_state.g3_answers = []
                for q, opts in [(g3_q1, ["Pilih...", "Menyebarkannya ke semua grup chat agar teman sekelas waspada.", "Mengabaikan karena merupakan hoaks lama tentang manipulasi kode IMEI.", "Panik dan membawa HP ke konter servis terdekat."]),
                                (g3_q2, ["Pilih...", "Mengunggah ulang video di status media sosial Anda.", "Mencari klarifikasi resmi dengan kata kunci 'hoaks video singkong beracun' di mesin pencari atau situs TurnBackHoax.id.", "Berhenti makan singkong seumur hidup."]),
                                (g3_q3, ["Pilih...", "Membuka tab baru (lateral reading) untuk memeriksa kebenaran artikel di situs medis terpercaya.", "Percaya karena tulisan terlihat meyakinkan.", "Menimbun daun pepaya untuk dijual."]),
                                (g3_q4, ["Pilih...", "Klik link dan mengisi data kartu ATM.", "Memeriksa pengirim (domain gratisan adalah phishing) dan menghubungi call center resmi bank.", "Mengabaikan saja tanpa melaporkan."]),
                                (g3_q5, ["Pilih...", "Menjadikannya referensi tugas sekolah.", "Meragukan keakuratannya karena tidak memenuhi aspek Authority dan Accuracy.", "Menyalin artikel tersebut."])]:
                    st.session_state.g3_answers.append(opts.index(q))
                st.session_state.g3_checked = True
                st.success("Jawaban Game 3 disimpan!")
                st.rerun()
                
        if st.session_state.g3_checked:
            st.markdown("#### **Evaluasi Jawaban Game 3:**")
            correct_keys = [2, 2, 1, 2, 2]
            feedback_g3 = [
                "Soal 1: BENAR. Pesan IMEI *#06# adalah hoaks teknis lama. Abaikan saja.",
                "Soal 2: BENAR. Cek kebenaran video ke situs verifikasi fakta resmi seperti TurnBackHoax.id.",
                "Soal 3: BENAR. Buka tab baru untuk membandingkan informasi kesehatan dengan artikel medis berlisensi dokter resmi.",
                "Soal 4: BENAR. Domain email gratisan seperti @gmail.com tidak pernah digunakan oleh admin bank resmi untuk urusan nasabah.",
                "Soal 5: BENAR. Sumber tanpa nama penulis dan referensi memiliki kredibilitas rendah dan tidak layak dijadikan referensi tugas akademik."
            ]
            for idx, ans_idx in enumerate(st.session_state.g3_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ {feedback_g3[idx]}")
                else:
                    st.error(f"❌ Soal {idx+1}: BELUM TEPAT. Pilih tindakan yang mengutamakan verifikasi kebenaran dan kehati-hatian data.")

    # --- GAME 4: CRAAP TEST ---
    with t4:
        st.subheader("🛡️ Game 4: Menerapkan Aspek CRAAP Test")
        st.write("Identifikasi aspek CRAAP Test mana yang menjadi masalah utama pada masing-masing kasus di bawah ini!")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("C = Currency (Waktu/Kebaruan), R = Relevance (Kesesuaian), A = Authority (Otoritas Penulis), A = Accuracy (Akurasi/Referensi), P = Purpose (Tujuan/Bias Penjualan).")
            
        opsi_craap = ["Pilih...", "Currency (Kebaruan)", "Relevance (Relevansi)", "Authority (Otoritas)", "Accuracy (Akurasi)", "Purpose (Tujuan)"]
        
        g4_q1 = st.selectbox("1. Sebuah artikel yang Anda temukan tentang teknologi AI tercanggih ternyata terbit pada tahun 2005 (tidak terupdate).", opsi_craap, index=st.session_state.g4_answers[0] if st.session_state.g4_answers[0] is not None else 0, key="g4_q1")
        g4_q2 = st.selectbox("2. Artikel kesehatan membahas keajaiban obat herbal, ditulis oleh produsen suplemen kesehatan herbal itu sendiri agar laku terjual.", opsi_craap, index=st.session_state.g4_answers[1] if st.session_state.g4_answers[1] is not None else 0, key="g4_q2")
        g4_q3 = st.selectbox("3. Anda mencari tutorial pemrograman Python terbaru, namun hasil pencarian hanya membahas sejarah perkembangan tabung hampa udara 1940-an.", opsi_craap, index=st.session_state.g4_answers[2] if st.session_state.g4_answers[2] is not None else 0, key="g4_q3")
        g4_q4 = st.selectbox("4. Teks medis tentang penanganan virus ditulis oleh seorang blogger anonim yang tidak memiliki latar belakang keilmuan kedokteran.", opsi_craap, index=st.session_state.g4_answers[3] if st.session_state.g4_answers[3] is not None else 0, key="g4_q4")
        g4_q5 = st.selectbox("5. Artikel mengeklaim alien menduduki gedung DPR, namun menyajikan data statistik yang acak-acakan serta tanpa ada tautan sumber data satu pun.", opsi_craap, index=st.session_state.g4_answers[4] if st.session_state.g4_answers[4] is not None else 0, key="g4_q5")
        
        if st.button("Periksa Jawaban Game 4"):
            choices = [g4_q1, g4_q2, g4_q3, g4_q4, g4_q5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua pertanyaan terlebih dahulu!")
            else:
                st.session_state.g4_answers = [opsi_craap.index(c) for c in choices]
                st.session_state.g4_checked = True
                st.success("Jawaban Game 4 disimpan!")
                st.rerun()
                
        if st.session_state.g4_checked:
            st.markdown("#### **Evaluasi Jawaban Game 4:**")
            correct_keys = [1, 5, 2, 3, 4]
            meanings = [
                "Currency (Kebaruan): Informasi usang tidak relevan dengan kebutuhan teknologi terbaru saat ini.",
                "Purpose (Tujuan): Artikel memiliki kepentingan finansial tersembunyi (bias promosi penjualan).",
                "Relevance (Relevansi): Isi materi tidak sesuai dengan kata kunci tutorial Python yang Anda butuhkan.",
                "Authority (Otoritas): Penulis anonim tidak memiliki kualifikasi keahlian di bidang tersebut.",
                "Accuracy (Akurasi): Data tidak logis dan tidak didukung oleh referensi ilmiah yang valid."
            ]
            for idx, ans_idx in enumerate(st.session_state.g4_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ Soal {idx+1}: BENAR! {meanings[idx]}")
                else:
                    st.error(f"❌ Soal {idx+1}: BELUM TEPAT. Anda memilih: **{opsi_craap[ans_idx]}**. Seharusnya: **{opsi_craap[correct_keys[idx]]}**.")

    st.divider()
    if st.button("Lanjut ke Puzzle & Teka-Teki Silang 🧩"):
        st.session_state.current_page = "7. 🧩 Puzzle / TTS"
        st.rerun()

# --- HALAMAN 7: PUZZLE & TTS (Game 5, 6, 7, 8) ---
elif st.session_state.current_page == "7. 🧩 Puzzle / TTS":
    st.markdown('<h1 class="title-gradient">🧩 Puzzle & Teka-Teki Silang</h1>', unsafe_allow_html=True)
    st.write("Selesaikan tantangan logika berpikir di bawah ini dengan membuka masing-masing tab:")
    
    t5, t6, t7, t8 = st.tabs([
        "🧩 Game 5: Urutan Web Scraping",
        "💻 Game 6: Tebak Kode Python",
        "💬 Game 7: Teka-Teki Silang (TTS)",
        "💼 Game 8: Kasus Kehidupan Siswa"
    ])
    
    # --- GAME 5: URUTAN WEB SCRAPING ---
    with t5:
        st.subheader("🧩 Game 5: Urutkan Proses Web Scraping")
        st.write("Tentukan langkah yang benar (Langkah 1 s.d. Langkah 4) untuk alur kerja web scraping:")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Alur program web scraping: Tentukan URL -> Kirim HTTP Request ke URL -> Analisis HTML (BeautifulSoup) -> Tampilkan/simpan data bersih.")
            
        opsi_langkah = ["Pilih...", "Langkah 1", "Langkah 2", "Langkah 3", "Langkah 4"]
        
        g5_a = st.selectbox("Langkah A: Mengolah/parsing data HTML mentah menggunakan BeautifulSoup.", opsi_langkah, index=st.session_state.g5_answers[0] if st.session_state.g5_answers[0] is not None else 0, key="g5_a")
        g5_b = st.selectbox("Langkah B: Menentukan URL alamat website target yang ingin diambil datanya.", opsi_langkah, index=st.session_state.g5_answers[1] if st.session_state.g5_answers[1] is not None else 0, key="g5_b")
        g5_c = st.selectbox("Langkah C: Menampilkan data hasil scraping atau menyimpannya dalam format Excel.", opsi_langkah, index=st.session_state.g5_answers[2] if st.session_state.g5_answers[2] is not None else 0, key="g5_c")
        g5_d = st.selectbox("Langkah D: Mengirim GET request untuk mengunduh konten kode HTML halaman web.", opsi_langkah, index=st.session_state.g5_answers[3] if st.session_state.g5_answers[3] is not None else 0, key="g5_d")
        
        if st.button("Periksa Urutan Langkah"):
            choices = [g5_a, g5_b, g5_c, g5_d]
            if "Pilih..." in choices:
                st.warning("⚠️ Tentukan nomor urutan untuk seluruh langkah!")
            elif len(set(choices)) < 4:
                st.warning("⚠️ Urutan tidak boleh ada yang kembar!")
            else:
                st.session_state.g5_answers = [opsi_langkah.index(c) for c in choices]
                st.session_state.g5_checked = True
                st.success("Susunan langkah web scraping disimpan!")
                st.rerun()
                
        if st.session_state.g5_checked:
            st.markdown("#### **Evaluasi Susunan Langkah:**")
            correct_keys = [3, 1, 4, 2] # A=3, B=1, C=4, D=2
            steps = ["Langkah A (Parsing)", "Langkah B (Menentukan URL)", "Langkah C (Menyimpan)", "Langkah D (GET Request)"]
            for idx, ans_idx in enumerate(st.session_state.g5_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ {steps[idx]}: BENAR (Langkah ke-{correct_keys[idx]})")
                else:
                    st.error(f"❌ {steps[idx]}: BELUM TEPAT. Anda memilih Langkah ke-{ans_idx}. Seharusnya Langkah ke-{correct_keys[idx]}")

    # --- GAME 6: TEBAK KODE PYTHON ---
    with t6:
        st.subheader("💻 Game 6: Tebak Fungsi Kode Scraping Python")
        st.write("Jawablah fungsi dari perintah baris kode Python BeautifulSoup berikut ini:")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Pikirkan library mana yang mengirim request (requests) vs mengurai HTML (BeautifulSoup). Pahami juga perbedaan mencari satu tag pertama (find) vs semua tag (find_all).")
            
        g6_q1 = st.radio(
            "1. Apa fungsi dari baris kode: 'import requests'?",
            ["Pilih...", "Mengurai struktur dokumen HTML website.", "Mengimpor library requests untuk mengirim permintaan panggilan HTTP ke server web.", "Membuat grafik visualisasi batang."],
            index=st.session_state.g6_answers[0] if st.session_state.g6_answers[0] is not None else 0,
            key="g6_q1"
        )
        g6_q2 = st.radio(
            "2. Apa fungsi dari baris kode: 'response = requests.get(url)'?",
            ["Pilih...", "Mengirim permintaan GET ke URL target untuk mengunduh kode HTML halaman web.", "Mengubah file teks menjadi file Excel secara offline.", "Mencari semua tautan hyperlink dalam halaman web."],
            index=st.session_state.g6_answers[1] if st.session_state.g6_answers[1] is not None else 0,
            key="g6_q2"
        )
        g6_q3 = st.radio(
            "3. Apa fungsi dari baris kode: 'soup = BeautifulSoup(html_mentah, \"html.parser\")'?",
            ["Pilih...", "Mencari tautan gambar.", "Mengolah atau mengurai (parsing) HTML mentah agar struktur tag-nya dapat diakses dengan mudah.", "Menghapus tag HTML agar teks hilang."],
            index=st.session_state.g6_answers[2] if st.session_state.g6_answers[2] is not None else 0,
            key="g6_q3"
        )
        g6_q4 = st.radio(
            "4. Apa fungsi dari baris kode: 'title = soup.find(\"h1\").text'?",
            ["Pilih...", "Mencari tag <h1> pertama di halaman HTML lalu mengambil isi teks bersihnya saja.", "Menghapus semua tag h1 di halaman tersebut.", "Mengubah warna tulisan h1 menjadi tebal."],
            index=st.session_state.g6_answers[3] if st.session_state.g6_answers[3] is not None else 0,
            key="g6_q4"
        )
        g6_q5 = st.radio(
            "5. Apa fungsi dari baris kode: 'links = soup.find_all(\"a\")'?",
            ["Pilih...", "Mencari semua tautan/hyperlink (tag <a>) di seluruh halaman HTML tersebut.", "Membuat tautan link baru agar terhubung ke Google.", "Menghitung jumlah kata dalam website."],
            index=st.session_state.g6_answers[4] if st.session_state.g6_answers[4] is not None else 0,
            key="g6_q5"
        )
        
        if st.button("Periksa Jawaban Game 6"):
            choices = [g6_q1, g6_q2, g6_q3, g6_q4, g6_q5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua pertanyaan terlebih dahulu!")
            else:
                st.session_state.g6_answers = []
                for q, opts in [(g6_q1, ["Pilih...", "Mengurai struktur dokumen HTML website.", "Mengimpor library requests untuk mengirim permintaan panggilan HTTP ke server web.", "Membuat grafik visualisasi batang."]),
                                (g6_q2, ["Pilih...", "Mengirim permintaan GET ke URL target untuk mengunduh kode HTML halaman web.", "Mengubah file teks menjadi file Excel secara offline.", "Mencari semua tautan hyperlink dalam halaman web."]),
                                (g6_q3, ["Pilih...", "Mencari tautan gambar.", "Mengolah atau mengurai (parsing) HTML mentah agar struktur tag-nya dapat diakses dengan mudah.", "Menghapus tag HTML agar teks hilang."]),
                                (g6_q4, ["Pilih...", "Mencari tag <h1> pertama di halaman HTML lalu mengambil isi teks bersihnya saja.", "Menghapus semua tag h1 di halaman tersebut.", "Mengubah warna tulisan h1 menjadi tebal."]),
                                (g6_q5, ["Pilih...", "Mencari semua tautan/hyperlink (tag <a>) di seluruh halaman HTML tersebut.", "Membuat tautan link baru agar terhubung ke Google.", "Menghitung jumlah kata dalam website."])]:
                    st.session_state.g6_answers.append(opts.index(q))
                st.session_state.g6_checked = True
                st.success("Jawaban Game 6 disimpan!")
                st.rerun()
                
        if st.session_state.g6_checked:
            st.markdown("#### **Evaluasi Jawaban Game 6:**")
            correct_keys = [2, 1, 2, 1, 1]
            feedback_g6 = [
                "Soal 1: BENAR. import requests digunakan untuk memanggil library pengirim request internet.",
                "Soal 2: BENAR. requests.get() digunakan untuk mendownload konten HTML halaman web target.",
                "Soal 3: BENAR. BeautifulSoup parsing mengonversi string HTML menjadi objek pohon tag agar mudah diekstrak.",
                "Soal 4: BENAR. find('h1') mengambil tag h1 pertama, dan .text mengambil teks di dalamnya.",
                "Soal 5: BENAR. find_all('a') mengekstrak seluruh tag tautan (hyperlink) pada halaman web."
            ]
            for idx, ans_idx in enumerate(st.session_state.g6_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ {feedback_g6[idx]}")
                else:
                    st.error(f"❌ Soal {idx+1}: BELUM TEPAT. Pelajari lagi sintaksis BeautifulSoup dasar Python.")

    # --- GAME 7: TEKA-TEKI SILANG (TTS) ---
    with t7:
        st.subheader("💬 Game 7: Teka-Teki Silang Bab 1")
        st.write("Isilah kolom jawaban untuk memecahkan Teka-Teki Silang di bawah ini! Huruf akan terisi di grid jika jawaban Anda benar.")
        
        st.markdown(render_crossword_html(), unsafe_allow_html=True)
        
        col_clues_1, col_clues_2 = st.columns(2)
        with col_clues_1:
            st.markdown("""
            **➡️ MENDATAR:**
            * **[2]** Fakta mentah yang belum diolah dan belum memiliki arti yang jelas. (4 huruf)
            * **[3]** Bahasa pemrograman yang populer digunakan pemula untuk web scraping. (6 huruf)
            * **[5]** Uji validitas informasi internet (Currency, Relevance, Authority, Accuracy, Purpose). (5 huruf)
            """)
        with col_clues_2:
            st.markdown("""
            **⬇️ MENURUN:**
            * **[1]** Teknik mengambil data dari halaman website secara otomatis menggunakan kode program. (8 huruf)
            * **[4]** Informasi palsu atau berita bohong di internet yang merugikan. (5 huruf)
            """)
            
        with st.expander("💡 Petunjuk Pengerjaan (TTS)"):
            st.info("Jawaban menggunakan huruf kapital semua. Petunjuk Istilah: DATA, PYTHON, CRAAP, SCRAPING, HOAKS.")
            
        ans_tts_1 = st.text_input("1. Menurun (8 huruf):", value=st.session_state.g7_answers[0], placeholder="Tulis jawaban di sini...").strip().upper()
        ans_tts_2 = st.text_input("2. Mendatar (4 huruf):", value=st.session_state.g7_answers[1], placeholder="Tulis jawaban di sini...").strip().upper()
        ans_tts_3 = st.text_input("3. Mendatar (6 huruf):", value=st.session_state.g7_answers[2], placeholder="Tulis jawaban di sini...").strip().upper()
        ans_tts_4 = st.text_input("4. Menurun (5 huruf):", value=st.session_state.g7_answers[3], placeholder="Tulis jawaban di sini...").strip().upper()
        ans_tts_5 = st.text_input("5. Mendatar (5 huruf):", value=st.session_state.g7_answers[4], placeholder="Tulis jawaban di sini...").strip().upper()
        
        if st.button("Periksa Jawaban TTS"):
            st.session_state.g7_answers = [ans_tts_1, ans_tts_2, ans_tts_3, ans_tts_4, ans_tts_5]
            st.session_state.g7_checked = True
            st.success("Jawaban TTS disimpan!")
            st.rerun()
            
        if st.session_state.g7_checked:
            st.markdown("#### **Status Jawaban TTS:**")
            correct_words = ["SCRAPING", "DATA", "PYTHON", "HOAKS", "CRAAP"]
            labels = ["1. Menurun (Scraping)", "2. Mendatar (Data)", "3. Mendatar (Python)", "4. Menurun (Hoaks)", "5. Mendatar (CRAAP)"]
            for idx, word in enumerate(correct_words):
                student_word = st.session_state.g7_answers[idx]
                if student_word == word:
                    st.success(f"🎉 {labels[idx]}: BENAR!")
                else:
                    if student_word == "":
                        st.warning(f"⚠️ {labels[idx]}: Kosong / belum terisi.")
                    else:
                        st.error(f"❌ {labels[idx]}: BELUM TEPAT.")

    # --- GAME 8: TANTANGAN KASUS ---
    with t8:
        st.subheader("💼 Game 8: Tantangan Kasus Dunia Nyata")
        st.write("Selesaikan studi kasus yang sangat dekat dengan kehidupan sehari-hari berikut:")
        
        with st.expander("💡 Petunjuk Pengerjaan"):
            st.info("Gunakan konsep periksa fakta, CRAAP Test, web scraping otomatis, dan perlindungan privasi data pribadi untuk menentukan tindakan terbaik.")
            
        g8_q1 = st.radio(
            "1. Teman sekelas Anda membagikan info di grup WhatsApp: 'Mulai besok pagi, Kemendikbud resmi menghapus Ujian Sekolah untuk semua jenjang kelas X.' Apa langkah Anda?",
            ["Pilih...",
             "Langsung percaya dan meneruskan pesan tersebut ke grup chat keluarga.",
             "Melakukan pencarian literal di Google dengan mengetik kata kunci 'Ujian sekolah dihapus kemdikbud' dan membandingkannya dengan berita di situs resmi Kemendikbud.",
             "Menuntut guru wali kelas Anda untuk mengonfirmasi kebenaran info tersebut secara paksa."],
            index=st.session_state.g8_answers[0] if st.session_state.g8_answers[0] is not None else 0,
            key="g8_q1"
        )
        g8_q2 = st.radio(
            "2. Seorang anggota grup online mengunggah foto kartu pelajar teman-teman sekelas Anda yang berisi data alamat rumah, nomor HP, dan NIK. Bagaimana tanggapan Anda?",
            ["Pilih...",
             "Menganggapnya biasa saja karena data tersebut adalah informasi publik.",
             "Mengingatkannya bahwa data identitas tersebut adalah data pribadi sensitif yang dilindungi undang-undang privasi dan tidak boleh disebarluaskan sembarangan.",
             "Menyalin data tersebut ke catatan pribadi Anda."],
            index=st.session_state.g8_answers[1] if st.session_state.g8_answers[1] is not None else 0,
            key="g8_q2"
        )
        g8_q3 = st.radio(
            "3. Guru meminta Anda merekap data harga 100 jenis produk alat tulis di 5 e-commerce berbeda dalam waktu 1 jam. Metode apa yang paling efisien?",
            ["Pilih...",
             "Menyalin satu per satu harga tersebut secara manual (copy-paste).",
             "Menggunakan teknik Web Scraping otomatis menggunakan Python BeautifulSoup untuk mengekstrak data langsung dalam hitungan menit.",
             "Meminta bantuan seluruh teman sekelas agar dibagi rata."],
            index=st.session_state.g8_answers[2] if st.session_state.g8_answers[2] is not None else 0,
            key="g8_q3"
        )
        g8_q4 = st.radio(
            "4. Anda membaca artikel sains yang mencantumkan kutipan nomor registrasi DOI jurnal ilmiah resmi, namun isi kesimpulan artikel tersebut sangat melenceng dan bias. Apa tindakan kritis Anda?",
            ["Pilih...",
             "Percaya begitu saja karena mencantumkan registrasi DOI ilmiah.",
             "Membuka link DOI artikel ilmiah asli tersebut, membaca abstraknya secara lateral, dan mencocokkan apakah kesimpulan artikel web tersebut benar atau dibelokkan.",
             "Langsung melaporkan website tersebut ke Kominfo."],
            index=st.session_state.g8_answers[3] if st.session_state.g8_answers[3] is not None else 0,
            key="g8_q4"
        )
        g8_q5 = st.radio(
            "5. Anda mengevaluasi artikel kesehatan mata dengan CRAAP Test. Hasilnya: terbit 15 tahun lalu (Currency) dan ditulis sales produk suplemen (Purpose). Apa kesimpulan Anda?",
            ["Pilih...",
             "Artikel tersebut sangat layak dijadikan bahan tugas ilmiah.",
             "Artikel tersebut memiliki kredibilitas sangat rendah untuk medis masa kini karena data usang dan memiliki bias promosi penjualan.",
             "Artikel tersebut harus segera disebarkan ke grup keluarga."],
            index=st.session_state.g8_answers[4] if st.session_state.g8_answers[4] is not None else 0,
            key="g8_q5"
        )
        
        if st.button("Periksa Jawaban Game 8"):
            choices = [g8_q1, g8_q2, g8_q3, g8_q4, g8_q5]
            if "Pilih..." in choices:
                st.warning("⚠️ Jawab semua kasus terlebih dahulu!")
            else:
                st.session_state.g8_answers = []
                for q, opts in [(g8_q1, ["Pilih...", "Langsung percaya dan meneruskan pesan tersebut ke grup chat keluarga.", "Melakukan pencarian literal di Google dengan mengetik kata kunci 'Ujian sekolah dihapus kemdikbud' dan membandingkannya dengan berita di situs resmi Kemendikbud.", "Menuntut guru wali kelas Anda untuk mengonfirmasi kebenaran info tersebut secara paksa."]),
                                (g8_q2, ["Pilih...", "Menganggapnya biasa saja karena data tersebut adalah informasi publik.", "Mengingatkannya bahwa data identitas tersebut adalah data pribadi sensitif yang dilindungi undang-undang privasi dan tidak boleh disebarluaskan sembarangan.", "Menyalin data tersebut ke catatan pribadi Anda."]),
                                (g8_q3, ["Pilih...", "Menyalin satu per satu harga tersebut secara manual (copy-paste).", "Menggunakan teknik Web Scraping otomatis menggunakan Python BeautifulSoup untuk mengekstrak data langsung dalam hitungan menit.", "Meminta bantuan seluruh teman sekelas agar dibagi rata."]),
                                (g8_q4, ["Pilih...", "Percaya begitu saja karena mencantumkan registrasi DOI ilmiah.", "Membuka link DOI artikel ilmiah asli tersebut, membaca abstraknya secara lateral, dan mencocokkan apakah kesimpulan artikel web tersebut benar atau dibelokkan.", "Langsung melaporkan website tersebut ke Kominfo."]),
                                (g8_q5, ["Pilih...", "Artikel tersebut sangat layak dijadikan bahan tugas ilmiah.", "Artikel tersebut memiliki kredibilitas sangat rendah untuk medis masa kini karena data usang dan memiliki bias promosi penjualan.", "Artikel tersebut harus segera disebarkan ke grup keluarga."])]:
                    st.session_state.g8_answers.append(opts.index(q))
                st.session_state.g8_checked = True
                st.success("Jawaban Game 8 disimpan!")
                st.rerun()
                
        if st.session_state.g8_checked:
            st.markdown("#### **Evaluasi Jawaban Game 8:**")
            correct_keys = [2, 2, 2, 2, 2]
            feedback_g8 = [
                "Soal 1: BENAR. Melakukan pencarian literal di Google adalah tindakan cerdas memeriksa kebenaran hoaks di internet.",
                "Soal 2: BENAR. Melindungi data identitas diri teman kelas adalah bagian dari etika privasi informasi digital.",
                "Soal 3: BENAR. Web scraping otomatis sangat efisien dibandingkan menyalin data manual dalam jumlah besar.",
                "Soal 4: BENAR. Melakukan Lateral Reading pada sumber DOI ilmiah asli mencegah manipulasi informasi.",
                "Soal 5: BENAR. Data medis usang dan bias komersial adalah ciri artikel berkredibilitas rendah."
            ]
            for idx, ans_idx in enumerate(st.session_state.g8_answers):
                if ans_idx == correct_keys[idx]:
                    st.success(f"✔️ {feedback_g8[idx]}")
                else:
                    st.error(f"❌ Soal {idx+1}: BELUM TEPAT. Pilih tindakan kritis sesuai prinsip validasi informasi.")

    st.divider()
    if st.button("Lanjut ke Kuis Akhir ✍️"):
        st.session_state.current_page = "8. ✍️ Kuis Akhir"
        st.rerun()

# --- HALAMAN 8: KUIS AKHIR ---
elif st.session_state.current_page == "8. ✍️ Kuis Akhir":
    st.markdown('<h1 class="title-gradient">✍️ Kuis Akhir</h1>', unsafe_allow_html=True)
    st.write("Kerjakan 10 soal evaluasi berikut untuk menguji pemahaman Anda mengenai seluruh isi materi Bab 1!")
    
    st.divider()
    
    for i, q in enumerate(QUIZ_QUESTIONS):
        st.markdown(f"#### {q['question']}")
        opts = ["Pilih jawaban Anda"] + q["options"]
        
        cur_ans = st.session_state.quiz_answers[i]
        default_idx = 0 if cur_ans is None else (cur_ans + 1)
        
        ans = st.radio(
            f"Jawaban Soal {i+1}:",
            opts,
            index=default_idx,
            key=f"quiz_radio_q_{i}",
            label_visibility="collapsed"
        )
        
        if ans == "Pilih jawaban Anda":
            st.session_state.quiz_answers[i] = None
        else:
            st.session_state.quiz_answers[i] = opts.index(ans) - 1
            
        st.write("")
        
    st.divider()
    
    if st.button("Kirim Kuis Akhir"):
        if None in st.session_state.quiz_answers:
            st.warning("⚠️ Mohon lengkapi semua jawaban kuis terlebih dahulu!")
        else:
            st.session_state.quiz_checked = True
            st.success("Kuis Akhir berhasil dikirim! Silakan lihat hasil evaluasi di bawah ini atau buka halaman Rapor.")
            st.rerun()
            
    if st.session_state.quiz_checked:
        st.subheader("📝 Rincian Pembahasan Kuis Akhir:")
        total_quiz_correct = 0
        for i, q in enumerate(QUIZ_QUESTIONS):
            student_ans = st.session_state.quiz_answers[i]
            correct_ans = q["answer"]
            explanation = q["explanation"]
            
            if student_ans == correct_ans:
                total_quiz_correct += 1
                st.success(f"✔️ **Soal {i+1}: BENAR!** \n\n {explanation}")
            else:
                wrong_opt_text = q["options"][student_ans] if student_ans is not None else "(Kosong)"
                correct_opt_text = q["options"][correct_ans]
                st.error(f"❌ **Soal {i+1}: BELUM TEPAT.** \n\n Anda memilih: *{wrong_opt_text}* \n\n Seharusnya: *{correct_opt_text}* \n\n {explanation}")
                
        st.info(f"🏆 Skor Kuis Akhir Anda: **{total_quiz_correct * 10} / 100**")

    st.divider()
    if st.button("Lanjut ke Hasil & Skor 🏆"):
        st.session_state.current_page = "9. 🏆 Hasil dan Skor"
        st.rerun()

# --- HALAMAN 9: HASIL DAN SKOR ---
elif st.session_state.current_page == "9. 🏆 Hasil dan Skor":
    st.markdown('<h1 class="title-gradient">🏆 Rapor Hasil Belajar</h1>', unsafe_allow_html=True)
    st.write("Berikut adalah kartu rangkuman hasil jawaban Anda dari seluruh aktivitas belajar di LKPD ini.")
    
    skor_aktivitas = get_score_aktivitas()
    skor_games = get_score_games()
    skor_puzzle = get_score_puzzle()
    skor_kuis = get_score_kuis()
    
    total_skor = (skor_aktivitas + skor_games + skor_puzzle + skor_kuis) / 4.0
    
    predikat = "Perlu Belajar Lagi"
    if total_skor >= 90:
        predikat = "Sangat Baik"
    elif total_skor >= 80:
        predikat = "Baik"
    elif total_skor >= 70:
        predikat = "Cukup"
        
    st.balloons()
    
    st.markdown("### 🗂️ Kartu Hasil Belajar")
    
    pesan_motivasi = ""
    if total_skor >= 90:
        pesan_motivasi = "🎉 Luar biasa sempurna! Kamu memahami seluruh konsep data, informasi, web scraping, dan validasi dengan sangat baik!"
    elif total_skor >= 80:
        pesan_motivasi = "👍 Kerja bagus! Pemahaman kamu sudah sangat baik. Terus tingkatkan ketelitianmu!"
    elif total_skor >= 70:
        pesan_motivasi = "📐 Cukup! Kamu sudah memahami konsep dasar. Silakan pelajari lagi materi yang belum tepat."
    else:
        pesan_motivasi = "📚 Tetap semangat! Coba baca kembali ringkasan materi dan ulangi game-game yang masih salah."
        
    st.info(f"""
    **IDENTITAS SISWA:**
    * 👤 Nama: **{st.session_state.nama_siswa}**
    * 🏫 Kelas: **{st.session_state.kelas_siswa}**
    * 🔢 Nomor Absen: **{st.session_state.absen_siswa}**
    
    **PREDIKAT KELULUSAN:**
    ## 🏆 {predikat}
    
    **TOTAL SKOR AKHIR:**
    # {total_skor:.1f} / 100
    
    *{pesan_motivasi}*
    """)
    
    st.markdown("### 📊 Rincian Skor Per Aktivitas:")
    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    with col_s1:
        st.metric(label="⚡ Aktivitas Interaktif", value=f"{skor_aktivitas}/100")
    with col_s2:
        st.metric(label="🎮 Games & Tantangan", value=f"{skor_games}/100")
    with col_s3:
        st.metric(label="🧩 Puzzles & TTS", value=f"{skor_puzzle:.1f}/100")
    with col_s4:
        st.metric(label="✍️ Kuis Akhir", value=f"{skor_kuis}/100")
        
    st.divider()
    
    refleksi_text = st.session_state.j_refleksi.strip() if st.session_state.j_refleksi else "(Belum diisi)"
    st.write(f"📝 **Refleksi Siswa:** \n\n *{refleksi_text}*")
    
    st.divider()
    st.subheader("💾 Ekspor Laporan Hasil Belajar")
    
    st.write("### Opsi A: Format Excel (.xlsx) - Sangat Direkomendasikan")
    
    if not openpyxl_installed:
        st.error("❌ Pustaka `openpyxl` belum terinstall di lingkungan Python Anda.")
        st.info("Silakan jalankan perintah ini di terminal Anda untuk menginstalnya: \n`pip install openpyxl` \n\nSetelah selesai instalasi, muat ulang halaman browser Anda.")
    else:
        data_excel = {
            "Kategori Data": [
                "Identitas Siswa", "Identitas Siswa", "Identitas Siswa", 
                "Nilai Aktivitas Interaktif", "Nilai Games & Tantangan", 
                "Nilai Puzzles & TTS", "Nilai Kuis Akhir", "Nilai Akhir Rata-rata",
                "Predikat Kelulusan", "Refleksi Siswa"
            ],
            "Variabel / Pertanyaan": [
                "Nama Siswa", "Kelas", "Nomor Absen",
                "Skor Aktivitas (Halaman 5)",
                "Skor Games (Halaman 6)",
                "Skor Puzzles (Halaman 7)",
                "Skor Kuis Akhir (Halaman 8)",
                "Total Nilai LKPD",
                "Predikat Hasil Belajar",
                "Kesimpulan Refleksi"
            ],
            "Jawaban / Hasil": [
                st.session_state.nama_siswa,
                st.session_state.kelas_siswa,
                st.session_state.absen_siswa,
                f"{skor_aktivitas} / 100",
                f"{skor_games} / 100",
                f"{skor_puzzle:.1f} / 100",
                f"{skor_kuis} / 100",
                f"{total_skor:.1f} / 100",
                predikat,
                st.session_state.j_refleksi if st.session_state.j_refleksi else "(Tidak diisi)"
            ],
            "Evaluasi": [
                "-", "-", "-",
                "LULUS" if skor_aktivitas >= 70 else "BELUM LULUS",
                "LULUS" if skor_games >= 70 else "BELUM LULUS",
                "LULUS" if skor_puzzle >= 70 else "BELUM LULUS",
                "LULUS" if skor_kuis >= 70 else "BELUM LULUS",
                "LULUS" if total_skor >= 70 else "BELUM LULUS",
                "-", "-"
            ]
        }
        
        df = pd.DataFrame(data_excel)
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Hasil_LKPD_Siswa')
            
        st.download_button(
            label="📁 Unduh Laporan format Excel (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"Laporan_LKPD_Lengkap_{st.session_state.nama_siswa.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_btn"
        )
        
    st.write("### Opsi B: Format Teks (.txt)")
    teks_laporan = f"""=== LAPORAN HASIL BELAJAR LKPD INTERAKTIF ===
BAB 1: Data, Informasi, dan Validasinya

IDENTITAS SISWA:
- Nama: {st.session_state.nama_siswa}
- Kelas: {st.session_state.kelas_siswa}
- Nomor Absen: {st.session_state.absen_siswa}

PERHITUNGAN SKOR LKPD:
1. Skor Aktivitas Interaktif: {skor_aktivitas} / 100
2. Skor Games & Tantangan  : {skor_games} / 100
3. Skor Puzzles & TTS       : {skor_puzzle:.1f} / 100
4. Skor Kuis Akhir          : {skor_kuis} / 100

=============================================
TOTAL SKOR AKHIR            : {total_skor:.1f} / 100
PREDIKAT KELULUSAN          : {predikat}
=============================================

REFLEKSI SISWA:
{st.session_state.j_refleksi if st.session_state.j_refleksi else '(Tidak diisi)'}
=============================================
"""
    st.download_button(
        label="📄 Unduh Laporan format Teks (.txt)",
        data=teks_laporan,
        file_name=f"Laporan_LKPD_Lengkap_{st.session_state.nama_siswa.replace(' ', '_')}.txt",
        mime="text/plain",
        key="download_text_btn"
    )

    st.divider()
    if st.button("Lanjut ke Refleksi Siswa 💬"):
        st.session_state.current_page = "10. 💬 Refleksi"
        st.rerun()

# --- HALAMAN 10: REFLEKSI ---
elif st.session_state.current_page == "10. 💬 Refleksi":
    st.markdown('<h1 class="title-gradient">💬 Refleksi Belajar</h1>', unsafe_allow_html=True)
    st.write("Tuliskan apa saja kesimpulan yang dapat Anda ambil atau hal baru apa yang Anda pelajari setelah menyelesaikan seluruh aktivitas pembelajaran ini.")
    
    st.divider()
    
    refleksi_siswa = st.text_area(
        "Tanggapan/Refleksi Siswa:",
        value=st.session_state.j_refleksi,
        placeholder="Tulis di sini... (Contoh: Setelah menyelesaikan LKPD ini, saya belajar cara membedakan data mentah dengan informasi yang matang. Saya juga belajar cara memvalidasi informasi dengan CRAAP Test dan melakukan pencarian Google dengan operator pencarian agar lebih terarah.)",
        height=200
    )
    
    st.session_state.j_refleksi = refleksi_siswa
    
    if st.button("Simpan Refleksi"):
        st.success("Refleksi Anda berhasil disimpan! Silakan unduh laporan lengkap hasil belajar Anda di Halaman 9.")
        st.rerun()
