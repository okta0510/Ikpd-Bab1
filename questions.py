# ==========================================
# FILE: questions.py
# Berisi daftar soal kuis untuk LKPD Interaktif
# Pemrogram pemula dapat mengubah, menghapus, atau menambah soal di sini.
# ==========================================

QUIZ_QUESTIONS = [
    {
        "question": "1. Apa perbedaan utama antara data dan informasi?",
        "options": [
            "A. Data adalah fakta mentah, sedangkan informasi adalah data yang sudah diolah dan memiliki arti.",
            "B. Data sudah pasti benar, sedangkan informasi selalu salah.",
            "C. Data hanya berupa angka, sedangkan informasi hanya berupa teks.",
            "D. Tidak ada perbedaan sama sekali antara data dan informasi."
        ],
        "answer": 0, # Pilihan A (indeks 0) adalah jawaban yang benar
        "explanation": "Pembahasan: Data adalah fakta mentah yang belum diolah (contoh: angka-angka acak). Setelah diolah, dikelompokkan, dan diberi konteks sehingga memiliki makna bagi penerimanya, barulah disebut sebagai Informasi."
    },
    {
        "question": "2. Apa yang dimaksud dengan teknik Web Scraping?",
        "options": [
            "A. Teknik menghapus data website agar tidak bisa diakses orang lain.",
            "B. Teknik mengambil data dari halaman website secara otomatis menggunakan kode program.",
            "C. Mengunggah file Microsoft Excel dari komputer ke Google Drive.",
            "D. Menulis kode HTML untuk merancang desain tampilan website."
        ],
        "answer": 1, # Pilihan B (indeks 1) adalah jawaban yang benar
        "explanation": "Pembahasan: Web Scraping adalah proses pengumpulan atau pengambilan data dari halaman web secara otomatis menggunakan script/program komputer (seperti Python)."
    },
    {
        "question": "3. Mengapa Google Colab sangat cocok digunakan oleh pemula untuk belajar pemrograman Python?",
        "options": [
            "A. Karena bisa digunakan untuk membuat game 3D dengan kualitas grafis sangat tinggi.",
            "B. Karena gratis, berjalan secara online di browser, dan tidak memerlukan instalasi Python di komputer.",
            "C. Karena Google Colab tidak membutuhkan koneksi internet sama sekali saat digunakan.",
            "D. Karena Google Colab bisa menggantikan seluruh fungsi dari sistem operasi Windows."
        ],
        "answer": 1, # Pilihan B (indeks 1) adalah jawaban yang benar
        "explanation": "Pembahasan: Google Colab berjalan di cloud (server Google). Kita hanya perlu browser dan internet untuk menulis serta menjalankan kode Python, tanpa perlu repot menginstal software apa pun di komputer lokal."
    },
    {
        "question": "4. Pustaka (library) Python yang digunakan untuk mengirim permintaan (request) mengambil halaman web adalah...",
        "options": [
            "A. BeautifulSoup",
            "B. Pandas",
            "C. requests",
            "D. math"
        ],
        "answer": 2, # Pilihan C (indeks 2) adalah jawaban yang benar
        "explanation": "Pembahasan: Library 'requests' berfungsi untuk melakukan panggilan HTTP (seperti GET request) ke server website untuk mengunduh konten mentah (HTML) dari alamat URL tersebut."
    },
    {
        "question": "5. Apa fungsi dari library BeautifulSoup dalam proses web scraping di Python?",
        "options": [
            "A. Mengirimkan sinyal internet ke server website tujuan.",
            "B. Menyimpan data langsung ke dalam file Microsoft Excel.",
            "C. Membaca dan menganalisis (parsing) struktur HTML untuk memudahkan pencarian data tertentu.",
            "D. Membuat grafik visualisasi data yang berwarna-warni."
        ],
        "answer": 2, # Pilihan C (indeks 2) adalah jawaban yang benar
        "explanation": "Pembahasan: BeautifulSoup digunakan untuk memotong-motong dan membaca (parsing) dokumen HTML yang rumit agar kita bisa mencari tag tertentu (seperti mencari judul, harga, dll.) berdasarkan nama tag dan kelasnya."
    },
    {
        "question": "6. Dalam struktur HTML, data buku dibungkus dalam tag <article> dengan kelas (class) 'product_pod'. Perintah BeautifulSoup yang tepat untuk mengambil SEMUA buku tersebut adalah...",
        "options": [
            "A. soup.find('article', class_='product_pod')",
            "B. soup.find_all('article', class_='product_pod')",
            "C. soup.get('product_pod')",
            "D. requests.get('product_pod')"
        ],
        "answer": 1, # Pilihan B (indeks 1) adalah jawaban yang benar
        "explanation": "Pembahasan: Perintah 'find_all' digunakan untuk mencari SEMUA tag yang cocok di seluruh halaman web dan mengembalikannya dalam bentuk list, sedangkan 'find' hanya mencari satu tag pertama yang cocok."
    },
    {
        "question": "7. Jika kita memiliki data berupa list, cara paling mudah untuk menampilkan data tersebut dalam bentuk tabel baris dan kolom di Python adalah menggunakan library...",
        "options": [
            "A. requests dengan kode requests.get()",
            "B. BeautifulSoup dengan kode soup.find()",
            "C. Pandas dengan kode pd.DataFrame()",
            "D. time dengan kode time.sleep()"
        ],
        "answer": 2, # Pilihan C (indeks 2) adalah jawaban yang benar
        "explanation": "Pembahasan: Library Pandas memiliki fungsi 'pd.DataFrame()' yang sangat populer untuk mengubah kumpulan data mentah (list/dictionary) menjadi tabel terstruktur (DataFrame) yang rapi dan mudah dianalisis."
    },
    {
        "question": "8. Apa kepanjangan dari CRAAP dalam metode validasi informasi (CRAAP Test)?",
        "options": [
            "A. Currency, Relevance, Authority, Accuracy, Purpose",
            "B. Critical, Rational, Analytical, Active, Professional",
            "C. Computer, Network, Access, Application, Protocol",
            "D. Creativity, Reading, Action, Ability, Practice"
        ],
        "answer": 0, # Pilihan A (indeks 0) adalah jawaban yang benar
        "explanation": "Pembahasan: CRAAP Test adalah metode evaluasi informasi yang terdiri dari: Currency (Kebaruan), Relevance (Relevansi), Authority (Sumber/Otoritas), Accuracy (Akurasi), dan Purpose (Tujuan pembuatan informasi)."
    },
    {
        "question": "9. Apa yang dimaksud dengan teknik Lateral Reading (Membaca Lateral)?",
        "options": [
            "A. Membaca sebuah website dari halaman paling atas sampai bawah secara berulang-ulang.",
            "B. Membaca teks dengan cepat (skimming) untuk mencari kata kunci tertentu saja.",
            "C. Memverifikasi kebenaran informasi suatu website dengan cara membuka tab baru untuk mencari tahu apa kata sumber lain tentang website tersebut.",
            "D. Menerjemahkan bahasa asing pada website ke dalam bahasa Indonesia menggunakan kamus."
        ],
        "answer": 2, # Pilihan C (indeks 2) adalah jawaban yang benar
        "explanation": "Pembahasan: Membaca Lateral dilakukan dengan keluar dari website asal (membuka tab baru di browser) untuk memeriksa kebenaran klaim website tersebut berdasarkan sudut pandang situs-situs berita atau pakar lainnya."
    },
    {
        "question": "10. Mengapa data yang kita peroleh dari hasil scraping di internet tetap harus divalidasi?",
        "options": [
            "A. Karena program web scraping selalu merusak data yang diunduh.",
            "B. Karena informasi di internet tidak semuanya benar; bisa saja mengandung bias, data usang, atau berasal dari sumber yang tidak tepercaya.",
            "C. Karena koneksi internet akan terputus jika data tidak segera divalidasi.",
            "D. Karena aturan hukum mewajibkan semua file hasil scraping diubah formatnya menjadi PDF."
        ],
        "answer": 1, # Pilihan B (indeks 1) adalah jawaban yang benar
        "explanation": "Pembahasan: Internet adalah ruang bebas di mana siapa saja bisa mengunggah informasi. Oleh karena itu, data hasil scraping belum tentu otomatis valid/benar, sehingga kita wajib menyaringnya dengan metode validasi informasi."
    }
]
