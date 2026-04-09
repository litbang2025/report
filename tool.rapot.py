import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re

st.set_page_config(page_title="Analisis Kekuatan & Kelemahan Siswa", layout="wide")

st.title("📊 Analisis Kekuatan & Kelemahan Siswa")
st.markdown("### Otomatis memperbaiki format angka (95.00.00 → 95.00)")

# File uploader
uploaded_file = st.sidebar.file_uploader(
    "Upload File Excel", 
    type=["xlsx", "xls"],
    help="Upload file Excel dengan data nilai siswa"
)

def fix_number_format(value):
    """Fix number format like 95.00.00 to 95.00"""
    if pd.isna(value):
        return value
    
    # Convert to string
    str_value = str(value)
    
    # Pattern: numbers with multiple dots (e.g., 95.00.00)
    if '.' in str_value and str_value.count('.') > 1:
        # Take only first two parts (integer and first decimal)
        parts = str_value.split('.')
        if len(parts) >= 2:
            fixed = f"{parts[0]}.{parts[1]}"
            try:
                return float(fixed)
            except:
                return value
    
    # Try to convert to float normally
    try:
        return float(str_value)
    except:
        return value

def clean_dataframe(df):
    """Clean entire dataframe: fix numbers and identify columns"""
    # Fix number formats in all columns
    for col in df.columns:
        df[col] = df[col].apply(fix_number_format)
    
    return df

# Data from your input (as backup)
default_data = {
    "NISN": [3143539590, 3150333935, 3161328188, 3157710996, 3158913526, 3157089416, 3150492070, 3154615260, 3161409449, 3162906560, 3159758093, 3156456673, 3164882755, 3169680431, 3152014029, 3157612630, 3149765954, 3159261937, 3158866436, 158718478, 133692902, 159673218, 3168078382, 158633055, 136408533, 3153379248, 3158669297, 3161542999],
    "Adab dan Akhlak": [95,94,94,96,92,99,99,93,93,92,96,97,96,95,92,100,97,98,97,96,99,95,85,99,98,98,88,98],
    "Al-Qur'an": [91,97,96,98,93,96,94,97,92,92,97,91,91,93,98,97,91,90,99,92,95,99,90,96,95,95,92,91],
    "Aqidah": [76,76,90,93,86,88,88,79,77,80,86,79,94,94,86,93,82,77,99,80,88,89,62,95,97,95,75,92],
    "Bahasa Arab": [88,91,95,98,94,99,94,96,85,88,91,96,99,95,99,100,92,86,100,92,100,98,84,96,96,95,86,93],
    "Bahasa Indonesia": [80,80,80,83,81,84,89,87,83,80,83,87,91,84,82,94,87,80,93,80,91,85,80,90,89,86,80,81],
    "Bahasa Inggris": [87,98,94,95,95,99,98,93,97,98,97,100,96,98,96,97,96,87,99,92,96,99,86,98,98,98,89,97],
    "Fiqih": [93,95,96,100,91,98,98,98,91,92,95,96,91,96,96,99,95,92,100,93,99,98,92,99,100,96,89,92],
    "Ilmu Pengetahuan Alam dan Sosial": [68,79,92,90,72,89,92,94,92,72,91,89,94,90,88,96,89,82,98,84,88,87,75,96,96,94,73,85],
    "Kosa Kata Bahasa Arab": [87,97,97,100,98,100,97,97,93,96,97,94,96,96,94,99,94,84,100,95,97,99,59,100,100,97,92,98],
    "Matematika": [65,67,90,85,86,82,92,80,78,76,88,85,80,87,78,98,73,67,97,88,75,87,88,88,97,85,70,80],
    "Pendidikan Jasmani Olahraga dan Kesehatan": [95,93,96,93,94,95,94,96,93,93,95,94,96,95,96,95,94,95,96,95,95,95,93,94,95,97,94,94],
    "Pendidikan Lingkungan dan Budaya Jakarta": [80,85,87,91,86,97,93,92,95,78,92,96,93,91,91,96,87,88,99,85,95,94,76,95,97,92,79,96],
    "Pendidikan Pancasila": [79,83,84,85,84,87,94,84,89,84,90,86,82,89,89,96,92,87,94,83,91,90,84,93,95,94,87,92],
    "Praktek Ibadah": [88,91,95,98,94,99,94,96,85,88,91,98,99,95,99,100,92,86,100,92,100,98,84,96,96,95,86,93],
    "Seni Budaya": [93,93,96,98,93,93,96,93,94,96,96,94,96,93,97,98,95,93,97,93,95,98,93,94,96,93,93,95],
    "Siroh": [91,87,94,99,94,100,96,100,90,91,96,85,90,93,99,99,93,85,100,96,99,95,84,100,100,98,84,94]
}

# Subject columns (for reference)
subject_columns = [
    "Adab dan Akhlak", "Al-Qur'an", "Aqidah", "Bahasa Arab", "Bahasa Indonesia",
    "Bahasa Inggris", "Fiqih", "Ilmu Pengetahuan Alam dan Sosial", "Kosa Kata Bahasa Arab",
    "Matematika", "Pendidikan Jasmani Olahraga dan Kesehatan",
    "Pendidikan Lingkungan dan Budaya Jakarta", "Pendidikan Pancasila", "Praktek Ibadah",
    "Seni Budaya", "Siroh"
]

# Option to use default data
use_default = st.sidebar.checkbox("Gunakan Data Contoh (Jika tidak upload file)", value=True)

if uploaded_file is not None:
    # Read Excel file
    try:
        # Try to read with different engines
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        except:
            try:
                df = pd.read_excel(uploaded_file, engine='xlrd')
            except:
                df = pd.read_excel(uploaded_file)
        
        st.success("✅ File berhasil diupload")
        
        # Clean the data
        df = clean_dataframe(df)
        st.info("🔧 Format angka telah diperbaiki (95.00.00 → 95.00)")
        
        # Display preview
        st.subheader("📋 Preview Data (Setelah Perbaikan)")
        st.dataframe(df.head(10))
        
        # Identify columns
        if 'Nama Siswa' in df.columns:
            name_col = 'Nama Siswa'
        else:
            name_col = df.columns[0]
        
        # Get subject columns (numeric columns except NISN)
        subject_cols = []
        for col in df.columns:
            if col not in [name_col, 'NISN'] and pd.api.types.is_numeric_dtype(df[col]):
                subject_cols.append(col)
        
    except Exception as e:
        st.error(f"Error membaca file: {e}")
        st.stop()

elif use_default:
    # Use default data from your input
    df = pd.DataFrame(default_data)
    df.insert(0, 'Nama Siswa', [
        "ABDUL HAFIDZ", "ABDULLAH HAFIZ DHIAURRAHMAN", "ABDULLAH ZAIN", "ACHMAD ZAKY AL HAFIZ", 
        "ADELLA KAYLIFA RAMADHANI", "ADINDA ZAHRA MAHARANI", "AISYAH AULIA RAMADHANI", 
        "ALYA RAHMANIA", "ANDHIKA RIZKI PRATAMA", "ANINDYA PUTRI RAMADHANI", 
        "ANNAFI ARKANANTA", "AQILA NURUL AULIA", "ASIAH NURUL JANNAH", "AULIA RAMADHANI",
        "AZKA NAILUFAR", "BINTANG RAMADHAN", "CHIKA AMELIA PUTRI", "DZAKIYYAH AULIYAH",
        "FARHAN ALFARIZI", "FAUZAN ABDILLAH", "FELISHA AZ-ZAHRA", "FIKRI HAIKAL", 
        "GHINA HANIFA", "HANA NURUL AZIZAH", "HANIFAH ZAHRA", "HARUN ARRASYID", 
        "IRENE AULIA", "JIHAN NURUL AULIA"
    ])
    df = df[['Nama Siswa', 'NISN'] + subject_columns]
    
    st.info("📚 Menggunakan data contoh (28 siswa)")
    st.subheader("📋 Preview Data")
    st.dataframe(df.head(10))
    
    subject_cols = subject_columns
    name_col = 'Nama Siswa'
    
else:
    st.info("👈 Silakan upload file Excel atau centang 'Gunakan Data Contoh'")
    st.stop()

# Check if we have data
if df is not None and len(subject_cols) > 0:
    
    # Sidebar for student selection
    st.sidebar.header("🎯 Pilih Siswa")
    selected_student = st.sidebar.selectbox("Nama Siswa", df[name_col].unique())
    
    # Get student data
    student_data = df[df[name_col] == selected_student].iloc[0]
    
    # Main content
    st.header(f"📊 Analisis: {selected_student}")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("📋 Data Siswa")
        st.markdown(f"**Nama:** {student_data[name_col]}")
        st.markdown(f"**NISN:** {int(student_data['NISN'])}")
        
        # Statistics
        scores = [student_data[subject] for subject in subject_cols]
        avg_score = np.mean(scores)
        max_score = max(scores)
        min_score = min(scores)
        
        st.markdown("---")
        st.subheader("📈 Statistik Nilai")
        st.metric("📊 Rata-rata", f"{avg_score:.2f}")
        st.metric("🏆 Nilai Tertinggi", f"{max_score:.2f}")
        st.metric("⚠️ Nilai Terendah", f"{min_score:.2f}")
        
        # Grade category
        if avg_score >= 90:
            st.success("✅ Kategori: **ISTIMEWA** (A)")
        elif avg_score >= 85:
            st.success("✅ Kategori: **SANGAT BAIK** (B+)")
        elif avg_score >= 75:
            st.info("📘 Kategori: **BAIK** (B)")
        elif avg_score >= 60:
            st.warning("⚠️ Kategori: **CUKUP** (C)")
        else:
            st.error("❌ Kategori: **PERLU BIMBINGAN** (D)")
    
    with col2:
        st.subheader("💪 KEKUATAN & ⚠️ KELEMAHAN")
        
        # Determine strengths and weaknesses
        subject_scores = {subject: student_data[subject] for subject in subject_cols}
        sorted_subjects = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)
        
        strengths = sorted_subjects[:5]  # Top 5
        weaknesses = sorted_subjects[-5:]  # Bottom 5
        
        col_strength, col_weakness = st.columns(2)
        
        with col_strength:
            st.markdown("##### ✅ 5 NILAI TERTINGGI")
            for subject, score in strengths:
                if score >= 90:
                    st.markdown(f"- **{subject}:** {score:.0f} 🏆")
                elif score >= 75:
                    st.markdown(f"- **{subject}:** {score:.0f} ✨")
                else:
                    st.markdown(f"- **{subject}:** {score:.0f}")
        
        with col_weakness:
            st.markdown("##### ⚠️ 5 NILAI TERENDAH")
            for subject, score in weaknesses:
                if score < 70:
                    st.markdown(f"- **{subject}:** {score:.0f} 🔴")
                elif score < 75:
                    st.markdown(f"- **{subject}:** {score:.0f} 🟡")
                else:
                    st.markdown(f"- **{subject}:** {score:.0f}")
    
    # Visualization
    st.subheader("📊 GRAFIK NILAI SEMUA MATA PELAJARAN")
    
    fig, ax = plt.subplots(figsize=(12, 8))
    scores_display = [student_data[sub] for sub in subject_cols]
    
    # Color coding
    colors = []
    for s in scores_display:
        if s >= 90:
            colors.append('darkgreen')
        elif s >= 75:
            colors.append('green')
        elif s >= 60:
            colors.append('orange')
        else:
            colors.append('red')
    
    bars = ax.barh(subject_cols, scores_display, color=colors)
    ax.axvline(x=75, color='blue', linestyle='--', linewidth=2, label='KKM (75)')
    ax.axvline(x=90, color='gold', linestyle='--', linewidth=2, label='Istimewa (90)')
    ax.set_xlabel('Nilai', fontsize=12)
    ax.set_title(f'Analisis Nilai {selected_student}', fontsize=14, fontweight='bold')
    ax.legend(loc='lower right')
    
    # Add value labels
    for i, (bar, score) in enumerate(zip(bars, scores_display)):
        ax.text(score + 1, bar.get_y() + bar.get_height()/2, f'{score:.0f}', 
                va='center', fontsize=9)
    
    plt.tight_layout()
    st.pyplot(fig)
    
    # Recommendations
    st.subheader("💡 REKOMENDASI")
    
    weak_subjects = [(sub, score) for sub, score in subject_scores.items() if score < 75]
    very_weak_subjects = [(sub, score) for sub, score in subject_scores.items() if score < 70]
    
    if very_weak_subjects:
        st.warning(f"⚠️ **Perlu Bimbingan Khusus** - Nilai di bawah 70 pada: {', '.join([f'{sub} ({score:.0f})' for sub, score in very_weak_subjects])}")
        
        # Specific recommendations
        for subject, score in very_weak_subjects:
            if "Matematika" in subject:
                st.markdown(f"- 📐 **{subject}**: Perbanyak latihan soal dan bimbingan privat")
            elif "Bahasa" in subject:
                st.markdown(f"- 📖 **{subject}**: Tingkatkan kosakata dan praktek percakapan")
            elif "Ilmu Pengetahuan" in subject:
                st.markdown(f"- 🔬 **{subject}**: Pelajari ulang konsep dasar dan praktikum")
            else:
                st.markdown(f"- 📚 **{subject}**: Tambah jam belajar dan diskusi kelompok")
    
    elif weak_subjects:
        st.info(f"📚 **Perlu Peningkatan** - Nilai di bawah 75 pada: {', '.join([f'{sub} ({score:.0f})' for sub, score in weak_subjects])}")
    else:
        st.success("🎉 **SELAMAT!** Semua nilai di atas standar KKM. Pertahankan prestasi ini!")
    
    # Strengths recognition
    strong_subjects = [(sub, score) for sub, score in subject_scores.items() if score >= 90]
    if strong_subjects:
        st.balloons()
        st.success(f"🏆 **KEUNGGULAN** - Sangat menguasai: {', '.join([f'{sub} ({score:.0f})' for sub, score in strong_subjects[:3]])}")
    
    # Download report
    st.markdown("---")
    col_download1, col_download2 = st.columns(2)
    
    with col_download1:
        # Generate report for this student
        report_data = {
            "Nama Siswa": [selected_student],
            "NISN": [int(student_data['NISN'])],
            "Rata-rata": [f"{avg_score:.2f}"],
            "Kategori": ["ISTIMEWA" if avg_score >= 90 else "SANGAT BAIK" if avg_score >= 85 else "BAIK" if avg_score >= 75 else "CUKUP" if avg_score >= 60 else "PERLU BIMBINGAN"]
        }
        report_df = pd.DataFrame(report_data)
        csv = report_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Laporan Singkat (CSV)",
            data=csv,
            file_name=f"laporan_{selected_student.replace(' ', '_')}.csv",
            mime="text/csv"
        )
    
    with col_download2:
        # Full scores
        full_scores = pd.DataFrame([{**{"Nama": selected_student, "NISN": int(student_data['NISN'])}, **{sub: student_data[sub] for sub in subject_cols}}])
        csv_full = full_scores.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Semua Nilai (CSV)",
            data=csv_full,
            file_name=f"nilai_lengkap_{selected_student.replace(' ', '_')}.csv",
            mime="text/csv"
        )

else:
    st.error("Tidak ditemukan kolom nilai yang valid")