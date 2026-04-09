import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm

# Page config
st.set_page_config(
    page_title="Analisis Kekuatan & Kelemahan Siswa", 
    layout="wide",
    initial_sidebar_state="auto",
    page_icon="📊"
)

# ============ CUSTOM CSS - RESPONSIVE MOBILE ============
st.markdown("""
<style>
    /* Main container */
    .stApp {
        background-color: #ffffff !important;
    }
    
    .main {
        background-color: #ffffff !important;
    }
    
    .block-container {
        padding: 1rem 1rem 2rem 1rem !important;
        background-color: #ffffff !important;
    }
    
    /* Responsive untuk mobile */
    @media (max-width: 768px) {
        .block-container {
            padding: 0.8rem 0.8rem 1.5rem 0.8rem !important;
        }
        
        .premium-header {
            padding: 20px 15px !important;
        }
        
        .premium-header h1 {
            font-size: 1.2rem !important;
        }
        
        .premium-header p {
            font-size: 0.7rem !important;
        }
        
        .saldo-nilai {
            font-size: 1.5rem !important;
        }
        
        .info-value {
            font-size: 1rem !important;
        }
        
        .progress-label {
            font-size: 0.7rem !important;
        }
        
        .score-chip {
            font-size: 0.6rem !important;
            padding: 2px 8px !important;
        }
    }
    
    /* Premium Header */
    .premium-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #0f2b4d 50%, #1a4a7a 100%);
        padding: 25px 30px;
        border-radius: 20px;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15);
        position: relative;
        overflow: hidden;
    }
    
    .premium-header::before {
        content: '';
        position: absolute;
        top: -30%;
        right: -10%;
        width: 200px;
        height: 200px;
        background: rgba(255,255,255,0.05);
        border-radius: 50%;
    }
    
    .premium-header h1 {
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
        color: #ffffff !important;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .premium-header p {
        margin: 8px 0 0 0;
        font-size: 0.85rem;
        color: #cbd5e1 !important;
    }
    
    /* Saldo Card */
    .premium-card {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        border-radius: 20px;
        padding: 20px;
        color: #ffffff !important;
        margin-bottom: 20px;
        box-shadow: 0 8px 20px rgba(37,99,235,0.25);
        position: relative;
        overflow: hidden;
    }
    
    .saldo-label {
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #bfdbfe !important;
    }
    
    .saldo-nilai {
        font-size: 2rem;
        font-weight: 800;
        margin: 8px 0;
        color: #ffffff !important;
    }
    
    /* Info Cards */
    .info-card {
        background: #f8fafc;
        border-radius: 14px;
        padding: 12px;
        border-left: 3px solid #2563eb;
        margin-bottom: 10px;
    }
    
    .info-title {
        font-size: 0.65rem;
        color: #64748b !important;
        text-transform: uppercase;
    }
    
    .info-value {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a !important;
        margin-top: 4px;
    }
    
    /* Progress Container */
    .premium-progress {
        background: #f8fafc;
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 10px;
        border: 1px solid #e2e8f0;
    }
    
    .progress-label {
        display: flex;
        justify-content: space-between;
        margin-bottom: 6px;
        font-size: 0.75rem;
        color: #0f172a !important;
    }
    
    .progress-label strong {
        color: #0f172a !important;
    }
    
    .progress-bar-bg {
        background: #e2e8f0;
        border-radius: 10px;
        height: 6px;
        overflow: hidden;
    }
    
    .progress-bar-fill {
        background: linear-gradient(90deg, #2563eb, #1d4ed8);
        height: 100%;
        border-radius: 10px;
    }
    
    /* Score Chips */
    .score-chip {
        display: inline-flex;
        align-items: center;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.65rem;
        font-weight: 600;
    }
    
    .chip-excellent { background: #dcfce7; color: #166534 !important; }
    .chip-good { background: #dbeafe; color: #1e40af !important; }
    .chip-average { background: #fed7aa; color: #9a3412 !important; }
    .chip-poor { background: #fee2e2; color: #991b1b !important; }
    
    /* Metric styling */
    [data-testid="stMetric"] {
        background: #f8fafc;
        border-radius: 14px;
        padding: 12px;
        border: 1px solid #e2e8f0;
    }
    
    [data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-size: 0.7rem !important;
    }
    
    [data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 1.2rem !important;
    }
    
    /* Tabs untuk mobile */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #f1f5f9;
        padding: 4px;
        border-radius: 30px;
        flex-wrap: wrap;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 25px;
        padding: 4px 12px;
        color: #64748b !important;
        font-size: 0.7rem !important;
    }
    
    .stTabs [aria-selected="true"] {
        background: #ffffff !important;
        color: #2563eb !important;
    }
    
    /* Select Box */
    .stSelectbox > div > div {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        border: none !important;
        border-radius: 40px !important;
        padding: 5px 12px !important;
    }
    
    .stSelectbox > div > div > div {
        color: white !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
    }
    
    /* File uploader */
    .stFileUploader > div > div {
        background: #f8fafc !important;
        border: 1px dashed #2563eb !important;
        border-radius: 12px !important;
    }
    
    /* Button */
    .stButton button {
        background: linear-gradient(135deg, #2563eb, #1d4ed8);
        color: white !important;
        border: none;
        border-radius: 12px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    [data-testid="stSidebar"] {
        background: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    
    @media (max-width: 768px) {
        [data-testid="stSidebar"] {
            min-width: 250px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ============ HEADER PREMIUM ============
st.markdown("""
<div class="premium-header">
    <h1>
        <span>🎓</span> Analisis Kekuatan & Kelemahan Siswa
    </h1>
    <p>Platform Analisis Akademik | Identifikasi Potensi | Rekomendasi Belajar</p>
</div>
""", unsafe_allow_html=True)

# ============ FUNCTIONS ============
def fix_number_format(value):
    if pd.isna(value):
        return value
    str_value = str(value)
    if '.' in str_value and str_value.count('.') > 1:
        parts = str_value.split('.')
        if len(parts) >= 2:
            fixed = f"{parts[0]}.{parts[1]}"
            try:
                return float(fixed)
            except:
                return value
    try:
        return float(str_value)
    except:
        return value

def clean_dataframe(df):
    for col in df.columns:
        df[col] = df[col].apply(fix_number_format)
    return df

def get_grade_info(score, kkm=80):
    if score >= 95:
        return "A+ (Sangat Istimewa)", "chip-excellent", "#166534"
    elif score >= 90:
        return "A (Istimewa)", "chip-excellent", "#166534"
    elif score >= 85:
        return "B+ (Sangat Baik)", "chip-good", "#1e40af"
    elif score >= kkm:
        return "B (Baik)", "chip-good", "#1e40af"
    elif score >= 70:
        return "C+ (Cukup)", "chip-average", "#9a3412"
    elif score >= 60:
        return "C (Kurang)", "chip-poor", "#991b1b"
    else:
        return "D (Perbaikan)", "chip-poor", "#991b1b"

def create_radar_chart(scores, subjects, student_name, kkm=80):
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=scores,
        theta=subjects,
        fill='toself',
        name=student_name,
        line_color='#2563eb',
        line_width=2,
        fillcolor='rgba(37, 99, 235, 0.2)'
    ))
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True, 
                range=[0, 100], 
                gridcolor='#cbd5e1',
                tickfont=dict(color='#0f172a', size=8)
            ),
            angularaxis=dict(
                tickfont=dict(color='#0f172a', size=7),
                gridcolor='#e2e8f0'
            )
        ),
        showlegend=True,
        height=400,
        title=dict(
            text=f"Radar Chart - {student_name}", 
            font=dict(color='#0f172a', size=14, weight='bold')
        ),
        template='plotly_white',
        paper_bgcolor='white',
        plot_bgcolor='white',
        font=dict(color='#0f172a')
    )
    return fig

def create_gauge_chart(value, title, color="#2563eb", kkm=80):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={'text': title, 'font': {'size': 12, 'color': '#0f172a'}},
        number={'font': {'color': '#0f172a', 'size': 28, 'weight': 'bold'}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': '#64748b', 'tickfont': {'size': 8}},
            'bar': {'color': color},
            'bgcolor': 'white',
            'borderwidth': 1,
            'bordercolor': '#e2e8f0',
            'steps': [
                {'range': [0, 60], 'color': '#fee2e2'},
                {'range': [60, kkm], 'color': '#fed7aa'},
                {'range': [kkm, 90], 'color': '#dcfce7'},
                {'range': [90, 100], 'color': '#bbf7d0'}
            ],
            'threshold': {
                'line': {'color': "#ef4444", 'width': 3},
                'thickness': 0.75,
                'value': kkm
            }
        }
    ))
    fig.update_layout(
        height=250, 
        margin=dict(l=20, r=20, t=40, b=20), 
        paper_bgcolor='white'
    )
    return fig

def create_pdf_report(student_data, name_col, subject_cols, scores, avg_score, max_score, min_score, strengths, weaknesses, recommendations, kkm=80):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story = []
    
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=20, textColor=colors.HexColor('#0f172a'), alignment=1)
    header_style = ParagraphStyle('Header', parent=styles['Heading2'], fontSize=12, spaceAfter=10, textColor=colors.HexColor('#1e293b'))
    
    story.append(Paragraph("LAPORAN ANALISIS AKADEMIK SISWA", title_style))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(f"Nama: {student_data[name_col]}", styles['Heading3']))
    story.append(Paragraph(f"NISN: {int(student_data['NISN'])}", styles['Normal']))
    story.append(Paragraph(f"Tanggal: {datetime.now().strftime('%d %B %Y')}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    stats_data = [
        ['Statistik', 'Nilai', 'Kategori'],
        ['Rata-rata', f'{avg_score:.2f}', 'ISTIMEWA' if avg_score >= 90 else 'SANGAT BAIK' if avg_score >= 85 else 'BAIK' if avg_score >= kkm else 'CUKUP' if avg_score >= 70 else 'PERLU BIMBINGAN'],
        ['Nilai Tertinggi', f'{max_score:.0f}', ''],
        ['Nilai Terendah', f'{min_score:.0f}', ''],
        ['KKM', f'{kkm}', '']
    ]
    stats_table = Table(stats_data, colWidths=[4*inch, 2*inch, 3*inch])
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2563eb')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke)
    ]))
    story.append(stats_table)
    story.append(Spacer(1, 0.3*inch))
    
    story.append(Paragraph("✅ KEKUATAN (5 Nilai Tertinggi)", header_style))
    strengths_data = [['No', 'Mata Pelajaran', 'Nilai', 'Status']]
    for i, (sub, score) in enumerate(strengths[:5], 1):
        status = "✓ Lulus" if score >= kkm else "✗ Perbaikan"
        strengths_data.append([str(i), sub, f'{score:.0f}', status])
    strengths_table = Table(strengths_data, colWidths=[0.8*inch, 3.5*inch, 1.2*inch, 2*inch])
    strengths_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#10b981')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1'))
    ]))
    story.append(strengths_table)
    story.append(Spacer(1, 0.2*inch))
    
    story.append(Paragraph("⚠️ KELEMAHAN (5 Nilai Terendah)", header_style))
    weaknesses_data = [['No', 'Mata Pelajaran', 'Nilai', 'Status']]
    for i, (sub, score) in enumerate(weaknesses[-5:], 1):
        status = "✓ Lulus" if score >= kkm else "✗ Perbaikan"
        weaknesses_data.append([str(i), sub, f'{score:.0f}', status])
    weaknesses_table = Table(weaknesses_data, colWidths=[0.8*inch, 3.5*inch, 1.2*inch, 2*inch])
    weaknesses_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#ef4444')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1'))
    ]))
    story.append(weaknesses_table)
    story.append(Spacer(1, 0.3*inch))
    
    story.append(Paragraph("💡 REKOMENDASI BELAJAR", header_style))
    for i, rec in enumerate(recommendations[:5], 1):
        story.append(Paragraph(f"{i}. {rec}", styles['Normal']))
        story.append(Spacer(1, 0.05*inch))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# ============ DATA SISWA (LENGKAP 28 SISWA) ============
# Buat DataFrame langsung dari dictionary
data = {
    "Nama Siswa": [
        "ABDUL HAFIDZ", "ABDULLAH HAFIZ DHIAURRAHMAN", "ABDULLAH ZAIN", "AFNAN FIRDAUS",
        "AHMADUN AL HAQQI", "ALDRICO PASHA FATEEH NUGROHO", "ALFATIH ARSHAKA AKBAR", "BARRA KAUTSAR FIRDAUS",
        "EMIER ARJUNA AL FATIH", "FIKRI ALVARO TAJUZZAMAN", "HAFIDZ NAUFAL ALFATHAN", "HAMMAM AZZUHRI PERMANA",
        "IMAM MAHDI", "KHALID FATIH ALBASTIAH", "MUHAMMAD FATHAN HAISYAM", "MUHAMMAD KHALEEV AR RAYYAN KURNIAWAN",
        "MUHAMMAD MIKEL ALASTA EMIL", "MUHAMMAD NABIL HAMZAH AL-FATTAH", "MUHAMMAD ZHAFIF MUSLIM", "NABIL ALGHAZI SETIAWAN",
        "RAFFI ADLI SHIDQI", "RAYHAN RIZAN AL FATIH", "SA'AD ABDULLAH THOBARONY", "SABIQ SHIRATULLAH WIJAYA",
        "TAJUNA MAZRA IBRAHIM", "TSANII MUHAMMAD KHALIFIDZIKRI", "YAHYA AMRI ABDILLAH", "YUSUF AKMAL RIFAI"
    ],
    "NISN": [
        3143539590, 3150333935, 3161328188, 3157710996, 3158913526, 3157089416, 3150492070, 3154615260,
        3161409449, 3162906560, 3159758093, 3156456673, 3164882755, 3169680431, 3152014029, 3157612630,
        3149765954, 3159261937, 3158866436, 158718478, 133692902, 159673218, 3168078382, 158633055,
        136408533, 3153379248, 3158669297, 3161542999
    ],
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

subject_columns = [
    "Adab dan Akhlak", "Al-Qur'an", "Aqidah", "Bahasa Arab", "Bahasa Indonesia",
    "Bahasa Inggris", "Fiqih", "Ilmu Pengetahuan Alam dan Sosial", "Kosa Kata Bahasa Arab",
    "Matematika", "Pendidikan Jasmani Olahraga dan Kesehatan",
    "Pendidikan Lingkungan dan Budaya Jakarta", "Pendidikan Pancasila", "Praktek Ibadah",
    "Seni Budaya", "Siroh"
]

# Buat DataFrame
df_default = pd.DataFrame(data)

# ============ SIDEBAR MENU ============
with st.sidebar:
    st.markdown("### 👨‍🎓 Menu")
    st.markdown("---")
    
    st.markdown("#### 📂 Upload Data")
    uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx", "xls"], key="file_uploader")
    
    st.markdown("---")
    
    st.markdown("#### 📊 Sumber Data")
    use_default = st.radio("Pilih Sumber Data:", ["Data Contoh (28 Siswa)", "Upload File Excel"], index=0)
    
    st.markdown("---")
    st.info(f"🎯 **KKM: 80**")
    st.markdown("---")
    st.caption("📱 EduAnalytics Mobile v3.0")
    st.caption("© 2024 All Rights Reserved")

# ============ LOAD DATA ============
if uploaded_file is not None and use_default == "Upload File Excel":
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df = clean_dataframe(df)
        if 'Nama Siswa' in df.columns:
            name_col = 'Nama Siswa'
        else:
            name_col = df.columns[0]
        subject_cols = []
        for col in df.columns:
            if col not in [name_col, 'NISN'] and pd.api.types.is_numeric_dtype(df[col]):
                subject_cols.append(col)
        st.success("✅ Data berhasil dimuat!")
    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Menggunakan data contoh...")
        df = df_default.copy()
        name_col = 'Nama Siswa'
        subject_cols = subject_columns
else:
    df = df_default.copy()
    name_col = 'Nama Siswa'
    subject_cols = subject_columns
    if use_default == "Data Contoh (28 Siswa)":
        st.info("📚 Menggunakan data contoh (28 siswa)")

# ============ SUB HEADER ============
st.markdown(f"""
<div style="background: #f8fafc; padding: 10px 15px; border-radius: 12px; margin-bottom: 20px; border: 1px solid #e2e8f0;">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
        <div style="font-weight: 600; color: #0f172a;">🎯 Dashboard Akademik</div>
        <div style="background: #2563eb; color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.7rem;">Semester Ganjil 2024/2025</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============ MAIN CONTENT ============
if df is not None and len(subject_cols) > 0:
    
    KKM = 80
    
    # Student Selection
    st.markdown("### 🎓 Pilih Siswa")
    selected_student = st.selectbox("", df[name_col].unique(), label_visibility="collapsed")
    
    student_data = df[df[name_col] == selected_student].iloc[0]
    scores = [student_data[subject] for subject in subject_cols]
    avg_score = np.mean(scores)
    max_score = max(scores)
    min_score = min(scores)
    grade_info, chip_class, grade_color = get_grade_info(avg_score, KKM)
    
    # ============ PREMIUM SALDO CARD ============
    st.markdown(f"""
    <div class="premium-card">
        <div class="saldo-label">🏆 RATA-RATA NILAI AKADEMIK</div>
        <div class="saldo-nilai">{avg_score:.1f} <span style="font-size: 0.9rem;">/ 100</span></div>
        <div style="display: flex; justify-content: space-between; margin-top: 12px; flex-wrap: wrap; gap: 10px;">
            <div>
                <div class="saldo-label">Nilai Tertinggi</div>
                <div style="font-size: 1.1rem; font-weight: 700;">{max_score:.0f}</div>
            </div>
            <div>
                <div class="saldo-label">Nilai Terendah</div>
                <div style="font-size: 1.1rem; font-weight: 700;">{min_score:.0f}</div>
            </div>
            <div>
                <div class="saldo-label">KKM</div>
                <div style="font-size: 1.1rem; font-weight: 700;">{KKM}</div>
            </div>
            <div>
                <div class="saldo-label">Kategori</div>
                <div style="font-size: 0.9rem; font-weight: 600;">{grade_info.split(' ')[0]}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # ============ INFO RINGKASAN ============
    st.markdown("### 📊 Ringkasan Akademik")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total_mapel = len(subject_cols)
    lulus = sum(1 for s in scores if s >= KKM)
    perlu_bimbingan = sum(1 for s in scores if s < 70)
    istimewa = sum(1 for s in scores if s >= 90)
    
    with col1:
        st.metric("📚 Total Mapel", total_mapel)
    with col2:
        st.metric(f"✅ Lulus (≥{KKM})", f"{lulus} / {total_mapel}")
    with col3:
        st.metric("⚠️ Perlu Bimbingan", f"{perlu_bimbingan} / {total_mapel}")
    with col4:
        st.metric("🏆 Istimewa (≥90)", f"{istimewa} / {total_mapel}")
    
    # ============ KEKUATAN & KELEMAHAN ============
    subject_scores = {subject: student_data[subject] for subject in subject_cols}
    sorted_subjects = sorted(subject_scores.items(), key=lambda x: x[1], reverse=True)
    strengths = sorted_subjects[:5]
    weaknesses = sorted_subjects[-5:]
    
    col_str, col_weak = st.columns(2)
    
    with col_str:
        st.markdown("### ✅ TOP 5 KEKUATAN")
        for subject, score in strengths:
            grade, chip, _ = get_grade_info(score, KKM)
            st.markdown(f"""
            <div class="premium-progress">
                <div class="progress-label">
                    <span><strong>{subject}</strong></span>
                    <span class="score-chip {chip}">{score:.0f} | {grade.split(' ')[0]}</span>
                </div>
                <div class="progress-bar-bg">
                    <div class="progress-bar-fill" style="width: {score}%;"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    with col_weak:
        st.markdown("### ⚠️ TOP 5 KELEMAHAN")
        for subject, score in weaknesses:
            grade, chip, _ = get_grade_info(score, KKM)
            bar_color = "linear-gradient(90deg, #ef4444, #dc2626)" if score < KKM else "linear-gradient(90deg, #f59e0b, #d97706)"
            st.markdown(f"""
            <div class="premium-progress">
                <div class="progress-label">
                    <span><strong>{subject}</strong></span>
                    <span class="score-chip {chip}">{score:.0f} | {grade.split(' ')[0]}</span>
                </div>
                <div class="progress-bar-bg">
                    <div class="progress-bar-fill" style="width: {score}%; background: {bar_color};"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    # ============ VISUALISASI INTERAKTIF ============
    st.markdown("---")
    st.markdown("### 📈 Visualisasi Interaktif")
    
    tab1, tab2, tab3 = st.tabs(["📊 Bar Chart", "🎯 Radar Chart", "📊 Gauge Meter"])
    
    with tab1:
        fig = go.Figure()
        colors_bar = ['#10b981' if s >= KKM else '#f59e0b' if s >= 60 else '#ef4444' for s in scores]
        fig.add_trace(go.Bar(
            x=subject_cols,
            y=scores,
            marker_color=colors_bar,
            text=scores,
            textposition='auto',
            textfont=dict(color='white', size=9),
            hovertemplate='<b>%{x}</b><br>Nilai: %{y}<extra></extra>'
        ))
        fig.add_hline(y=KKM, line_dash="dash", line_color="#ef4444", line_width=2, 
                     annotation_text=f"KKM ({KKM})", annotation_font_color="#ef4444", annotation_font_size=10)
        fig.update_layout(
            title=dict(text=f"Analisis Nilai {selected_student}", font=dict(color='#0f172a', size=14)),
            xaxis_title="Mata Pelajaran",
            yaxis_title="Nilai",
            height=400,
            template='plotly_white',
            paper_bgcolor='white',
            plot_bgcolor='white',
            font=dict(color='#0f172a'),
            xaxis={'tickangle': -45, 'tickfont': dict(color='#0f172a', size=8)},
            yaxis={'tickfont': dict(color='#0f172a'), 'gridcolor': '#e2e8f0', 'range': [0, 100]}
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        radar_fig = create_radar_chart(scores, subject_cols, selected_student, KKM)
        st.plotly_chart(radar_fig, use_container_width=True)
    
    with tab3:
        col_g1, col_g2, col_g3 = st.columns(3)
        with col_g1:
            gauge1 = create_gauge_chart(avg_score, "Rata-rata", "#2563eb", KKM)
            st.plotly_chart(gauge1, use_container_width=True)
        with col_g2:
            gauge2 = create_gauge_chart(max_score, "Tertinggi", "#10b981", KKM)
            st.plotly_chart(gauge2, use_container_width=True)
        with col_g3:
            gauge3 = create_gauge_chart(min_score, "Terendah", "#ef4444", KKM)
            st.plotly_chart(gauge3, use_container_width=True)
    
    # ============ REKOMENDASI ============
    st.markdown("---")
    st.markdown("### 💡 Rekomendasi Belajar")
    
    weak_subjects_list = [(sub, score) for sub, score in subject_scores.items() if score < KKM]
    recommendations = []
    
    if weak_subjects_list:
        for subject, score in weak_subjects_list[:5]:
            if "Matematika" in subject:
                rec = f"{subject} (Nilai: {score:.0f}) - Perbanyak latihan soal"
            elif "Bahasa" in subject:
                rec = f"{subject} (Nilai: {score:.0f}) - Tingkatkan kosakata"
            elif "Ilmu Pengetahuan" in subject:
                rec = f"{subject} (Nilai: {score:.0f}) - Pelajari ulang konsep"
            else:
                rec = f"{subject} (Nilai: {score:.0f}) - Tambah jam belajar"
            recommendations.append(rec)
            st.warning(f"📚 {rec}")
    else:
        recommendations.append("Pertahankan prestasi yang sudah baik!")
        st.success("🎉 **SELAMAT!** Semua nilai di atas KKM!")
    
    # ============ PRESTASI ============
    if istimewa > 0:
        st.markdown("#### 🏆 Prestasi Akademik")
        for subject, score in strengths[:3]:
            if score >= 90:
                st.info(f"🌟 **{subject}** (Nilai: {score:.0f}) - Pertahankan!")
    
    # ============ DOWNLOAD PDF ============
    st.markdown("---")
    st.markdown("### 📄 Export Laporan")
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("📑 Download Laporan PDF", use_container_width=True):
            with st.spinner("Membuat laporan PDF..."):
                pdf_buffer = create_pdf_report(
                    student_data, name_col, subject_cols, scores, 
                    avg_score, max_score, min_score, 
                    strengths, weaknesses, recommendations, KKM
                )
                b64 = base64.b64encode(pdf_buffer.getvalue()).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Laporan_{selected_student.replace(" ", "_")}.pdf" style="text-decoration: none;"><div style="background: linear-gradient(135deg, #10b981, #059669); color: white; text-align: center; padding: 10px; border-radius: 12px; font-weight: 600;">✅ Download PDF</div></a>'
                st.markdown(href, unsafe_allow_html=True)
                st.balloons()
                st.success("✅ Laporan PDF siap!")

else:
    st.error("Tidak ditemukan data nilai yang valid")
