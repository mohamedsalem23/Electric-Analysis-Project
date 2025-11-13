# ====== 1. إعدادات سريعة ======
import os, streamlit as st, pandas as pd, re, io, base64
from typing import List
from PIL import Image
from langchain_core.documents import Document
from langchain_core.messages import HumanMessage
from langchain import PromptTemplate
from langchain.chains.question_answering import load_qa_chain
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import HuggingFaceEmbeddings

# ✅ إضافات PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, black, grey
from datetime import datetime

# ✅ دعم العربية الكامل
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from bidi import algorithm as bidi_algorithm
from arabic_reshaper import reshape 

os.environ["CUDA_VISIBLE_DEVICES"] = "-1"

# ✅ قراءة API Key من Streamlit Secrets
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
    if not GEMINI_API_KEY:
        st.error("⚠️ يرجى إضافة GEMINI_API_KEY في Settings → Secrets على Streamlit Cloud")
        st.info("للتطوير المحلي: أنشئ ملف .streamlit/secrets.toml وضع فيه: GEMINI_API_KEY = 'your_key'")
        st.stop()

# ====== 2. دوال مساعدة ======
def pil_to_base64_uri(image: Image.Image, fmt="PNG") -> str:
    buf = io.BytesIO()
    image.save(buf, format=fmt)
    img_bytes = buf.getvalue()
    return f"data:image/{fmt.lower()};base64,{base64.b64encode(img_bytes).decode()}"

@st.cache_data(show_spinner=False)
def load_excel() -> pd.DataFrame:
    """قراءة ملف Excel من المسار النسبي"""
    try:
        # ✅ مسار نسبي للـ Deploy
        excel_path = os.path.join(os.path.dirname(__file__), "data", "جميع_بنود_فحص_الكهرباء.xlsx")
        
        if not os.path.exists(excel_path):
            st.error(f"❌ الملف غير موجود: {excel_path}")
            st.info("تأكد من وجود الملف في مجلد data/")
            # بيانات تجريبية
            return pd.DataFrame({
                "رقم البند": [5]*4,
                "اسم البند": ["جودة التشطيب حول الأفياش الكهربائية"]*4,
                "المتطلب": ["..."],
                "التعريف حسب الكود السعودي": ["..."],
                "التوصيات": ["يجب التأكد من تثبيت المفتاح جيداً.; يجب معالجة الفراغات حول الإطار."],
                "طريقة الإصلاح": ["استخدام السيليكون لملء الفراغات.; إعادة تثبيت الأفياش بشكل مستقيم."],
                "التكلفة التقديرية (ريال)": [35,30,40,25]
            })
        
        return pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"❌ خطأ في قراءة الملف: {e}")
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def df_to_docs(df: pd.DataFrame) -> List[Document]:
    return [Document(page_content=f"اسم البند: {r['اسم البند']}. المتطلب: {r['المتطلب']}.", metadata=r.to_dict())
            for _, r in df.iterrows()]

def filter_best_doc(similar_docs: List[Document], query: str) -> int:
    best_doc = None
    best_score = 0.0
    for doc in similar_docs:
        name = doc.metadata.get('اسم البند', '')
        match_score = len(set(re.findall(r'\w+', query.lower())) & set(re.findall(r'\w+', name.lower()))) / max(len(set(re.findall(r'\w+', query.lower()))), 1)
        if match_score > best_score:
            best_score = match_score
            best_doc = doc
    return int(best_doc.metadata.get('رقم البند', 0)) if best_doc else int(similar_docs[0].metadata.get('رقم البند', 0))

def build_table_from_band(dataframe: pd.DataFrame, band_num: int, query: str) -> str:
    band_rows = dataframe[dataframe['رقم البند'] == band_num].copy()
    if band_rows.empty:
        return "| لا توجد بيانات |"
    def match_score(row):
        req = str(row.get('المتطلب', '')).lower()
        q_words = set(re.findall(r'\w+', query.lower()))
        row_words = set(re.findall(r'\w+', req))
        return len(q_words & row_words) / max(len(q_words), 1)
    band_rows['match_score'] = band_rows.apply(match_score, axis=1)
    best_row = band_rows.loc[band_rows['match_score'].idxmax()]
    cols = ['رقم البند', 'اسم البند', 'المتطلب', 'التعريف حسب الكود السعودي', 'التوصيات', 'طريقة الإصلاح', 'التكلفة التقديرية (ريال)']
    best_row = best_row[cols].to_frame().T
    return best_row.to_markdown(index=False)

@st.cache_resource(show_spinner=False)
def get_models():
    embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    chat = ChatGoogleGenerativeAI(google_api_key=GEMINI_API_KEY, model="gemini-2.5-flash", temperature=0)
    return embeddings, chat

embeddings, chat = get_models()

@st.cache_resource(show_spinner=False)
def get_vector_db(_docs: List[Document]):
    persist_dir = "chroma_db"
    if os.path.isdir(persist_dir):
        return Chroma(persist_directory=persist_dir, embedding_function=embeddings)
    db = Chroma.from_documents(_docs, embeddings, persist_directory=persist_dir)
    return db

@st.cache_data(show_spinner=False)
def batch_analyze(images_bytes: List[bytes]) -> List[str]:
    prompt = """
    أنت نظام رؤية حاسوبية متخصص. مهمتك هي تحليل الصورة المرفقة وتحديد **جميع أسماء العيوب الرئيسية** اللي تظهر (حتى لو أكثر من واحدة، مثل فراغات + ميلان + بروز). 
    **لكل عيب، أعطِ اسم البند المطابق (أو الأقرب) من جدول الجودة**، وفصلها بـ ';' (مثل: 'جودة التشطيب حول الأفياش الكهربائية; استقامة الأفياش الكهربائية أفقيًا').
    لو عيب واحد، أعطِ اسمه بس. لا تضف تفسير أو شرح، ناتجك نص واحد مفصول بـ ';'.
    """
    content = [{"type": "text", "text": prompt}]
    for img_bytes in images_bytes:
        img = Image.open(io.BytesIO(img_bytes))
        uri = pil_to_base64_uri(img)
        content.append({"type": "image_url", "image_url": {"url": uri}})
    msg = HumanMessage(content=content)
    resp = chat.invoke([msg])
    lines = resp.content.strip().splitlines()
    defects = []
    for line in lines:
        defects.extend([x.strip() for x in line.split(";") if x.strip()])
    return defects

# ✅ دالة معالجة النص العربي المُحسّنة
def process_arabic_text(text: str) -> str:
    """معالجة النص العربي بشكل صحيح مع دعم RTL"""
    if not text or text == "nan" or pd.isna(text):
        return "—"
    text = str(text).strip()
    reshaped = reshape(text)
    bidi_text = bidi_algorithm.get_display(reshaped)
    return bidi_text

# ✅ دالة تنظيف Markdown المُحسّنة
def clean_markdown_text(text: str) -> str:
    """تنظيف نص Markdown وتحويله لنص عادي"""
    text = re.sub(r'\*{1,2}([^*]+)\*{1,2}', r'\1', text)
    text = re.sub(r'#{1,6}\s+', '', text)
    text = re.sub(r'[_`~\[\]]', '', text)
    text = re.sub(r'^\s*([•\-*+]|\d+\.)\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()

# ✅ دالة تسجيل الخطوط من المسار النسبي
def register_fonts():
    """تسجيل خطوط Tahoma من مجلد fonts"""
    try:
        fonts_dir = os.path.join(os.path.dirname(__file__), "fonts")
        tahoma_path = os.path.join(fonts_dir, "Tahoma.ttf")
        tahoma_bold_path = os.path.join(fonts_dir, "Tahomabd.ttf")
        
        if os.path.exists(tahoma_path):
            pdfmetrics.registerFont(TTFont("Tahoma", tahoma_path))
        else:
            st.warning(f"⚠️ ملف Tahoma.ttf غير موجود في {fonts_dir}")
            
        if os.path.exists(tahoma_bold_path):
            pdfmetrics.registerFont(TTFont("Tahoma-Bold", tahoma_bold_path))
        else:
            st.warning(f"⚠️ ملف Tahomabd.ttf غير موجود في {fonts_dir}")
            
    except Exception as e:
        st.warning(f"⚠️ تحذير: مشكلة في تسجيل الخط: {e}")

# ✅ دالة إنشاء أنماط مُحسّنة
def create_custom_styles():
    """إنشاء أنماط مخصصة للتقرير"""
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName='Tahoma-Bold',
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        textColor=HexColor('#1a1a1a'),
        spaceAfter=20,
        spaceBefore=10
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontName='Tahoma-Bold',
        fontSize=14,
        leading=20,
        alignment=TA_RIGHT,
        textColor=HexColor('#2c3e50'),
        spaceAfter=12,
        spaceBefore=15
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontName='Tahoma',
        fontSize=11,
        leading=18,
        alignment=TA_RIGHT,
        textColor=HexColor('#333333'),
        spaceAfter=10,
        spaceBefore=5,
        rightIndent=10,
        leftIndent=10,
        wordWrap='RTL'
    )
    
    summary_style = ParagraphStyle(
        'CustomSummary',
        parent=body_style,
        fontName='Tahoma',
        fontSize=11,
        leading=20,
        backColor=HexColor('#f8f9fa'),
        borderWidth=1,
        borderColor=HexColor('#dee2e6'),
        borderPadding=10,
        borderRadius=3,
        spaceAfter=8,
        spaceBefore=5
    )
    
    defect_title_style = ParagraphStyle(
        'DefectTitle',
        parent=styles['Heading3'],
        fontName='Tahoma-Bold',
        fontSize=12,
        leading=16,
        alignment=TA_RIGHT,
        textColor=HexColor('#e74c3c'),
        spaceAfter=8,
        spaceBefore=12
    )
    
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Tahoma',
        fontSize=10,
        leading=14,
        alignment=TA_RIGHT,
        textColor=HexColor('#2c3e50'),
        wordWrap='RTL',
        rightIndent=5,
        leftIndent=5
    )
    
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=table_cell_style,
        fontName='Tahoma-Bold',
        fontSize=10,
        textColor=HexColor('#ffffff'),
        backColor=HexColor('#34495e')
    )
    
    return {
        'title': title_style,
        'subtitle': subtitle_style,
        'body': body_style,
        'summary': summary_style,
        'defect_title': defect_title_style,
        'table_cell': table_cell_style,
        'table_header': table_header_style
    }

# ✅ دالة تحويل Markdown table لـTable object المُحسّنة
def markdown_to_enhanced_table(md_text: str, styles_dict: dict) -> Table:
    """تحويل جدول Markdown إلى Table object مع تنسيق محسّن"""
    lines = [line.strip() for line in md_text.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        empty_para = Paragraph(process_arabic_text("لا توجد بيانات"), styles_dict['body'])
        return Table([[empty_para]], colWidths=[6*inch])
    
    header_cells = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
    rows = []
    for line in lines[2:]:
        row_cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        if len(row_cells) == len(header_cells):
            rows.append(row_cells)
    
    num_cols = len(header_cells)
    total_width = 6.5 * inch
    
    col_width_ratios = {
        'رقم البند': 0.5,
        'اسم البند': 1.2,
        'المتطلب': 1.3,
        'التعريف حسب الكود السعودي': 1.2,
        'التوصيات': 1.5,
        'طريقة الإصلاح': 1.5,
        'التكلفة التقديرية (ريال)': 0.8
    }
    
    col_widths = []
    for header in header_cells:
        ratio = col_width_ratios.get(header, 1.0)
        col_widths.append(ratio * inch)
    
    processed_data = []
    
    header_row = []
    for cell in header_cells:
        processed_text = process_arabic_text(cell)
        para = Paragraph(processed_text, styles_dict['table_header'])
        header_row.append(para)
    processed_data.append(header_row)
    
    for row in rows:
        row_processed = []
        for col_idx, cell in enumerate(row):
            col_name = header_cells[col_idx]
            
            if col_name in ['التوصيات', 'طريقة الإصلاح', 'المتطلب', 'التعريف حسب الكود السعودي']:
                items = [item.strip() for item in re.split(r'[;.]', str(cell)) if item.strip()]
                
                if len(items) > 1:
                    bullet_text = ""
                    for i, item in enumerate(items, 1):
                        if item:
                            bullet_text += f"• {item}<br/>"
                    
                    processed_text = process_arabic_text(bullet_text)
                    para = Paragraph(processed_text, styles_dict['table_cell'])
                    row_processed.append(para)
                else:
                    processed_text = process_arabic_text(cell)
                    para = Paragraph(processed_text, styles_dict['table_cell'])
                    row_processed.append(para)
            else:
                processed_text = process_arabic_text(cell)
                para = Paragraph(processed_text, styles_dict['table_cell'])
                row_processed.append(para)
        
        processed_data.append(row_processed)
    
    table = Table(processed_data, colWidths=col_widths, repeatRows=1)
    
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#34495e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#ffffff')),
        ('FONTNAME', (0, 0), (-1, 0), 'Tahoma-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, 0), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        
        ('BACKGROUND', (0, 1), (-1, -1), HexColor('#ffffff')),
        ('TEXTCOLOR', (0, 1), (-1, -1), HexColor('#2c3e50')),
        ('FONTNAME', (0, 1), (-1, -1), 'Tahoma'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 1), (-1, -1), 'TOP'),
        
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#bdc3c7')),
        ('LINEBELOW', (0, 0), (-1, 0), 2, HexColor('#2c3e50')),
        
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#ffffff'), HexColor('#f8f9fa')]),
    ])
    
    table.setStyle(table_style)
    return table

# ✅ دالة توليد PDF المُحسّنة
def generate_enhanced_pdf_report(images: List[Image.Image], summary: str, tables: List[tuple], defects: List[str]):
    """توليد تقرير PDF محسّن مع تنسيق احترافي"""
    buffer = io.BytesIO()
    
    # تسجيل الخطوط
    register_fonts()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=50,
        leftMargin=50,
        topMargin=60,
        bottomMargin=40
    )
    
    custom_styles = create_custom_styles()
    
    story = []
    
    # العنوان الرئيسي
    title_text = process_arabic_text("تقرير فحص العيوب الكهربائية")
    story.append(Paragraph(title_text, custom_styles['title']))
    story.append(Spacer(1, 30))
    
    # معلومات التقرير
    date_text = process_arabic_text(f"تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d')}")
    time_text = process_arabic_text(f"وقت التقرير: {datetime.now().strftime('%H:%M:%S')}")
    
    info_style = custom_styles['body']
    story.append(Paragraph(date_text, info_style))
    story.append(Paragraph(time_text, info_style))
    story.append(Spacer(1, 20))
    
    # خط فاصل
    from reportlab.platypus import HRFlowable
    story.append(HRFlowable(width="100%", thickness=2, color=HexColor('#3498db'), spaceAfter=20))
    
    # الملخص العام
    summary_title = process_arabic_text("الملخص العام")
    story.append(Paragraph(summary_title, custom_styles['subtitle']))
    story.append(Spacer(1, 10))
    
    cleaned_summary = clean_markdown_text(summary)
    summary_points = [p.strip() for p in cleaned_summary.split('\n') if p.strip()]
    
    for point in summary_points:
        processed_point = process_arabic_text(f"• {point}")
        story.append(Paragraph(processed_point, custom_styles['summary']))
        story.append(Spacer(1, 5))
    
    story.append(Spacer(1, 20))
    
    # الصور المرفقة
    images_title = process_arabic_text("الصور المرفوعة")
    story.append(Paragraph(images_title, custom_styles['subtitle']))
    story.append(Spacer(1, 15))
    
    for idx, img in enumerate(images, 1):
        img_resized = img.copy()
        img_resized.thumbnail((350, 350))
        
        img_buffer = io.BytesIO()
        img_resized.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        img_caption = process_arabic_text(f"صورة رقم {idx}")
        story.append(Paragraph(img_caption, custom_styles['body']))
        story.append(Spacer(1, 5))
        
        rl_img = RLImage(img_buffer, width=3.5*inch, height=3.5*inch)
        story.append(rl_img)
        story.append(Spacer(1, 15))
    
    story.append(PageBreak())
    
    # تفاصيل البنود
    details_title = process_arabic_text("تفاصيل البنود")
    story.append(Paragraph(details_title, custom_styles['subtitle']))
    story.append(Spacer(1, 15))
    
    for defect_name, table_md in tables:
        defect_title = process_arabic_text(f"العيب: {defect_name}")
        story.append(Paragraph(defect_title, custom_styles['defect_title']))
        story.append(Spacer(1, 10))
        
        table_obj = markdown_to_enhanced_table(table_md, custom_styles)
        story.append(KeepTogether([table_obj]))
        story.append(Spacer(1, 20))
    
    try:
        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"❌ خطأ في بناء التقرير: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

# ====== 4. واجهة Streamlit ======
st.set_page_config(page_title="⚡ محلل العيوب", layout="wide")
hide = """<style>#MainMenu{visibility:hidden;}footer{visibility:hidden;}header{visibility:hidden;}</style>"""
st.markdown(hide, unsafe_allow_html=True)

st.markdown("<h1 style='text-align:center;'>⚡ محلل العيوب الكهربائية</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align:center;color:grey;'>حمّلي صورك واطّلعي على تقرير مُجمّع في ثواني</h4>", unsafe_allow_html=True)

df = load_excel()

if df.empty:
    st.error("❌ لا يمكن تحميل ملف البيانات. يرجى التحقق من المسار.")
    st.stop()

docs = df_to_docs(df)
vector_db = get_vector_db(docs)

uploaded = st.file_uploader("📷 ارفعي صور العيوب (متعددة):", accept_multiple_files=True, type=["jpg", "jpeg", "png"])
if uploaded:
    cols = st.columns(4)
    images = []
    for idx, file in enumerate(uploaded):
        with cols[idx % 4]:
            img = Image.open(file)
            st.image(img, caption=f"صورة {idx+1}", use_column_width=True)
            images.append(img)

    if st.button("🚀 ابدأ التحليل", type="primary", use_container_width=True):
        bar = st.progress(0)
        images_bytes = [f.getvalue() for f in uploaded]
        all_defects = batch_analyze(images_bytes)
        bar.progress(30)

        unique = list(set(all_defects))
        st.success(f"✅ تم التعرف على {len(unique)} عيب فريد: {', '.join(unique)}")

        seen = set()
        tables = []
        results = []
        for d in unique:
            sim = vector_db.similarity_search(d, k=3)
            band = filter_best_doc(sim, d)
            if band and band not in seen:
                seen.add(band)
                tbl = build_table_from_band(df, band, d)
                tables.append((d, tbl))
                results.append({'query': d, 'doc': sim[0]})
        bar.progress(60)

        combined_queries = '; '.join([r['query'] for r in results])
        qna_template = """
أنت خبير في العيوب الكهربائية. قدم **ملخص عام قصير** للعيوب، مع **أولوية لكل بند** (قصوى: مخاطر سلامة، متوسطة: أداء/تشطيب، عادية: جمالي). **قسّم الملخص إلى جمل واضحة ومستقلة، كل جملة في سطر جديد** لوصف عيب واحد فقط، ولا تضع ترقيم أو بوليت.

### السياق:
{context}

### السؤال:
{question}

### الملخص:
"""
        qna_prompt = PromptTemplate(template=qna_template, input_variables=["context", "question"])
        stuff_chain = load_qa_chain(chat, chain_type="stuff", prompt=qna_prompt)

        context_docs = [r['doc'] for r in results]
        answer = stuff_chain({"input_documents": context_docs, "question": combined_queries}, return_only_outputs=True)
        summary = answer["output_text"]
        bar.progress(90)

        st.subheader("📋 الملخص العام والأولويات")
        st.markdown(summary)

        st.subheader("📊 تفاصيل البنود")
        for defect, tbl in tables:
            with st.expander(f"🔍 {defect}"):
                st.markdown(tbl)

        # ✅ توليد التقرير المحسّن
        pdf_buffer = generate_enhanced_pdf_report(images, summary, tables, unique)
        bar.progress(100)
        
        if pdf_buffer:
            st.success("✅ تم توليد التقرير بنجاح!")
            st.download_button(
                label="📥 تحميل التقرير المحسّن (PDF)",
                data=pdf_buffer,
                file_name=f"تقرير_العيوب_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        else:
            st.error("❌ فشل في توليد التقرير. يرجى المحاولة مرة أخرى.")

# ====== 5. Footer ======
st.markdown("---")
st.markdown("<p style='text-align:center;color:grey;'>⚡ نظام محلل العيوب الكهربائية | تم التطوير بواسطة AI</p>", unsafe_allow_html=True)