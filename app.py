import streamlit as st
import fitz
import os
import re
import tempfile
import io
from pdf2docx import Converter
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Thêm thư viện xử lý PPTX
from pdf2image import convert_from_path
from pptx import Presentation

# ── Kiểm tra PDF có footer "Trang X/Y" không ────────────────────
def pdf_has_page_marker(pdf_path, password=None):
    page_pattern = re.compile(r'Trang\s+\d+(/\d+)?', re.IGNORECASE)
    try:
        doc = fitz.open(pdf_path)
        if password:
            doc.authenticate(password)
        for page in doc:
            text = page.get_text()
            if page_pattern.search(text):
                doc.close()
                return True
        doc.close()
    except Exception:
        pass
    return False

# ── Gộp nhiều section về 1 (Cho Word) ───────────────────────────
def merge_sections_to_one(doc):
    body = doc.element.body
    removed = 0
    for para in body.findall('.//' + qn('w:p')):
        pPr = para.find(qn('w:pPr'))
        if pPr is not None:
            sectPr = pPr.find(qn('w:sectPr'))
            if sectPr is not None:
                pPr.remove(sectPr)
                removed += 1
    return removed

# ── Xóa dòng trống cuối file (Cho Word) ─────────────────────────
def remove_trailing_empty_paragraphs(doc):
    for p in reversed(doc.paragraphs):
        if not p.text.strip():
            p._element.getparent().remove(p._element)
        else:
            break

# ── Tạo field PAGE / NUMPAGES (Cho Word) ────────────────────────
def add_complex_field(run, instruction):
    r = run._r
    fc_begin = OxmlElement('w:fldChar')
    fc_begin.set(qn('w:fldCharType'), 'begin')
    r.append(fc_begin)

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = instruction
    r.append(instr)

    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    r.append(fc_sep)

    wt = OxmlElement('w:t')
    wt.text = '1'
    r.append(wt)

    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r.append(fc_end)

# ── Thiết lập footer (Cho Word) ──────────────────────────────────
def set_footer(doc):
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        footer = section.footer

        for para in footer.paragraphs:
            para._element.getparent().remove(para._element)

        fp = footer.add_paragraph()
        fp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        fp.add_run("Trang ")
        run_page = fp.add_run()
        add_complex_field(run_page, " PAGE ")
        fp.add_run("/")
        run_total = fp.add_run()
        add_complex_field(run_total, " NUMPAGES ")

        for run in fp.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.italic = True
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)
            rFonts.set(qn("w:ascii"),    "Times New Roman")
            rFonts.set(qn("w:hAnsi"),    "Times New Roman")
            rFonts.set(qn("w:eastAsia"), "Times New Roman")

# ── Xóa toàn bộ footer (Cho Word) ───────────────────────────────
def clear_all_footers(doc):
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        footer = section.footer
        for para in footer.paragraphs:
            para._element.getparent().remove(para._element)

# ── Hậu xử lý Word ──────────────────────────────────────────────
def post_process_word(docx_path, has_marker):
    doc = Document(docx_path)
    page_pattern = re.compile(r'^Trang\s+\d+(/\d+)?$', re.IGNORECASE)
    to_remove = [p for p in doc.paragraphs if page_pattern.match(p.text.strip())]
    for para in to_remove:
        para._element.getparent().remove(para._element)

    remove_trailing_empty_paragraphs(doc)
    merge_sections_to_one(doc)

    if has_marker:
        set_footer(doc)
    else:
        clear_all_footers(doc)

    doc.save(docx_path)

# ── HÀM CHUYỂN PDF SANG PPTX GIỮ NGUYÊN LAYOUT ĐỊNH DẠNG ────────
def convert_pdf_to_pptx(pdf_path, start_page, end_page, password=None):
    """
    Chuyển đổi từng trang PDF thành hình ảnh chất lượng cao 
    và chèn vào các slide trống của PowerPoint.
    """
    # Thiết lập khoảng trang để chuyển đổi ảnh
    # pdf2image dùng index từ 1
    first_p = start_page if start_page > 0 else 1
    last_p = end_page if end_page > 0 else None
    
    # Chuyển PDF sang danh sách ảnh (DPI 150-200 là tối ưu để cân bằng dung lượng và độ nét)
    images = convert_from_path(
        pdf_path, 
        dpi=150, 
        first_page=first_p, 
        last_page=last_p, 
        userpw=password
    )
    
    prs = Presentation()
    
    # Lặp qua từng trang ảnh để đưa vào slide
    for img in images:
        # Chọn layout số 6 (Slide trống hoàn toàn - Blank layout)
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Chuyển đổi ảnh sang dạng bytes trong bộ nhớ để không phải lưu file tạm
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        
        # Chèn ảnh phủ kín toàn bộ bề mặt slide PowerPoint
        slide.shapes.add_picture(
            img_byte_arr, 
            0, 0, 
            width=prs.slide_width, 
            height=prs.slide_height
        )
        
    # Lưu file PowerPoint vào luồng dữ liệu bytes
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io.getvalue()


# ── Cấu hình trang Streamlit ────────────────────────────────────
st.set_page_config(
    page_title="Bộ chuyển đổi PDF đa năng",
    page_icon="🛠️",
    layout="centered"
)

# ── Thanh Menu bên cạnh (Sidebar) để chọn chức năng ──────────────
st.sidebar.title("⚙️ Bảng điều khiển")
conversion_type = st.sidebar.radio(
    "Chọn định dạng đầu ra:",
    ("Chuyển sang Word (.docx)", "Chuyển sang PowerPoint (.pptx)")
)

st.title("🛠️ Bộ chuyển đổi PDF đa năng")
st.write(f"Đang chọn chế độ: **{conversion_type}**")

# ── Upload file ──────────────────────────────────────────────────
uploaded_file = st.file_uploader("Chọn file PDF cần xử lý", type=["pdf"])
password = st.text_input("Mật khẩu PDF (nếu có, để trống nếu không)", type="password")

with st.expander("📐 Tùy chọn nâng cao (Giới hạn số trang)"):
    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input("Trang bắt đầu (0 = từ đầu)", min_value=0, value=0)
    with col2:
        end_page = st.number_input("Trang kết thúc (0 = đến cuối)", min_value=0, value=0)

# ── Nút xử lý chính ──────────────────────────────────────────────
if uploaded_file is not None:
    st.info(f"✅ Đã nhận file: **{uploaded_file.name}**")

    if st.button("🚀 Bắt đầu chuyển đổi", use_container_width=True):
        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path = os.path.join(tmp_dir, uploaded_file.name)

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.read())

            # Kiểm tra tính hợp lệ và đếm số trang của PDF
            try:
                doc = fitz.open(pdf_path)
                if doc.is_encrypted:
                    if not password or not doc.authenticate(password):
                        st.error("❌ PDF được bảo vệ — mật khẩu sai hoặc chưa nhập.")
                        st.stop()
                total_pages = len(doc)
                doc.close()
                st.write(f"📑 Tổng số trang PDF phát hiện: **{total_pages}**")
            except Exception as e:
                st.error(f"❌ Không đọc được file PDF: {e}")
                st.stop()

            # ── XỬ LÝ NHÁNH 1: CHUYỂN SANG WORD ────────────────────────
            if conversion_type == "Chuyển sang Word (.docx)":
                docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
                docx_path = os.path.join(tmp_dir, docx_name)
                
                with st.spinner("⏳ Đang phân tích cấu trúc và chuyển sang Word..."):
                    try:
                        has_marker = pdf_has_page_marker(pdf_path, password if password else None)
                        start_0 = (start_page - 1) if start_page > 0 else 0
                        end_0 = end_page if end_page > 0 else None

                        cv = Converter(pdf_path, password=password if password else None)
                        cv.convert(docx_path, start=start_0, end=end_0)
                        cv.close()

                        post_process_word(docx_path, has_marker)

                        with open(docx_path, "rb") as f:
                            docx_bytes = f.read()

                        st.success("✅ Đã chuyển đổi sang Word thành công!")
                        st.download_button(
                            label="⬇️ Tải file Word về",
                            data=docx_bytes,
                            file_name=docx_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Lỗi khi chuyển đổi sang Word: {e}")

            # ── XỬ LÝ NHÁNH 2: CHUYỂN SANG POWERPOINT ──────────────────
            elif conversion_type == "Chuyển sang PowerPoint (.pptx)":
                pptx_name = os.path.splitext(uploaded_file.name)[0] + ".pptx"
                
                with st.spinner("⏳ Đang chụp màn hình PDF và dựng slide PowerPoint..."):
                    try:
                        pptx_bytes = convert_pdf_to_pptx(
                            pdf_path=pdf_path,
                            start_page=start_page,
                            end_page=end_page,
                            password=password if password else None
                        )
                        
                        st.success("✅ Đã chuyển đổi sang PowerPoint giữ nguyên Layout thành công!")
                        st.download_button(
                            label="⬇️ Tải file PowerPoint về",
                            data=pptx_bytes,
                            file_name=pptx_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Lỗi khi chuyển đổi sang PowerPoint: {e}")
else:
    st.warning("👆 Hãy tải file PDF lên để hệ thống bắt đầu làm việc.")
