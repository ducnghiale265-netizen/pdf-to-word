import streamlit as st
import fitz
import os
import re
import tempfile
from pdf2docx import Converter
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ── Kiểm tra PDF có footer "Trang X/Y" không ────────────────────
def pdf_has_page_marker(pdf_path, password=None):
    """
    Quét tất cả trang PDF, kiểm tra có dòng 'Trang X/Y' hay không.
    Trả về True nếu có, False nếu không.
    """
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


# ── Gộp nhiều section về 1 ──────────────────────────────────────
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


# ── Xóa dòng trống cuối file ────────────────────────────────────
def remove_trailing_empty_paragraphs(doc):
    for p in reversed(doc.paragraphs):
        if not p.text.strip():
            p._element.getparent().remove(p._element)
        else:
            break


# ── Tạo field PAGE / NUMPAGES ───────────────────────────────────
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


# ── Thiết lập footer ────────────────────────────────────────────
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


# ── Xóa toàn bộ footer (dùng khi PDF không có số trang) ─────────
def clear_all_footers(doc):
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        footer = section.footer
        for para in footer.paragraphs:
            para._element.getparent().remove(para._element)


# ── Pipeline chính ──────────────────────────────────────────────
def post_process(docx_path, has_marker):
    doc = Document(docx_path)

    # Bước 1: Xóa dòng "Trang X/Y" trong body (nếu có)
    page_pattern = re.compile(r'^Trang\s+\d+(/\d+)?$', re.IGNORECASE)
    to_remove = [p for p in doc.paragraphs if page_pattern.match(p.text.strip())]
    for para in to_remove:
        para._element.getparent().remove(para._element)

    # Bước 2: Xóa dòng trống cuối
    remove_trailing_empty_paragraphs(doc)

    # Bước 3: Gộp về 1 section
    merge_sections_to_one(doc)

    # Bước 4: Footer — chỉ gắn nếu PDF gốc có "Trang X/Y"
    if has_marker:
        set_footer(doc)
    else:
        clear_all_footers(doc)

    doc.save(docx_path)
    return len(to_remove)


# ── Cấu hình trang ──────────────────────────────────────────────
st.set_page_config(
    page_title="Chuyển PDF → Word",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Chuyển PDF sang Word")
st.write("Upload file PDF, nhận về file Word (.docx) giữ nguyên định dạng.")

# ── Upload file ──────────────────────────────────────────────────
uploaded_file = st.file_uploader("Chọn file PDF", type=["pdf"])
password = st.text_input("Mật khẩu PDF (nếu có, để trống nếu không)", type="password")

with st.expander("⚙️ Tùy chọn nâng cao (chọn trang)"):
    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input("Trang bắt đầu (0 = từ đầu)", min_value=0, value=0)
    with col2:
        end_page = st.number_input("Trang kết thúc (0 = đến cuối)", min_value=0, value=0)

# ── Nút chuyển đổi ──────────────────────────────────────────────
if uploaded_file is not None:
    st.info(f"✅ Đã chọn: **{uploaded_file.name}**")

    if st.button("🚀 Bắt đầu chuyển đổi", use_container_width=True):
        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path  = os.path.join(tmp_dir, uploaded_file.name)
            docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
            docx_path = os.path.join(tmp_dir, docx_name)

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.read())

            # Kiểm tra PDF
            try:
                doc = fitz.open(pdf_path)
                if doc.is_encrypted:
                    if not password or not doc.authenticate(password):
                        st.error("❌ PDF được bảo vệ — mật khẩu sai hoặc chưa nhập.")
                        st.stop()
                total_pages = len(doc)
                doc.close()
                st.write(f"📑 Tổng số trang PDF: **{total_pages}**")
            except Exception as e:
                st.error(f"❌ Không đọc được file PDF: {e}")
                st.stop()

            with st.spinner("⏳ Đang chuyển đổi, vui lòng chờ..."):
                try:
                    # Phát hiện PDF có footer "Trang X/Y" không
                    has_marker = pdf_has_page_marker(
                        pdf_path,
                        password if password else None
                    )

                    start_0 = (start_page - 1) if start_page > 0 else 0
                    end_0   = end_page if end_page > 0 else None

                    cv = Converter(pdf_path, password=password if password else None)
                    cv.convert(docx_path, start=start_0, end=end_0)
                    cv.close()

                    count_marker = post_process(docx_path, has_marker)

                    if has_marker:
                        st.info(f"🔧 Phát hiện số trang trong PDF → đã chuyển vào footer Word.")
                    else:
                        st.info("📄 PDF không có số trang → file Word cũng không có footer.")

                    with open(docx_path, "rb") as f:
                        docx_bytes = f.read()

                    st.success("✅ Chuyển đổi thành công!")
                    st.download_button(
                        label="⬇️ Tải file Word về",
                        data=docx_bytes,
                        file_name=docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Lỗi khi chuyển đổi: {e}")
else:
    st.warning("👆 Hãy upload file PDF để bắt đầu.")
