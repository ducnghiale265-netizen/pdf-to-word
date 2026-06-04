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

# ── Lấy nguyên từ notebook Merge_Hợp_Đồng ───────────────────────
def add_complex_field(run, instruction):
    r = run._r

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar_begin)

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = instruction
    r.append(instrText)

    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    r.append(fldChar_separate)

    text = OxmlElement('w:t')
    text.text = "1"
    r.append(text)

    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    r.append(fldChar_end)


def fix_and_style_footer(doc_path):
    doc = Document(doc_path)

    # Xóa dòng "Trang X/Y" trong body trước
    page_pattern = re.compile(r'^Trang\s+\d+(/\d+)?$', re.IGNORECASE)
    to_remove = [p for p in doc.paragraphs if page_pattern.match(p.text.strip())]
    for para in to_remove:
        para._element.getparent().remove(para._element)

    # Thiết lập lại footer cho tất cả sections
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        footer = section.footer

        # Xóa footer cũ
        for para in footer.paragraphs:
            para._element.getparent().remove(para._element)

        # Footer mới
        para = footer.add_paragraph()
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        run1 = para.add_run("Trang ")
        run_page = para.add_run()
        add_complex_field(run_page, " PAGE ")
        run2 = para.add_run("/")
        run_total = para.add_run()
        add_complex_field(run_total, " NUMPAGES ")

        # Format font
        for run in para.runs:
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

    doc.save(doc_path)
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
                st.write(f"📑 Tổng số trang: **{total_pages}**")
            except Exception as e:
                st.error(f"❌ Không đọc được file PDF: {e}")
                st.stop()

            # Chuyển đổi
            with st.spinner("⏳ Đang chuyển đổi, vui lòng chờ..."):
                try:
                    start_0 = (start_page - 1) if start_page > 0 else 0
                    end_0   = end_page if end_page > 0 else None

                    cv = Converter(pdf_path, password=password if password else None)
                    cv.convert(docx_path, start=start_0, end=end_0)
                    cv.close()

                    # Fix footer (xóa "Trang X/Y" trong body + đưa vào footer thật)
                    moved = fix_and_style_footer(docx_path)
                    if moved > 0:
                        st.info(f"🔧 Đã tự động chuyển {moved} dòng số trang vào footer.")

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
