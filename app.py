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

# Các thư viện phục vụ cho PowerPoint
from pptx import Presentation
from pptx.util import Inches, Pt as PptxPt

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

# ── HÀM CHUYỂN PDF SANG PPTX NÂNG CAO (THUẬT TOÁN CHỐNG ĐÈ CHỮ) ──
def convert_pdf_to_pptx(pdf_path, start_page, end_page, password=None):
    prs = Presentation()
    blank_layout = prs.slide_layouts[6] 
    
    doc = fitz.open(pdf_path)
    if password:
        doc.authenticate(password)
        
    total_p = len(doc)
    p_start = max(0, start_page - 1) if start_page > 0 else 0
    p_end = min(total_p, end_page) if end_page > 0 else total_p
    
    for page_idx in range(p_start, p_end):
        pdf_page = doc[page_idx]
        page_width = pdf_page.rect.width
        page_height = pdf_page.rect.height
        
        prs.slide_width = Inches(page_width / 72)
        prs.slide_height = Inches(page_height / 72)
        
        slide = prs.slides.add_slide(blank_layout)
        
        # 1. TRÍCH XUẤT VÀ DỰNG BẢNG NATIVE
        tables = pdf_page.find_tables()
        table_bboxes = []
        
        for t in tables:
            data = t.extract() 
            if not data:
                continue
            rows = len(data)
            cols = len(data[0]) if rows > 0 else 0
            if rows == 0 or cols == 0:
                continue
            
            table_bboxes.append(t.bbox)
            
            left = Inches(t.bbox[0] / 72)
            top = Inches(t.bbox[1] / 72)
            width = Inches((t.bbox[2] - t.bbox[0]) / 72)
            height = Inches((t.bbox[3] - t.bbox[1]) / 72)
            
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
            pptx_table = table_shape.table
            
            for r_idx, row_data in enumerate(data):
                for c_idx, cell_value in enumerate(row_data):
                    cell = pptx_table.cell(r_idx, c_idx)
                    cell.text = str(cell_value) if cell_value is not None else ""
                    for p in cell.text_frame.paragraphs:
                        p.font.name = "Times New Roman"
                        p.font.size = PptxPt(10)
                        
        # 2. TRÍCH XUẤT HÌNH ẢNH / LOGO
        image_infos = pdf_page.get_image_info(hashes=False, xrefs=True)
        for img_info in image_infos:
            xref = img_info['xref']
            if xref > 0:
                try:
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    bbox = img_info['bbox']
                    slide.shapes.add_picture(
                        io.BytesIO(image_bytes), 
                        Inches(bbox[0] / 72), 
                        Inches(bbox[1] / 72), 
                        Inches((bbox[2] - bbox[0]) / 72), 
                        Inches((bbox[3] - bbox[1]) / 72)
                    )
                except Exception:
                    pass
                    
        # 3. LỌC BỎ CHỮ TRONG BẢNG & SẮP XẾP KHỐI CHỮ NGOÀI BẢNG
        page_dict = pdf_page.get_text("dict")
        text_blocks = []
        
        for block in page_dict.get("blocks", []):
            if "lines" in block:
                bbox = block["bbox"]
                cx = (bbox[0] + bbox[2]) / 2
                cy = (bbox[1] + bbox[3]) / 2
                
                inside_table = False
                for t_box in table_bboxes:
                    if t_box[0] <= cx <= t_box[2] and t_box[1] <= cy <= t_box[3]:
                        inside_table = True
                        break
                
                if not inside_table:
                    text_blocks.append(block)
        
        # Sắp xếp các khối văn bản từ trên xuống dưới
        text_blocks.sort(key=lambda b: b["bbox"][1])
        
        # Nhóm các khối văn bản liên tiếp lại với nhau để đưa chung vào 1 TextBox
        grouped_boxes = []
        current_group = []
        
        for b in text_blocks:
            if not current_group:
                current_group.append(b)
            else:
                prev_b = current_group[-1]
                prev_bottom = prev_b["bbox"][3]
                curr_top = b["bbox"][1]
                
                # Kiểm tra xem có bảng biểu nào nằm chen giữa 2 khối chữ này không
                has_table_between = False
                for t_box in table_bboxes:
                    if prev_bottom <= t_box[1] and t_box[3] <= curr_top:
                        has_table_between = True
                        break
                
                # Nếu khoảng cách gần (< 45 pt) và không bị bảng ngăn cách -> Gộp nhóm
                if (curr_top - prev_bottom < 45) and not has_table_between:
                    current_group.append(b)
                else:
                    grouped_boxes.append(current_group)
                    current_group = [b]
        if current_group:
            grouped_boxes.append(current_group)
            
        # Dựng các nhóm văn bản lên Slide PowerPoint
        for group in grouped_boxes:
            g_x0 = min(b["bbox"][0] for b in group)
            g_y0 = min(b["bbox"][1] for b in group)
            g_x1 = max(b["bbox"][2] for b in group)
            g_y1 = max(b["bbox"][3] for b in group)
            
            # Thêm buffer 35 điểm chiều ngang để hộp chữ rộng rãi, chống rớt dòng lỗi
            margin_buffer = 35 
            b_left = Inches(g_x0 / 72)
            b_top = Inches(g_y0 / 72)
            b_width = Inches((g_x1 - g_x0 + margin_buffer) / 72)
            b_height = Inches((g_y1 - g_y0) / 72)
            
            txBox = slide.shapes.add_textbox(b_left, b_top, b_width, b_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            
            is_first_para = True
            for b in group:
                for line in b["lines"]:
                    if not is_first_para:
                        p = tf.add_paragraph()
                    else:
                        p = tf.paragraphs[0]
                        is_first_para = False
                        
                    for span in line.get("spans", []):
                        run = p.add_run()
                        run.text = span.get("text", "")
                        # Thu nhỏ font nhẹ 5% để chữ nằm gọn gàng lý tưởng trong khung vẽ
                        run.font.size = PptxPt(max(span.get("size", 11) * 0.95, 8))
                        run.font.name = "Times New Roman"
                        # Tự động giữ định dạng chữ ĐẬM nếu bản gốc có
                        if "bold" in span.get("font", "").lower():
                            run.font.bold = True
                        
    doc.close()
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

st.sidebar.title("⚙️ Bảng điều khiển")
conversion_type = st.sidebar.radio(
    "Chọn định dạng đầu ra:",
    ("Chuyển sang Word (.docx)", "Chuyển sang PowerPoint (.pptx)")
)

st.title("🛠️ Bộ chuyển đổi PDF đa năng")
st.write(f"Đang chọn chế độ: **{conversion_type}**")

uploaded_file = st.file_uploader("Chọn file PDF cần xử lý", type=["pdf"])
password = st.text_input("Mật khẩu PDF (nếu có, để trống nếu không)", type="password")

with st.expander("📐 Tùy chọn nâng cao (Giới hạn số trang)"):
    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input("Trang bắt đầu (0 = từ đầu)", min_value=0, value=0)
    with col2:
        end_page = st.number_input("Trang kết thúc (0 = đến cuối)", min_value=0, value=0)

if uploaded_file is not None:
    st.info(f"✅ Đã nhận file: **{uploaded_file.name}**")

    if st.button("🚀 Bắt đầu chuyển đổi", use_container_width=True):
        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path = os.path.join(tmp_dir, uploaded_file.name)

            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.read())

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

            # Chuyển sang Word
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

            # Chuyển sang PowerPoint
            elif conversion_type == "Chuyển sang PowerPoint (.pptx)":
                pptx_name = os.path.splitext(uploaded_file.name)[0] + ".pptx"
                
                with st.spinner("⏳ Đang tối ưu khối chữ và dựng slide chống đè chữ..."):
                    try:
                        pptx_bytes = convert_pdf_to_pptx(
                            pdf_path=pdf_path,
                            start_page=start_page,
                            end_page=end_page,
                            password=password if password else None
                        )
                        
                        st.success("✅ Đã chuyển đổi sang PowerPoint thành công!")
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
