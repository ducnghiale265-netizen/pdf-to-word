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

# ── HÀM CHUYỂN PDF SANG PPTX SIÊU CẤP CHỐNG ĐÈ CHỮ (FLOW LAYOUT) ──
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
        
        # 1. TRÍCH XUẤT HÌNH ẢNH / LOGO NỀN (Vẽ độc lập tuyệt đối)
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

        # 2. GOM TẤT CẢ PHẦN TỬ NỘI DUNG (BẢNG + CHỮ) ĐỂ XẾP HÀNG THEO DÒNG CHẢY
        tables = pdf_page.find_tables()
        all_elements = []
        table_bboxes = []
        
        # Lấy danh sách các bảng biểu
        for t in tables:
            data = t.extract() 
            if not data or len(data) == 0 or len(data[0]) == 0:
                continue
            table_bboxes.append(t.bbox)
            all_elements.append({
                "type": "table",
                "bbox": t.bbox,
                "data": data
            })
            
        # Lấy danh sách các khối chữ nằm ngoài bảng
        page_dict = pdf_page.get_text("dict")
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
                    # Kiểm tra chuỗi text thực tế xem có rỗng không
                    text_content = "".join([span.get("text", "") for line in block["lines"] for span in line.get("spans", "")]).strip()
                    if text_content:
                        all_elements.append({
                            "type": "text_block",
                            "bbox": bbox,
                            "block_data": block
                        })
                        
        # SẮP XẾP TẤT CẢ PHẦN TỬ THEO THỨ TỰ TỪ TRÊN XUỐNG DƯỚI (QUAN TRỌNG)
        all_elements.sort(key=lambda e: e["bbox"][1])
        
        # 3. DỰNG CÁC PHẦN TỬ THEO CƠ CHẾ CUỐN CHIẾU CHỐNG ĐÈ CHỮ
        last_bottom_y = 20  # Lề trên mặc định ban đầu
        
        for elem in all_elements:
            orig_x0, orig_y0, orig_x1, orig_y1 = elem["bbox"]
            
            # Tính toán vị trí top mới: Không bao giờ được đè lên phần tử phía trên nó
            if last_bottom_y == 20:
                render_y0 = orig_y0
            else:
                # Đảm bảo cách phần tử ngay trên nó ít nhất 12 điểm (points)
                render_y0 = max(orig_y0, last_bottom_y + 12)
            
            # --- Trường hợp dựng Khối chữ ---
            if elem["type"] == "text_block":
                block = elem["block_data"]
                num_lines = len(block["lines"])
                
                # Tìm cỡ chữ lớn nhất trong khối để tính toán độ co dãn chiều cao
                max_size = 11
                for line in block["lines"]:
                    for span in line.get("spans", []):
                        max_size = max(max_size, span.get("size", 11))
                        
                # Ước lượng chiều cao thực tế hộp chữ sau khi render paragraph
                estimated_height = (num_lines * max_size * 1.3) + 10
                
                b_left = Inches(orig_x0 / 72)
                b_top = Inches(render_y0 / 72)
                b_width = Inches((orig_x1 - orig_x0 + 45) / 72)  # Cộng 45pt tránh rớt dòng lỗi bề ngang
                b_height = Inches(estimated_height / 72)
                
                txBox = slide.shapes.add_textbox(b_left, b_top, b_width, b_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
                
                is_first_para = True
                for line in block["lines"]:
                    if not is_first_para:
                        p = tf.add_paragraph()
                    else:
                        p = tf.paragraphs[0]
                        is_first_para = False
                        
                    # Triệt tiêu margin thừa của Office gây nở chữ đè hàng
                    p.space_before = PptxPt(0)
                    p.space_after = PptxPt(2)
                    p.line_spacing = 1.05
                    
                    for span in line.get("spans", []):
                        run = p.add_run()
                        run.text = span.get("text", "")
                        run.font.size = PptxPt(max(span.get("size", 11) * 0.95, 8)) # Hạ font defensive 5%
                        run.font.name = "Times New Roman"
                        if "bold" in span.get("font", "").lower():
                            run.font.bold = True
                            
                # Lưu lại vị trí đáy thực tế của khối chữ này
                last_bottom_y = render_y0 + estimated_height
                
            # --- Trường hợp dựng Bảng biểu ---
            elif elem["type"] == "table":
                data = elem["data"]
                rows = len(data)
                cols = len(data[0])
                
                left = Inches(orig_x0 / 72)
                top = Inches(render_y0 / 72)
                width = Inches((orig_x1 - orig_x0) / 72)
                
                estimated_table_height = max(orig_y1 - orig_y0, rows * 24)
                height = Inches(estimated_table_height / 72)
                
                table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
                pptx_table = table_shape.table
                
                for r_idx, row_data in enumerate(data):
                    for c_idx, cell_value in enumerate(row_data):
                        cell = pptx_table.cell(r_idx, c_idx)
                        cell.text = str(cell_value) if cell_value is not None else ""
                        
                        # Khóa chặt biên ô bảng không cho tràn chữ tự do
                        cell.margin_left = Inches(0.04)
                        cell.margin_right = Inches(0.04)
                        cell.margin_top = Inches(0.04)
                        cell.margin_bottom = Inches(0.04)
                        
                        for p in cell.text_frame.paragraphs:
                            p.font.name = "Times New Roman"
                            p.font.size = PptxPt(9.5)
                            p.space_before = PptxPt(0)
                            p.space_after = PptxPt(0)
                            p.line_spacing = 1.0
                            
                # Lưu lại vị trí đáy thực tế của bảng biểu này
                last_bottom_y = render_y0 + estimated_table_height
                        
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
                
                with st.spinner("⏳ Đang tối ưu hóa dòng chảy layout, vui lòng đợi..."):
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
