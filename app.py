import streamlit as st
import fitz
import os
import tempfile
from pdf2docx import Converter

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

# ── Tùy chọn trang ──────────────────────────────────────────────
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
            # Lưu PDF tạm thời
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

                    # Đọc file kết quả để cho download
                    with open(docx_path, "rb") as f:
                        docx_bytes = f.read()

                    st.success(f"✅ Chuyển đổi thành công!")
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