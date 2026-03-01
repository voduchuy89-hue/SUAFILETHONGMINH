import streamlit as st
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
import io
import os
from openai import OpenAI
import docx
from docx.shared import Mm, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# ========================================================================================
# CẤU HÌNH TRANG
# ========================================================================================
MAX_FILES = 20  # Tối đa 20 file cùng lúc

st.set_page_config(
    page_title="Trợ lý OCR Thông minh",
    page_icon="📄",
    layout="wide"
)

# ========================================================================================
# HÀM HỖ TRỢ (LOGIC XỬ LÝ)
# ========================================================================================

@st.cache_data  # Sử dụng cache để không xử lý lại file đã xử lý
def process_file(file_bytes, file_extension, show_progress=True):
    """
    Hàm trung tâm xử lý file đầu vào (ảnh hoặc PDF) và trả về văn bản được trích xuất.
    Mặc định sử dụng chế độ song ngữ Việt + Anh.
    show_progress=False dùng khi xử lý hàng loạt nhiều file.
    """
    lang_code = "vie+eng"
    extracted_text = ""
    try:
        if file_extension == 'pdf':
            images = convert_from_bytes(file_bytes)
            all_text = []
            progress_bar = st.progress(0, text="Đang xử lý file PDF...") if show_progress else None
            for i, img in enumerate(images):
                all_text.append(pytesseract.image_to_string(img, lang=lang_code))
                if progress_bar is not None:
                    progress_bar.progress((i + 1) / len(images))
            extracted_text = "\n\n--- Hết trang ---\n\n".join(all_text)
        elif file_extension in ['png', 'jpg', 'jpeg']:
            image = Image.open(io.BytesIO(file_bytes))
            extracted_text = pytesseract.image_to_string(image, lang=lang_code)
        return extracted_text, None
    except Exception as e:
        return None, f"Đã xảy ra lỗi trong quá trình xử lý: {e}"


def call_openai_proofread(text: str) -> str:
    """
    Gọi OpenAI để kiểm tra lỗi, chỉnh sửa văn bản sau OCR.
    Ưu tiên lấy khóa từ Streamlit secrets (OPENAI_API_KEY).
    """
    api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY", None)
    if not api_key:
        raise RuntimeError(
            "Chưa cấu hình OpenAI. Vào Manage app → Settings → Secrets và thêm OPENAI_API_KEY."
        )

    # Đảm bảo thư viện OpenAI đọc được API key từ biến môi trường
    os.environ["OPENAI_API_KEY"] = api_key

    client = OpenAI()
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": (
                    "Bạn là trợ lý giúp chỉnh sửa văn bản tiếng Việt và tiếng Anh "
                    "được trích xuất từ OCR. Hãy sửa lỗi chính tả, dấu câu và cách "
                    "xuống dòng hợp lý. Chỉ trả về văn bản đã chỉnh sửa, không giải thích gì thêm."
                ),
            },
            {"role": "user", "content": text},
        ],
        temperature=0.2,
    )
    return response.choices[0].message.content


def build_docx(text: str) -> bytes:
    """
    Tạo file Word (.docx) khổ A4 từ văn bản.
    """
    buffer = io.BytesIO()
    document = docx.Document()

    # Thiết lập khổ giấy A4
    section = document.sections[0]
    section.page_width = Mm(210)   # A4 ngang 210mm
    section.page_height = Mm(297)  # A4 dọc 297mm

    # Thiết lập font mặc định Times New Roman, cỡ 12
    normal_style = document.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

    for line in text.splitlines():
        document.add_paragraph(line)

    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_excel(text: str) -> bytes:
    """
    Tạo file Excel (.xlsx) với nội dung ở cột A, thiết lập in trên khổ A4.
    """
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "OCR"

    # Thiết lập khổ giấy A4 khi in
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # Font chung cho toàn bộ nội dung
    base_font = Font(name="Times New Roman", size=12)

    lines = text.splitlines()
    for idx, line in enumerate(lines, start=1):
        cell = ws[f"A{idx}"]
        cell.value = line
        cell.alignment = Alignment(wrap_text=True)
        cell.font = base_font

    ws.column_dimensions["A"].width = 100

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ========================================================================================
# GIAO DIỆN CHÍNH CỦA ỨNG DỤNG
# ========================================================================================

st.title("📄 Trợ lý OCR Thông minh")
st.write("Trích xuất văn bản từ file ảnh hoặc PDF. Mặc định xử lý song ngữ Tiếng Việt và Tiếng Anh.")

# Cột cho phần tải lên và hướng dẫn
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_files_all = st.file_uploader(
        f"Tải lên nhiều file (tối đa {MAX_FILES} file cùng lúc)",
        type=['pdf', 'png', 'jpg', 'jpeg'],
        accept_multiple_files=True
    )

with col2:
    with st.expander("💡 Mẹo sử dụng", expanded=True):
        st.info(f"""
        - Tối đa **{MAX_FILES} file** mỗi lần; hỗ trợ song ngữ Việt + Anh.
        - Có thể chọn **Xử lý AI cho tất cả** sau khi OCR xong.
        - Mỗi file: TXT, Word (A4), Excel (A4); font Times New Roman.
        """)

# Giới hạn 20 file, giữ thứ tự
uploaded_files = list(uploaded_files_all)[:MAX_FILES] if uploaded_files_all else []
if uploaded_files_all and len(uploaded_files_all) > MAX_FILES:
    st.warning(f"Chỉ xử lý {MAX_FILES} file đầu tiên. Tổng số file chọn: {len(uploaded_files_all)}.")

# Khởi tạo session state cho kết quả OCR và AI
if "ocr_results" not in st.session_state:
    st.session_state.ocr_results = []
if "ocr_file_keys" not in st.session_state:
    st.session_state.ocr_file_keys = ()
if "ai_results" not in st.session_state:
    st.session_state.ai_results = {}

# Xử lý OCR hàng loạt khi có file mới hoặc đổi danh sách
file_keys = tuple((f.name, f.size) for f in uploaded_files) if uploaded_files else ()
if uploaded_files and file_keys != st.session_state.ocr_file_keys:
    st.session_state.ocr_file_keys = file_keys
    st.session_state.ocr_results = []
    progress_bar = st.progress(0, text="Đang OCR...")
    for idx, uf in enumerate(uploaded_files):
        progress_bar.progress((idx + 1) / len(uploaded_files), text=f"Đang xử lý file {idx + 1}/{len(uploaded_files)}: {uf.name}")
        file_bytes = uf.getvalue()
        ext = uf.name.split('.')[-1].lower()
        text, err = process_file(file_bytes, ext, show_progress=False)
        st.session_state.ocr_results.append({"name": uf.name, "text": text, "error": err})
    progress_bar.empty()
    st.session_state.ai_results = {}  # Reset AI khi đổi bộ file
    st.rerun()

# Hiển thị kết quả từng file
if uploaded_files and st.session_state.ocr_results:
    st.markdown("---")
    st.header("Kết quả trích xuất")

    # Nút xử lý AI cho tất cả file
    run_ai_all = st.button("✨ Xử lý AI cho tất cả các file", type="primary", use_container_width=True)
    if run_ai_all:
        try:
            bar = st.progress(0, text="Đang gọi OpenAI...")
            n = len(st.session_state.ocr_results)
            for i, res in enumerate(st.session_state.ocr_results):
                if res["error"]:
                    continue
                bar.progress((i + 1) / n, text=f"AI đang xử lý {i + 1}/{n}: {res['name']}")
                fixed = call_openai_proofread(res["text"])
                st.session_state.ai_results[res["name"]] = fixed
            bar.empty()
            st.success("Đã xử lý AI xong tất cả file.")
            st.rerun()
        except Exception as e:
            st.error(f"Lỗi khi gọi OpenAI: {e}")

    for i, uploaded_file in enumerate(uploaded_files):
        res = st.session_state.ocr_results[i] if i < len(st.session_state.ocr_results) else None
        if not res:
            continue
        name, text, error = res["name"], res["text"], res["error"]

        with st.expander(f"📄 {name}", expanded=(i < 3)):
            if error:
                st.error(error)
            else:
                st.text_area("Văn bản OCR:", text, height=220, key=f"text_{name}_{i}")

                col_txt, col_docx, col_xlsx = st.columns(3)
                with col_txt:
                    st.download_button(
                        label="📥 Tải TXT",
                        data=text.encode('utf-8'),
                        file_name=f"ket_qua_{name}.txt",
                        mime="text/plain",
                        key=f"dl_txt_{name}_{i}"
                    )
                with col_docx:
                    st.download_button(
                        label="📄 Tải Word (A4)",
                        data=build_docx(text),
                        file_name=f"ket_qua_{name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{name}_{i}"
                    )
                with col_xlsx:
                    st.download_button(
                        label="📊 Tải Excel (A4)",
                        data=build_excel(text),
                        file_name=f"ket_qua_{name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xlsx_{name}_{i}"
                    )

                # Xử lý AI từng file
                use_ai_one = st.button("✨ Dùng AI cho file này", key=f"ai_one_{name}_{i}")
                if use_ai_one:
                    try:
                        with st.spinner("OpenAI đang xử lý..."):
                            fixed_text = call_openai_proofread(text)
                        st.session_state.ai_results[name] = fixed_text
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi OpenAI: {e}")

                # Hiển thị kết quả AI nếu đã có
                fixed_text = st.session_state.ai_results.get(name)
                if fixed_text:
                    st.markdown("---")
                    st.subheader("Văn bản đã được AI hiệu đính")
                    st.text_area("", fixed_text, height=220, key=f"text_ai_{name}_{i}")

                    ai_txt, ai_docx, ai_xlsx = st.columns(3)
                    with ai_txt:
                        st.download_button(
                            label="📥 Tải TXT (AI)",
                            data=fixed_text.encode('utf-8'),
                            file_name=f"ket_qua_AI_{name}.txt",
                            mime="text/plain",
                            key=f"dl_ai_txt_{name}_{i}"
                        )
                    with ai_docx:
                        st.download_button(
                            label="📄 Tải Word (A4, AI)",
                            data=build_docx(fixed_text),
                            file_name=f"ket_qua_AI_{name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_ai_docx_{name}_{i}"
                        )
                    with ai_xlsx:
                        st.download_button(
                            label="📊 Tải Excel (A4, AI)",
                            data=build_excel(fixed_text),
                            file_name=f"ket_qua_AI_{name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_ai_xlsx_{name}_{i}"
                        )
