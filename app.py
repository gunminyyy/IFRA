import streamlit as st
import pdfplumber
from docxtpl import DocxTemplate
import os
import re
from io import BytesIO

# --- 헬퍼 함수 ---

def extract_text_between(text, start_keyword, end_keyword):
    """두 키워드 사이의 텍스트를 추출하는 함수"""
    
    def flexible_escape(kw):
        # 정규식 특수문자를 이스케이프 처리한 뒤, 띄어쓰기를 모든 공백문자 허용으로 변경
        escaped = re.escape(kw)
        escaped = escaped.replace(r'\ ', r'\s+')
        return escaped

    start_pattern = flexible_escape(start_keyword)
    end_pattern = flexible_escape(end_keyword)
    
    pattern = f"{start_pattern}(.*?){end_pattern}"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    
    if match:
        extracted = match.group(1).strip()
        return process_value(extracted)
    
    return "Not Permitted"

def process_value(val_str):
    """추출된 텍스트에서 숫자, Not Permitted, Not Restricted를 분류하여 변환"""
    if not val_str:
        return "Not Permitted"
        
    val_lower = val_str.lower()
    
    # 1. 'Not permitted' 처리
    if "not" in val_lower and "permitted" in val_lower:
        return "Not Permitted"
        
    # 2. 'Not Restricted' 처리 (원본 PDF의 제한 없음 문구 보존)
    if "not" in val_lower and "restricted" in val_lower:
        return "Not Restricted"

    # 3. 숫자 추출 및 소수점 2자리 포맷팅
    num_match = re.search(r'\d+\.?\d*', val_str)
    
    if not num_match:
        return "Not Permitted"
        
    try:
        clean_str = num_match.group(0)
        val_float = float(clean_str)
        
        if val_float == 0.0:
            return "Not Permitted"
        else:
            s = str(val_float)
            if '.' in s:
                int_part, dec_part = s.split('.')
                dec_part = dec_part[:2]
                dec_part = dec_part.ljust(2, '0')
                return f"{int_part}.{dec_part}"
            else:
                return f"{s}.00"
    except ValueError:
        return "Not Permitted"

# --- 메인 로직 ---

def process_pdf_to_word(pdf_file, customer_name, product_name, mode):
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"

    # 모드별 Context 딕셔너리 분리 (요청하신 정확한 키워드 반영)
    if mode == "CFF":
        context = {
            "CUSTOMER": customer_name,
            "PRODUCT": product_name,
            "CATEGORY1": extract_text_between(full_text, "Category 1", "Category 2"),
            "CATEGORY2": extract_text_between(full_text, "Category 2", "Category 3"),
            "CATEGORY3": extract_text_between(full_text, "Category 3", "Category 4"),
            "CATEGORY4": extract_text_between(full_text, "Category 4", "Category 5.A"),
            "CATEGORY5_A": extract_text_between(full_text, "Category 5.A", "Category 5.B"),
            "CATEGORY5_B": extract_text_between(full_text, "Category 5.B", "Category 5.C"),
            "CATEGORY5_C": extract_text_between(full_text, "Category 5.C", "Category 5.D"),
            "CATEGORY5_D": extract_text_between(full_text, "Category 5.D", "Category 6"),
            "CATEGORY6": extract_text_between(full_text, "Category 6", "Category 7.A"),
            "CATEGORY7_A": extract_text_between(full_text, "Category 7.A", "Category 7.B"),
            "CATEGORY7_B": extract_text_between(full_text, "Category 7.B", "Category 8"),
            "CATEGORY8": extract_text_between(full_text, "Category 8", "Category 9"),
            "CATEGORY9": extract_text_between(full_text, "Category 9", "Category 10.A"),
            "CATEGORY10_A": extract_text_between(full_text, "Category 10.A", "Category 10.B"),
            "CATEGORY10_B": extract_text_between(full_text, "Category 10.B", "Category 11.A"),
            "CATEGORY11_A": extract_text_between(full_text, "Category 11.A", "Category 11.B"),
            "CATEGORY11_B": extract_text_between(full_text, "Category 11.B", "Category 12"),
            "CATEGORY12": extract_text_between(full_text, "Category 12", "For other")
        }
    elif mode == "HP":
        context = {
            "CUSTOMER": customer_name,
            "PRODUCT": product_name,
            "CATEGORY1": extract_text_between(full_text, "Category 1*", "Category 2"),
            "CATEGORY2": extract_text_between(full_text, "Category 2", "Category 3"),
            "CATEGORY3": extract_text_between(full_text, "Category 3", "Category 4"),
            "CATEGORY4": extract_text_between(full_text, "Category 4", "Category 5.A"),
            "CATEGORY5_A": extract_text_between(full_text, "Category 5.A", "Category 5.B"),
            "CATEGORY5_B": extract_text_between(full_text, "Category 5.B", "Category 5.C"),
            "CATEGORY5_C": extract_text_between(full_text, "Category 5.C", "Category 5.D"),
            "CATEGORY5_D": extract_text_between(full_text, "Category 5.D", "Category 6*"),
            "CATEGORY6": extract_text_between(full_text, "Category 6*", "Category 7.A"),
            "CATEGORY7_A": extract_text_between(full_text, "Category 7.A", "Category 7.B"),
            "CATEGORY7_B": extract_text_between(full_text, "Category 7.B", "Category 8"),
            "CATEGORY8": extract_text_between(full_text, "Category 8", "Category 9"),
            "CATEGORY9": extract_text_between(full_text, "Category 9", "Category 10.A"),
            "CATEGORY10_A": extract_text_between(full_text, "Category 10.A", "Category 10.B"),
            "CATEGORY10_B": extract_text_between(full_text, "Category 10.B", "Category 11.A"),
            "CATEGORY11_A": extract_text_between(full_text, "Category 11.A", "Category 11.B"),
            "CATEGORY11_B": extract_text_between(full_text, "Category 11.B", "Category 12"),
            "CATEGORY12": extract_text_between(full_text, "Category 12", "*Only fragrance")
        }

    template_path = "templates/IFRA.docx"
    
    if not os.path.exists(template_path):
        st.error(f"오류: 템플릿 파일이 없습니다. '{template_path}' 경로를 확인해주세요.")
        return None

    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        output_io = BytesIO()
        doc.save(output_io)
        output_io.seek(0)
        
        return output_io
    except Exception as e:
        st.error(f"템플릿 렌더링 중 오류가 발생했습니다: {e}")
        return None

# --- Streamlit UI 구성 ---

st.set_page_config(page_title="PDF to Word Converter", layout="wide")
st.title("IFRA PDF -> Word 자동 변환기")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 업로드")
    uploaded_pdf = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])

with col2:
    st.subheader("2. 정보 입력")
    customer_input = st.text_input("고객사명")
    product_input = st.text_input("제품명")
    mode_selection = st.selectbox("모드 선택", options=["CFF", "HP"])

st.divider()

col_btn_left, col_btn_right = st.columns(2)

with col_btn_left:
    convert_clicked = st.button("변환 시작", type="primary", use_container_width=True)

with col_btn_right:
    download_placeholder = st.empty()

# --- 변환 버튼 클릭 시 동작 ---
if convert_clicked:
    if uploaded_pdf is None:
        st.warning("PDF 파일을 먼저 업로드해주세요.")
    elif not customer_input or not product_input:
        st.warning("고객사명과 제품명을 모두 입력해주세요.")
    else:
        with st.spinner("변환 작업 진행 중..."):
            result_docx = process_pdf_to_word(uploaded_pdf, customer_input, product_input, mode_selection)
            
            if result_docx:
                st.success("변환 성공!")
                
                # 파일명을 "제품명 IFRA 51TH.docx" 로 지정
                file_name = f"{product_input} IFRA 51TH.docx"
                download_placeholder.download_button(
                    label="결과물 다운로드",
                    data=result_docx,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

