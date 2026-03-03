import streamlit as st
import pdfplumber
from docxtpl import DocxTemplate
import os
import re
from io import BytesIO

# --- 헬퍼 함수 ---

def extract_text_between(text, start_keyword, end_keyword):
    """두 키워드 사이의 텍스트를 추출하는 함수"""
    # 이스케이프 처리를 통해 정규표현식 오류 방지
    start_esc = re.escape(start_keyword)
    end_esc = re.escape(end_keyword)
    
    # re.DOTALL을 사용하여 줄바꿈 포함 검색
    pattern = f"{start_esc}(.*?){end_esc}"
    match = re.search(pattern, text, re.DOTALL)
    
    if match:
        extracted = match.group(1).strip()
        return process_value(extracted)
    return "Not Permitted" # 매칭 안 되면 기본값

def process_value(val_str):
    """추출된 값을 분석하여 소수점 둘째 자리 또는 Not Permitted로 변환"""
    try:
        # 숫자 추출 (쉼표 제거 등)
        clean_str = re.sub(r'[^\d.]', '', val_str)
        if not clean_str:
             return "Not Permitted"
             
        val_float = float(clean_str)
        
        # 0이거나 0.0일 경우 Not Permitted
        if val_float == 0.0:
            return "Not Permitted"
        else:
            # 반올림 없이(절사) 소수점 2자리까지만 표시
            s = str(val_float)
            if '.' in s:
                int_part, dec_part = s.split('.')
                dec_part = dec_part[:2] # 소수점 2자리까지만 절사
                dec_part = dec_part.ljust(2, '0') # 1자리일 경우 0을 채움
                return f"{int_part}.{dec_part}"
            else:
                return f"{s}.00"
    except ValueError:
        # 숫자로 변환할 수 없는 경우 (예: 이미 텍스트로 Not Permitted 등이 적혀있을 때)
        if "Not" in val_str or "permitted" in val_str.lower():
            return "Not Permitted"
        return val_str # 그 외 텍스트는 그대로 반환

# --- 메인 로직 ---

def process_pdf_to_word(pdf_file, customer_name, product_name, mode):
    # 1. PDF 텍스트 추출
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"

    # 2. Context 딕셔너리 생성 (공통) - 키 값의 마침표를 언더바로 변경
    context = {
        "CUSTOMER": customer_name,
        "PRODUCT": product_name,
        "CATEGORY1": extract_text_between(full_text, "Category 1*", "Category 2"),
        "CATEGORY2": extract_text_between(full_text, "Category 2*", "Category 3"),
        "CATEGORY3": extract_text_between(full_text, "Category 3*", "Category 4"),
        "CATEGORY4": extract_text_between(full_text, "Category 4*", "Category 5.A"),
        "CATEGORY5_A": extract_text_between(full_text, "Category 5.A*", "Category 5.B"),
        "CATEGORY5_B": extract_text_between(full_text, "Category 5.B*", "Category 5.C"),
        "CATEGORY5_C": extract_text_between(full_text, "Category 5.C*", "Category 5.D"),
        "CATEGORY7_A": extract_text_between(full_text, "Category 7.A*", "Category 7.B"),
        "CATEGORY7_B": extract_text_between(full_text, "Category 7.B*", "Category 8"),
        "CATEGORY8": extract_text_between(full_text, "Category 8*", "Category 9"),
        "CATEGORY9": extract_text_between(full_text, "Category 9*", "Category 10.A"),
        "CATEGORY10_A": extract_text_between(full_text, "Category 10.A*", "Category 10.B"),
        "CATEGORY10_B": extract_text_between(full_text, "Category 10.B*", "Category 11.A"),
        "CATEGORY11_A": extract_text_between(full_text, "Category 11.A*", "Category 11.B"),
        "CATEGORY11_B": extract_text_between(full_text, "Category 11.B*", "Category 12")
    }

    # 3. 모드별 분기 처리 (차이점만 반영) - 키 값의 마침표를 언더바로 변경
    if mode == "CFF":
        context["CATEGORY5_D"] = extract_text_between(full_text, "Category 5.D*", "Category 6")
        context["CATEGORY6"] = extract_text_between(full_text, "Category 6", "Category 7.A")
        context["CATEGORY12"] = extract_text_between(full_text, "Category 12*", "For other")
    elif mode == "HP":
        context["CATEGORY5_D"] = extract_text_between(full_text, "Category 5.D*", "Category 6*")
        context["CATEGORY6"] = extract_text_between(full_text, "Category 6*", "Category 7.A")
        context["CATEGORY12"] = extract_text_between(full_text, "Category 12*", "*Only fragrance")

    # 4. Word 템플릿 렌더링
    template_path = "templates/IFRA.docx"
    
    # 템플릿 파일 존재 여부 확인
    if not os.path.exists(template_path):
        st.error(f"오류: 템플릿 파일이 없습니다. '{template_path}' 경로를 확인해주세요.")
        return None

    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        # 메모리 버퍼에 저장 (직접 다운로드를 위해)
        output_io = BytesIO()
        doc.save(output_io)
        output_io.seek(0)
        
        return output_io
    except Exception as e:
        st.error(f"템플릿 렌더링 중 오류가 발생했습니다: {e}")
        st.error("Word 템플릿 파일 내의 변수명에 마침표(.)가 없는지 확인해주세요. (예: {{CATEGORY5.A}} -> {{CATEGORY5_A}})")
        return None

# --- Streamlit UI 구성 ---

st.set_page_config(page_title="PDF to Word Converter", layout="wide")
st.title("IFRA PDF -> Word 자동 변환기")

# 레이아웃 나누기 (왼쪽: 파일 업로드, 오른쪽: 설정 입력)
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

# 하단 레이아웃 (왼쪽: 변환 버튼, 오른쪽: 다운로드 영역)
col_btn_left, col_btn_right = st.columns(2)

with col_btn_left:
    convert_clicked = st.button("변환 시작", type="primary", use_container_width=True)

with col_btn_right:
    # 다운로드 버튼을 담을 빈 공간(placeholder) 생성
    download_placeholder = st.empty()

# --- 변환 버튼 클릭 시 동작 ---
if convert_clicked:
    if uploaded_pdf is None:
        st.warning("PDF 파일을 먼저 업로드해주세요.")
    elif not customer_input or not product_input:
        st.warning("고객사명과 제품명을 모두 입력해주세요.")
    else:
        with st.spinner("변환 작업 진행 중..."):
            # 로직 실행
            result_docx = process_pdf_to_word(uploaded_pdf, customer_input, product_input, mode_selection)
            
            if result_docx:
                st.success("변환 성공!")
                
                # 다운로드 버튼 렌더링
                file_name = f"spec_{customer_input}_{product_input}.docx"
                download_placeholder.download_button(
                    label="결과물 다운로드",
                    data=result_docx,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
