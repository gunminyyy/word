import streamlit as st
import pdfplumber
import re
from docxtpl import DocxTemplate
from datetime import datetime
import io
import os
import sys

# --- 1. 경로 탐색 헬퍼 함수 ---
def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 2. Streamlit UI 구성 ---
st.title("📄 PDF to Word 자동 변환기")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 업로드")
    uploaded_pdf = st.file_uploader("PDF 파일을 여기에 끌어다 놓으세요.", type=["pdf"])

with col2:
    st.subheader("2. 정보 입력 및 옵션")
    product_name = st.text_input("제품명")
    mode = st.selectbox("모드 선택", ["CFF", "HP", "HPD"])

st.divider()

col3, col4 = st.columns(2)

with col3:
    convert_btn = st.button("변환 실행", use_container_width=True)

# --- 3. 데이터 추출 및 변환 로직 ---
if convert_btn:
    if not uploaded_pdf:
        st.error("원본 PDF 파일을 업로드해주세요.")
    elif not product_name:
        st.error("제품명을 입력해주세요.")
    else:
        with st.spinner("파일을 변환하는 중입니다..."):
            try:
                # PDF 텍스트 추출
                pdf_text = ""
                with pdfplumber.open(uploaded_pdf) as pdf:
                    for page in pdf.pages:
                        extracted = page.extract_text()
                        if extracted:
                            pdf_text += extracted + "\n"
                
                # 워드 템플릿에 들어갈 기본값 세팅 (PDF에서 값을 못 찾을 경우 이 값이 들어감)
                context = {
                    "PRODUCT": product_name,
                    "COLOR": "PALE YELLOW TO YELLOW",
                    "SG": "0.902 ~ 0.922",
                    "RI": "1.466 ~ 1.476",
                    "DATE": datetime.now().strftime("%d. %b. %Y").upper()
                }
                
                # 모드별 로직
                if mode == "CFF":
                    # COLOR 추출
                    color_match = re.search(r'COLOR\s*:(.*?)APPEARANCE\s*:', pdf_text, re.DOTALL | re.IGNORECASE)
                    if color_match:
                        context["COLOR"] = color_match.group(1).strip().upper()
                    
                    # SPECIFIC GRAVITY 계산
                    sg_match = re.search(r'SPECIFIC GRAVITY.*?\(\d+°C\)\s*:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
                    if sg_match:
                        sg_base = float(sg_match.group(1))
                        context["SG"] = f"{sg_base - 0.01:.3f} ~ {sg_base + 0.01:.3f}"
                        
                    # REFRACTIVE INDEX 계산
                    ri_match = re.search(r'REFRACTIVE INDEX.*?\(\d+°C\)\s*:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
                    if ri_match:
                        ri_base = float(ri_match.group(1))
                        context["RI"] = f"{ri_base - 0.01:.3f} ~ {ri_base + 0.01:.3f}"

                elif mode == "HP":
                    # TODO: HP 모드 로직 작성
                    st.info("HP 모드 로직이 아직 구현되지 않았습니다. 기본값이 적용됩니다.")
                    
                elif mode == "HPD":
                    # TODO: HPD 모드 로직 작성
                    st.info("HPD 모드 로직이 아직 구현되지 않았습니다. 기본값이 적용됩니다.")

                # 워드 템플릿 불러오기 및 데이터 렌더링
                doc_path = get_resource_path("templates/spec.docx")
                doc = DocxTemplate(doc_path)
                
                # context 딕셔너리의 데이터를 템플릿의 {{태그}} 위치에 쏙 맞춰 넣음
                doc.render(context)
                
                # 결과물 저장 및 다운로드
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                
                st.success("변환이 완료되었습니다! 우측에서 다운로드하세요.")
                
                with col4:
                    st.download_button(
                        label="결과물 다운로드 (.docx)",
                        data=bio,
                        file_name=f"{product_name}_변환결과.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
                st.info("템플릿 폴더에 태그가 적용된 spec.docx 파일이 있는지 확인해주세요.")
