import io
import re
import time
import requests
from bs4 import BeautifulSoup
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ----------------------------------------------------
# 1. 페이지 기본 설정
# ----------------------------------------------------
st.set_page_config(
    page_title="삼성생명 기획팀 동향분석 Agent",
    page_icon="📊",
    layout="wide"
)

# ----------------------------------------------------
# 2. 기사 본문 자동 추출 크롤러 유틸리티
# ----------------------------------------------------
def extract_article_content(user_text):
    """
    입력 텍스트에서 URL을 감지하면 실제 웹페이지 본문을 긁어오고,
    일반 텍스트면 그대로 반환합니다.
    """
    url_pattern = re.compile(r'https?://[^\s]+')
    urls = url_pattern.findall(user_text.strip())
    
    if not urls:
        return user_text, None

    target_url = urls[0]
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    try:
        res = requests.get(target_url, headers=headers, timeout=8)
        res.raise_for_status()
        res.encoding = res.apparent_encoding
        
        soup = BeautifulSoup(res.text, 'html.parser')
        
        for tag in soup(['script', 'style', 'header', 'footer', 'nav', 'aside', 'iframe']):
            tag.decompose()
            
        naver_news = soup.find('article', id='dic_area') or soup.find('div', id='newsct_article') or soup.find('div', id='articeBody')
        if naver_news:
            return naver_news.get_text(separator='\n', strip=True), target_url
            
        article_tag = soup.find('article') or soup.find('div', class_=re.compile(r'article|news|content|body', re.I))
        if article_tag:
            paragraphs = [p.get_text(strip=True) for p in article_tag.find_all(['p', 'div']) if len(p.get_text(strip=True)) > 20]
            if paragraphs:
                return '\n'.join(paragraphs), target_url

        text_lines = [p.get_text(strip=True) for p in soup.find_all('p') if len(p.get_text(strip=True)) > 20]
        if text_lines:
            return '\n'.join(text_lines), target_url
            
        return user_text, target_url
    except Exception:
        return user_text, target_url

# ----------------------------------------------------
# 3. Gem 시스템 지침 (Instructions) 설정
# ----------------------------------------------------
SYSTEM_INSTRUCTION = """
# Role
당신은 삼성생명 기획팀의 "동향분석 및 보고서 작성 전문 Agent"입니다.
국내 1위 생명보험사인 삼성생명의 기획팀 직원 관점에서 기사 본문을 분석하여, 군더더기 없이 고밀도로 압축된 2페이지 규격 보고서를 작성합니다.

# Core Workflow
[Step 1. 메인 기사 분석]
- 입력된 기사 본문의 사실관계, 핵심 수치, 이슈 맥락을 정확히 파악합니다.

[Step 2. 심층 데이터 분석 및 배경 맥락 연계]
- 메인 이슈 관련 업계/경쟁사 동향, 규제/제도 기조 및 시장 수치를 연계 분석합니다.

[Step 3. 삼성생명 맞춤형 전략적 시사점 도출]
- 입력된 기사의 핵심 내용에 맞춰 삼성생명 관점의 전략적 영향 및 대응 방향('So What')을 맞춤형으로 도출합니다.

[Step 4. 규격화된 보고서 작성]
- 아래 작성 원칙 및 출력 포맷을 엄격히 준수합니다.

# Output Structure & Formatting Rules (작성 원칙)
1. 문체 및 길이 규칙 (장황함 배제):
   - 문체: 완결된 서술어(~입니다, ~함)를 배제하고 명확한 개조식 명사형 종결문(~확대, ~추진, ~견지, ~구축, ~강화, ~대응, ~도모 등)으로만 작성
   - 분량: 워드 15pt 기준 각 문장은 반드시 **1줄~최대 2줄 이내**로 작성 (3줄 이상 장황한 서술 엄격 금지)
   - 위계 구조:
     * 대항목 (네모): `□` -> 대주제 명시 (반드시 `□ [요약]` 및 `□ [시사점]` 표기 준수)
     * 중항목 (찍): `-` -> 하위 소항목(·)들을 포괄하는 압축적 한 문장 요약 (명사형 종결)
     * 소항목 (땡): `·` -> 상위 중항목(-)을 뒷받침하는 핵심 수치 및 구체적 사실관계 (명사형 종결)

2. 보고서 출력 포맷:
---
□ [요약] (분석된 기사의 핵심 제목)
  - (하위 팩트를 포괄하는 핵심 사실관계 1줄 요약)
    · 메인 기사 핵심 사실관계 및 주요 수치
    · 삼성생명 및 업계 관련 세부 지표
  - (하위 시장 변화를 포괄하는 업계 판도 1줄 요약)
    · 경쟁사 대응 현황 및 주요 영업 동향
    · 금융당국 제도/규제 기조 및 시장 영향

□ [시사점] (해당 이슈 기반 삼성생명 전략 테마)
  - (직접적 영향 및 핵심 대응 과제 1줄 요약)
    · 단기 사업/영업/상품/자본 측면 세부 실행 방안
    · 시장 지배력 유지를 위한 차별화 과제
  - (리스크 관리 및 중장기 대응 과제 1줄 요약)
    · 시장 변화에 따른 선제적 리스크 관리 방향
    · 미래 경쟁력 확보를 위한 연계 전략

# 참고 자료 및 심층 출처 리스트 (References & Deep Dive Data)

1. [메인 기사 출처]
  - 출처명: 기사 제목
    · 핵심 요약: 1줄 컴팩트 요약

2. [심층 배경 및 비교 데이터]
  - 주요 지표 및 제도적 맥락
    · 핵심 데이터: 1줄 핵심 수치 요약

3. [기획팀 종합 평가 및 향후 모니터링 포인트]
  - 시장 영향도: 업계 및 1위사에 미치는 파급력 평가
  - 리스크 대응: 중점 모니터링 요소 및 리스크 관리 방향
  - 전략 제언: 기획팀 차원의 핵심 추진 과제
"""

# ----------------------------------------------------
# 4. Word 문서 생성 유틸리티 (바탕체, 15pt 서식 적용)
# ----------------------------------------------------
def set_font_style(run, name="바탕체", size_pt=15, bold=False, color_rgb=None):
    run.font.name = name
    run._r.get_or_add_rPr().set(qn('w:rFonts'), qn('w:eastAsia'))
    run._r.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb

def create_docx(text_content):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    lines = text_content.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
            
        if line.startswith("# "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(16)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(line.replace("# ", "").strip())
            set_font_style(run, name="바탕체", size_pt=16, bold=True, color_rgb=RGBColor(0, 51, 102))
        elif line.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line.replace("## ", "").strip())
            set_font_style(run, name="바탕체", size_pt=15, bold=True)
        elif stripped.startswith("□") or stripped.startswith("ㅁ"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.25
            run = p.add_run(stripped)
            set_font_style(run, name="바탕체", size_pt=15, bold=True, color_rgb=RGBColor(0, 51, 102))
        elif stripped.startswith("-"):
            p = doc.add_paragraph(stripped)
            p.paragraph_format.left_indent = Inches(0.2)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.25
            run = p.add_run(stripped)
            set_font_style(run, name="바탕체", size_pt=15, bold=False)
        elif stripped.startswith("·") or stripped.startswith("."):
            p = doc.add_paragraph(stripped)
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.25
            run = p.add_run(stripped)
            set_font_style(run, name="바탕체", size_pt=15, bold=False)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.25
            run = p.add_run(stripped)
            set_font_style(run, name="바탕체", size_pt=15, bold=False)
                
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------------------------------------------
# 5. 부서원 비밀번호 인증
# ----------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 기획팀 전용 분석 Agent")
    pwd = st.text_input("부서 접속 비밀번호를 입력하세요", type="password")
    correct_pwd = st.secrets.get("AUTH_PASSWORD", "1234")
    
    if st.button("접속", use_container_width=True):
        if pwd == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")
    st.stop()

# ----------------------------------------------------
# 6. 메인 앱 화면 및 분석 실행
# ----------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("GEMINI_API_KEY가 설정되지 않았습니다. Secrets를 확인해주세요.")
    st.stop()

client = genai.Client(api_key=api_key)

st.title("📊 삼성생명 기획팀 동향분석 Agent")
st.caption("기사 URL 또는 본문 텍스트를 입력하면 실제 기사 본문을 정밀 분석하여 사내 표준 보고서 및 Word/TXT 파일을 생성합니다.")

user_input = st.text_area(
    "분석할 기사 내용 또는 기사 URL을 입력하세요:",
    height=150,
    placeholder="예: 기사 전문을 붙여넣거나 분석할 기사의 웹 링크(URL)를 입력하세요."
)

if st.button("보고서 생성 시작", type="primary", use_container_width=True):
    if not user_input.strip():
        st.warning("분석할 기사 내용이나 URL을 입력해주세요.")
    else:
        with st.spinner("기사 본문을 추출하고 기획팀 규격 보고서를 작성 중입니다..."):
            extracted_text, detected_url = extract_article_content(user_input)
            
            if detected_url:
                prompt_content = f"다음은 입력된 URL({detected_url})에서 추출한 기사 본문입니다. 이 내용을 바탕으로 심층 분석 보고서를 작성하십시오:\n\n{extracted_text}"
            else:
                prompt_content = f"다음 기사 내용을 바탕으로 심층 분석 보고서를 작성하십시오:\n\n{extracted_text}"

            response_success = False
            last_error_msg = ""

            for attempt in range(3):
                try:
                    response = client.models.generate_content(
                        model="gemini-3.6-flash",
                        contents=prompt_content,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            temperature=0.2
                        )
                    )
                    st.session_state["last_report"] = response.text
                    response_success = True
                    break
                except Exception as e:
                    last_error_msg = str(e)
                    if any(err in last_error_msg for err in ["503", "UNAVAILABLE", "429"]):
                        time.sleep(3)
                        continue
                    else:
                        break

            if not response_success:
                st.error(f"보고서 생성 중 오류가 발생했습니다: {last_error_msg}")

# ----------------------------------------------------
# 7. 생성된 보고서 표시 및 다운로드
# ----------------------------------------------------
if "last_report" in st.session_state:
    st.divider()
    st.markdown("### 📄 작성된 동향분석 보고서")
    st.markdown(st.session_state["last_report"])
    
    st.divider()
    col1, col2 = st.columns(2)
    
    with col1:
        docx_file = create_docx(st.session_state["last_report"])
        st.download_button(
            label="📥 Word 보고서 (.docx) 다운로드",
            data=docx_file,
            file_name="삼성생명_동향분석보고서.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with col2:
        st.download_button(
            label="📥 텍스트 보고서 (.txt) 다운로드",
            data=st.session_state["last_report"].encode("utf-8"),
            file_name="삼성생명_동향분석보고서.txt",
            mime="text/plain",
            use_container_width=True
        )
