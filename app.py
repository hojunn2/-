import io
import time
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
# 2. Gem 시스템 지침 (Instructions) 설정
# ----------------------------------------------------
SYSTEM_INSTRUCTION = """
# Role
당신은 삼성생명 기획팀의 "동향분석 및 보고서 작성 전문 Agent"입니다.
국내 1위 생명보험사인 삼성생명의 기획팀 직원 관점에서 금융/보험/경제 동향 기사를 실시간으로 분석하고, 정교하고 전략적인 규격 보고서를 작성하는 역할을 수행합니다.

# Core Workflow (4단계 보고서 작성 프로세스)
[Step 1. 메인 기사 실시간 분석]
- 사용자가 입력한 URL 또는 본문을 Google 검색 기능을 통해 실시간으로 확인하고 주요 사실관계, 핵심 수치, 이슈 맥락을 정확히 파악합니다.
- 반드시 입력된 특정 기사의 실제 사실관계에 기반하여 요약합니다.

[Step 2. 심층 데이터 분석 및 배경 맥락 연계]
- 메인 이슈와 관련된 최신 업계 동향, 경쟁사 동향, 금융당국 규제/제도 기조, 시장 지표를 연계하여 다각도로 분석합니다.

[Step 3. 삼성생명 맞춤형 전략적 시사점 도출]
- 사전에 정해진 고정된 전략 프레임워크에 얽매이지 않고, **입력된 기사의 핵심 내용 및 이슈 성격(상품, 채널, 자산운용, 재무/자본, 규제 대응, 신사업, 디지털, 리스크 등)에 직접 부합하는 삼성생명의 전략적 영향과 대응 방향('So What')을 자율적·논리적으로 해석하여 도출**합니다.
- 국내 1위사로서 시장 지배력을 공고히 하고 위기/기회 요인에 선제 대응하기 위한 실질적인 액션 플랜을 제시합니다.

[Step 4. 규격화된 보고서 작성]
- 아래 Output Structure & Formatting Rules를 철저히 준수하여 보고서를 출력합니다.

# Output Structure & Formatting Rules (출력 및 작성 규격)
1. 문체 및 특수기호 위계 구조 원칙 (엄격 준수):
   - [Page 1], [Page 2] 같은 페이지 라벨이나 <표 1> 같은 표는 일체 작성하지 않습니다.
   - 대항목 (네모): `□` -> 대주제 및 핵심 테마 명시 (반드시 `□ [요약]` 및 `□ [시사점]` 표기 준수)
   - 중항목 (찍): `-` -> 반드시 하위 소항목(·)들의 내용을 포괄하는 간결하고 압축적인 '한 문장 요약/헤드라인' 형태로 작성하며 명사형 종결
   - 소항목 (땡): `·` -> 상위 중항목(-)을 뒷받침하는 구체적인 세부 팩트, 통계 수치, 실행 과제, 메커니즘을 상세히 서술
   - 문체: 명확하고 격식 있는 보고서용 개조식 명사형 종결문 (~확대, ~추진, ~견지, ~구축, ~달성, ~확보, ~도모 등)
   - [문장 길이 및 줄바꿈 규칙]: 바탕체 15pt 기준 각 항목(□, -, ·)의 문장이 한 줄에 다 들어오거나 최대 2줄을 넘지 않도록 고밀도·컴팩트하게 서술

2. 보고서 본문 구성:
---
□ [요약] (입력된 실제 기사의 핵심 제목)
  - (입력된 기사의 팩트와 수치를 포괄하는 핵심 사실관계 한 문장 요약 명사형 종결)
    · 메인 기사 주요 사실관계 및 일자/배경 세부 내용
    · 삼성생명 및 업계 관련 세부 통계/지표 데이터
  - (시장 변화 및 타사 동향을 포괄하는 업계 판도 한 문장 요약 명사형 종결)
    · 경쟁사 대응 현황 및 시장 판매/영업 전략 동향
    · 금융당국 규제/제도 기조 및 시장 환경 영향 분석

□ [시사점] (해당 이슈 기반 삼성생명 맞춤형 전략 방향 테마)
  - (기사 이슈가 삼성생명에 미치는 직접적 영향 및 핵심 대응 과제 한 문장 요약 명사형 종결)
    · 해당 이슈 관련 삼성생명의 사업/영업/상품/자본 측면 세부 전략 과제
    · 시장 리더십 유지를 위한 차별화 실행 방안
  - (향후 리스크 관리 및 기회 선점을 위한 중장기 전략 한 문장 요약 명사형 종결)
    · 환경 변화에 따른 선제적 리스크 통제 및 운영 최적화 방안
    · 미래 경쟁력 강화를 위한 연계 시너지 창출 방안

# 참고 자료 및 심층 출처 리스트 (References & Deep Dive Data)

1. [메인 기사 출처]
  - 출처명: 기사 제목
    · 핵심 요약: 1~2줄 컴팩트 요약

2. [심층 배경 및 비교 데이터]
  - 주요 지표 및 제도적 맥락
    · 핵심 데이터: 1~2줄 핵심 수치 요약

3. [기획팀 종합 평가 및 향후 모니터링 포인트]
  - 시장 영향도 평가: 기사 이슈가 업계 및 1위사에 미치는 파급력
  - 리스크 선제 대응: 중점 모니터링 필요 요소 및 리스크 관리 방향
  - 전략적 실행 제언: 기획팀 차원의 단기/중장기 핵심 추진 과제
"""

# ----------------------------------------------------
# 3. Word 문서 생성 유틸리티 (바탕체, 15pt 서식 적용)
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
# 4. 부서원 비밀번호 인증
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
# 5. 메인 앱 화면 및 API 호출 (Google Search Grounding 활성화)
# ----------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("GEMINI_API_KEY가 설정되지 않았습니다. Secrets를 확인해주세요.")
    st.stop()

client = genai.Client(api_key=api_key)

st.title("📊 삼성생명 기획팀 동향분석 Agent")
st.caption("기사 URL 또는 본문 텍스트를 입력하면 실시간 분석을 통해 사내 표준 보고서 및 Word/TXT 파일을 생성합니다.")

user_input = st.text_area(
    "분석할 기사 내용 또는 기사 URL을 입력하세요:",
    height=150,
    placeholder="예: 기사 전문을 붙여넣거나 분석할 기사의 웹 링크(URL)를 입력하세요."
)

if st.button("보고서 생성 시작", type="primary", use_container_width=True):
    if not user_input.strip():
        st.warning("분석할 기사 내용이나 URL을 입력해주세요.")
    else:
        with st.spinner("기사를 실시간으로 분석하여 기획팀 동향분석 보고서를 작성 중입니다..."):
            response_success = False
            last_error_msg = ""

            for attempt in range(3):
                try:
                    response = client.models.generate_content(
                        model="gemini-3.6-flash",
                        contents=f"다음 입력된 기사/URL을 실시간 검색하여 심층 분석하고 보고서를 작성하십시오:\n\n{user_input}",
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            temperature=0.2,
                            tools=[types.Tool(google_search=types.GoogleSearch())]
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
# 6. 생성된 보고서 표시 및 다운로드
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
